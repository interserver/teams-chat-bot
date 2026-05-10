// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// POST /api/message
//
// Proactive plain-text send. Replaces the original hardcoded
// `int-dev-private` target with a body-driven channel name resolved
// against the shared channels map. Retry logic now lives in
// server/lib/retry.js.

const { BotFrameworkAdapter } = require('botbuilder');
const Redis = require('ioredis');
const { runWithRetry } = require('../lib/retry');
const { resolve: resolveChannel } = require('../queue/channels');

// Channels to skip sending messages to. To re-enable a channel, remove it
// from this array or set SKIP_CHANNELS="" in your .env file.
const SKIP_CHANNELS = [];

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

adapter.onTurnError = async (context, error) => {
    console.error(`\n [msg onTurnError] unhandled error: ${ error }`);
};

const redis = new Redis({
    host: process.env.REDIS_HOST || '67.217.60.234',
    port: parseInt(process.env.REDIS_PORT || '6379', 10)
});
redis.on('connect', () => console.log('✅ Connected to Redis (msgController)'));
redis.on('error', (err) => console.error('❌ Redis error (msgController):', err.message));

async function sendProactiveMessage(conversationReference, messageText) {
    return runWithRetry(async () => {
        await adapter.continueConversation(conversationReference, async (proactiveContext) => {
            await proactiveContext.sendActivity(messageText);
        });
    }, {
        label: 'msgController',
        serviceUrl: conversationReference && conversationReference.serviceUrl,
        maxRetries: 3
    });
}

const msgHandler = async (req, res) => {
    const { message, channel } = req.body || {};
    if (!message || typeof message !== 'string') {
        return res.status(400).json({ message: 'missing or non-string "message"' });
    }
    const roomName = channel || 'int-dev-private';
    if (SKIP_CHANNELS.includes(roomName)) {
        return res.status(200).json({ message: 'skipped', channel: roomName, reason: 'channel disabled' });
    }
    const targetConversationId = resolveChannel(roomName);
    if (!targetConversationId) {
        return res.status(400).json({ message: `unknown channel "${ roomName }"` });
    }
    try {
        const stored = await redis.get(`convref:${ targetConversationId }`);
        if (!stored) {
            return res.status(404).json({ message: 'no conversation reference found for ' + roomName });
        }
        const conversationReference = JSON.parse(stored);
        await sendProactiveMessage(conversationReference, message);
        res.json({ message: 'sent', channel: roomName });
    } catch (err) {
        console.error('[msgHandler] send failed:', err.message);
        res.status(500).json({ message: err.message });
    }
};

module.exports = msgHandler;
