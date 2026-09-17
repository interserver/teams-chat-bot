// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// POST /api/message
//
// Proactive plain-text send. Replaces the original hardcoded
// `int-dev-private` target with a body-driven channel name resolved
// against the shared channels map. Retry logic now lives in
// server/lib/retry.js.

const { getAdapter } = require('../lib/adapter');
const { createBotRedis } = require('../lib/redis');
const { runWithRetry } = require('../lib/retry');
const { resolve: resolveChannel } = require('../queue/channels');
const { MAX_ACTIVITY_BYTES, activityBytes, fitTextActivity } = require('../lib/messageSize');

// Channels to skip sending messages to. To re-enable a channel, remove it
// from this array or set SKIP_CHANNELS="" in your .env file.
const SKIP_CHANNELS = [];

const redis = createBotRedis();

async function sendProactiveMessage(conversationReference, messageText) {
    return runWithRetry(async () => {
        await getAdapter().continueConversation(conversationReference, async (proactiveContext) => {
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
    // Clamp to Teams' ~25 KB serialized-activity limit up front. If we let an
    // oversized payload through, the SDK swallows the resulting
    // `Message size too large` RestError in onTurnError and the caller would
    // get a false 200 while nothing lands. Clamp — never reject.
    let outboundText = message;
    let truncated = false;
    const requestedBytes = activityBytes({ type: 'message', text: message });
    if (requestedBytes > MAX_ACTIVITY_BYTES) {
        outboundText = fitTextActivity(message, MAX_ACTIVITY_BYTES);
        truncated = true;
        console.log(`[msg] oversized request from ${ req.ip } clamped: ${ requestedBytes } → ${ activityBytes({ type: 'message', text: outboundText }) } bytes`);
    }
    try {
        const stored = await redis.get(`convref:${ targetConversationId }`);
        if (!stored) {
            return res.status(404).json({ message: 'no conversation reference found for ' + roomName });
        }
        const conversationReference = JSON.parse(stored);
        await sendProactiveMessage(conversationReference, outboundText);
        res.json(truncated ? { message: 'sent', channel: roomName, truncated: true } : { message: 'sent', channel: roomName });
    } catch (err) {
        console.error('[msgHandler] send failed:', err.message);
        res.status(500).json({ message: err.message });
    }
};

module.exports = msgHandler;
