// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// POST /api/dailyrecap
//
// Proactively post an Adaptive Card built from an Adaptive Cards
// Templating template + data pair (https://learn.microsoft.com/en-us/adaptive-cards/templating/).
// Microsoft Teams' renderer does not expand templating syntax on its own,
// so we bind it server-side via the official `adaptivecards-templating`
// SDK before forwarding the bound card via `adapter.continueConversation`.
//
// Request body:
//   {
//     "template": { ... AdaptiveCard template with ${...}, $data, $when ... },
//     "data":     { ... substitution dictionary ... },
//     "channel":  "int-dev"   // friendly name from CHANNELS map below
//   }
//
// Response:
//   200 { ok: true,  activityId, conversationId, channel, bytes_sent }
//   4xx { ok: false, error }
//
// The bot only knows how to send proactively to a conversation whose
// reference it has previously stored in Redis (`convref:{conversationId}`).
// botActivityHandler stores those refs on every onMessage, so the bot
// must have observed at least one message in the target channel since
// the bot was added.

const { BotFrameworkAdapter } = require('botbuilder');
const { Template } = require('adaptivecards-templating');
const Redis = require('ioredis');
const { runWithRetry } = require('../lib/retry');
const { CHANNELS } = require('../queue/channels');

const RECAP_REDIS_TTL = 60 * 60 * 24 * 2; // 2 days — long enough for the chat user to interact with the card
// Auto-delete the recap card from the channel after this many ms.  The
// recap is meant as a transient digest; leaving it in the channel feed
// is noise.  Five minutes gives admins enough time to read it.
const AUTO_DELETE_MS = 5 * 60 * 1000;

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

adapter.onTurnError = async (context, error) => {
    console.error(`[dailyrecap onTurnError] ${ error }`);
};

const redis = new Redis({
    host: process.env.REDIS_HOST || '67.217.60.234',
    port: parseInt(process.env.REDIS_PORT || '6379', 10)
});
redis.on('error', (err) => console.error('[dailyrecap redis] ', err.message));

/**
 * Bind template + data via the official Adaptive Cards Templating SDK.
 * Wraps data in `$root` per SDK convention so `${foo}` resolves to
 * `data.foo` and `${$root.foo}` resolves the same.
 */
function bindCard(template, data) {
    const tpl = new Template(template);
    return tpl.expand({ $root: data });
}

/**
 * Schedule a one-shot deletion of an activity from the conversation
 * after `delayMs` milliseconds.  Uses `adapter.continueConversation`
 * so the timer can run after the original turn / proactive send has
 * completed.  Best-effort — failures are logged, never thrown, and
 * the timer is unrefed so it doesn't block process shutdown.
 *
 * Note: Microsoft Teams allows a bot to delete only its own messages,
 * and only within the bot's app-permission window.  The delete will
 * silently fail if the bot lacks permission or the message has already
 * been removed.
 */
function scheduleAutoDelete(conversationReference, activityId, delayMs) {
    if (!activityId) {
        return;
    }
    const timer = setTimeout(async () => {
        try {
            await adapter.continueConversation(conversationReference, async (proactiveContext) => {
                await proactiveContext.deleteActivity(activityId);
            });
        } catch (err) {
            console.warn(`[dailyrecap auto-delete] activity ${ activityId } failed: ${ err.message }`);
        }
        try {
            await redis.del(`dailyrecap:${ activityId }`);
        } catch (err) {
            console.warn(`[dailyrecap auto-delete] redis cleanup for ${ activityId } failed: ${ err.message }`);
        }
    }, delayMs);
    if (typeof timer.unref === 'function') {
        timer.unref();
    }
}

/**
 * Send an attachment proactively to a stored conversation reference,
 * with retry on transient/auth failures.  Returns the outgoing
 * activity ID so callers can update it later via Action.Submit.
 */
async function sendProactiveCard(conversationReference, card) {
    let activityId = null;
    await runWithRetry(async () => {
        await adapter.continueConversation(conversationReference, async (proactiveContext) => {
            const sent = await proactiveContext.sendActivity({
                type: 'message',
                attachments: [{
                    contentType: 'application/vnd.microsoft.card.adaptive',
                    content: card
                }]
            });
            activityId = sent && sent.id ? sent.id : null;
        });
    }, {
        label: 'dailyrecap',
        serviceUrl: conversationReference && conversationReference.serviceUrl,
        maxRetries: 3
    });
    return activityId;
}

const dailyRecapHandler = async (req, res) => {
    const { template, data, channel } = req.body || {};
    if (!template || typeof template !== 'object') {
        return res.status(400).json({ ok: false, error: 'missing or non-object "template"' });
    }
    if (!data || typeof data !== 'object') {
        return res.status(400).json({ ok: false, error: 'missing or non-object "data"' });
    }
    if (!channel || !CHANNELS[channel]) {
        return res.status(400).json({
            ok: false,
            error: `unknown channel "${ channel }".  Known: ${ Object.keys(CHANNELS).join(', ') }`
        });
    }

    const conversationId = CHANNELS[channel];
    const stored = await redis.get(`convref:${ conversationId }`);
    if (!stored) {
        return res.status(404).json({
            ok: false,
            error: `no conversation reference for channel "${ channel }" (${ conversationId }).  The bot has to observe at least one message in the channel before it can post proactively.`
        });
    }

    let card;
    try {
        card = bindCard(template, data);
    } catch (err) {
        return res.status(400).json({ ok: false, error: `template binding failed: ${ err.message }` });
    }

    const cardJson = JSON.stringify(card);
    if (cardJson.length > 25 * 1024) {
        console.warn(`[dailyrecap] bound card is ${ Math.round(cardJson.length / 1024) } KB — Teams' 25 KB Adaptive Card limit may drop it silently`);
    }

    let activityId;
    let conversationReference;
    try {
        conversationReference = JSON.parse(stored);
        activityId = await sendProactiveCard(conversationReference, card);
    } catch (err) {
        console.error('[dailyrecap] send failed:', err.message);
        return res.status(502).json({ ok: false, error: err.message });
    }

    // Stash the template + data so a later Action.Submit toggle can
    // re-bind the card with different show_pie / show_details / show_month
    // flags and updateActivity in place.
    if (activityId) {
        try {
            await redis.set(
                `dailyrecap:${ activityId }`,
                JSON.stringify({ template, data, conversationId }),
                'EX',
                RECAP_REDIS_TTL
            );
        } catch (err) {
            console.warn('[dailyrecap] failed to cache template/data in Redis:', err.message);
        }
        scheduleAutoDelete(conversationReference, activityId, AUTO_DELETE_MS);
    }

    return res.json({
        ok:             true,
        activityId,
        conversationId,
        channel,
        bytes_sent:     cardJson.length
    });
};

module.exports = dailyRecapHandler;
module.exports.CHANNELS = CHANNELS;
module.exports.bindCard = bindCard;
module.exports.scheduleAutoDelete = scheduleAutoDelete;
module.exports.AUTO_DELETE_MS = AUTO_DELETE_MS;
