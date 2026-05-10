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
// Headers:
//   X-Daily-Recap-Token: <DAILY_RECAP_TOKEN>
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

const { getAdapter } = require('../lib/adapter');
const { createBotRedis } = require('../lib/redis');
const { Template } = require('adaptivecards-templating');
const { runWithRetry } = require('../lib/retry');
const { CHANNELS } = require('../queue/channels');

const RECAP_REDIS_TTL = 60 * 60 * 24 * 2; // 2 days — long enough for the chat user to interact with the card
// Auto-delete the recap card from the channel after this many ms.  The
// recap is meant as a transient digest; leaving it in the channel feed
// is noise.  Five minutes gives admins enough time to read it.
const AUTO_DELETE_MS = 5 * 60 * 1000;
const MAX_CARD_BYTES = 25 * 1024; // Teams hard limit

// Track pending auto-delete timers for debugging/monitoring
const _pendingDeletes = new Map();

const redis = createBotRedis();

// ---------------------------------------------------------------------------
// Auth middleware — validate X-Daily-Recap-Token header
// ---------------------------------------------------------------------------
function validateToken(req) {
    const expected = process.env.DAILY_RECAP_TOKEN;
    if (!expected) {
        // Token not configured — allow (bot may be in dev mode)
        return true;
    }
    const provided = req.headers['x-daily-recap-token'];
    return provided && provided === expected;
}

// ---------------------------------------------------------------------------
// Bind template + data via the official Adaptive Cards Templating SDK.
// Wraps data in `$root` per SDK convention so `${foo}` resolves to
// `data.foo` and `${$root.foo}` resolves the same.
// ---------------------------------------------------------------------------
function bindCard(template, data) {
    const tpl = new Template(template);
    return tpl.expand({ $root: data });
}

// ---------------------------------------------------------------------------
// Schedule a one-shot deletion of an activity after `delayMs` milliseconds.
// Uses `adapter.continueConversation` so the timer can run after the
// original turn / proactive send has completed.  Best-effort — failures
// are logged, never thrown, and the timer is unrefed so it doesn't block
// process shutdown.
//
// Note: Microsoft Teams allows a bot to delete only its own messages,
// and only within the bot's app-permission window.  The delete will
// silently fail if the bot lacks permission or the message has already
// been removed.
// ---------------------------------------------------------------------------
function scheduleAutoDelete(conversationReference, activityId, delayMs) {
    if (!activityId) {
        return;
    }
    const timer = setTimeout(async () => {
        _pendingDeletes.delete(activityId);
        try {
            await getAdapter().continueConversation(conversationReference, async (proactiveContext) => {
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
    _pendingDeletes.set(activityId, { timer, conversationReference, delayMs });
}

// ---------------------------------------------------------------------------
// Send an attachment proactively to a stored conversation reference,
// with retry on transient/auth failures.  Returns the outgoing
// activity ID so callers can update it later via Action.Submit.
// ---------------------------------------------------------------------------
async function sendProactiveCard(conversationReference, card) {
    let activityId = null;
    await runWithRetry(async () => {
        await getAdapter().continueConversation(conversationReference, async (proactiveContext) => {
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

// ---------------------------------------------------------------------------
// Attempt to reduce card size to fit within Teams' 25 KB limit.
// Returns the modified card object (may be the same reference if already small).
// Strips large image URLs, truncates long text fields, and removes empty arrays.
// ---------------------------------------------------------------------------
function enforceCardSizeLimit(card) {
    const clone = JSON.parse(JSON.stringify(card));
    _enforceRecursive(clone, 0);
    return clone;
}

function _enforceRecursive(obj, depth) {
    if (depth > 20) return; // prevent stack overflow on deeply nested cards
    if (!obj || typeof obj !== 'object') return;

    if (Array.isArray(obj)) {
        for (const item of obj) _enforceRecursive(item, depth + 1);
        return;
    }

    // Truncate long text values
    if (typeof obj.text === 'string' && obj.text.length > 2000) {
        obj.text = obj.text.slice(0, 2000) + '…';
    }
    if (typeof obj.speak === 'string' && obj.speak.length > 1000) {
        obj.speak = obj.speak.slice(0, 1000) + '…';
    }

    // Remove base64 image URLs (they bloat the card)
    if (typeof obj.url === 'string' && obj.url.startsWith('data:')) {
        obj.url = '';
    }
    if (typeof obj.image === 'string' && obj.image.startsWith('data:')) {
        obj.image = '';
    }

    // Trim empty arrays to save space
    for (const [key, val] of Object.entries(obj)) {
        if (Array.isArray(val) && val.length === 0) {
            delete obj[key];
        } else {
            _enforceRecursive(val, depth + 1);
        }
    }
}

const dailyRecapHandler = async (req, res) => {
    // Auth check
    if (!validateToken(req)) {
        return res.status(401).json({ ok: false, error: 'invalid or missing X-Daily-Recap-Token header' });
    }

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
    const cardSize = cardJson.length;

    // Enforce Teams' 25 KB Adaptive Card limit
    if (cardSize > MAX_CARD_BYTES) {
        const originalSizeKB = Math.round(cardSize / 1024);
        card = enforceCardSizeLimit(card);
        const newSize = JSON.stringify(card).length;
        const newSizeKB = Math.round(newSize / 1024);
        console.warn(`[dailyrecap] card ${ originalSizeKB } KB exceeded limit — trimmed to ${ newSizeKB } KB`);
        if (newSize > MAX_CARD_BYTES) {
            return res.status(400).json({
                ok: false,
                error: `bound card is ${ originalSizeKB } KB (limit 25 KB) and could not be trimmed enough`
            });
        }
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
        bytes_sent:     JSON.stringify(card).length
    });
};

// Expose for tests
module.exports = dailyRecapHandler;
module.exports.CHANNELS = CHANNELS;
module.exports.bindCard = bindCard;
module.exports.scheduleAutoDelete = scheduleAutoDelete;
module.exports.AUTO_DELETE_MS = AUTO_DELETE_MS;
module.exports.enforceCardSizeLimit = enforceCardSizeLimit;
module.exports.validateToken = validateToken;
