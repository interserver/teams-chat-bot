// `!notif ...` admin commands for inspecting and managing the
// notification queue from a Teams chat. Stays in the same command-
// registry shape as the other commands in this directory.
//
// Subcommands:
//   !notif status            — queue depth, processing depth, dead count, last tick
//   !notif rooms             — list known room → conversationId mappings,
//                              flagging which have a stored conversation reference
//   !notif test <room> <msg> — enqueue a probe envelope (uses dedup_key 'admin:test')
//   !notif drain-dead        — move all entries from notif:dead back onto notif:queue
//   !notif seed-room <room>  — post a hidden-noise message into the named room so
//                              the bot captures the conversation reference for
//                              future proactive sends

const { MessageFactory } = require('botbuilder');
const { CHANNELS, knownRooms, resolve: resolveChannel } = require('../queue/channels');
const { getNotifRedis, DOWNSTREAM_REPOS, wfactiveKey, EDIT_WINDOW_MS } = require('../queue/notificationConsumer');

const KEY_PREFIX = process.env.NOTIF_KEY_PREFIX || 'notif:';
function k(name) { return KEY_PREFIX + name; }

module.exports = {
    match(text) {
        const m = text.match(/^!notif(?:\s+(\S+)(?:\s+(.+))?)?$/i);
        return m ? { sub: (m[1] || '').toLowerCase(), rest: (m[2] || '').trim() } : null;
    },
    async execute({ sub, rest }, deps) {
        const { context } = deps;
        // notif:* keys live on the InterServer Redis (consumer's client),
        // separate from the bot's convref:* Redis. Use the consumer's client.
        const notifRedis = getNotifRedis();
        const botRedis = deps.redis;
        if (!notifRedis) {
            await context.sendActivity(MessageFactory.text('notif consumer is not running — try again after the bot has finished startup'));
            return;
        }
        try {
            switch (sub) {
            case 'status':
                return await statusCmd(context, notifRedis);
            case 'rooms':
                return await roomsCmd(context, botRedis);
            case 'test':
                return await testCmd(context, notifRedis, rest);
            case 'drain-dead':
                return await drainDeadCmd(context, notifRedis);
            case 'seed-room':
                return await seedRoomCmd(context, botRedis, rest);
            case 'wfactive':
                return await wfactiveCmd(context, notifRedis);
            case '':
            case 'help':
            default:
                return await helpCmd(context);
            }
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`!notif ${ sub } failed: ${ err.message }`));
        }
    }
};

async function statusCmd(context, redis) {
    const [queueDepth, processingDepth, deadDepth, enqueued, sent, edited, coalesced, redirected, fallback, dead] = await Promise.all([
        redis.llen(k('queue')),
        redis.llen(k('processing')),
        redis.llen(k('dead')),
        redis.get(k('metrics:enqueued')).catch(() => '0'),
        redis.get(k('metrics:sent')).catch(() => '0'),
        redis.get(k('metrics:edited')).catch(() => '0'),
        redis.get(k('metrics:coalesced')).catch(() => '0'),
        redis.get(k('metrics:redirected')).catch(() => '0'),
        redis.get(k('metrics:fallback')).catch(() => '0'),
        redis.get(k('metrics:dead')).catch(() => '0')
    ]);
    const lines = [
        '**notif queue status**',
        '```',
        `queue:      ${ queueDepth }`,
        `processing: ${ processingDepth }`,
        `dead:       ${ deadDepth }`,
        '',
        `enqueued:   ${ enqueued || 0 }`,
        `sent:       ${ sent || 0 }`,
        `edited:     ${ edited || 0 }`,
        `coalesced:  ${ coalesced || 0 }`,
        `redirected: ${ redirected || 0 }`,
        `fallback:   ${ fallback || 0 }`,
        `dead-meter: ${ dead || 0 }`,
        '```'
    ];
    await context.sendActivity(MessageFactory.text(lines.join('\n')));
}

async function roomsCmd(context, redis) {
    const rooms = knownRooms();
    const lines = ['**known rooms**', '```'];
    for (const r of rooms) {
        const cid = CHANNELS[r];
        const ref = await redis.get(`convref:${ cid }`).catch(() => null);
        lines.push(`${ ref ? '✅' : '❌' } ${ r.padEnd(18) } ${ cid }`);
    }
    lines.push('```');
    lines.push('_(❌ = bot has not observed inbound activity in this room yet — send anything in the channel or use `!notif seed-room <name>` to register the conversation reference)_');
    await context.sendActivity(MessageFactory.text(lines.join('\n')));
}

async function testCmd(context, redis, rest) {
    const m = rest.match(/^(\S+)\s+(.+)$/);
    if (!m) {
        await context.sendActivity(MessageFactory.text('usage: `!notif test <room> <message>`'));
        return;
    }
    const [, room, msg] = m;
    if (!resolveChannel(room)) {
        await context.sendActivity(MessageFactory.text(`unknown room "${ room }"`));
        return;
    }
    // Sanitize: trim, cap message length to 4000 chars (Teams message limit)
    const safeMsg = String(msg).slice(0, 4000);
    const envelope = {
        v: 1,
        id: 'test-' + Date.now(),
        ts: Math.floor(Date.now() / 1000),
        expires_at: Math.floor(Date.now() / 1000) + 300,
        room: room.trim(),
        type: 'msg',
        message: safeMsg,
        card: null,
        extra: { dedup_key: 'admin:test', level: 'info', source: 'commands/notifAdmin' },
        fallback_webhook_url: null
    };
    await redis.lpush(k('queue'), JSON.stringify(envelope));
    await context.sendActivity(MessageFactory.text(`✅ enqueued test envelope to room \`${ room }\` — should appear within one tick`));
}

async function drainDeadCmd(context, redis) {
    const items = await redis.lrange(k('dead'), 0, -1);
    if (!items.length) {
        await context.sendActivity(MessageFactory.text('dead list is empty'));
        return;
    }
    const pipe = redis.pipeline();
    for (const j of items) pipe.rpush(k('queue'), j);
    pipe.del(k('dead'));
    await pipe.exec();
    await context.sendActivity(MessageFactory.text(`drained ${ items.length } entries from dead → queue`));
}

async function seedRoomCmd(context, redis, rest) {
    const room = String(rest).trim().slice(0, 64);
    if (!room) {
        await context.sendActivity(MessageFactory.text('usage: `!notif seed-room <room>`'));
        return;
    }
    const cid = resolveChannel(room);
    if (!cid) {
        await context.sendActivity(MessageFactory.text(`unknown room "${ room }"`));
        return;
    }
    const existing = await redis.get(`convref:${ cid }`);
    if (existing) {
        await context.sendActivity(MessageFactory.text(`✅ \`${ room }\` already has a stored conversation reference`));
        return;
    }
    await context.sendActivity(MessageFactory.text(
        `❌ no stored conversation reference for \`${ room }\` (${ cid })\n\n` +
        'To seed it: have someone send any message in that channel while the bot is a member. ' +
        'The bot stores the conversation reference automatically on every onMessage.'
    ));
}

async function wfactiveCmd(context, redis) {
    // Discover wfactive keys without a `KEYS *` scan (Redis can be configured
    // to refuse it). The downstream-map enumerates the upstream repos we
    // expect to track; for everything else, fall back to a SCAN.
    const seen = new Set();
    for (const { upstream } of DOWNSTREAM_REPOS) seen.add(upstream);
    const cursorScan = async () => {
        const found = new Set();
        let cursor = '0';
        const match = wfactiveKey('*');
        // Cap iterations defensively — wfactive is not expected to grow large.
        for (let i = 0; i < 50; i++) {
            const [next, batch] = await redis.scan(cursor, 'MATCH', match, 'COUNT', 200);
            for (const k_ of batch) {
                const repo = k_.slice(wfactiveKey('').length);
                if (repo) found.add(repo);
            }
            cursor = next;
            if (cursor === '0') break;
        }
        return found;
    };
    try {
        const scanned = await cursorScan();
        for (const r of scanned) seen.add(r);
    } catch (_) { /* SCAN unsupported — fall back to map-only enumeration */ }

    const cutoff = Date.now() - EDIT_WINDOW_MS;
    const lines = ['**notif active workflows**', '```'];
    let total = 0;
    for (const repo of [...seen].sort()) {
        const entries = await redis.zrevrangebyscore(
            wfactiveKey(repo), '+inf', cutoff, 'WITHSCORES', 'LIMIT', 0, 10
        ).catch(() => []);
        if (!entries.length) continue;
        lines.push(`${ repo }`);
        for (let i = 0; i < entries.length; i += 2) {
            const sha = entries[i];
            const ts = parseFloat(entries[i + 1]);
            const ageSec = Math.floor((Date.now() - ts) / 1000);
            const age = ageSec < 60 ? `${ ageSec }s` : `${ Math.floor(ageSec / 60) }m ${ ageSec % 60 }s`;
            lines.push(`  ${ sha }  age=${ age }`);
            total++;
        }
    }
    if (total === 0) lines.push('(no active workflows within edit window)');
    lines.push('```');
    if (DOWNSTREAM_REPOS.length) {
        lines.push('_downstream-repo map:_');
        for (const { upstream, pattern } of DOWNSTREAM_REPOS) {
            lines.push(`  \`${ upstream }\` → \`${ pattern.source }\``);
        }
    }
    await context.sendActivity(MessageFactory.text(lines.join('\n')));
}

async function helpCmd(context) {
    const lines = [
        '**!notif** — notification queue admin',
        '',
        '`!notif status`            — queue/processing/dead depths and metrics',
        '`!notif rooms`             — list known rooms; ✅ = convref cached, ❌ = bot needs to observe inbound activity first',
        '`!notif test <room> <msg>` — enqueue a probe envelope',
        '`!notif drain-dead`        — re-queue everything in `notif:dead`',
        '`!notif seed-room <room>`  — show whether the bot has a conversation reference for the given room',
        '`!notif wfactive`          — show parent commits with active CI per repo (used for action-triggered push attribution)'
    ];
    await context.sendActivity(MessageFactory.text(lines.join('\n')));
}
