// Integration tests for the oversized-activity guard in the notification
// consumer. Reproduces the 2026-09-17 incident shape: a burst of large
// GitHub notifications destined for `int-dev-announce` whose merged/
// coalesed payloads blew past Teams' ~25 KB serialized-activity limit.
// The guard must never lose sends (chunking) and never silently drop
// updates (in-place truncation), and must leave `oversized_guard` trace
// evidence behind.

const { describe, it, beforeEach, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// Point the trace writer at a temp file (not '0') so oversized_guard
// events can be asserted. Must be set BEFORE requiring notifTrace,
// which reads the env var at module load.
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'notif-oversized-'));
const tracePath = path.join(tmpDir, 'trace.jsonl');
process.env.NOTIF_TRACE_LOG = tracePath;

const consumer = require('../server/queue/notificationConsumer');
const { _setInternalsForTest, handleCoalesced, tryEdit, guardActivity } = consumer;
const trace = require('../server/queue/notifTrace');
const { MAX_ACTIVITY_BYTES, activityBytes, TRUNCATION_MARKER } = require('../server/lib/messageSize');
const { resolve: resolveChannel } = require('../server/queue/channels');

const APPEND_BREAK = '\n\n— · —\n\n';
const ROOM = 'int-dev-announce';
const CONV_ID = resolveChannel(ROOM);
const CONVERSATION_REF = {
    serviceUrl: 'https://smba.test.example/sms/',
    bot: { id: 'bot-id' },
    user: { id: 'bot' },
    conversation: { id: CONV_ID },
    channelId: 'msteams'
};

// --- Mocks (mirror notifBatchMerge.test.js) --------------------------------

function makeRedisMock() {
    const hashes = new Map();
    const lists = new Map();
    return {
        _hashes: hashes,
        _lists: lists,
        async hget(key, field) {
            const h = hashes.get(key);
            return h ? (h[field] || null) : null;
        },
        async hset(key, field, value) {
            const h = hashes.get(key) || {};
            h[field] = value;
            hashes.set(key, h);
            return 1;
        },
        async incr() { return 1; },
        async get() { return null; },
        async expire() { return 1; },
        async rpush(key, ...values) {
            const arr = lists.get(key) || [];
            arr.push(...values);
            lists.set(key, arr);
            return arr.length;
        },
        async lrem(key, count, value) {
            const arr = lists.get(key) || [];
            const idx = arr.indexOf(value);
            if (idx >= 0) arr.splice(idx, 1);
            lists.set(key, arr);
            return idx >= 0 ? 1 : 0;
        },
        pipeline() {
            const ops = [];
            const self = this;
            const chain = {
                hset(key, field, value) { ops.push({ op: 'hset', key, field, value }); return chain; },
                expire() { return chain; },
                zadd() { return chain; },
                zremrangebyscore() { return chain; },
                async exec() {
                    for (const o of ops) {
                        if (o.op === 'hset') {
                            const h = self._hashes.get(o.key) || {};
                            h[o.field] = o.value;
                            self._hashes.set(o.key, h);
                        }
                    }
                    return [];
                }
            };
            return chain;
        }
    };
}

function makeBotRedisMock(convrefs = {}) {
    return {
        async get(key) {
            const m = key.match(/^convref:(.+)$/);
            if (!m) return null;
            const v = convrefs[m[1]];
            return v == null ? null : JSON.stringify(v);
        },
        async incr() { return 1; }
    };
}

function makeAdapterMock() {
    const calls = [];
    return {
        _calls: calls,
        async continueConversation(conversationRef, callback) {
            const record = { conversationRef, sentActivities: [], updatedActivities: [] };
            calls.push(record);
            const proactiveContext = {
                async sendActivity(activity) {
                    record.sentActivities.push(activity);
                    return { id: 'new-activity-' + calls.length };
                },
                async updateActivity(activity) {
                    record.updatedActivities.push(activity);
                    return null;
                }
            };
            await callback(proactiveContext);
        }
    };
}

async function readTraceEvents() {
    await trace.flush();
    if (!fs.existsSync(tracePath)) return [];
    return fs.readFileSync(tracePath, 'utf8')
        .split('\n')
        .filter(Boolean)
        .map(line => JSON.parse(line));
}

function stats() {
    return { drained: 0, sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
}

after(async () => {
    await trace.flush();
    fs.rmSync(tmpDir, { recursive: true, force: true });
});

// --- sendCombinedText chunking ----------------------------------------------

describe('coalesced oversized text is split into byte-fit chunks', () => {
    let redisMock, botRedisMock, adapterMock, originals;

    beforeEach(() => {
        redisMock = makeRedisMock();
        botRedisMock = makeBotRedisMock({ [CONV_ID]: CONVERSATION_REF });
        adapterMock = makeAdapterMock();
        redisMock._lists.set('notif:processing', []);
        originals = _setInternalsForTest({
            redis: redisMock,
            redisBot: botRedisMock,
            adapter: adapterMock
        });
    });

    after(() => {
        _setInternalsForTest(originals || {});
    });

    it('sends every chunk under the limit with zero content loss', async () => {
        // 8 pieces of ~6 K JS chars each, emoji-heavy so SERIALIZED bytes
        // overflow the 24 KB budget well before the 24 K-char COALESCE cap.
        const pieces = Array.from({ length: 8 }, (_, i) => `🚀 build ${ i } ` + '🚀'.repeat(2990));
        const items = pieces.map(message => {
            const env = { type: 'msg', room: ROOM, message };
            return { raw: JSON.stringify(env), env };
        });
        const s = stats();

        await handleCoalesced(ROOM, items, s);

        assert.equal(adapterMock._calls.length, 1, 'one proactive turn carries all chunk sends');
        const sent = adapterMock._calls[0].sentActivities;
        assert.ok(sent.length >= 2, `expected chunked sends, got ${ sent.length }`);

        // Hard backstop: nothing oversized ever reaches the adapter.
        for (const activity of sent) {
            assert.ok(activityBytes(activity) <= MAX_ACTIVITY_BYTES, 'each sent activity fits');
        }

        // The byte-cap included the first N pieces; the combined text must
        // survive across the chunks verbatim, in order.
        const included = [];
        let combined = '';
        for (const piece of pieces) {
            const next = included.length === 0 ? piece : combined + APPEND_BREAK + piece;
            if (next.length > 24000 && included.length > 0) break;
            combined = next;
            included.push(piece);
        }
        assert.equal(sent.map(a => a.text).join(''), combined, 'chunk concatenation is lossless');
        assert.ok(included.length < pieces.length, 'sanity: COALESCE_MAX_CHARS left overspill to re-queue');

        // Leftovers are re-queued, not dropped.
        const requeued = redisMock._lists.get('notif:queue') || [];
        assert.equal(requeued.length, pieces.length - included.length);

        const guardEvents = (await readTraceEvents()).filter(e => e.kind === 'oversized_guard' && e.room === ROOM);
        assert.ok(guardEvents.some(e => e.strategy === 'text_chunks'), 'oversized_guard trace emitted');
    });
});

// --- tryEdit truncation (chunking impossible on updates) --------------------

describe('oversized trackable edit is truncated in place', () => {
    let redisMock, botRedisMock, adapterMock, originals;

    beforeEach(() => {
        redisMock = makeRedisMock();
        botRedisMock = makeBotRedisMock({ [CONV_ID]: CONVERSATION_REF });
        adapterMock = makeAdapterMock();
        originals = _setInternalsForTest({
            redis: redisMock,
            redisBot: botRedisMock,
            adapter: adapterMock
        });
    });

    after(() => {
        _setInternalsForTest(originals || {});
    });

    it('updates the tracked activity with byte-fit truncated text', async () => {
        const sha = 'f22353b';
        const it2 = {
            raw: 'raw-envelope',
            env: { type: 'msg', room: ROOM, message: 'tiny update', extra: { dedup_key: `github:commit:${ sha }`, _commit_sha: sha } }
        };
        const recent = {
            activityId: 'act-42',
            ts: Date.now(),
            type: 'msg',
            text: 'stale',
            header: '📦 burst header',
            items: [],
            header_identity: null,
            appended_count: 0,
            conversationId: CONV_ID,
            commit_sha: sha
        };
        // Pre-merged text mimicking the incident: dozens of folded commits,
        // ~40 K chars — the updateActivity payload that kept getting rejected.
        const preMerged = {
            text: 'z'.repeat(40000),
            header: '📦 burst header',
            items: [],
            mergeMode: 'github_trackable_batch'
        };
        const s = stats();

        const edited = await tryEdit(ROOM, it2, recent, s, { preMerged });

        assert.ok(edited, 'tryEdit reports success');
        assert.equal(adapterMock._calls.length, 1, 'exactly one update');
        const updated = adapterMock._calls[0].updatedActivities[0];
        assert.equal(updated.id, 'act-42', 'tracked activity is updated, not re-sent');
        assert.ok(activityBytes(updated) <= MAX_ACTIVITY_BYTES, 'updated activity fits Teams');
        assert.ok(updated.text.endsWith(TRUNCATION_MARKER), 'reader sees the ✂ truncation marker');

        // Stored recent state must carry the visible (truncated) text so the
        // tracked activity stops inflating on every later merge.
        const stored = JSON.parse(redisMock._hashes.get(`notif:recent:${ ROOM }`)[`github:commit:${ sha }`]);
        assert.equal(stored.text, updated.text, 'truncated text is persisted, not the 40 KB original');
        assert.equal(s.edited, 1);

        const guardEvents = (await readTraceEvents()).filter(e => e.kind === 'oversized_guard' && e.room === ROOM);
        assert.ok(guardEvents.some(e => e.strategy === 'text_truncate'), 'oversized_guard trace emitted');
    });
});

describe('guardActivity strip_to_text_stub budgets its preserved id', () => {
    it('M1: stub activity with a tracked id still serializes under the cap', () => {
        const huge = {
            type: 'AdaptiveCard',
            body: Array.from({ length: 300 }, (_, i) => ({ type: 'TextBlock', text: `b${ i } ` + '🚀'.repeat(500) }))
        };
        const activity = {
            type: 'message',
            id: 'x'.repeat(200),
            text: 't'.repeat(MAX_ACTIVITY_BYTES - 10),
            attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: huge }]
        };
        const { activity: safe, wasGuarded } = guardActivity(activity, ROOM);
        assert.ok(wasGuarded);
        assert.ok(!('attachments' in safe), 'unfittable card content is stripped');
        assert.equal(safe.id, 'x'.repeat(200), 'tracked id survives the downgrade');
        assert.ok(
            activityBytes(safe) <= MAX_ACTIVITY_BYTES,
            'id bytes must be counted in the stub budget — otherwise the guard returns an oversized activity'
        );
        assert.ok(safe.text.includes('reduced'), 'reader is told the content was reduced');
    });
});
