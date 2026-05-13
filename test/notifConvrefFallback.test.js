// Regression tests for the "tryEdit must fall back to a constructed
// ConversationReference" bug.
//
// Original failure mode: when the bot has never received an inbound message
// in a channel, its `convref:{conversationId}` key is never written to
// Redis. `handleSingleNew` already handles that case by calling
// `tryConstructedConvRef`, so the FIRST event in such a channel ships
// successfully and `recent` gets saved with the channel's conversationId.
// But `tryEdit` used to give up at the first `loadConvRef → null`, so every
// SUBSEQUENT event for the same trackable bailed back to `handleSingleNew`
// and spawned a brand-new top-level message. The user observed this as
// "PR #415 events not grouping" — five events in one tick, five separate
// messages.
//
// These tests drive the full handleTrackable path with mock redis + adapter,
// so a future refactor that breaks the constructed-ref fallback fails CI.

const { describe, it, before, beforeEach, after } = require('node:test');
const assert = require('node:assert/strict');

process.env.NOTIF_TRACE_LOG = '0';

const consumer = require('../server/queue/notificationConsumer');
const { handleTrackable, _setInternalsForTest, buildConstructedConvRef } = consumer;

// Minimal in-memory mock of the ioredis subset used by handleTrackable /
// tryEdit / handleSingleNew. Tracks calls so tests can assert on them.
function makeRedisMock() {
    const hashes = new Map(); // key → { field: stringValue }
    const calls = { hget: [], hset: [], incr: [] };
    return {
        _hashes: hashes,
        _calls: calls,
        async hget(key, field) {
            calls.hget.push({ key, field });
            const h = hashes.get(key);
            return h ? h[field] || null : null;
        },
        async hset(key, field, value) {
            calls.hset.push({ key, field, value });
            const h = hashes.get(key) || {};
            h[field] = value;
            hashes.set(key, h);
            return 1;
        },
        async incr(key) {
            calls.incr.push(key);
            return 1;
        },
        async get() { return null; },
        async expire() { return 1; },
        async lrem() { return 1; },
        pipeline() {
            const ops = [];
            return {
                hset: (key, field, value) => { ops.push({ op: 'hset', key, field, value }); return this; },
                expire: () => this,
                zadd: () => this,
                zremrangebyscore: () => this,
                async exec() {
                    for (const o of ops) {
                        if (o.op === 'hset') {
                            const h = hashes.get(o.key) || {};
                            h[o.field] = o.value;
                            hashes.set(o.key, h);
                        }
                    }
                    return [];
                }
            };
        }
    };
}

// Mock the bot Redis specifically: convrefs map of conversationId → stored JSON.
// Use `null` for "no convref stored" to simulate a channel where the bot has
// never seen an inbound activity.
function makeBotRedisMock(convrefs = {}) {
    return {
        _gets: [],
        async get(key) {
            this._gets.push(key);
            const m = key.match(/^convref:(.+)$/);
            if (!m) return null;
            const v = convrefs[m[1]];
            return v == null ? null : JSON.stringify(v);
        },
        async incr() { return 1; }
    };
}

// Mock the BotFrameworkAdapter. Records every continueConversation call so
// tests can inspect which conversationRef was used and whether the proactive
// callback invoked sendActivity vs updateActivity.
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

describe('tryEdit constructed-convref fallback', () => {
    let originals;
    let redisMock;
    let botRedisMock;
    let adapterMock;

    const ROOM = 'int-dev-announce';
    const CONV_ID = '19:test-conv-id@thread.v2';
    const DEDUP = 'github:pr:detain/sugarcraft:415';

    before(() => {
        // Suppress noisy console output that the consumer emits for the
        // happy-path edit log line and friends.
        // (Tests can still inspect failures via assertions.)
    });

    beforeEach(() => {
        redisMock = makeRedisMock();
        // No convref stored — this is the "bot never saw inbound" scenario.
        botRedisMock = makeBotRedisMock({});
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

    function seedRecent(extra = {}) {
        // Pre-populate the recent hash as if a previous tick had already
        // sent an opening activity via tryConstructedConvRef.
        const key = `notif:recent:${ ROOM }`;
        const value = {
            activityId: 'act-1',
            ts: Date.now() - 1000,
            type: 'msg',
            text: '🔀 PR opened: detain/sugarcraft#415',
            header: '🔀 PR opened: detain/sugarcraft#415',
            items: [],
            header_identity: 'pr:detain/sugarcraft:415:opened:aaaaaaa',
            appended_count: 0,
            conversationId: CONV_ID,
            commit_sha: 'pr:detain/sugarcraft:415',
            ...extra
        };
        redisMock._hashes.set(key, { [DEDUP]: JSON.stringify(value) });
    }

    function prItem(action, headSha, message) {
        return {
            raw: `raw-${ action }`,
            env: {
                type: 'msg',
                room: ROOM,
                message,
                extra: {
                    event_type: 'pull_request',
                    repo: 'detain/sugarcraft',
                    dedup_key: DEDUP,
                    _commit_sha: 'pr:detain/sugarcraft:415',
                    data: { action, pull_request: { number: 415, head: { sha: headSha } } }
                }
            }
        };
    }

    it('edits via a constructed convref when loadConvRef returns null', async () => {
        seedRecent();
        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
        await handleTrackable(ROOM, prItem('synchronize', 'bbbbbbb', '🔁 PR synchronized: detain/sugarcraft#415'), stats);

        assert.equal(stats.edited, 1, 'should count as an edit, not a new send');
        assert.equal(stats.sent, 0, 'must NOT spawn a new top-level message');

        // adapter.continueConversation was called exactly once …
        assert.equal(adapterMock._calls.length, 1);
        const call = adapterMock._calls[0];
        // … and used a CONSTRUCTED ref (the fallback path).
        assert.equal(call.conversationRef._constructed, true,
            'fallback must use a constructed convref when loadConvRef misses');
        // … and called updateActivity (not sendActivity).
        assert.equal(call.updatedActivities.length, 1, 'must call updateActivity to edit');
        assert.equal(call.sentActivities.length, 0, 'must NOT call sendActivity');
        // The activity id being updated is the one stored in `recent`.
        assert.equal(call.updatedActivities[0].id, 'act-1');
    });

    it('still prefers a stored convref when one exists', async () => {
        seedRecent();
        const realRef = {
            serviceUrl: 'https://smba.trafficmanager.net/teams/',
            conversation: { id: CONV_ID },
            channelId: 'msteams',
            bot: { id: 'real-bot' }
            // note: no _constructed flag
        };
        botRedisMock = makeBotRedisMock({ [CONV_ID]: realRef });
        _setInternalsForTest({ redisBot: botRedisMock });

        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
        await handleTrackable(ROOM, prItem('closed', 'ccccccc', '✅ PR closed: detain/sugarcraft#415'), stats);

        assert.equal(stats.edited, 1);
        assert.equal(adapterMock._calls.length, 1);
        assert.notEqual(adapterMock._calls[0].conversationRef._constructed, true,
            'stored convref must be used as-is when present');
        assert.equal(adapterMock._calls[0].updatedActivities.length, 1);
    });

    it('persists conversationId so the next event edits the same activity', async () => {
        // Simulate: tick has two PR events for the same PR. The first one
        // has no recent yet, sends a new activity via constructed ref. The
        // second one looks up recent, finds it, must edit — and the edit
        // must land on the SAME activity id, not spawn another send.
        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };

        // First event — no recent yet.
        await handleTrackable(ROOM, prItem('opened', 'aaaaaaa', '🔀 PR opened: detain/sugarcraft#415'), stats);
        assert.equal(stats.sent, 1);
        assert.equal(stats.edited, 0);

        // Second event arrives in the same tick. After the first call, a
        // recent entry should exist in redisMock._hashes.
        await handleTrackable(ROOM, prItem('synchronize', 'bbbbbbb', '🔁 PR synchronized: detain/sugarcraft#415'), stats);

        // One send (the first event) + one edit (the second event).
        assert.equal(stats.sent, 1, 'second event must not become a second send');
        assert.equal(stats.edited, 1, 'second event must edit the first activity');

        // Both adapter calls used a constructed ref (no real convref stored).
        assert.equal(adapterMock._calls.length, 2);
        for (const c of adapterMock._calls) {
            assert.equal(c.conversationRef._constructed, true);
        }
        // First call sent, second call updated.
        assert.equal(adapterMock._calls[0].sentActivities.length, 1);
        assert.equal(adapterMock._calls[0].updatedActivities.length, 0);
        assert.equal(adapterMock._calls[1].sentActivities.length, 0);
        assert.equal(adapterMock._calls[1].updatedActivities.length, 1);
        // The edit targets the activity id that the send returned.
        assert.equal(adapterMock._calls[1].updatedActivities[0].id, 'new-activity-1');
    });
});

describe('buildConstructedConvRef', () => {
    it('builds a ref with the marker + the channel conversation id', () => {
        const ref = buildConstructedConvRef('int-dev-announce', '19:abc@thread.v2');
        assert.equal(ref._constructed, true);
        assert.equal(ref.conversation.id, '19:abc@thread.v2');
        assert.equal(ref.conversation.name, 'int-dev-announce');
        assert.equal(ref.conversation.isGroup, true);
        assert.equal(ref.channelId, 'msteams');
        assert.ok(ref.serviceUrl, 'serviceUrl must be set so the adapter can authenticate');
    });
});
