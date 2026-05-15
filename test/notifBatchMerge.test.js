// Tests for the in-memory batch-merge optimization: when N notifications
// sharing a dedup_key are pulled from the queue in the same tick, the
// consumer should fold them in memory and issue ONE API call (1 edit if a
// recent activity is in the edit window, 1 send otherwise) instead of one
// API call per envelope.

const { describe, it, beforeEach, after } = require('node:test');
const assert = require('node:assert/strict');

process.env.NOTIF_TRACE_LOG = '0';

const consumer = require('../server/queue/notificationConsumer');
const {
    groupTrackableByDedup,
    canBatchMergeGroup,
    handleTrackableBatch,
    _setInternalsForTest
} = consumer;

// --- Mocks (mirror notifConvrefFallback.test.js) ---------------------------

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

// --- Pure function tests ---------------------------------------------------

describe('groupTrackableByDedup', () => {
    it('groups items by dedup_key preserving order', () => {
        const items = [
            { env: { extra: { dedup_key: 'a' } } },
            { env: { extra: { dedup_key: 'b' } } },
            { env: { extra: { dedup_key: 'a' } } },
            { env: { extra: { dedup_key: 'a' } } }
        ];
        const groups = groupTrackableByDedup(items);
        assert.equal(groups.size, 2);
        assert.equal(groups.get('a').length, 3);
        assert.equal(groups.get('b').length, 1);
        // First-seen order is preserved within a group
        assert.equal(groups.get('a')[0], items[0]);
        assert.equal(groups.get('a')[1], items[2]);
        assert.equal(groups.get('a')[2], items[3]);
    });

    it('skips items without a dedup_key', () => {
        const items = [
            { env: { extra: { dedup_key: 'a' } } },
            { env: { extra: {} } },
            { env: {} }
        ];
        const groups = groupTrackableByDedup(items);
        assert.equal(groups.size, 1);
        assert.equal(groups.get('a').length, 1);
    });

    it('handles an empty input', () => {
        const groups = groupTrackableByDedup([]);
        assert.equal(groups.size, 0);
    });
});

describe('canBatchMergeGroup', () => {
    function msgItem(sha) {
        return { env: { type: 'msg', extra: { _commit_sha: sha } } };
    }
    function cardItem(sha) {
        return { env: { type: 'card', extra: { _commit_sha: sha } } };
    }

    it('returns false for a group of size 1', () => {
        assert.equal(canBatchMergeGroup([msgItem('abc')]), false);
    });

    it('returns true for two msg items sharing a SHA', () => {
        assert.equal(canBatchMergeGroup([msgItem('abc'), msgItem('abc')]), true);
    });

    it('returns false when SHAs differ', () => {
        assert.equal(canBatchMergeGroup([msgItem('abc'), msgItem('def')]), false);
    });

    it('returns false when any item is a card', () => {
        assert.equal(canBatchMergeGroup([msgItem('abc'), cardItem('abc')]), false);
    });

    it('returns true when none of the group items carry a SHA (non-github)', () => {
        const a = { env: { type: 'msg', extra: {} } };
        const b = { env: { type: 'msg', extra: {} } };
        assert.equal(canBatchMergeGroup([a, b]), true);
    });

    it('returns false when one item has a SHA and the other does not', () => {
        const a = { env: { type: 'msg', extra: { _commit_sha: 'abc' } } };
        const b = { env: { type: 'msg', extra: {} } };
        assert.equal(canBatchMergeGroup([a, b]), false);
    });
});

// --- Integration tests: handleTrackableBatch end-to-end --------------------

describe('handleTrackableBatch', () => {
    let originals;
    let redisMock;
    let botRedisMock;
    let adapterMock;

    const ROOM = 'int-dev-announce';
    const CONV_ID = '19:test-conv-id@thread.v2';
    const SHA = '4fd48e4';
    const DEDUP = `github:commit:${ SHA }`;

    function jobItem(jobName, status, conclusion = '') {
        return {
            raw: `raw-${ jobName }-${ status }`,
            env: {
                type: 'msg',
                room: ROOM,
                message: `${ status === 'in_progress' ? '⏳' : status === 'queued' ? '🔄' : '✅' } Workflow **${ jobName }** ${ status } for detain/scoop-emulators on \`master\` (view run)`,
                extra: {
                    event_type: 'workflow_job',
                    repo: 'detain/scoop-emulators',
                    dedup_key: DEDUP,
                    _commit_sha: SHA,
                    data: { workflow_job: { name: jobName, workflow_name: jobName, status, conclusion, html_url: `https://example/${ jobName }` } }
                }
            }
        };
    }

    function seedRecent(extra = {}) {
        const key = `notif:recent:${ ROOM }`;
        const value = {
            activityId: 'act-1',
            ts: Date.now() - 1000,
            type: 'msg',
            text: 'commit 4fd48e4 updates to grouping',
            header: 'commit 4fd48e4 updates to grouping',
            items: [],
            header_identity: null,
            appended_count: 0,
            conversationId: CONV_ID,
            commit_sha: SHA,
            ...extra
        };
        redisMock._hashes.set(key, { [DEDUP]: JSON.stringify(value) });
    }

    beforeEach(() => {
        redisMock = makeRedisMock();
        // No stored convref → batch send/edit uses constructed ref
        botRedisMock = makeBotRedisMock({});
        adapterMock = makeAdapterMock();
        // Seed the processing list with raw values so ackOne can lrem them
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

    it('folds 3 in-window events into ONE edit', async () => {
        seedRecent();
        const group = [
            jobItem('build', 'queued'),
            jobItem('test', 'queued'),
            jobItem('deploy', 'queued')
        ];
        // Track processing-list state so we can assert ack-all behavior.
        for (const it of group) redisMock._lists.get('notif:processing').push(it.raw);

        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
        await handleTrackableBatch(ROOM, group, stats);

        // Critical: 1 edit, 0 sends — that's the whole point of the batch.
        assert.equal(stats.edited, 1, 'batch must collapse to a single edit');
        assert.equal(stats.sent, 0, 'no new sends when in-window recent exists');

        // Adapter received exactly one call, and it was an updateActivity.
        assert.equal(adapterMock._calls.length, 1, 'exactly one adapter call');
        assert.equal(adapterMock._calls[0].updatedActivities.length, 1);
        assert.equal(adapterMock._calls[0].sentActivities.length, 0);
        // The edit targets the activity we seeded.
        assert.equal(adapterMock._calls[0].updatedActivities[0].id, 'act-1');

        // All three items were ack'd (removed from notif:processing).
        assert.deepEqual(redisMock._lists.get('notif:processing'), []);

        // Persisted recent should carry all three job identities in items[].
        const stored = JSON.parse(redisMock._hashes.get(`notif:recent:${ ROOM }`)[DEDUP]);
        assert.equal(stored.activityId, 'act-1', 'edits do not change activityId');
        assert.equal(stored.items.length, 3, 'all three jobs appear in items');
        // appended_count tracks the number of folded events for this edit.
        assert.equal(stored.appended_count, 3, 'appended_count increments by group size');
    });

    it('with no recent, folds 3 events into ONE send', async () => {
        const group = [
            jobItem('build', 'queued'),
            jobItem('test', 'in_progress'),
            jobItem('deploy', 'in_progress')
        ];
        for (const it of group) redisMock._lists.get('notif:processing').push(it.raw);

        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
        await handleTrackableBatch(ROOM, group, stats);

        // No recent existed → batch becomes 1 send (carrying the folded content).
        assert.equal(stats.sent, 1, 'batch must collapse to a single send');
        assert.equal(stats.edited, 0, 'no edits when there is no recent activity');

        assert.equal(adapterMock._calls.length, 1, 'exactly one adapter call');
        assert.equal(adapterMock._calls[0].sentActivities.length, 1);
        assert.equal(adapterMock._calls[0].updatedActivities.length, 0);

        // All three items were ack'd.
        assert.deepEqual(redisMock._lists.get('notif:processing'), []);

        // Recent should be saved with the merged state under the dedup_key.
        const stored = JSON.parse(redisMock._hashes.get(`notif:recent:${ ROOM }`)[DEDUP]);
        assert.ok(stored.activityId, 'send response activityId persisted');
        assert.equal(stored.commit_sha, SHA);
        assert.ok(stored.items.length >= 1, 'merged state captures multiple jobs');
    });

    it('falls back to a fresh send when the in-window edit fails', async () => {
        seedRecent();
        // Adapter that throws on update but succeeds on send.
        adapterMock.continueConversation = async (conversationRef, callback) => {
            const record = { conversationRef, sentActivities: [], updatedActivities: [] };
            adapterMock._calls.push(record);
            const ctx = {
                async sendActivity(a) {
                    record.sentActivities.push(a);
                    return { id: 'new-after-fail-' + adapterMock._calls.length };
                },
                async updateActivity() {
                    throw new Error('simulated edit failure');
                }
            };
            await callback(ctx);
        };

        const group = [jobItem('a', 'queued'), jobItem('b', 'queued')];
        for (const it of group) redisMock._lists.get('notif:processing').push(it.raw);

        const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
        await handleTrackableBatch(ROOM, group, stats);

        // Edit failed → batch must have fallen through to a new send with
        // the merged content. Stats reflect 1 send.
        assert.equal(stats.sent, 1, 'edit failure → 1 send with merged content');
        // edited counter is incremented inside tryEdit only after a successful
        // updateActivity; it must NOT increment when update throws. (The retry
        // helper will surface the error and tryEdit returns false.)
        assert.equal(stats.edited, 0);

        // Each adapter call records its outcomes; assert that updateActivity
        // attempts happened (one or more, depending on retries) and that
        // exactly one sendActivity succeeded.
        const updateAttempts = adapterMock._calls.reduce((n, c) => n + c.updatedActivities.length, 0);
        const sendAttempts = adapterMock._calls.reduce((n, c) => n + c.sentActivities.length, 0);
        assert.equal(updateAttempts, 0, 'failed updates throw before recording');
        assert.equal(sendAttempts, 1, 'one new send carries the merged content');

        // All items in the group still get ack'd.
        assert.deepEqual(redisMock._lists.get('notif:processing'), []);
    });
});
