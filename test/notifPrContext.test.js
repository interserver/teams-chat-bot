// Regression tests for branch → PR routing.
//
// `push`, `delete`, and `issue_comment` events have no direct link back to
// the PR they belong to. Without help they each spawn their own top-level
// Teams message. attachPrContext bridges that gap:
//   - records `pull_request.head.ref → pr_number` in a Redis index
//   - rewrites `push`/`delete`-on-branch envelopes to the PR's dedup_key
//   - rewrites `issue_comment`-on-PR envelopes to the PR's dedup_key
//   - condenses regular-issue comments onto a shared issue dedup_key
//   - rewrites the verbose message of a `delete` event so it says *what*
//     was deleted instead of "triggered a delete event"

const { describe, it, before, beforeEach, after } = require('node:test');
const assert = require('node:assert/strict');

process.env.NOTIF_TRACE_LOG = '0';

const consumer = require('../server/queue/notificationConsumer');
const {
    attachPrContext, recordPrBranch, lookupPrByBranch,
    prBranchKey, prNumberFromIssue, branchFromRef,
    _setInternalsForTest
} = consumer;

// In-memory Redis mock with the minimal surface attachPrContext needs:
// set/get with millisecond TTL.
function makeRedisMock() {
    const kv = new Map();
    return {
        _kv: kv,
        async set(key, value, mode, ttlMs) {
            kv.set(key, { value: String(value), expiresAt: mode === 'PX' ? Date.now() + ttlMs : Infinity });
            return 'OK';
        },
        async get(key) {
            const entry = kv.get(key);
            if (!entry) return null;
            if (entry.expiresAt < Date.now()) { kv.delete(key); return null; }
            return entry.value;
        }
    };
}

describe('attachPrContext', () => {
    let originals;
    let redisMock;

    before(() => {
        redisMock = makeRedisMock();
        originals = _setInternalsForTest({ redis: redisMock });
    });
    after(() => {
        _setInternalsForTest(originals || {});
    });

    beforeEach(() => {
        redisMock._kv.clear();
    });

    it('pull_request event seeds the branch→PR index', async () => {
        const env = {
            type: 'msg',
            message: '🔀 opened PR #415',
            extra: {
                event_type: 'pull_request',
                repo: 'detain/sugarcraft',
                data: {
                    action: 'opened',
                    pull_request: { number: 415, head: { sha: 'aaa', ref: 'ai/sugar-dash-batch5' } }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false, 'pull_request itself is not rerouted');
        // The branch index now resolves the head branch back to the PR number.
        const lookup = await lookupPrByBranch('detain/sugarcraft', 'ai/sugar-dash-batch5');
        assert.equal(lookup, 415);
    });

    it('push to a recorded PR head branch is rerouted to the PR trackable', async () => {
        await recordPrBranch('detain/sugarcraft', 'ai/sugar-dash-batch5', 415);
        const env = {
            type: 'msg',
            message: '📦 detain pushed 1 commit',
            extra: {
                event_type: 'push',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:commit:eaa9f3d',
                _commit_sha: 'eaa9f3d',
                data: {
                    ref: 'refs/heads/ai/sugar-dash-batch5',
                    pusher: { name: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, true);
        assert.equal(result.reason, 'branch_is_pr_head');
        assert.equal(result.pr_number, 415);
        assert.equal(env.extra.dedup_key, 'github:pr:detain/sugarcraft:415');
        assert.equal(env.extra._commit_sha, 'pr:detain/sugarcraft:415');
    });

    it('push to an unrelated branch is NOT rerouted', async () => {
        await recordPrBranch('detain/sugarcraft', 'ai/sugar-dash-batch5', 415);
        const env = {
            type: 'msg',
            message: '📦 detain pushed 1 commit',
            extra: {
                event_type: 'push',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:commit:zzz9999',
                _commit_sha: 'zzz9999',
                data: { ref: 'refs/heads/master' }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false);
        assert.equal(env.extra.dedup_key, 'github:commit:zzz9999', 'untouched');
    });

    it('delete event with a known PR branch is rerouted AND its message is upgraded', async () => {
        await recordPrBranch('detain/sugarcraft', 'ai/sugar-dash-batch5', 415);
        const env = {
            type: 'msg',
            message: 'ℹ️ detain triggered a **delete** event on [detain/sugarcraft](url).',
            extra: {
                event_type: 'delete',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:delete:detain/sugarcraft',
                data: {
                    ref: 'ai/sugar-dash-batch5',
                    ref_type: 'branch',
                    sender: { login: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, true);
        assert.equal(result.pr_number, 415);
        assert.equal(env.extra.dedup_key, 'github:pr:detain/sugarcraft:415');

        // Message now says what was deleted.
        assert.ok(env.message.includes('deleted'));
        assert.ok(env.message.includes('branch'));
        assert.ok(env.message.includes('ai/sugar-dash-batch5'),
            'message must mention the deleted branch name');
        assert.ok(!env.message.includes('triggered a **delete** event'),
            'generic stub must be replaced');
    });

    it('create event rewrites the generic stub to say what was created', async () => {
        // create normally fires BEFORE the pull_request opened event (the
        // branch has to exist before a PR can target it), so no prbranch
        // mapping exists yet — the value here is the message rewrite.
        const env = {
            type: 'msg',
            message: 'ℹ️ detain triggered a **create** event on [detain/sugarcraft](url).',
            extra: {
                event_type: 'create',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:create:detain/sugarcraft',
                data: {
                    ref: 'ai/doc-review-components',
                    ref_type: 'branch',
                    sender: { login: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false, 'no PR yet → not rerouted');
        assert.ok(env.message.includes('created'));
        assert.ok(env.message.includes('branch'));
        assert.ok(env.message.includes('ai/doc-review-components'));
        assert.ok(!env.message.includes('triggered a **create** event'),
            'generic stub must be replaced');
    });

    it('create event for a branch already tied to a PR routes to the PR trackable', async () => {
        // Defensive: if a PR somehow opens before its create event lands
        // (e.g. webhook reordering), reuse the PR's trackable instead of
        // spawning a new top-level message.
        await recordPrBranch('detain/sugarcraft', 'pre-opened-branch', 999);
        const env = {
            type: 'msg',
            message: 'ℹ️ detain triggered a **create** event on [detain/sugarcraft](url).',
            extra: {
                event_type: 'create',
                repo: 'detain/sugarcraft',
                data: {
                    ref: 'pre-opened-branch',
                    ref_type: 'branch',
                    sender: { login: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, true);
        assert.equal(result.reason, 'branch_became_pr_head');
        assert.equal(env.extra.dedup_key, 'github:pr:detain/sugarcraft:999');
    });

    it('create event for a TAG is not rerouted (only branches map to PRs)', async () => {
        const env = {
            type: 'msg',
            message: 'ℹ️ detain triggered a **create** event on [detain/sugarcraft](url).',
            extra: {
                event_type: 'create',
                repo: 'detain/sugarcraft',
                data: {
                    ref: 'v1.2.3',
                    ref_type: 'tag',
                    sender: { login: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false);
        assert.ok(env.message.includes('created tag `v1.2.3`'),
            'tag creation still gets a message rewrite');
    });

    it('delete event with NO known PR branch still gets the message upgrade', async () => {
        const env = {
            type: 'msg',
            message: 'ℹ️ detain triggered a **delete** event on [detain/sugarcraft](url).',
            extra: {
                event_type: 'delete',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:delete:detain/sugarcraft',
                data: {
                    ref: 'some-stale-branch',
                    ref_type: 'branch',
                    sender: { login: 'detain' }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false, 'no PR match → not rerouted');
        assert.ok(env.message.includes('deleted'));
        assert.ok(env.message.includes('some-stale-branch'),
            'still tells the user what was deleted');
    });

    it('issue_comment on a PR (per html_url) routes to the PR trackable', async () => {
        // The producer occasionally drops issue.pull_request, but
        // html_url containing /pull/{n} is a reliable PR marker.
        const env = {
            type: 'msg',
            message: 'ℹ️ bot triggered a issue_comment event',
            extra: {
                event_type: 'issue_comment',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:issue_comment:detain/sugarcraft',
                data: {
                    action: 'created',
                    issue: {
                        number: 418,
                        pull_request: null,
                        html_url: 'https://github.com/detain/sugarcraft/pull/418'
                    }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, true);
        assert.equal(result.reason, 'issue_is_pr');
        assert.equal(result.pr_number, 418);
        assert.equal(env.extra.dedup_key, 'github:pr:detain/sugarcraft:418');
    });

    it('issue_comment on a regular issue groups by issue number', async () => {
        const env = {
            type: 'msg',
            message: 'comment on issue',
            extra: {
                event_type: 'issue_comment',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:issue_comment:detain/sugarcraft',
                data: {
                    action: 'created',
                    issue: {
                        number: 99,
                        html_url: 'https://github.com/detain/sugarcraft/issues/99'
                    }
                }
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, true);
        assert.equal(result.reason, 'group_by_issue');
        assert.equal(env.extra.dedup_key, 'github:issue:detain/sugarcraft:99');
    });

    it('events for other types pass through unchanged', async () => {
        const env = {
            type: 'msg',
            message: 'release notes',
            extra: {
                event_type: 'release',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:commit:abc1234',
                _commit_sha: 'abc1234',
                data: {}
            }
        };
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false);
        assert.equal(env.extra.dedup_key, 'github:commit:abc1234');
    });

    it('lookupPrByBranch returns null after the TTL window elapses', async () => {
        // The mock's TTL math is millisecond-precise. Setting a 1ms TTL and
        // waiting 5ms is enough for the entry to be considered expired.
        await redisMock.set(prBranchKey('x/y', 'b'), '7', 'PX', 1);
        await new Promise(r => setTimeout(r, 5));
        const got = await lookupPrByBranch('x/y', 'b');
        assert.equal(got, null);
    });
});

describe('prNumberFromIssue + branchFromRef helpers', () => {
    it('prNumberFromIssue prefers pull_request marker when set', () => {
        assert.equal(prNumberFromIssue({ number: 12, pull_request: { url: 'x' } }), 12);
    });
    it('prNumberFromIssue falls back to /pull/{n} in html_url', () => {
        assert.equal(prNumberFromIssue({ number: 12, html_url: 'https://github.com/o/r/pull/12' }), 12);
    });
    it('prNumberFromIssue returns null for /issues/ urls', () => {
        assert.equal(prNumberFromIssue({ number: 9, html_url: 'https://github.com/o/r/issues/9' }), null);
    });
    it('prNumberFromIssue returns null for missing data', () => {
        assert.equal(prNumberFromIssue(null), null);
        assert.equal(prNumberFromIssue({}), null);
    });
    it('branchFromRef strips refs/heads/ prefix', () => {
        assert.equal(branchFromRef('refs/heads/feature/foo'), 'feature/foo');
        assert.equal(branchFromRef('refs/heads/main'), 'main');
        assert.equal(branchFromRef(''), '');
        assert.equal(branchFromRef(null), '');
    });
});
