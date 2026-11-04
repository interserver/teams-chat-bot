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
        assert.ok(env.message.includes('created tag **v1.2.3**'),
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
        // Standalone message: bold ref (not a code span) and the repo link,
        // since no parent header names the repo.
        assert.equal(env.message,
            '🗑️ detain deleted branch **some-stale-branch** in [detain/sugarcraft](https://github.com/detain/sugarcraft).');
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

// ---------------------------------------------------------------------------
// Release event rendering
// ---------------------------------------------------------------------------

const RELEASE_STUB = 'ℹ️ github-actions[bot] triggered a **release** event (edited) on [detain/sugarcraft](https://github.com/detain/sugarcraft).';

describe('describeReleaseEvent', () => {
    const { describeReleaseEvent } = consumer;

    it('describes an edited release with name + tag note + links', () => {
        const line = describeReleaseEvent({
            action: 'edited',
            release: {
                tag_name: 'v1.2.3',
                name: 'Sugarcraft v1.2.3',
                html_url: 'https://github.com/detain/sugarcraft/releases/tag/v1.2.3'
            }
        }, 'detain/sugarcraft');
        assert.equal(line, '🚀 edited [**Sugarcraft v1.2.3**](https://github.com/detain/sugarcraft/releases/tag/v1.2.3) — v1.2.3 on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('maps released → published and shows bare tag when name matches tag', () => {
        const line = describeReleaseEvent({
            action: 'released',
            release: { tag_name: 'v2.0.0', name: 'v2.0.0', html_url: 'https://github.com/o/r/releases/tag/v2.0.0' }
        }, 'o/r');
        assert.equal(line, '🚀 published [**v2.0.0**](https://github.com/o/r/releases/tag/v2.0.0) on [o/r](https://github.com/o/r).');
    });

    it('marks pre-releases with the test-tube emoji and note', () => {
        const line = describeReleaseEvent({
            action: 'prereleased',
            release: { tag_name: 'v2.0.0-rc1', prerelease: true }
        }, 'detain/sugarcraft');
        assert.equal(line, '🧪 prereleased [**v2.0.0-rc1**](https://github.com/detain/sugarcraft/releases/tag/v2.0.0-rc1) _(pre-release)_ on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('builds the tag URL when html_url is missing', () => {
        const line = describeReleaseEvent({
            release: { tag_name: 'v 1.0 & beta' }
        }, 'o/r');
        assert.ok(line.includes('https://github.com/o/r/releases/tag/v%201.0%20%26%20beta'), line);
    });

    it('sanitizes markdown control characters out of titles', () => {
        const line = describeReleaseEvent({
            release: { tag_name: 'v1', name: 'A [bug] *fix* (wip)' }
        }, 'o/r');
        assert.ok(line.includes('[**A bug fix wip**]'), line);
        assert.ok(!/\[[^\]]*\[[^\]]*\]/.test(line), 'no nested brackets in link label');
    });

    it('caps long names with an ellipsis', () => {
        const line = describeReleaseEvent({
            release: { tag_name: 'v1', name: 'x'.repeat(120) }
        }, 'o/r');
        assert.ok(line.includes('…'), line);
        assert.ok(!line.includes('x'.repeat(81)), 'name capped at 80 chars');
    });

    it('collapses every multi-space run in a title, not just the first', () => {
        const line = describeReleaseEvent({
            release: { tag_name: 'v1', name: 'Foo   Bar    Baz' }
        }, 'o/r');
        assert.ok(line.includes('[**Foo Bar Baz**]'), line);
    });

    it('renders deleted releases with the waste-basket and no dead link', () => {
        const line = describeReleaseEvent({
            action: 'deleted',
            release: { tag_name: 'v1.2.3', html_url: 'https://github.com/detain/sugarcraft/releases/tag/v1.2.3' }
        }, 'detain/sugarcraft');
        assert.equal(line, '🗑️ deleted **v1.2.3** on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('never leaks prototype members through forged actions', () => {
        const line = describeReleaseEvent({ action: 'constructor', release: { tag_name: 'v1' } }, 'o/r');
        assert.ok(line.includes('constructor'), line);
        assert.ok(!line.includes('[native code]'), line);
    });

    it('falls back to the constructed tag URL for unsafe html_url values', () => {
        const js = describeReleaseEvent({ action: 'published', release: { tag_name: 'v1', html_url: 'javascript:alert(1)' } }, 'o/r');
        assert.ok(js.includes('](https://github.com/o/r/releases/tag/v1)'), js);
        const spaced = describeReleaseEvent({ action: 'published', release: { tag_name: 'v1', html_url: 'https://example.com/a b' } }, 'o/r');
        assert.ok(spaced.includes('](https://github.com/o/r/releases/tag/v1)'), spaced);
    });

    it('returns null without a tag_name or with stripped data', () => {
        assert.equal(describeReleaseEvent({ release: {} }, 'o/r'), null);
        assert.equal(describeReleaseEvent({ release: { tag_name: '   ' } }, 'o/r'), null);
        assert.equal(describeReleaseEvent({ __stripped: true }, 'o/r'), null);
        assert.equal(describeReleaseEvent({}, 'o/r'), null);
    });
});

describe('attachPrContext — release stub rewrite', () => {
    let originals;
    let redisMock;

    before(() => {
        redisMock = makeRedisMock();
        originals = _setInternalsForTest({ redis: redisMock });
    });
    after(() => {
        _setInternalsForTest(originals || {});
    });

    function releaseEnv(message, data) {
        return {
            type: 'msg',
            message,
            extra: {
                event_type: 'release',
                repo: 'detain/sugarcraft',
                dedup_key: 'github:release:detain/sugarcraft:v1.2.3',
                data
            }
        };
    }

    it('rewrites the release stub to the described line', async () => {
        const env = releaseEnv(RELEASE_STUB, {
            action: 'edited',
            release: {
                tag_name: 'v1.2.3',
                name: 'Sugarcraft v1.2.3',
                html_url: 'https://github.com/detain/sugarcraft/releases/tag/v1.2.3'
            }
        });
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false);
        assert.equal(env.message, '🚀 edited [**Sugarcraft v1.2.3**](https://github.com/detain/sugarcraft/releases/tag/v1.2.3) — v1.2.3 on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('keeps the stub when the envelope data was stripped (fail-soft)', async () => {
        const env = releaseEnv(RELEASE_STUB, { __stripped: true, release: 'oversized' });
        await attachPrContext(env);
        assert.equal(env.message, RELEASE_STUB);
    });

    it('never clobbers a non-stub producer message', async () => {
        const custom = 'Custom release announcement: v1.2.3 is out!';
        const env = releaseEnv(custom, {
            action: 'published',
            release: { tag_name: 'v1.2.3', html_url: 'https://github.com/detain/sugarcraft/releases/tag/v1.2.3' }
        });
        await attachPrContext(env);
        assert.equal(env.message, custom);
    });
});

describe('condensedBulletText — release branch', () => {
    const { condensedBulletText } = consumer;

    it('condenses a release event to emoji + title + verb', () => {
        const bullet = condensedBulletText({
            extra: {
                event_type: 'release',
                repo: 'detain/sugarcraft',
                data: { action: 'edited', release: { tag_name: 'v1.2.3', name: 'Sugarcraft v1.2.3' } }
            }
        });
        assert.equal(bullet, '🚀 **Sugarcraft v1.2.3** (v1.2.3) — edited');
    });

    it('renders bare tag when no distinct name', () => {
        const bullet = condensedBulletText({
            extra: { event_type: 'release', data: { action: 'published', release: { tag_name: 'v9.9.9' } } }
        });
        assert.equal(bullet, '🚀 **v9.9.9** — published');
    });

    it('marks a deleted release with the waste-basket', () => {
        const bullet = condensedBulletText({
            extra: { event_type: 'release', data: { action: 'deleted', release: { tag_name: 'v9.9.9' } } }
        });
        assert.equal(bullet, '🗑️ **v9.9.9** — deleted');
    });

    it('returns null for stripped or tagless data', () => {
        assert.equal(condensedBulletText({ extra: { event_type: 'release', data: { __stripped: true } } }), null);
        assert.equal(condensedBulletText({ extra: { event_type: 'release', data: { action: 'edited', release: {} } } }), null);
    });
});

// ---------------------------------------------------------------------------
// Package event rendering
// ---------------------------------------------------------------------------

const PACKAGE_STUB = 'ℹ️ github-actions[bot] triggered a **registry_package** event (edited) on [detain/sugarcraft](https://github.com/detain/sugarcraft).';
const PACKAGE_STUB_SHORT = 'ℹ️ github-actions[bot] triggered a **package** event on [detain/sugarcraft](https://github.com/detain/sugarcraft).';

describe('describePackageEvent', () => {
    const { describePackageEvent } = consumer;

    it('describes a published npm package with version link and type note', () => {
        const line = describePackageEvent({
            action: 'published',
            registry_package: {
                name: '@detain/sugarcraft',
                package_type: 'npm',
                package_version: { name: '2.1.0' },
                html_url: 'https://github.com/orgs/detain/packages/@detain/sugarcraft/2.1.0'
            }
        }, 'detain/sugarcraft');
        assert.equal(line, '📦 published [**@detain/sugarcraft@2.1.0**](https://github.com/orgs/detain/packages/@detain/sugarcraft/2.1.0) (npm) on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('keeps scoped-name punctuation intact so the link stays balanced', () => {
        const line = describePackageEvent({
            action: 'updated',
            registry_package: { name: '@detain:sugar/craft', package_version: { name: '1.0.0:beta' } }
        }, 'o/r');
        assert.ok(line.includes('[**@detain:sugar/craft@1.0.0:beta**]'), line);
        assert.ok(!/\[[^\]]*\[[^\]]*\]/.test(line), 'no nested brackets in link label');
    });

    it('omits the version when the payload carries none', () => {
        const line = describePackageEvent({
            action: 'published',
            registry_package: { name: 'sugarcraft' }
        }, 'o/r');
        assert.equal(line, '📦 published [**sugarcraft**](https://github.com/o/r/packages) on [o/r](https://github.com/o/r).');
    });

    it('reads package_version.version and top-level version as fallbacks', () => {
        const viaVersionField = describePackageEvent({
            registry_package: { name: 'p', package_version: { version: '3.2.1' } }
        }, 'o/r');
        assert.ok(viaVersionField.includes('**p@3.2.1**'), viaVersionField);
        const viaTopLevel = describePackageEvent({
            registry_package: { name: 'p', version: '4.5.6' }
        }, 'o/r');
        assert.ok(viaTopLevel.includes('**p@4.5.6**'), viaTopLevel);
    });

    it('maps known package types and lowercases unknown ones', () => {
        const gems = describePackageEvent({ registry_package: { name: 'g', package_type: 'ruby_gems' } }, 'o/r');
        assert.ok(gems.includes('(gems)'), gems);
        const container = describePackageEvent({ registry_package: { name: 'c', package_type: 'CONTAINER_REGISTRY' } }, 'o/r');
        assert.ok(container.includes('(container)'), container);
        const unknown = describePackageEvent({ registry_package: { name: 'v', package_type: 'Vcpkg' } }, 'o/r');
        assert.ok(unknown.includes('(vcpkg)'), unknown);
        const absent = describePackageEvent({ registry_package: { name: 'v' } }, 'o/r');
        assert.ok(!absent.includes(' ('), absent);
    });

    it('accepts the legacy top-level package key', () => {
        const line = describePackageEvent({
            action: 'published',
            package: { name: 'legacy-pkg', package_type: 'nuget' }
        }, 'o/r');
        assert.ok(line.includes('**legacy-pkg**'), line);
        assert.ok(line.includes('(nuget)'), line);
    });

    it('falls back to plain bold when neither html_url nor repo can supply a link', () => {
        const line = describePackageEvent({ registry_package: { name: 'solo' } }, '');
        assert.equal(line, '📦 updated **solo**');
    });

    it('falls back to the repo packages page for unsafe html_url values', () => {
        const js = describePackageEvent({ registry_package: { name: 'p', html_url: 'javascript:alert(1)' } }, 'o/r');
        assert.ok(js.includes('](https://github.com/o/r/packages)'), js);
        const spaced = describePackageEvent({ registry_package: { name: 'p', html_url: 'https://example.com/a b' } }, 'o/r');
        assert.ok(spaced.includes('](https://github.com/o/r/packages)'), spaced);
    });

    it('caps long versions shorter than names', () => {
        const line = describePackageEvent({
            registry_package: { name: 'p', package_version: { name: 'v'.repeat(60) } }
        }, 'o/r');
        assert.ok(line.includes('…'), line);
        assert.ok(!line.includes('v'.repeat(41)), 'version capped at 40 chars');
    });

    it('defaults an empty action to updated and passes unknown actions verbatim', () => {
        assert.ok(describePackageEvent({ registry_package: { name: 'p' } }, 'o/r').includes('📦 updated'));
        assert.ok(describePackageEvent({ action: 'archived', registry_package: { name: 'p' } }, 'o/r').includes('📦 archived'));
    });

    it('never leaks prototype members through forged actions', () => {
        const line = describePackageEvent({ action: 'constructor', registry_package: { name: 'p' } }, 'o/r');
        assert.ok(line.includes('constructor'), line);
        assert.ok(!line.includes('[native code]'), line);
    });

    it('returns null without a usable name or with stripped/absent package data', () => {
        assert.equal(describePackageEvent({ registry_package: {} }, 'o/r'), null);
        assert.equal(describePackageEvent({ registry_package: { name: '   ' } }, 'o/r'), null);
        assert.equal(describePackageEvent({ __stripped: true }, 'o/r'), null);
        assert.equal(describePackageEvent({}, 'o/r'), null);
        assert.equal(describePackageEvent({ action: 'published' }, 'o/r'), null);
        assert.equal(describePackageEvent(null, 'o/r'), null);
    });
});

describe('attachPrContext — package stub rewrite', () => {
    let originals;
    let redisMock;

    before(() => {
        redisMock = makeRedisMock();
        originals = _setInternalsForTest({ redis: redisMock });
    });
    after(() => {
        _setInternalsForTest(originals || {});
    });

    function packageEnv(eventType, message, data) {
        return {
            type: 'msg',
            message,
            extra: {
                event_type: eventType,
                repo: 'detain/sugarcraft',
                dedup_key: 'github:package:detain/sugarcraft:@detain/sugarcraft:2.1.0',
                data
            }
        };
    }

    const PUBLISHED_DATA = {
        action: 'published',
        registry_package: {
            name: '@detain/sugarcraft',
            package_type: 'npm',
            package_version: { name: '2.1.0' },
            html_url: 'https://github.com/orgs/detain/packages/@detain/sugarcraft/2.1.0'
        }
    };

    it('rewrites the registry_package stub to the described line', async () => {
        const env = packageEnv('registry_package', PACKAGE_STUB, PUBLISHED_DATA);
        const result = await attachPrContext(env);
        assert.equal(result.rerouted, false);
        assert.equal(env.message, '📦 published [**@detain/sugarcraft@2.1.0**](https://github.com/orgs/detain/packages/@detain/sugarcraft/2.1.0) (npm) on [detain/sugarcraft](https://github.com/detain/sugarcraft).');
    });

    it('rewrites the package stub under the legacy event_type too', async () => {
        const env = packageEnv('package', PACKAGE_STUB_SHORT, PUBLISHED_DATA);
        await attachPrContext(env);
        assert.ok(env.message.startsWith('📦 published [**@detain/sugarcraft@2.1.0**]'), env.message);
    });

    it('keeps the stub when the payload lacks the package object (fail-soft)', async () => {
        // Today's producer strips registry_package; old queued envelopes
        // carry only action/repo/sender. Rendering must stay on the stub.
        const env = packageEnv('registry_package', PACKAGE_STUB, {
            action: 'published',
            repository: { full_name: 'detain/sugarcraft' },
            sender: { login: 'github-actions[bot]' }
        });
        await attachPrContext(env);
        assert.equal(env.message, PACKAGE_STUB);
    });

    it('keeps the stub when the envelope data was stripped (fail-soft)', async () => {
        const env = packageEnv('registry_package', PACKAGE_STUB, { __stripped: true, registry_package: 'oversized' });
        await attachPrContext(env);
        assert.equal(env.message, PACKAGE_STUB);
    });

    it('never clobbers a non-stub producer message', async () => {
        const custom = 'Hand-crafted package announcement for 2.1.0!';
        const env = packageEnv('registry_package', custom, PUBLISHED_DATA);
        await attachPrContext(env);
        assert.equal(env.message, custom);
    });
});

describe('container package labels + links', () => {
    const { describePackageEvent, packageLabel, packageUrl, condensedBulletText } = consumer;
    const DIGEST = 'sha256:00fdf23e8ee8fcfcf056c62aaab61030b68e2c8036fe3192276a98a6badde8d6';
    // Shape of a real ghcr.io publish delivery (detain/phlix-server, trimmed).
    function containerPkg(tag, extra = {}) {
        return {
            id: 12329883,
            name: 'phlix-hub',
            package_type: 'CONTAINER',
            html_url: 'https://github.com/detain/packages/12329883',
            package_version: {
                id: 1499622292,
                name: DIGEST,
                version: DIGEST,
                html_url: 'https://github.com/users/detain/packages/container/phlix-hub/1499622292',
                container_metadata: { tag: { name: tag, digest: DIGEST } },
                ...extra
            }
        };
    }

    it('labels a tagged container as name:tag', () => {
        assert.equal(packageLabel(containerPkg('latest')), 'phlix-hub:latest');
        assert.equal(packageLabel(containerPkg('7fb50998030b')), 'phlix-hub:7fb50998030b');
    });

    it('shortens a digest-only version to a 12-hex id', () => {
        assert.equal(packageLabel(containerPkg('')), 'phlix-hub@sha256:00fdf23e8ee8');
        // Older producer: no container_metadata, digest as version name.
        assert.equal(packageLabel({ name: 'phlix-hub', package_type: 'CONTAINER', package_version: { name: DIGEST } }),
            'phlix-hub@sha256:00fdf23e8ee8');
    });

    it('leaves non-container versions alone', () => {
        assert.equal(packageLabel({ name: '@detain/sugarcraft', package_type: 'npm', package_version: { name: '2.1.0' } }),
            '@detain/sugarcraft@2.1.0');
        assert.equal(packageLabel({ name: 'bare' }), 'bare');
        assert.equal(packageLabel({ name: '' }), '');
        assert.equal(packageLabel(null), '');
    });

    it('prefers the per-version html_url over the 404ing package html_url', () => {
        assert.equal(packageUrl(containerPkg('latest'), 'detain/phlix-server'),
            'https://github.com/users/detain/packages/container/phlix-hub/1499622292');
    });

    it('skips the known-broken /{owner}/packages/{id} url and falls back to the repo packages tab', () => {
        const pkg = { name: 'phlix-hub', html_url: 'https://github.com/detain/packages/12329883' };
        assert.equal(packageUrl(pkg, 'detain/phlix-server'), 'https://github.com/detain/phlix-server/packages');
    });

    it('keeps a working package-level html_url', () => {
        const pkg = { name: 'p', html_url: 'https://github.com/detain/sugarcraft/packages/8123456' };
        assert.equal(packageUrl(pkg, 'o/r'), 'https://github.com/detain/sugarcraft/packages/8123456');
    });

    it('rejects an unsafe per-version html_url', () => {
        const pkg = { name: 'p', package_version: { html_url: 'javascript:alert(1)' } };
        assert.equal(packageUrl(pkg, 'o/r'), 'https://github.com/o/r/packages');
    });

    it('describes a tagged container publish with tag + working link', () => {
        const line = describePackageEvent({ action: 'published', registry_package: containerPkg('latest') }, 'detain/phlix-server');
        assert.equal(line, '📦 published [**phlix-hub:latest**](https://github.com/users/detain/packages/container/phlix-hub/1499622292) (container) on [detain/phlix-server](https://github.com/detain/phlix-server).');
    });

    it('condenses a folded container bullet to name:tag', () => {
        const bullet = condensedBulletText({
            extra: { event_type: 'registry_package', data: { action: 'published', registry_package: containerPkg('latest') } }
        });
        assert.equal(bullet, '📦 [**phlix-hub:latest**](https://github.com/users/detain/packages/container/phlix-hub/1499622292) — published');
    });
});

describe('attachPackageCommit — nest SHA-tagged images under their commit', () => {
    const { attachPackageCommit, collectTickPackageShas, commitShaFromTag, pkgDigestKey, jobIdentity, mergeGithubTrackable } = consumer;
    const DIGEST = 'sha256:243f1c8ce7b2aa11bb22cc33dd44ee55ff6677889900aabbccddeeff00112233';
    const ROOM = 'int-dev-announce';
    let originals;
    let redisMock;

    function pkgEnv(tag, { digest = DIGEST, room = ROOM, action = 'published' } = {}) {
        return {
            type: 'msg',
            room,
            message: `📦 github-actions[bot] ${ action } container package phlix-hub on detain/phlix-server.`,
            extra: {
                event_type: 'registry_package',
                repo: 'detain/phlix-server',
                dedup_key: `github:package:detain/phlix-server:phlix-hub:${ digest }`,
                _commit_sha: `package:detain/phlix-server:phlix-hub:${ digest }`,
                data: {
                    action,
                    registry_package: {
                        name: 'phlix-hub',
                        package_type: 'CONTAINER',
                        package_version: {
                            name: digest,
                            html_url: 'https://github.com/users/detain/packages/container/phlix-hub/1499627176',
                            container_metadata: { tag: { name: tag, digest } }
                        }
                    }
                }
            }
        };
    }

    function seedCommitTrackable(sha7, { room = ROOM, ageMs = 60_000 } = {}) {
        redisMock._hashes.set(`notif:recent:${ room }`, new Map([[
            `github:commit:${ sha7 }`,
            JSON.stringify({ activityId: 'act-1', ts: Date.now() - ageMs, type: 'msg' })
        ]]));
    }

    before(() => {
        redisMock = makeRedisMock();
        redisMock._hashes = new Map();
        redisMock.hget = async (key, field) => redisMock._hashes.get(key)?.get(field) ?? null;
        // Sorted sets: key → Map(member → score); honours the min-score cutoff.
        redisMock._zsets = new Map();
        redisMock.zrevrangebyscore = async (key, _max, min) => [...(redisMock._zsets.get(key) || new Map())]
            .filter(([, score]) => score >= min)
            .sort((a, b) => b[1] - a[1])
            .map(([member]) => member);
        originals = _setInternalsForTest({ redis: redisMock });
    });
    after(() => {
        _setInternalsForTest(originals || {});
    });
    beforeEach(() => {
        redisMock._kv.clear();
        redisMock._hashes.clear();
        redisMock._zsets.clear();
    });

    function seedLiveCommits(shas, room = ROOM) {
        redisMock._hashes.set(`notif:recent:${ room }`, new Map(shas.map(sha => [
            `github:commit:${ sha }`,
            JSON.stringify({ activityId: `act-${ sha }`, ts: Date.now() - 60_000, type: 'msg' })
        ])));
    }
    function seedActiveCi(repo, shas, ageMs = 60_000) {
        redisMock._zsets.set(`notif:wfactive:${ repo }`, new Map(shas.map((sha, i) => [sha, Date.now() - ageMs - i])));
    }

    it('treats only 7–40 hex tags as commit SHAs', () => {
        assert.equal(commitShaFromTag('f2c9cc915d8b'), 'f2c9cc9');
        assert.equal(commitShaFromTag('F2C9CC915D8B'), 'f2c9cc9');
        assert.equal(commitShaFromTag('latest'), null);
        assert.equal(commitShaFromTag('buildcache'), null);
        assert.equal(commitShaFromTag('abc123'), null);
        assert.equal(commitShaFromTag('v1.2.3'), null);
        assert.equal(commitShaFromTag(''), null);
    });

    it('re-keys a SHA-tagged publish onto the live commit trackable', async () => {
        seedCommitTrackable('f2c9cc9');
        const env = pkgEnv('f2c9cc915d8b');
        const result = await attachPackageCommit(env);
        assert.equal(result.attached, true);
        assert.equal(result.source, 'tag');
        assert.equal(env.extra.dedup_key, 'github:commit:f2c9cc9');
        assert.equal(env.extra._commit_sha, 'f2c9cc9');
        assert.equal(env.extra._original_dedup_key, `github:package:detain/phlix-server:phlix-hub:${ DIGEST }`);
        // Digest index seeded for the `latest` tag that may follow.
        assert.equal(await redisMock.get(pkgDigestKey('phlix-hub', DIGEST)), 'f2c9cc9');
    });

    it('follows the digest index for a `latest` tag drained in a later tick', async () => {
        seedCommitTrackable('f2c9cc9');
        await attachPackageCommit(pkgEnv('f2c9cc915d8b'));
        const latest = pkgEnv('latest');
        const result = await attachPackageCommit(latest);
        assert.equal(result.attached, true);
        assert.equal(result.source, 'digest_index');
        assert.equal(latest.extra.dedup_key, 'github:commit:f2c9cc9');
    });

    it('resolves `latest` from the same tick even when it is drained first', async () => {
        seedCommitTrackable('f2c9cc9');
        const latest = pkgEnv('latest');
        const shaTagged = pkgEnv('f2c9cc915d8b');
        const tickShas = collectTickPackageShas([{ env: latest }, { env: shaTagged }]);
        const result = await attachPackageCommit(latest, tickShas);
        assert.equal(result.attached, true);
        assert.equal(result.source, 'tick_digest');
        assert.equal(latest.extra.dedup_key, 'github:commit:f2c9cc9');
    });

    it('keeps its own trackable when no commit trackable is live in the room', async () => {
        const env = pkgEnv('f2c9cc915d8b');
        const result = await attachPackageCommit(env);
        assert.equal(result.attached, false);
        assert.equal(result.reason, 'no_live_commit_trackable');
        assert.equal(env.extra.dedup_key, `github:package:detain/phlix-server:phlix-hub:${ DIGEST }`);
    });

    it('does not nest under a commit trackable in a different room or past the edit window', async () => {
        seedCommitTrackable('f2c9cc9', { room: 'notifications' });
        assert.equal((await attachPackageCommit(pkgEnv('f2c9cc915d8b'))).attached, false);
        seedCommitTrackable('f2c9cc9', { ageMs: 2 * 60 * 60 * 1000 });
        assert.equal((await attachPackageCommit(pkgEnv('f2c9cc915d8b'))).attached, false);
    });

    it('leaves non-SHA tags without a digest match alone', async () => {
        seedCommitTrackable('f2c9cc9');
        const env = pkgEnv('latest');
        const result = await attachPackageCommit(env);
        assert.equal(result.attached, false);
        assert.equal(result.reason, 'no_commit_tag');
    });

    it('ignores non-package events and stripped payloads', async () => {
        assert.deepEqual(await attachPackageCommit({ extra: { event_type: 'push', data: {} } }), { attached: false });
        assert.deepEqual(await attachPackageCommit({ extra: { event_type: 'registry_package', data: { __stripped: true } } }), { attached: false });
    });

    it('attaches an early `latest` to the single commit with CI running in the package repo', async () => {
        seedLiveCommits(['32cbd66']);
        seedActiveCi('detain/phlix-server', ['32cbd66']);
        const latest = pkgEnv('latest');
        const result = await attachPackageCommit(latest);
        assert.equal(result.attached, true);
        assert.equal(result.source, 'active_ci');
        assert.equal(latest.extra.dedup_key, 'github:commit:32cbd66');
    });

    it('does not guess when several commits have CI running (falls back to waiting)', async () => {
        const { shouldDeferPackage } = consumer;
        seedLiveCommits(['bf49678', 'f2c9cc9']);
        seedActiveCi('detain/phlix-server', ['bf49678', 'f2c9cc9']);
        const latest = pkgEnv('latest');
        const result = await attachPackageCommit(latest);
        assert.equal(result.attached, false);
        assert.equal(result.reason, 'no_commit_tag');
        assert.equal(result.candidates, 2);
        assert.equal(shouldDeferPackage(latest, result), true);
    });

    it('ignores active-CI commits without a live trackable in the room, or with stale CI', async () => {
        seedLiveCommits(['32cbd66'], 'notifications');
        seedActiveCi('detain/phlix-server', ['32cbd66']);
        assert.equal((await attachPackageCommit(pkgEnv('latest'))).candidates, 0);
        seedLiveCommits(['32cbd66']);
        seedActiveCi('detain/phlix-server', ['32cbd66'], consumer.PKG_ACTIVE_CI_MS + 60_000);
        assert.equal((await attachPackageCommit(pkgEnv('latest'))).attached, false);
    });

    it('learns the builder repo from a SHA-tagged publish and uses it for later floating tags', async () => {
        // phlix-hub images are filed under phlix-server but built by phlix-hub's CI.
        await redisMock.set(consumer.shaRepoKey('f2c9cc9'), 'detain/phlix-hub', 'PX', 60_000);
        seedLiveCommits(['f2c9cc9', '58530ec']);
        await attachPackageCommit(pkgEnv('f2c9cc915d8b'));
        assert.equal(await redisMock.get(consumer.pkgRepoKey('phlix-hub')), 'detain/phlix-hub');

        // Next build: phlix-server also has CI running, but only phlix-hub's commit counts.
        seedLiveCommits(['54cbf32', '58530ec']);
        seedActiveCi('detain/phlix-server', ['58530ec']);
        seedActiveCi('detain/phlix-hub', ['54cbf32']);
        const latest = pkgEnv('latest', { digest: 'sha256:' + 'ab'.repeat(32) });
        const result = await attachPackageCommit(latest);
        assert.equal(result.attached, true);
        assert.equal(result.source, 'active_ci');
        assert.equal(latest.extra.dedup_key, 'github:commit:54cbf32');
    });

    it('recordActiveWorkflow seeds the sha → repo index', async () => {
        redisMock.pipeline = () => {
            const chain = { zadd() { return chain; }, zremrangebyscore() { return chain; }, expire() { return chain; }, async exec() { return []; } };
            return chain;
        };
        await consumer.recordActiveWorkflow({ extra: { event_type: 'workflow_job', repo: 'detain/phlix-hub', _commit_sha: 'f2c9cc9' } });
        assert.equal(await redisMock.get(consumer.shaRepoKey('f2c9cc9')), 'detain/phlix-hub');
        delete redisMock.pipeline;
    });

    it('defers an unlinked floating tag until PKG_DEFER_MS has passed', () => {
        const { shouldDeferPackage, PKG_DEFER_MS } = consumer;
        const miss = { attached: false, reason: 'no_commit_tag' };
        const latest = pkgEnv('latest');
        assert.equal(shouldDeferPackage(latest, miss), true);
        latest.extra._pkg_defer_since = Date.now() - PKG_DEFER_MS - 1;
        assert.equal(shouldDeferPackage(latest, miss), false, 'gives up after the window');
        // Never defer: SHA tags, linked items, a missing commit trackable, no digest.
        assert.equal(shouldDeferPackage(pkgEnv('f2c9cc915d8b'), miss), false);
        assert.equal(shouldDeferPackage(pkgEnv('latest'), { attached: true }), false);
        assert.equal(shouldDeferPackage(pkgEnv('latest'), { attached: false, reason: 'no_live_commit_trackable' }), false);
        assert.equal(shouldDeferPackage(pkgEnv('latest', { digest: 'not-a-digest' }), miss), false);
        assert.equal(shouldDeferPackage({ extra: { event_type: 'push', data: {} } }, miss), false);
    });

    it('re-queues the original envelope with a defer stamp and drops the processing copy', async () => {
        const { deferPackageItem } = consumer;
        const ops = [];
        redisMock.multi = () => {
            const chain = {
                lpush(key, value) { ops.push(['lpush', key, value]); return chain; },
                lrem(key, count, value) { ops.push(['lrem', key, count, value]); return chain; },
                async exec() { return []; }
            };
            return chain;
        };
        const original = pkgEnv('latest');
        const raw = JSON.stringify(original);
        // Routing mutations on it.env must NOT leak into the re-queued copy.
        const mutated = JSON.parse(raw);
        mutated.room = 'mutated-room';
        assert.equal(await deferPackageItem({ raw, env: mutated }, 1000), true);
        assert.equal(ops[0][0], 'lpush');
        assert.equal(ops[0][1], 'notif:queue');
        const requeued = JSON.parse(ops[0][2]);
        assert.equal(requeued.room, ROOM);
        assert.equal(requeued.extra._pkg_defer_since, 1000);
        assert.deepEqual(ops[1], ['lrem', 'notif:processing', 1, raw]);

        // A second deferral keeps the first stamp so the wait stays capped.
        ops.length = 0;
        await deferPackageItem({ raw: ops[0]?.[2] || JSON.stringify(requeued), env: requeued }, 9000);
        assert.equal(JSON.parse(ops[0][2]).extra._pkg_defer_since, 1000);
        delete redisMock.multi;
    });

    it('gives each image reference + action its own bullet identity', () => {
        assert.equal(jobIdentity(pkgEnv('f2c9cc915d8b')), 'package:phlix-hub:f2c9cc915d8b:published');
        assert.equal(jobIdentity(pkgEnv('latest')), 'package:phlix-hub:latest:published');
        assert.equal(jobIdentity({ extra: { event_type: 'registry_package', data: { __stripped: true } } }), null);
    });

    it('renders as nested bullets under the commit header, re-delivery updates in place', () => {
        const commit = { header: '📦 Joe pushed 1 commit to detain/phlix-hub master', items: [], header_identity: null };
        const shaTagged = pkgEnv('f2c9cc915d8b');
        const step1 = mergeGithubTrackable(commit, shaTagged);
        const step2 = mergeGithubTrackable(step1, pkgEnv('latest'));
        const step3 = mergeGithubTrackable(step2, shaTagged);
        assert.equal(step3.header, commit.header);
        assert.equal(step3.items.length, 2);
        assert.ok(step3.items[0].text.includes('**phlix-hub:f2c9cc915d8b**'), step3.items[0].text);
        assert.ok(step3.items[1].text.includes('**phlix-hub:latest**'), step3.items[1].text);
        assert.ok(step3.text.includes('/users/detain/packages/container/phlix-hub/1499627176'), step3.text);
    });
});

describe('condensedBulletText — package branch', () => {
    const { condensedBulletText } = consumer;

    it('condenses a package event to emoji + linked name@version + verb', () => {
        const bullet = condensedBulletText({
            extra: {
                event_type: 'registry_package',
                repo: 'detain/sugarcraft',
                data: {
                    action: 'published',
                    registry_package: {
                        name: '@detain/sugarcraft',
                        package_type: 'npm',
                        package_version: { name: '2.1.0' }
                    }
                }
            }
        });
        assert.equal(bullet, '📦 [**@detain/sugarcraft@2.1.0**](https://github.com/detain/sugarcraft/packages) — published');
    });

    it('renders bare name when no version, under the legacy event_type', () => {
        const bullet = condensedBulletText({
            extra: { event_type: 'package', data: { action: 'updated', registry_package: { name: 'sugarcraft' } } }
        });
        assert.equal(bullet, '📦 **sugarcraft** — updated');
    });

    it('returns null for stripped, nameless or package-less data', () => {
        assert.equal(condensedBulletText({ extra: { event_type: 'registry_package', data: { __stripped: true } } }), null);
        assert.equal(condensedBulletText({ extra: { event_type: 'registry_package', data: { action: 'published', registry_package: {} } } }), null);
        assert.equal(condensedBulletText({ extra: { event_type: 'registry_package', data: { action: 'published' } } }), null);
    });
});
