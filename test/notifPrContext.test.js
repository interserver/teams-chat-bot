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

describe('condensedBulletText — package branch', () => {
    const { condensedBulletText } = consumer;

    it('condenses a package event to emoji + name@version + verb', () => {
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
        assert.equal(bullet, '📦 **@detain/sugarcraft@2.1.0** — published');
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
