const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const { mergeGithubTrackable } = require('../server/queue/notificationConsumer');

const PUSH_MSG = '📦 Joe Huss pushed 1 commit to interserver/teams-chat-bot main (compare)\n• 4fd48e4 updates to the message grouping logic (~1 files)';

function pushEnv(message = PUSH_MSG) {
    return { type: 'msg', message, extra: { event_type: 'push', _commit_sha: '4fd48e4', dedup_key: 'github:commit:4fd48e4' } };
}

function jobEnv(name, status, conclusion = '', htmlUrl = 'https://example/job') {
    return {
        type: 'msg',
        message: '',
        extra: {
            event_type: 'workflow_job',
            _commit_sha: '4fd48e4',
            dedup_key: 'github:commit:4fd48e4',
            data: { workflow_job: { name, status, conclusion, html_url: htmlUrl } }
        }
    };
}

describe('mergeGithubTrackable — duplicate push regression', () => {
    it('does NOT duplicate the push message when the same push arrives twice', () => {
        // Sequence: handleSingleNew saves recent.text = push message. A second
        // envelope for the same commit SHA arrives. The merger should NOT
        // produce `push_msg + — jobs — + push_msg` (the bug we are fixing).
        const merged = mergeGithubTrackable(PUSH_MSG, pushEnv());
        assert.equal(merged, PUSH_MSG, 'duplicate push must produce the push message exactly once, with no — jobs — section');
        assert.ok(!merged.includes('— jobs —'), 'no — jobs — marker should appear when there are no actual jobs');
    });

    it('treats a 📦 push prefix as a commit header (not a job emoji)', () => {
        // The 📦 (U+1F4E6) emoji shares a UTF-16 high surrogate (D83D) with
        // 🔄 (U+1F504). Without the `u` flag on the job-emoji regex, 📦 was
        // misclassified as a job status, which is how the duplication bug
        // arose. This test fails against the old non-`u` regex.
        const jobIncoming = jobEnv('build', 'queued');
        const merged = mergeGithubTrackable(PUSH_MSG, jobIncoming);
        assert.ok(merged.startsWith(PUSH_MSG), 'push message must remain the header, not be moved into the jobs section');
        assert.ok(merged.includes('\n\n— jobs —\n'), 'a real job event should add the — jobs — section');
        assert.ok(merged.includes('⏳ build queued'), 'job line should be appended after the marker');
    });

    it('appends a job line after the push header on first job event', () => {
        const merged = mergeGithubTrackable(PUSH_MSG, jobEnv('lint', 'in_progress'));
        assert.equal(merged, `${ PUSH_MSG }\n\n— jobs —\n🔄 lint in_progress (https://example/job)`);
    });

    it('replaces a job line in place when the same job updates', () => {
        const first = mergeGithubTrackable(PUSH_MSG, jobEnv('lint', 'queued'));
        const second = mergeGithubTrackable(first, jobEnv('lint', '', 'success'));
        // The single job line should be replaced, not duplicated.
        const lines = second.split('\n').filter(l => l.includes('lint'));
        assert.equal(lines.length, 1, 'same-named job must replace its line rather than duplicate');
        assert.ok(second.includes('✅ lint success'), 'final state should reflect the successful conclusion');
    });

    it('appends a second distinct job below the first', () => {
        const first = mergeGithubTrackable(PUSH_MSG, jobEnv('lint', 'in_progress'));
        const second = mergeGithubTrackable(first, jobEnv('build', 'queued'));
        assert.ok(second.includes('🔄 lint in_progress'));
        assert.ok(second.includes('⏳ build queued'));
        // Only one — jobs — marker, headers preserved.
        assert.equal(second.match(/— jobs —/g).length, 1);
    });

    it('promotes a push to header when an earlier job-only placeholder exists', () => {
        const jobOnly = '⏳ build queued (https://example/job)';
        const merged = mergeGithubTrackable(jobOnly, pushEnv());
        assert.ok(merged.startsWith(PUSH_MSG));
        assert.ok(merged.includes('\n\n— jobs —\n'));
        assert.ok(merged.endsWith(jobOnly), 'previous job placeholder should be preserved in the jobs section');
    });
});
