const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const { mergeGithubTrackable } = require('../server/queue/notificationConsumer');

const PUSH_MSG = '📦 Joe Huss pushed 1 commit to interserver/teams-chat-bot main (compare)\n• 4fd48e4 updates to the message grouping logic (~1 files)';
const CHECK_RUN_MSG = '⏳ Check **Excavate** in_progress for detain/scoop-emulators on `master` (details)';
const WORKFLOW_JOB_MSG = '🔄 Workflow **Excavator** queued for detain/scoop-emulators on `master` (view run)';
const STATUS_MSG = 'ℹ️ appveyor[bot] triggered a **status** event on detain/scoop-emulators.';

function emptyRecent() {
    return { header: '', items: [], header_identity: null };
}

function pushEnv(message = PUSH_MSG) {
    return { type: 'msg', message, extra: { event_type: 'push', _commit_sha: '4fd48e4', dedup_key: 'github:commit:4fd48e4' } };
}

function checkRunEnv(name = 'Excavate', status = 'in_progress', conclusion = '', message = CHECK_RUN_MSG, htmlUrl = 'https://example/cr') {
    return {
        type: 'msg',
        message,
        extra: {
            event_type: 'check_run',
            _commit_sha: '4fd48e4',
            dedup_key: 'github:commit:4fd48e4',
            data: { check_run: { name, status, conclusion, html_url: htmlUrl } }
        }
    };
}

function workflowJobEnv(jobName = 'deploy', workflowName = 'Excavator', status = 'queued', conclusion = '', message = WORKFLOW_JOB_MSG, htmlUrl = 'https://example/job') {
    return {
        type: 'msg',
        message,
        extra: {
            event_type: 'workflow_job',
            _commit_sha: '4fd48e4',
            dedup_key: 'github:commit:4fd48e4',
            data: { workflow_job: { name: jobName, workflow_name: workflowName, status, conclusion, html_url: htmlUrl } }
        }
    };
}

function statusEnv(context = 'continuous-integration/appveyor/branch', message = STATUS_MSG) {
    return {
        type: 'msg',
        message,
        extra: {
            event_type: 'status',
            _commit_sha: '4fd48e4',
            dedup_key: 'github:commit:4fd48e4',
            data: { context, sha: '4fd48e4cdea42a785e2ae20d8cd669fc3b512728' }
        }
    };
}

function prReviewCommentEnv(commentId, message) {
    return {
        type: 'msg',
        message,
        extra: {
            event_type: 'pull_request_review_comment',
            _commit_sha: '4fd48e4',
            dedup_key: 'github:commit:4fd48e4',
            data: { comment: { id: commentId, commit_id: '4fd48e4cdea42a785e2ae20d8cd669fc3b512728' } }
        }
    };
}

describe('mergeGithubTrackable — first-event becomes header', () => {
    it('seeds the header from the first event when recent is empty', () => {
        const merged = mergeGithubTrackable(emptyRecent(), pushEnv());
        // For a push, the first line becomes the header and each `•` commit
        // line becomes a top-level item so it renders at the same indent as
        // later check_run / workflow_job bullets.
        assert.equal(merged.header, '📦 Joe Huss pushed 1 commit to interserver/teams-chat-bot main (compare)');
        assert.equal(merged.items.length, 1);
        assert.equal(merged.items[0].text, '4fd48e4 updates to the message grouping logic (~1 files)');
        assert.ok(merged.text.includes(' - 4fd48e4 updates to the message grouping logic'));
    });

    it('uses verbose env.message for the header (not the condensed form)', () => {
        // First event is a check_run. The header must carry the verbose
        // message because the trackable has no other context yet.
        const merged = mergeGithubTrackable(emptyRecent(), checkRunEnv('Excavate', 'in_progress'));
        assert.equal(merged.header, CHECK_RUN_MSG);
    });

    it('appends subsequent events as condensed nested bullets under the header', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, checkRunEnv('build', 'completed', 'success', '✅ Check build success for repo on master (details)', 'https://example/cr-build'));
        const r3 = mergeGithubTrackable(r2, workflowJobEnv('deploy', 'pages build and deployment', 'completed', 'success', '✅ Workflow pages build and deployment success for repo on master (view run)', 'https://example/wf-run/1'));

        // Header is the push's first line; commits + check_run + workflow_job are all top-level bullets.
        assert.equal(r3.header, '📦 Joe Huss pushed 1 commit to interserver/teams-chat-bot main (compare)');

        const bulletLines = r3.text.split('\n').filter(l => l.startsWith(' - '));
        assert.equal(bulletLines.length, 3, 'commit + check_run + workflow_job = 3 top-level bullets');
        assert.ok(bulletLines.some(l => l.includes('4fd48e4 updates')), 'push commit promoted to a top-level bullet');
        assert.ok(bulletLines.some(l => l.includes('**build** Check success')), 'check_run bullet uses condensed format');
        assert.ok(bulletLines.some(l => l.includes('**pages build and deployment** Workflow success')), 'workflow_job bullet uses workflow_name');
    });

    it('collapses multiple workflow_job events for the same workflow_name into one bullet', () => {
        // GitHub fires one workflow_job per job in the workflow. They all
        // share workflow_name but have distinct `name`s. The displayed
        // message uses workflow_name, so all three would render identically
        // — collapse them.
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, workflowJobEnv('build', 'pages build and deployment', 'completed', 'success'));
        const r3 = mergeGithubTrackable(r2, workflowJobEnv('deploy', 'pages build and deployment', 'completed', 'success'));
        const r4 = mergeGithubTrackable(r3, workflowJobEnv('report', 'pages build and deployment', 'completed', 'success'));

        const bullets = r4.items.filter(i => (i.identity || '').startsWith('workflow_job:'));
        assert.equal(bullets.length, 1, 'three jobs in the same workflow must produce one bullet');
        assert.ok(r4.text.includes('**pages build and deployment** Workflow success'));
    });

    it('updates the same check_run bullet across queued → in_progress → success', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, checkRunEnv('build', 'queued'));
        const r3 = mergeGithubTrackable(r2, checkRunEnv('build', 'in_progress'));
        const r4 = mergeGithubTrackable(r3, checkRunEnv('build', 'completed', 'success'));

        const checkBullets = r4.items.filter(i => (i.identity || '').startsWith('check_run:'));
        assert.equal(checkBullets.length, 1, 'same check_run name must stay one bullet');
        assert.ok(r4.text.includes('**build** Check success'));
        assert.ok(!r4.text.includes('Check queued'));
        assert.ok(!r4.text.includes('Check in_progress'));
    });

    it('keeps distinct check_run names as separate condensed bullets', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, checkRunEnv('lint', 'completed', 'success'));
        const r3 = mergeGithubTrackable(r2, checkRunEnv('build', 'completed', 'failure'));

        const checkBullets = r3.items.filter(i => (i.identity || '').startsWith('check_run:'));
        assert.equal(checkBullets.length, 2);
        assert.ok(r3.text.includes('**lint** Check success'));
        assert.ok(r3.text.includes('**build** Check failure'));
    });

    it('updates the header in place when same-identity event arrives for header', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), checkRunEnv('build', 'in_progress', '', '⏳ Check build in_progress'));
        const r2 = mergeGithubTrackable(r1, checkRunEnv('build', 'completed', 'success', '✅ Check build success'));

        assert.equal(r2.header, '✅ Check build success');
        assert.deepEqual(r2.items, []);
    });

    it('does not duplicate a header re-delivered as the same event', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, pushEnv());
        // Same push delivered twice → header + items unchanged, no duplicate
        // commit bullets appended.
        assert.equal(r2.items.length, r1.items.length);
        assert.equal(r2.text, r1.text);
    });

    it('handles re-delivery when the stored header is the legacy unsplit env.message', () => {
        // Before initialTrackableState, handleSingleNew saved the full multi-
        // line push message as recent.header with items = []. The next
        // arrival of the same push would compare splitHeaderAndBullets(incoming)
        // (just the first line) against header.trim() (the full text),
        // miss the match, and append the entire push again as a bullet —
        // producing a duplicated push with a nested commit sub-bullet.
        const legacyRecent = { header: PUSH_MSG, items: [], header_identity: null, text: PUSH_MSG };
        const r = mergeGithubTrackable(legacyRecent, pushEnv());
        assert.equal(r.items.length, 0, 'legacy unsplit re-delivery must not append');
        assert.ok(!r.text.includes(' - 📦'), 'no nested push line should appear');
        assert.ok(!r.text.includes('  - 4fd48e4'), 'no doubly-indented commit sub-bullet should appear');
    });

    it('accepts a legacy string for `recent` and treats it as the existing header', () => {
        const merged = mergeGithubTrackable(CHECK_RUN_MSG, statusEnv());
        assert.equal(merged.header, CHECK_RUN_MSG);
        assert.equal(merged.items.length, 1);
        assert.ok(merged.text.startsWith(CHECK_RUN_MSG));
        assert.ok(merged.text.includes(' - ℹ️ appveyor[bot] triggered'));
    });

    it('keeps distinct pull_request_review_comment events as separate bullets', () => {
        // Two different PR review comments by the same bot produce the same
        // env.message text. Without per-comment identity they would collapse
        // (or hit the "duplicate text, skip" guard). The per-comment id
        // identity keeps them as separate bullets.
        const msg = 'ℹ️ bot triggered a **pull_request_review_comment** event (created)';
        const r1 = mergeGithubTrackable(emptyRecent(), pushEnv());
        const r2 = mergeGithubTrackable(r1, prReviewCommentEnv(101, msg));
        const r3 = mergeGithubTrackable(r2, prReviewCommentEnv(102, msg));

        const commentItems = r3.items.filter(i => (i.identity || '').startsWith('pr_review_comment:'));
        assert.equal(commentItems.length, 2, 'distinct comment ids → distinct bullets');
    });

    it('groups appveyor status, push, and check_run for the same SHA into one message', () => {
        const r1 = mergeGithubTrackable(emptyRecent(), checkRunEnv('Excavate'));
        const r2 = mergeGithubTrackable(r1, pushEnv());
        const r3 = mergeGithubTrackable(r2, statusEnv());

        const topBullets = r3.text.split('\n').filter(l => l.startsWith(' - '));
        assert.equal(topBullets.length, 2, 'push and status are bullets; check_run is the header');
        assert.ok(r3.text.startsWith(CHECK_RUN_MSG));
        assert.ok(topBullets.some(l => l.includes('pushed')));
        assert.ok(topBullets.some(l => l.includes('status')));
    });
});
