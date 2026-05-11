// Decide whether a queued envelope should actually be sent to Teams.
//
// The PHP github webhook now queues every event it receives, leaving the
// "is this signal worth posting?" decision to us. Filtered envelopes are
// silently ack'd by the consumer — they still count as drained, but never
// hit Bot Framework or the fallback webhook.
//
// Returning a non-null string means "skip with this reason"; null means
// "send it".

const REPO_OPT_OUT = new Set([
    'interserver/mailbaby-api-samples',
    'detain/interserver-api-samples'
]);

const LOW_SIGNAL_GITHUB_EVENTS = new Set(['star', 'watch', 'fork', 'ping']);
const SUCCESSFUL_CHECK_CONCLUSIONS = new Set(['success', 'neutral', 'skipped', 'cancelled']);
const SUCCESSFUL_WORKFLOW_CONCLUSIONS = new Set(['success', 'skipped', 'cancelled']);

function shouldSkip(envelope) {
    const extra = envelope.extra || {};
    const repo = extra.repo || '';
    const eventType = extra.event_type || '';
    const data = extra.data || envelope.data || {};

    if (repo && REPO_OPT_OUT.has(repo)) {
        return `repo opt-out: ${ repo }`;
    }

    // GitHub-specific noise filters — moved here from web/github.php's
    // decideSend(). Keeping the logic on the bot side means we can tune
    // verbosity without redeploying the webhook receiver.
    if (eventType) {
        // check_run and workflow_job are NOT filtered here — they are
        // passed through to the notification consumer where they are
        // grouped with their parent commit message via commit-scope
        // dedup keys (see normalizeGithubDedup in notificationConsumer).
        if (eventType === 'check_run') {
            return null; // pass through
        }
        if (eventType === 'workflow_job') {
            return null; // pass through
        }
        if (eventType === 'check_suite') {
            return '__SILENT__'; // aggregate — individual check_run events carry the detail
        }
        if (eventType === 'workflow_run') {
            return '__SILENT__'; // aggregate — individual workflow_job events carry the detail
        }
        if (LOW_SIGNAL_GITHUB_EVENTS.has(eventType)) {
            return `low-signal github event ${ eventType }`;
        }
    }

    return null;
}

module.exports = { shouldSkip };
