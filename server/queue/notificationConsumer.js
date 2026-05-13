// Notification queue consumer.
//
// Polls Redis (`notif:queue`) on a configurable interval, drains a tick's
// worth of envelopes, decides per-room edit/coalesce/new, sends via Bot
// Framework's adapter.continueConversation(), and falls back to the
// envelope's Power Automate webhook URL if Bot Framework cannot deliver.
//
// Envelope shape is produced by:
//   - PHP: MyAdmin\Notifications\Queue (mystage)
//   - PHP: NotificationQueue (webhooks.interserver.net)
// See those classes for the canonical schema (v=1).

const axios = require('axios');

const { runWithRetry, classify } = require('../lib/retry');
const { resolve: resolveChannel } = require('./channels');
const { shouldSkip } = require('./filters');
const { getAdapter } = require('../lib/adapter');
const { createNotifRedis, createBotRedis, TEAMS_SERVICE_URL } = require('../lib/redis');
const trace = require('./notifTrace');

const POLL_INTERVAL_MS = parseInt(process.env.NOTIF_POLL_MS || '5000', 10);
const POLL_INTERVAL_FAST_MS = parseInt(process.env.NOTIF_POLL_FAST_MS || '1000', 10);
const MAX_PER_TICK = parseInt(process.env.NOTIF_MAX_PER_TICK || '50', 10);
const COALESCE_MAX_CHARS = parseInt(process.env.NOTIF_COALESCE_MAX_CHARS || '24000', 10);
const COALESCE_MAX_ITEMS = parseInt(process.env.NOTIF_COALESCE_MAX_ITEMS || '8', 10);
const EDIT_WINDOW_MS = parseInt(process.env.NOTIF_EDIT_WINDOW_MS || String(30 * 60 * 1000), 10);
const KEY_PREFIX = process.env.NOTIF_KEY_PREFIX || 'notif:';
const HEARTBEAT_MS = parseInt(process.env.NOTIF_HEARTBEAT_MS || '60000', 10);
const APPEND_BREAK = '\n\n— · —\n\n';
const APPEND_TRAIL_SUMMARY_AFTER = 3;
// GitHub commit grouping window — if no event for a given commit SHA arrives
// within this many ms of the last edit, a new message is started.
const COMMIT_GROUP_WINDOW_MS = parseInt(process.env.NOTIF_COMMIT_GROUP_WINDOW_MS || '180000', 10); // 3 min

function k(name) { return KEY_PREFIX + name; }
function recentKey(room) { return k('recent:') + room; }
function wfactiveKey(repo) { return k('wfactive:') + repo; }
function metric(name) { return k('metrics:') + name; }
// Branch → PR number index. Populated when a `pull_request` event arrives
// so subsequent `push` and `delete` events on the same branch can be
// rerouted into the PR's trackable instead of spawning their own messages.
function prBranchKey(repo, branch) { return k('prbranch:') + repo + ':' + branch; }

// ---------------------------------------------------------------------------
// Action-triggered push attribution
// ---------------------------------------------------------------------------
//
// When a workflow on commit X commits + pushes a follow-up (e.g. the VHS
// workflow that re-renders demo GIFs and pushes "vhs: regenerate demo GIFs"
// back to master), GitHub fires that as a brand-new `push` event with a brand-
// new SHA — the webhook does not say "this push was caused by workflow run R
// on commit X". Without help, the consumer treats it as an unrelated commit
// and spawns a new trackable.
//
// We bridge that gap with two pieces:
//   1. A per-repo "active workflows" sorted set (`notif:wfactive:{repo}`,
//      member = parent commit SHA, score = ts), populated whenever a
//      workflow_run / workflow_job / check_run event arrives. An entry's
//      presence means "this commit currently has CI in flight".
//   2. On every push, decide if it was action-triggered (bot pusher, OR repo
//      lives in the downstream map). If so, look up the most recent active
//      parent SHA in this repo (or in the upstream repo via the map) and
//      override `dedup_key` so the merger nests the push under the existing
//      trackable instead of starting a fresh one.
//
// The upstream→downstream map is hardcoded to `detain/sugarcraft → sugarcraft/*`
// and overridable via NOTIF_DOWNSTREAM_REPOS=upstream:glob,upstream:glob,...

function parseDownstreamMap(spec) {
    const out = [];
    const raw = (spec || '').trim();
    if (!raw) return out;
    for (const pair of raw.split(',')) {
        const idx = pair.indexOf(':');
        if (idx <= 0) continue;
        const upstream = pair.slice(0, idx).trim();
        const glob = pair.slice(idx + 1).trim();
        if (!upstream || !glob) continue;
        // Convert simple glob to anchored regex: only `*` is meaningful
        const escaped = glob.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*');
        out.push({ upstream, pattern: new RegExp('^' + escaped + '$') });
    }
    return out;
}

const DOWNSTREAM_REPOS = parseDownstreamMap(
    process.env.NOTIF_DOWNSTREAM_REPOS || 'detain/sugarcraft:sugarcraft/*'
);

const WORKFLOW_EVENT_TYPES = new Set(['workflow_run', 'workflow_job', 'check_run', 'check_suite']);
const BOT_PUSHER_RE = /^github-actions(\[bot\])?$/i;

/**
 * Record the parent commit SHA into the per-repo active-workflow index for
 * any workflow-flavoured event. Idempotent — repeated events for the same
 * SHA just bump the score so the entry's TTL is effectively refreshed.
 */
async function recordActiveWorkflow(env) {
    const extra = env.extra || {};
    const eventType = extra.event_type || '';
    if (!WORKFLOW_EVENT_TYPES.has(eventType)) return;
    const repo = extra.repo || '';
    const sha = extra._commit_sha || '';
    if (!repo || !sha) return;
    const key = wfactiveKey(repo);
    const ts = Date.now();
    const cutoff = ts - EDIT_WINDOW_MS;
    try {
        const pipe = redis.pipeline();
        pipe.zadd(key, ts, sha);
        pipe.zremrangebyscore(key, 0, cutoff);
        pipe.expire(key, Math.ceil(EDIT_WINDOW_MS / 1000));
        await pipe.exec();
    } catch (err) {
        console.warn(`[notif] recordActiveWorkflow failed for ${ repo }/${ sha }: ${ err.message }`);
    }
}

/** Resolve any upstream repos for a child via the configured map. */
function upstreamCandidates(repo) {
    const out = [];
    for (const { upstream, pattern } of DOWNSTREAM_REPOS) {
        if (pattern.test(repo)) out.push(upstream);
    }
    return out;
}

/**
 * Decide whether a push envelope looks action-triggered. Returns true when
 * the pusher/sender is a bot, the head commit's author email is one of the
 * known GitHub Actions identities, OR the repo lives under an upstream's
 * downstream-glob (the only writes there in practice come from CI).
 */
function isActionTriggeredPush(env) {
    const extra = env.extra || {};
    if (extra.event_type !== 'push') return false;
    const data = extra.data || {};

    const pusherName = (data.pusher && data.pusher.name) || '';
    if (BOT_PUSHER_RE.test(pusherName)) return true;
    const pusherEmail = (data.pusher && data.pusher.email) || '';
    if (/github-actions(\[bot\])?@/i.test(pusherEmail)) return true;

    const senderLogin = (data.sender && data.sender.login) || '';
    if (/\[bot\]$/i.test(senderLogin)) return true;

    const commits = Array.isArray(data.commits) ? data.commits : [];
    const head = commits[commits.length - 1];
    const headAuthorEmail = (head && head.author && head.author.email) || '';
    if (/github-actions(\[bot\])?@/i.test(headAuthorEmail)) return true;

    const repo = extra.repo || '';
    if (repo && upstreamCandidates(repo).length > 0) return true;

    return false;
}

/**
 * Find the most recent active-workflow parent SHA that this push should be
 * attributed to. Searches the push's own repo first (covers in-repo bot
 * commits like the vhs.yml `commit` job), then any configured upstream repos
 * (covers downstream syncs like `detain/sugarcraft → sugarcraft/<lib>`).
 * Returns null when nothing within the window matches.
 */
async function findParentSha(env) {
    const extra = env.extra || {};
    const repo = extra.repo || '';
    if (!repo) return null;
    const ownSha = extra._commit_sha || '';
    const cutoff = Date.now() - EDIT_WINDOW_MS;

    const candidates = [repo, ...upstreamCandidates(repo)];
    let best = null;
    for (const candidate of candidates) {
        try {
            const recent = await redis.zrevrangebyscore(
                wfactiveKey(candidate), '+inf', cutoff,
                'WITHSCORES', 'LIMIT', 0, 5
            );
            for (let i = 0; i < recent.length; i += 2) {
                const sha = recent[i];
                const ts = parseFloat(recent[i + 1]);
                if (!sha || sha === ownSha) continue;
                if (!best || ts > best.ts) best = { sha, ts, repo: candidate };
            }
        } catch (err) {
            console.warn(`[notif] findParentSha lookup failed for ${ candidate }: ${ err.message }`);
        }
    }
    return best ? best.sha : null;
}

// ---------------------------------------------------------------------------
// Branch → PR routing
// ---------------------------------------------------------------------------
//
// A PR's lifecycle generates events that GitHub never explicitly links back
// to the PR: pushes to the PR branch arrive as plain `push` events, the
// branch deletion after merge as a `delete` event, and PR-conversation
// comments as `issue_comment` events. Without help, each of these spawns
// its own trackable. We bridge the gap with:
//
//   1. A `notif:prbranch:{repo}:{branch}` key written whenever a
//      `pull_request` event arrives, mapping the head branch back to the
//      PR number for the same EDIT_WINDOW_MS (30 min) life as the PR's
//      own trackable.
//   2. `attachPrContext(env)` inspects each envelope after dedup
//      normalisation and rewrites `dedup_key` / `_commit_sha` to the PR's
//      scoped values when:
//        - `issue_comment` event whose issue is actually a PR (per
//          `issue.html_url` containing `/pull/{n}` or `issue.pull_request`
//          set), even when the comment producer didn't tag it.
//        - `push` or `delete` event for a branch in the prbranch index.
//      It also rewrites the verbose message for `delete` so the user sees
//      *what* was deleted, not just "triggered a delete event".

async function recordPrBranch(repo, branch, prNumber) {
    if (!redis || !repo || !branch || !prNumber) return;
    try {
        await redis.set(
            prBranchKey(repo, branch),
            String(prNumber),
            'PX', EDIT_WINDOW_MS
        );
    } catch (err) {
        console.warn(`[notif] recordPrBranch failed for ${ repo }/${ branch }: ${ err.message }`);
    }
}

async function lookupPrByBranch(repo, branch) {
    if (!redis || !repo || !branch) return null;
    try {
        const val = await redis.get(prBranchKey(repo, branch));
        if (!val) return null;
        const n = parseInt(val, 10);
        return Number.isFinite(n) ? n : null;
    } catch (err) {
        console.warn(`[notif] lookupPrByBranch failed for ${ repo }/${ branch }: ${ err.message }`);
        return null;
    }
}

/** Pull the bare branch name out of a `refs/heads/...` ref string. */
function branchFromRef(ref) {
    return String(ref || '').replace(/^refs\/heads\//, '');
}

/**
 * Pull the PR number out of a GitHub issue object that's actually a PR.
 * GitHub normally sets `issue.pull_request` (an object with `url`,
 * `html_url`, etc.) to mark PRs, but the producer occasionally drops that
 * field; `issue.html_url` ending in `/pull/{n}` is the more reliable
 * indicator. Returns null when the issue isn't a PR.
 */
function prNumberFromIssue(issue) {
    if (!issue) return null;
    if (issue.pull_request) return issue.number || null;
    const m = String(issue.html_url || '').match(/\/pull\/(\d+)(?:[/?#]|$)/);
    return m ? parseInt(m[1], 10) : null;
}

/**
 * Reroute the envelope into a PR trackable when the event is implicitly
 * tied to a PR. Mutates `env.extra` (and `env.message` for `delete`
 * events). Returns a small descriptor that callers can use for tracing:
 * `{ rerouted: bool, reason, pr_number }`.
 */
async function attachPrContext(env) {
    const extra = env.extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';
    const repo = extra.repo || '';
    if (!repo) return { rerouted: false };

    // pull_request → seed the branch→PR index for future push/delete.
    if (eventType === 'pull_request') {
        const prNumber = data.pull_request?.number;
        const headBranch = data.pull_request?.head?.ref;
        if (prNumber && headBranch) {
            await recordPrBranch(repo, headBranch, prNumber);
        }
        return { rerouted: false };
    }

    // issue_comment on a PR → route to the PR's trackable.
    if (eventType === 'issue_comment') {
        const prNumber = prNumberFromIssue(data.issue);
        if (prNumber) {
            extra.dedup_key = `github:pr:${ repo }:${ prNumber }`;
            extra._commit_sha = `pr:${ repo }:${ prNumber }`;
            return { rerouted: true, reason: 'issue_is_pr', pr_number: prNumber };
        }
        // Regular issue (not a PR) — still group multiple comments on the
        // same issue together so they don't each spawn their own message.
        const issueNumber = data.issue?.number;
        if (issueNumber) {
            extra.dedup_key = `github:issue:${ repo }:${ issueNumber }`;
            extra._commit_sha = `issue:${ repo }:${ issueNumber }`;
            return { rerouted: true, reason: 'group_by_issue', pr_number: null };
        }
        return { rerouted: false };
    }

    // create / delete → rewrite message to say *what* was created or
    // deleted, and (for delete only — see note below) reroute to the
    // PR's trackable if the branch was the head of a recent PR.
    //
    // For `create`: the PR-branch index isn't populated yet at create
    // time (the create event fires when the branch first appears; the
    // pull_request opened event arrives later), so a branch→PR lookup
    // is normally a miss. We still try it for symmetry — if a PR
    // somehow opened before the create event landed, the routing kicks
    // in — but the main value here is the message rewrite.
    if (eventType === 'create' || eventType === 'delete') {
        const ref = data.ref || '';
        const refType = data.ref_type || '';
        const sender = data.sender?.login || 'someone';
        const stubRe = new RegExp(`triggered\\s+a\\s+\\*?\\*?${ eventType }\\*?\\*?\\s+event`, 'i');
        if (ref && refType && stubRe.test(env.message || '')) {
            const verb = eventType === 'create' ? 'created' : 'deleted';
            const emoji = eventType === 'create' ? '🌱' : '🗑️';
            env.message = `${ emoji } ${ sender } ${ verb } ${ refType } \`${ ref }\` in [${ repo }](https://github.com/${ repo }).`;
        }
        if (refType === 'branch' && ref) {
            const prNumber = await lookupPrByBranch(repo, ref);
            if (prNumber) {
                extra.dedup_key = `github:pr:${ repo }:${ prNumber }`;
                extra._commit_sha = `pr:${ repo }:${ prNumber }`;
                const reason = eventType === 'create' ? 'branch_became_pr_head' : 'branch_was_pr_head';
                return { rerouted: true, reason, pr_number: prNumber };
            }
        }
        return { rerouted: false };
    }

    // push to a PR branch → route to the PR trackable.
    if (eventType === 'push') {
        const branch = branchFromRef(data.ref);
        if (branch) {
            const prNumber = await lookupPrByBranch(repo, branch);
            if (prNumber) {
                extra.dedup_key = `github:pr:${ repo }:${ prNumber }`;
                extra._commit_sha = `pr:${ repo }:${ prNumber }`;
                return { rerouted: true, reason: 'branch_is_pr_head', pr_number: prNumber };
            }
        }
        return { rerouted: false };
    }

    return { rerouted: false };
}

// ---------------------------------------------------------------------------
// GitHub commit SHA extraction
// ---------------------------------------------------------------------------

/**
 * Extract the commit SHA from a GitHub envelope.
 * Returns null if the SHA cannot be determined.
 */
function extractCommitSha(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';

    let sha = null;

    if (eventType === 'check_run' || eventType === 'check_suite') {
        sha = data.check_run?.check_suite?.head_sha
            || data.check_run?.head_sha
            || data.check_suite?.head_sha
            || null;
    } else if (eventType === 'workflow_job' || eventType === 'workflow_run') {
        sha = data.workflow_job?.head_sha
            || data.workflow_run?.head_sha
            || null;
    } else if (eventType === 'push') {
        // For a push, `after` is the HEAD SHA after the push and is what the
        // build/check_run will reference. `commits[0]` is the OLDEST commit
        // in the push (commits are in chronological order), so reading it
        // would produce a SHA that no downstream event references.
        const commitsArr = Array.isArray(data.commits) ? data.commits : [];
        sha = data.after
            || data.head_commit?.id
            || commitsArr[commitsArr.length - 1]?.id
            || commitsArr[0]?.id
            || null;
    } else if (eventType === 'status') {
        sha = data.sha || null;
    } else if (eventType === 'commit_comment') {
        sha = data.comment?.commit_id || null;
    } else if (eventType === 'pull_request') {
        // PR events are grouped by PR number, not commit SHA
        sha = null;
    } else if (eventType === 'pull_request_review') {
        // PR reviews are anchored to the commit being reviewed (review.commit_id)
        // and fall back to the PR's current HEAD sha.
        sha = data.review?.commit_id || data.pull_request?.head?.sha || null;
    } else if (eventType === 'pull_request_review_comment') {
        sha = data.comment?.commit_id || data.pull_request?.head?.sha || null;
    } else if (eventType === 'release') {
        sha = data.release?.target_commitish || null;
    }

    // Normalize: take short SHA if available (first 7 chars)
    if (sha && sha.length > 7) sha = sha.slice(0, 7);
    return sha;
}

// ---------------------------------------------------------------------------
// GitHub job status line builder
// ---------------------------------------------------------------------------

/**
 * Build a single status line for a GitHub job event (check_run / workflow_job).
 * Format: {emoji} {job_name} {status_or_conclusion}
 * Returns null if this envelope doesn't represent a runnable job.
 */
function buildGithubJobLine(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';

    let name = '';
    let conclusion = '';
    let status = '';
    let htmlUrl = '';

    if (eventType === 'check_run') {
        const cr = data.check_run || {};
        name = cr.name || 'check_run';
        conclusion = cr.conclusion || '';
        status = cr.status || '';
        htmlUrl = cr.html_url || '';
    } else if (eventType === 'workflow_job') {
        const wj = data.workflow_job || {};
        name = wj.name || wj.workflow_name || 'workflow_job';
        conclusion = wj.conclusion || '';
        status = wj.status || '';
        htmlUrl = wj.html_url || '';
    } else {
        return null;
    }

    // Determine emoji and status text
    let emoji = '';
    let statusText = '';

    if (conclusion) {
        // Has conclusion - job completed
        switch (conclusion) {
        case 'success': emoji = '✅'; break;
        case 'failure': emoji = '❌'; break;
        case 'skipped': emoji = '⏭️'; break;
        case 'cancelled': emoji = '⚠️'; break;
        case 'neutral': emoji = 'ℹ️'; break;
        default: emoji = '❓';
        }
        statusText = conclusion;
    } else if (status) {
        // No conclusion yet - in progress
        switch (status) {
        case 'queued': emoji = '⏳'; statusText = 'queued'; break;
        case 'in_progress': emoji = '🔄'; statusText = 'in_progress'; break;
        case 'waiting': emoji = '⏸️'; statusText = 'waiting'; break;
        default: emoji = '❓'; statusText = status;
        }
    } else {
        return null; // Can't build line without status or conclusion
    }

    let line = `${ emoji } ${ name }`;
    if (statusText) line += ` ${ statusText }`;
    if (htmlUrl) line += ` (${ htmlUrl })`;

    return line;
}

/** Return a human-readable duration string from an ISO timestamp, or null. */
function elapsedMs(isoTimestamp) {
    if (!isoTimestamp) return null;
    const diff = Date.now() - new Date(isoTimestamp).getTime();
    if (isNaN(diff) || diff < 0) return null;
    const s = Math.floor(diff / 1000);
    if (s < 60) return `${ s }s`;
    const m = Math.floor(s / 60);
    const h = Math.floor(m / 60);
    if (h > 0) return `${ h }h ${ m % 60 }m`;
    return `${ m }m ${ s % 60 }s`;
}

// ---------------------------------------------------------------------------
// GitHub dedup normalisation
// ---------------------------------------------------------------------------

/**
 * For GitHub envelopes that carry a commit SHA but have no dedup_key,
 * inject one scoped to the commit so that all events for the same commit
 * are routed to the same trackable message (commit message + job list).
 *
 * This means a push creates the trackable message (as the parent commit
 * notification), and all subsequent check_run / workflow_job events for
 * that same SHA edit it in-place, appending/updating their job status lines.
 */
function normalizeGithubDedup(it) {
    const extra = it.env.extra = it.env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';

    // pull_request events are grouped by PR number, not commit SHA.
    // Handle this first since extractCommitSha returns null for PRs.
    // We also set _commit_sha to a PR-scoped value so the isGithubAppend
    // check in tryEdit works (it compares commit SHAs to decide whether
    // to merge into existing trackable or use standard append mode).
    if (eventType === 'pull_request') {
        const prNumber = data.pull_request?.number;
        const repo = extra.repo;
        if (prNumber && repo) {
            extra.dedup_key = `github:pr:${ repo }:${ prNumber }`;
            // Set _commit_sha to a stable PR-scoped value so isGithubAppend
            // in tryEdit is true for events in the same PR trackable.
            extra._commit_sha = `pr:${ repo }:${ prNumber }`;
        }
        return;
    }

    // For all other GitHub events: ensure we have _commit_sha set
    // even if dedup_key was already set by PHP. This enables secondary SHA-based lookup.
    const sha = extractCommitSha(it.env);
    if (!sha) return; // no SHA — can't group

    extra._commit_sha = sha;

    // pull_request_review and pull_request_review_comment events have a
    // pull_request.number — if present, route them to the PR trackable
    // instead of the commit trackable so they nest under the PR "opened"
    // header rather than scattering across commit-based trackables.
    if (eventType === 'pull_request_review' || eventType === 'pull_request_review_comment') {
        const prNumber = data.pull_request?.number;
        const repo = extra.repo;
        if (prNumber && repo) {
            extra.dedup_key = `github:pr:${ repo }:${ prNumber }`;
            // Set _commit_sha to PR-scoped value so isGithubAppend works
            extra._commit_sha = `pr:${ repo }:${ prNumber }`;
        } else {
            extra.dedup_key = `github:commit:${ sha }`;
        }
        return;
    }

    // For events that carry a commit SHA we force the dedup key to
    // be commit-scoped so that pushes, workflow events, status events
    // and PR-review events for the same commit all merge into one trackable.
    if (eventType === 'push' || eventType === 'workflow_job' ||
        eventType === 'check_run' || eventType === 'check_suite' ||
        eventType === 'workflow_run' || eventType === 'status' ||
        eventType === 'commit_comment') {
        extra.dedup_key = `github:commit:${ sha }`;
        return;
    }

    // For other events with existing dedup_key, preserve it
    if (extra.dedup_key) return;

    // For events without dedup_key, set commit-based key for grouping
    extra.dedup_key = `github:commit:${ sha }`;
}

// ---------------------------------------------------------------------------
// GitHub trackable message merger (pure)
// ---------------------------------------------------------------------------
//
// Format: the first event for a commit SHA becomes the header. Every later
// event for the same SHA becomes a nested bullet ` - ` under the header.
// Push events keep their commit list as ` - ` sub-bullets indented one level
// deeper. Job events (check_run / workflow_job) carry an identity derived
// from event_type+name so that queued → in_progress → success edits one
// bullet rather than appending three.

/**
 * Identity for events whose successive states should overwrite the same
 * bullet (or the header, if the header was set by that identity). Returns
 * null for events that should just append on each occurrence.
 */
function jobIdentity(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';
    if (eventType === 'check_run') {
        const name = data.check_run?.name;
        return name ? `check_run:${ name }` : null;
    }
    if (eventType === 'workflow_job') {
        // Multiple jobs in the same workflow share `workflow_name`. The
        // displayed message uses workflow_name (not the per-job name), so
        // keying identity by workflow_name collapses jobs of the same
        // workflow to one bullet rather than producing N identical lines.
        const wj = data.workflow_job || {};
        const name = wj.workflow_name || wj.name;
        return name ? `workflow_job:${ name }` : null;
    }
    if (eventType === 'workflow_run') {
        const wr = data.workflow_run || {};
        const name = wr.name;
        return name ? `workflow_run:${ name }` : null;
    }
    if (eventType === 'status') {
        // Status events from a CI bot identify themselves by `context`
        // (e.g. "continuous-integration/appveyor/branch"). Coalesce by it
        // so queued / pending / success states for the same CI bot share
        // one bullet rather than spawning new ones.
        const context = data.context;
        return context ? `status:${ context }` : 'status:unknown';
    }
    if (eventType === 'pull_request_review') {
        // Each review has its own id — distinct reviews should be distinct
        // bullets, but if the same review is later edited the same id
        // appears again and updates in place.
        const id = data.review?.id;
        return id ? `pr_review:${ id }` : null;
    }
    if (eventType === 'pull_request_review_comment') {
        const id = data.comment?.id;
        return id ? `pr_review_comment:${ id }` : null;
    }
    if (eventType === 'commit_comment') {
        const id = data.comment?.id;
        return id ? `commit_comment:${ id }` : null;
    }
    if (eventType === 'pull_request') {
        // PRs use prIdentity for grouping - each PR action (opened, closed,
        // synchronize, etc.) is a distinct event that should be a sub-bullet
        // under the PR header, not a header update. jobIdentity returning
        // prIdentity enables re-delivery detection via the headerIdentity
        // check at line 945, while the isPullRequest check ensures we
        // always append as sub-bullet rather than updating the header.
        return prIdentity(env);
    }
    return null;
}

/**
 * Per-event identity for `pull_request` events. Distinct per action +
 * head_sha so each PR action (opened, synchronize, closed, reopened…)
 * occupies its own slot in items[], while re-delivery of the same event
 * is deduped (same action + same head_sha → same identity → in-place
 * overwrite with identical text).
 *
 * The trackable as a whole is held together by `dedup_key`
 * (`github:pr:{repo}:{number}`) and `_commit_sha` (`pr:{repo}:{number}`)
 * — both PR-scoped — so all events for the same PR land in the same
 * trackable. This identity only determines which bullet within that
 * trackable an event maps to.
 */
function prIdentity(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';
    if (eventType !== 'pull_request') return null;
    const prNumber = data.pull_request?.number;
    const repo = extra.repo;
    if (!prNumber || !repo) return null;
    const action = data.action || 'unknown';
    const headSha = (data.pull_request?.head?.sha || '').slice(0, 7) || 'nosha';
    return `pr:${ repo }:${ prNumber }:${ action }:${ headSha }`;
}

/**
 * Pick the emoji + status-text for a check_run / workflow_job conclusion or
 * in-flight status. Mirrors the table in `buildGithubJobLine` so verbose
 * (env.message) and condensed (condensedBulletText) renderings stay
 * consistent.
 */
function pickStatusEmoji(conclusion, status) {
    if (conclusion) {
        switch (conclusion) {
        case 'success': return { emoji: '✅', statusText: 'success' };
        case 'failure': return { emoji: '❌', statusText: 'failure' };
        case 'skipped': return { emoji: '⏭️', statusText: 'skipped' };
        case 'cancelled': return { emoji: '⚠️', statusText: 'cancelled' };
        case 'neutral': return { emoji: 'ℹ️', statusText: 'neutral' };
        default: return { emoji: '❓', statusText: String(conclusion) };
        }
    }
    if (status) {
        switch (status) {
        case 'queued': return { emoji: '⏳', statusText: 'queued' };
        case 'in_progress': return { emoji: '🔄', statusText: 'in_progress' };
        case 'waiting': return { emoji: '⏸️', statusText: 'waiting' };
        case 'completed': return { emoji: '✅', statusText: 'completed' };
        default: return { emoji: '❓', statusText: String(status) };
        }
    }
    return { emoji: '', statusText: '' };
}

/**
 * Build a compact bullet line for a known GitHub event. The aim is to drop
 * the repository link and branch suffix from the PHP-generated env.message
 * (the parent push header already carries that context) so that bullets
 * stay scannable. Returns null when the event type has no condensed form,
 * letting the caller fall back to env.message.
 */
function condensedBulletText(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';

    if (eventType === 'check_run') {
        const cr = data.check_run || {};
        const name = cr.name;
        const { emoji, statusText } = pickStatusEmoji(cr.conclusion, cr.status);
        if (!name || !statusText) return null;
        const url = cr.html_url || cr.details_url || '';
        const link = url ? ` ([details](${ url } "${ url }"))` : '';
        return `${ emoji } **${ name }** Check ${ statusText }${ link }`;
    }
    if (eventType === 'workflow_job') {
        const wj = data.workflow_job || {};
        const name = wj.workflow_name || wj.name;
        const { emoji, statusText } = pickStatusEmoji(wj.conclusion, wj.status);
        if (!name || !statusText) return null;
        const url = wj.html_url || '';
        const link = url ? ` ([view run](${ url } "${ url }"))` : '';
        return `${ emoji } **${ name }** Workflow ${ statusText }${ link }`;
    }
    if (eventType === 'pull_request') {
        // PR event bodies (## Summary / ## Test plan / commit list) are the
        // PR *description*, which is identical across opened / synchronize /
        // closed events for the same PR. The header already shows it once;
        // sub-bullets only need the first line — the verb + actor + PR
        // title + branches + (✅ Merged.) state — so a closed-after-opened
        // bullet doesn't duplicate the full description body underneath
        // its own parent.
        const firstLine = String(env.message || '')
            .split('\n')
            .find(l => l.trim().length > 0);
        return firstLine ? firstLine.trim() : null;
    }
    return null;
}

// Teams' Markdown renderer only visually nests Markdown bullets one level
// deep — `    - foo` collapses back to the same depth as `- foo`. To keep
// deeper hierarchies legible, levels 2+ are emitted as plain text indented
// with NBSPs (U+00A0) and a literal `- ` prefix. NBSP is not Markdown
// whitespace, so the leading `- ` is rendered verbatim rather than parsed
// as a nested list marker, and Teams preserves the indent visually.
const NBSP = ' ';

function bulletPrefix(level) {
    if (level <= 1) return '- ';
    return NBSP.repeat((level - 1) * 2) + '- ';
}

/**
 * Render an event's `message` as a nested bullet. The first non-empty line
 * gets a level-1 Markdown bullet (`- <text>`). Embedded `•`/`*`/`-`
 * sub-bullets are promoted to level-2 NBSP-indented text bullets so Teams
 * still shows the nesting under the parent. Other continuation lines align
 * under the text that follows the dash.
 */
function indentAsBullet(messageText) {
    const lines = String(messageText || '')
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map(l => l.trim())
        .filter(l => l.length > 0);
    if (lines.length === 0) return bulletPrefix(1);
    const out = [];
    let firstHandled = false;
    for (const line of lines) {
        const m = line.match(/^[•*\-]\s+(.+)/);
        if (m) {
            out.push(bulletPrefix(2) + m[1]);
        } else if (!firstHandled) {
            out.push(bulletPrefix(1) + line);
            firstHandled = true;
        } else {
            out.push(NBSP.repeat(2) + line);
        }
    }
    return out.join('\n');
}

/**
 * Decompose a structured check_run name into a list of grouping segments
 * plus a leaf. Recognises the GitHub-Actions naming patterns we see in
 * practice and returns null if none of them fit:
 *
 *   "render (candy-vcr)"             → segments=["render"],            leaf="candy-vcr"
 *   "Windows · PHP 8.3 · candy-core" → segments=["Windows","PHP 8.3"], leaf="candy-core"
 *   "Test PHP 8.4 · candy-vt"        → segments=["Test","PHP 8.4"],    leaf="candy-vt"
 *   "build (linux)"                  → segments=["build"],             leaf="linux"
 *   "changed"                        → null (no decomposition possible)
 *
 * "PHP X.Y" embedded inside a `·`-segment is split into its own segment
 * so all PHP X.Y siblings naturally collapse under one bucket.
 */
function decomposeCheckSegments(name) {
    let segments = [];
    let leaf = String(name || '').trim();
    if (!leaf) return null;

    if (leaf.includes(' · ')) {
        const parts = leaf.split(' · ').map(s => s.trim()).filter(Boolean);
        segments = parts.slice(0, -1);
        leaf = parts[parts.length - 1] || '';
    }

    // Split any "X PHP 8.Y" segment into "X" + "PHP 8.Y" (e.g. "Test PHP 8.4").
    const phpRe = /^(.*?)\s*(PHP \d+\.\d+)\s*(.*)$/;
    const expanded = [];
    for (const seg of segments) {
        const m = seg.match(phpRe);
        if (m && m[2]) {
            if (m[1]) expanded.push(m[1].trim());
            expanded.push(m[2]);
            if (m[3]) expanded.push(m[3].trim());
        } else {
            expanded.push(seg);
        }
    }
    segments = expanded;

    // "render (candy-vcr)" — promote "render" to a segment.
    const parenMatch = leaf.match(/^(.+?)\s+\((.+)\)$/);
    if (parenMatch) {
        segments.push(parenMatch[1].trim());
        leaf = parenMatch[2].trim();
    }

    if (segments.length === 0) return null;
    if (!leaf) return null;
    return { segments, leaf };
}

/**
 * Parse the condensed text produced by `condensedBulletText` for a
 * check_run back into its components so renderer code can recompose it in
 * a different shape (e.g. share the status row across leaves).
 *
 *   "✅ **changed** Check success ([details](url \"url\"))"
 *      → { emoji: '✅', name: 'changed', statusText: 'success',
 *          urlLink: '([details](url "url"))' }
 *
 * Returns null when the text doesn't have the expected check_run shape
 * (callers fall back to the verbatim text).
 */
function parseCheckRunText(text) {
    const m = String(text || '').match(/^(\S+)\s+\*\*([^*]+)\*\*\s+Check\s+(\S+)\s*(.*)$/);
    if (!m) return null;
    return {
        emoji: m[1],
        name: m[2].trim(),
        statusText: m[3],
        urlLink: m[4].trim()
    };
}

/**
 * Render the grouped section of a trackable for a list of `check_run`
 * items that have been parsed and segmented. Recursive: each call peels
 * off one prefix segment.
 *
 * Rules:
 *  - A group with a single item is rendered flat (no extra nesting) so a
 *    lone "render (foo)" stays as one bullet rather than spawning a one-
 *    item "render" sub-tree.
 *  - When a group's items have no further segments (we're at the leaf
 *    level for that prefix), apply status compression:
 *      * If every leaf shares the same status, the prefix and status are
 *        combined into one header line and the leaves render with just
 *        `**leaf** ([details](url))`.
 *      * If statuses differ, the prefix header is bold-only and the
 *        leaves are sub-grouped by status — a status sub-group of 2+
 *        gets its own `emoji Check statusText` header with bare leaves
 *        below, while a singleton status keeps the full inline form.
 *  - When a group still has deeper segments, the prefix becomes a
 *    bold-only header and recursion handles the rest.
 */
function renderChecksGrouped(checkItems, level, lines, depth = 0) {
    if (checkItems.length === 0) return;

    const groups = new Map();
    const noPrefix = [];
    for (const ci of checkItems) {
        if (ci.segments.length === 0) {
            noPrefix.push(ci);
        } else {
            const firstSeg = ci.segments[0];
            if (!groups.has(firstSeg)) groups.set(firstSeg, []);
            groups.get(firstSeg).push(ci);
        }
    }

    for (const ci of noPrefix) {
        lines.push(bulletPrefix(level) + leafInline(ci));
    }

    for (const [seg, group] of groups) {
        if (group.length === 1) {
            const ci = group[0];
            if (depth === 0) {
                // Top level — preserve the original full-name bullet so
                // paren-style names like "render (candy-vcr)" stay readable
                // as-is.
                lines.push(bulletPrefix(level) + ci.originalText);
            } else {
                // Inside a recursive subtree — the parent header already
                // carries the consumed segments, so rebuild the bullet from
                // the remaining segments + leaf to avoid repeating them.
                const remainingName = [...ci.segments, ci.leaf].join(' · ');
                const link = ci.urlLink ? ' ' + ci.urlLink : '';
                lines.push(bulletPrefix(level) + `${ ci.emoji } **${ remainingName }** Check ${ ci.statusText }${ link }`);
            }
            continue;
        }

        const stripped = group.map(ci => ({ ...ci, segments: ci.segments.slice(1) }));
        const allLeaf = stripped.every(s => s.segments.length === 0);

        if (!allLeaf) {
            // Deeper nesting still possible — peel and recurse.
            lines.push(bulletPrefix(level) + `**${ seg }**`);
            renderChecksGrouped(stripped, level + 1, lines, depth + 1);
            continue;
        }

        const allSameStatus = stripped.every(s => s.statusKey === stripped[0].statusKey);
        if (allSameStatus) {
            const first = stripped[0];
            lines.push(bulletPrefix(level) + `${ first.emoji } **${ seg }** Check ${ first.statusText }`);
            emitLeafBucket(lines, level + 1, stripped);
            continue;
        }

        // Mixed statuses at the leaf level: prefix-only header, then group by
        // status (compress only when a status has 2+ items).
        lines.push(bulletPrefix(level) + `**${ seg }**`);
        const byStatus = new Map();
        for (const s of stripped) {
            if (!byStatus.has(s.statusKey)) byStatus.set(s.statusKey, []);
            byStatus.get(s.statusKey).push(s);
        }
        for (const statusGroup of byStatus.values()) {
            if (statusGroup.length >= 2) {
                const first = statusGroup[0];
                lines.push(bulletPrefix(level + 1) + `${ first.emoji } Check ${ first.statusText }`);
                emitLeafBucket(lines, level + 2, statusGroup);
            } else {
                lines.push(bulletPrefix(level + 1) + leafInline(statusGroup[0]));
            }
        }
    }
}

function leafInline(ci) {
    const link = ci.urlLink ? ' ' + ci.urlLink : '';
    return `${ ci.emoji } **${ ci.leaf }** Check ${ ci.statusText }${ link }`;
}

function leafOnly(ci) {
    const link = ci.urlLink ? ' ' + ci.urlLink : '';
    return `**${ ci.leaf }**${ link }`;
}

/**
 * Compact form of a leaf for inline comma-joined lists. Drops the per-link
 * title attribute and shortens "details" → "d" so a row of 9 leaves on one
 * line stays scannable rather than turning into a wall of repeated link
 * markup.
 */
function leafCompact(ci) {
    if (!ci.urlLink) return `**${ ci.leaf }**`;
    const compactLink = ci.urlLink
        .replace('[details]', '[d]')
        .replace(/\s"[^"]*"\)\)$/, '))');
    return `**${ ci.leaf }** ${ compactLink }`;
}

// When a status sub-bucket has more than this many same-status leaves, fold
// them onto a single comma-separated line via `leafCompact` instead of one
// bullet per leaf. Tuned to keep small groups skimmable while preventing
// long matrix runs (e.g. 30+ test packages all green) from dominating the
// trackable.
const LEAVES_COMPACT_THRESHOLD = 3;

/**
 * Emit a flat list of leaves either as one bullet per leaf (small bucket)
 * or as a single comma-joined compact bullet when over the threshold.
 */
function emitLeafBucket(lines, level, leaves) {
    if (leaves.length > LEAVES_COMPACT_THRESHOLD) {
        lines.push(bulletPrefix(level) + leaves.map(leafCompact).join(', '));
        return;
    }
    for (const leaf of leaves) {
        lines.push(bulletPrefix(level) + leafOnly(leaf));
    }
}

/**
 * Render a trackable to its final text. Non-check_run items render as
 * flat top-level bullets in insertion order; check_run items whose name
 * decomposes are collected and rendered as a single adaptively-nested
 * tree below the flat items (see `renderChecksGrouped`).
 */
function renderTrackable(header, items) {
    const headerText = String(header || '').trim();
    const flatBullets = [];
    const checkItems = [];

    for (const item of items) {
        if (item.identity && item.identity.startsWith('check_run:')) {
            const name = item.identity.slice('check_run:'.length);
            const decomp = decomposeCheckSegments(name);
            const parsed = parseCheckRunText(item.text);
            if (decomp && parsed) {
                checkItems.push({
                    segments: decomp.segments,
                    leaf: decomp.leaf,
                    emoji: parsed.emoji,
                    statusText: parsed.statusText,
                    statusKey: parsed.emoji + ':' + parsed.statusText,
                    urlLink: parsed.urlLink,
                    originalText: item.text
                });
                continue;
            }
        }
        const bullet = indentAsBullet(item.text);
        if (bullet) flatBullets.push(bullet);
    }

    const lines = [];
    if (headerText) lines.push(headerText);
    lines.push(...flatBullets);
    renderChecksGrouped(checkItems, 1, lines);

    if (lines.length === 0) return '';

    // Teams natively renders two visual levels of Markdown nested bullets
    // (level 1 plus one indent). When the tree we built never goes deeper
    // than that — no line starts with the level-3 NBSP prefix (4 NBSPs +
    // `-`) — swap the level-2 NBSPs on each bullet back to regular spaces
    // and join with plain `\n`, so Teams uses its native nested-list
    // rendering. As soon as we have depth 3+ the NBSPs stay; in that mode
    // the deeper lines aren't Markdown list items (the NBSP prefix is not
    // Markdown whitespace, so the leading `-` is literal text inside the
    // parent list-item continuation), and Teams would otherwise concat
    // them onto a single line. Joining with `  \n` (two trailing spaces
    // before each newline) emits the CommonMark hard line break so the
    // lines render separately.
    const L3_PREFIX = NBSP.repeat(4) + '-';
    const hasDepth3 = lines.some(l => l.startsWith(L3_PREFIX));
    if (hasDepth3) {
        // Deep mode: the renderer emits 2 NBSPs of indent per level, which
        // worked for two levels but made depth 3 vs depth 4 hard to tell
        // apart visually in Teams (2 vs 4 char widths of indent looks
        // similar at a glance). Double the leading NBSP run on every line
        // before serialising so each level steps 4 NBSPs further than its
        // parent. The trailing `  \n` is CommonMark's hard line break,
        // forcing each NBSP-text bullet onto its own row.
        return lines
            .map(l => l.replace(new RegExp('^' + NBSP + '+'), m => NBSP.repeat(m.length * 2)))
            .join('  \n');
    }
    return lines.join('\n').replace(new RegExp('^' + NBSP + '+', 'gm'), m => ' '.repeat(m.length));
}

/**
 * Split an incoming env.message into a header line plus any leading bullet
 * items it already contains. Bullet markers `•`, `*`, and `-` (with at
 * least one trailing space) are recognised. Used when seeding a brand-new
 * trackable so a push's commit-list lines become top-level bullets at the
 * same indent as later check_run/workflow_job bullets, instead of staying
 * stuck inside the verbose header as raw `•` lines.
 */
function splitHeaderAndBullets(text) {
    const lines = String(text || '').replace(/\r\n/g, '\n').split('\n');
    const headerLines = [];
    const bulletItems = [];
    let bulletsStarted = false;
    for (const raw of lines) {
        const trimmed = raw.trim();
        if (!trimmed) continue; // collapse blank lines
        const m = trimmed.match(/^[•*\-]\s+(.+)/);
        if (m) {
            bulletsStarted = true;
            bulletItems.push({ identity: null, text: m[1] });
        } else if (bulletsStarted) {
            // Trailing prose after a bullet list — keep it as another item
            // rather than tacking it back onto the header (which would
            // re-introduce a multi-paragraph header).
            bulletItems.push({ identity: null, text: trimmed });
        } else {
            headerLines.push(raw);
        }
    }
    return { headerText: headerLines.join('\n').trim(), bulletItems };
}

/**
 * Return everything after the first non-empty line of a PR event's
 * verbose message. The first line is the verb + actor + PR title + branch
 * pair (which varies per action: opened, closed, synchronize…); the rest
 * is the PR description body (## Summary, ## Test plan, commit list) and
 * is identical across every event of the same PR.
 */
function prBodyAfterFirstLine(text) {
    const lines = String(text || '').split('\n');
    let foundFirst = false;
    const body = [];
    for (const line of lines) {
        if (!foundFirst) {
            if (line.trim().length > 0) foundFirst = true;
            continue;
        }
        body.push(line);
    }
    return body.join('\n').trim();
}

/**
 * Decide whether a PR sub-bullet's body is already shown elsewhere in
 * the trackable (in the header or as items extracted from a prior PR
 * event's split). When true, the caller can safely condense the bullet
 * to its first line so a `closed` sub-bullet under an `opened` header
 * doesn't duplicate the description body.
 *
 * We compare on whitespace-collapsed, bullet-marker-stripped text so
 * cosmetic differences (extra blank lines, the `- ` prefix that
 * splitHeaderAndBullets strips, trailing spaces) don't defeat the match.
 * If the body has been edited between events, normalised strings won't
 * match and the caller falls back to the verbose form — preserving the
 * new content for the reader.
 */
function parentAlreadyShowsPrBody(env, header, items) {
    if ((env.extra || {}).event_type !== 'pull_request') return false;
    const body = prBodyAfterFirstLine(env.message || '');
    if (!body) return false;
    const stored = [header, ...items.map(it => (it && it.text) || '')].join('\n');
    const norm = s => String(s)
        .split('\n')
        .map(l => l.replace(/^[ \t]*[•*\-][ \t]+/, '').trim())
        .filter(l => l)
        .join(' ');
    return norm(stored).includes(norm(body));
}

/**
 * Merge an incoming envelope into an existing trackable message.
 *
 *   recent — either the full recent object (with `header`, `items`,
 *            `header_identity`) or a legacy plain text string. The string
 *            form is preserved for backward compat (it is treated as the
 *            header with no prior items).
 *   env    — the incoming envelope.
 *
 * Returns `{ text, header, items, header_identity }`. Callers persist the
 * full object so the structure survives across edits.
 */
function mergeGithubTrackable(recent, env) {
    const recentObj = typeof recent === 'string'
        ? { header: recent, items: [], header_identity: null }
        : (recent || { header: '', items: [], header_identity: null });

    // Header keeps the verbose PHP-generated message so the trackable
    // carries repo + branch context up top. Bullets prefer a condensed
    // rendering so identical workflow-level lines don't clutter the
    // message — fall back to env.message when there is no condensed form.
    const verboseMessage = String(env.message || '').trim();
    const condensed = condensedBulletText(env);
    const headerMessage = verboseMessage || condensed;
    const incomingIdentity = jobIdentity(env);

    let header = recentObj.header || recentObj.text || '';
    let items = Array.isArray(recentObj.items) ? recentObj.items.map(it => ({ ...it })) : [];
    let headerIdentity = recentObj.header_identity || null;

    // Choose the bullet text. For most event types we always prefer the
    // condensed form (check_run/workflow_job have noisy verbose text and
    // the condensed line is strictly better). For `pull_request` events,
    // the condensed form drops the PR description body — safe only when
    // the parent already shows it. If the trackable doesn't carry the
    // body yet (this is the first event, or the PR description was
    // edited between events), keep the verbose text so the reader sees
    // it instead of a content-less stub.
    const eventType = (env.extra || {}).event_type || '';
    let bulletMessage;
    if (eventType === 'pull_request' && condensed) {
        bulletMessage = parentAlreadyShowsPrBody(env, header, items)
            ? condensed
            : verboseMessage;
    } else {
        bulletMessage = condensed || verboseMessage;
    }

    // First event for this commit becomes the header. If the verbose
    // message already contains its own bullet list (push events emit one
    // line per commit), split those out as initial items so they render at
    // the same indent as later check_run / workflow_job bullets.
    if (!header) {
        if (headerMessage) {
            const split = splitHeaderAndBullets(headerMessage);
            header = split.headerText;
            items = split.bulletItems;
            headerIdentity = incomingIdentity;
        }
        return { text: renderTrackable(header, items), header, items, header_identity: headerIdentity };
    }

    // Nothing usable to add — emit current state.
    if (!bulletMessage) {
        return { text: renderTrackable(header, items), header, items, header_identity: headerIdentity };
    }

    // Same identity as the header → in-place update of the header (e.g.
    // check_run "Excavate" queued → in_progress → success edits the header
    // rather than appending a new bullet for each transition).
    //
    // For pull_request events this only fires on webhook re-delivery (same
    // action + same head SHA → same identity), which is the exact case
    // where overwriting the header with the identical text is correct. A
    // *different* PR action (synchronize after opened, etc.) gets its own
    // identity from prIdentity() and falls through to the bullet-append
    // branch below.
    if (incomingIdentity && headerIdentity && incomingIdentity === headerIdentity) {
        if (headerMessage) header = headerMessage;
        return { text: renderTrackable(header, items), header, items, header_identity: headerIdentity };
    }

    // Re-delivery of the same event. Compare on the first non-empty line of
    // each side so the check works whether the stored header was already
    // split (header = first line, items hold the commit list) or is still
    // the unsplit verbose message saved by handleSingleNew on the very
    // first send (header = full multi-line). Without this two-sided first-
    // line compare, a push arriving twice would render as
    //   📦 …push…
    //   • sha …
    //    - 📦 …push…
    //      - sha …
    if (!incomingIdentity) {
        const firstLine = (s) => String(s || '')
            .split('\n')
            .map(l => l.trim())
            .find(l => l.length > 0) || '';
        const incomingFirst = firstLine(bulletMessage);
        const headerFirst = firstLine(header);
        if (incomingFirst && incomingFirst === headerFirst) {
            return { text: renderTrackable(header, items), header, items, header_identity: headerIdentity };
        }
    }

    if (incomingIdentity) {
        const idx = items.findIndex(it => it.identity === incomingIdentity);
        if (idx >= 0) {
            items[idx] = { identity: incomingIdentity, text: bulletMessage };
        } else {
            items.push({ identity: incomingIdentity, text: bulletMessage });
        }
    } else {
        // No identity — append unless the same text is already present
        // (defensive guard against accidental re-delivery from the queue).
        const already = items.some(it => it.text === bulletMessage);
        if (!already) items.push({ identity: null, text: bulletMessage });
    }

    return { text: renderTrackable(header, items), header, items, header_identity: headerIdentity };
}

let stopping = false;
let tickInFlight = false;
let timer = null;
let firstTick = true;
let lastTickStats = null;
let lastHeartbeatAt = 0;
let totalTicks = 0;
// Notification queue + dedup state lives on the InterServer Redis (no auth).
let redis = null;
// `convref:{conversationId}` is stored by botActivityHandler.js on the
// existing bot Redis (Dragonfly, possibly behind auth) and we need it to
// resolve where to send proactive activities.
let redisBot = null;
// Bot Framework adapter — assigned in startConsumer() from getAdapter().
// Declared explicitly so the assignment doesn't create an implicit global.
let adapter = null;

function startConsumer() {
    if (timer) return;
    if (process.env.NOTIF_CONSUMER_ENABLED === '0') {
        console.log('[notif] consumer disabled via NOTIF_CONSUMER_ENABLED=0');
        return;
    }
    redis = createNotifRedis();
    redisBot = createBotRedis();
    adapter = getAdapter();
    const notifHost = process.env.REDIS_HOST_MY || '67.217.60.234';
    const notifPort = process.env.REDIS_PORT_MY || '6379';
    const botHost = process.env.REDIS_HOST || '67.217.60.234';
    const botPort = process.env.REDIS_PORT || '6379';
    console.log(`[notif] consumer starting; poll=${ POLL_INTERVAL_MS }ms fast=${ POLL_INTERVAL_FAST_MS }ms heartbeat=${ HEARTBEAT_MS }ms key-prefix=${ KEY_PREFIX }`);
    console.log(`[notif]   queue redis : ${ notifHost }:${ notifPort } (no auth)`);
    console.log(`[notif]   convref redis: ${ botHost }:${ botPort }`);
    scheduleNextTick(POLL_INTERVAL_MS);
}

function scheduleNextTick(delayMs) {
    if (stopping) return;
    timer = setTimeout(() => {
        runTick().catch(err => console.error('[notif] tick error:', err));
    }, delayMs);
    timer.unref();
}

async function stopConsumer() {
    stopping = true;
    if (timer) {
        clearTimeout(timer);
        timer = null;
    }
    // Wait briefly for in-flight tick to complete; abandoned items remain in
    // notif:processing and are recovered on next startup.
    const deadline = Date.now() + 5000;
    while (tickInFlight && Date.now() < deadline) {
        await new Promise(r => setTimeout(r, 100));
    }
    if (redis) {
        try { await redis.quit(); } catch (_) { /* noop */ }
        redis = null;
    }
    if (redisBot) {
        try { await redisBot.quit(); } catch (_) { /* noop */ }
        redisBot = null;
    }
    console.log('[notif] consumer stopped');
}

function getNotifRedis() { return redis; }

// First N chars of a message, newlines collapsed, ellipsis if truncated.
// Used purely for debug log readability.
function preview(text, n = 100) {
    if (text === null || text === undefined) return '';
    const s = String(text).replace(/\s+/g, ' ').trim();
    return s.length > n ? s.slice(0, n) + '…' : s;
}

function shortConv(conversationId) {
    if (!conversationId) return '';
    // conversation IDs look like 19:<32-hex>@thread.v2 — last 8 hex chars
    // are unique enough for at-a-glance correlation in logs.
    const m = conversationId.match(/19:([0-9a-f]+)@thread/i);
    return m ? m[1].slice(-8) : conversationId.slice(-8);
}

async function runTick() {
    if (stopping || tickInFlight || !redis) return;
    tickInFlight = true;
    const t0 = Date.now();
    totalTicks++;
    const tickNum = trace.newTick();
    const stats = { drained: 0, sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
    try {
        // Recover anything left over from a prior crashed tick BEFORE we drain new items.
        if (firstTick) {
            firstTick = false;
            await recoverProcessing();
        }

        const items = await drainBatch();
        if (items.length === 0) {
            // Silence is confusing. Emit a heartbeat at most every HEARTBEAT_MS
            // so an operator can confirm the consumer is alive even when the
            // queue is empty. Reports queue depth from the producer side too,
            // so a non-zero depth here means producer and consumer disagree.
            const now = Date.now();
            if (now - lastHeartbeatAt >= HEARTBEAT_MS) {
                lastHeartbeatAt = now;
                try {
                    const depth = await redis.llen(k('queue'));
                    console.log(`[notif] heartbeat alive ticks=${ totalTicks } queue_depth=${ depth }`);
                } catch (err) {
                    console.warn(`[notif] heartbeat — redis llen failed: ${ err.message }`);
                }
            }
            return;
        }
        stats.drained = items.length;
        console.log(`[notif] picked up ${ items.length } item(s) from ${ k('queue') }`);
        trace.emit('drained', {
            count: items.length,
            queue_key: k('queue'),
            order: items.map((it, idx) => ({
                idx,
                room: it.env.room,
                type: it.env.type,
                event_type: it.env.extra?.event_type,
                dedup_key: it.env.extra?.dedup_key,
                repo: it.env.extra?.repo,
                _commit_sha: it.env.extra?._commit_sha
            }))
        });
        for (const it of items) {
            const ev = (it.env.extra && it.env.extra.event_type) ? ` ev=${ it.env.extra.event_type }` : '';
            const dk = (it.env.extra && it.env.extra.dedup_key) ? ` dedup=${ it.env.extra.dedup_key }` : '';
            console.log(`[notif]   • room=${ it.env.room } type=${ it.env.type }${ ev }${ dk } msg="${ preview(it.env.message) }"`);
            trace.emit('item_drained', {
                room: it.env.room,
                envelope: trace.snapshotEnvelope(it.env)
            });
        }

        const valid = [];
        const nowSec = Math.floor(Date.now() / 1000);
        for (const it of items) {
            if (it.env.expires_at && it.env.expires_at < nowSec) {
                stats.expired++;
                trace.emit('expired', {
                    room: it.env.room,
                    expires_at: it.env.expires_at,
                    now_sec: nowSec,
                    envelope: trace.snapshotEnvelope(it.env)
                });
                await deadLetter(it, 'expired');
                continue;
            }
            // Filtering: events that match a skip rule used to be silently
            // dropped. Now we redirect them to the int-dev-announce channel
            // so they remain visible (and tunable) without flooding the
            // primary channels. Set NOTIF_FILTER_ENABLED=0 in .env to
            // bypass the filter entirely (original room is preserved).
            if (process.env.NOTIF_FILTER_ENABLED !== '0') {
                const skipReason = shouldSkip(it.env);
                if (skipReason) {
                    if (skipReason === '__SILENT__') {
                        // Truly silent skip - just ack and don't send anywhere
                        trace.emit('filter_silent', {
                            room: it.env.room,
                            event_type: it.env.extra?.event_type,
                            envelope: trace.snapshotEnvelope(it.env)
                        });
                        await ackOne(it);
                        stats.redirected = (stats.redirected || 0) + 1;
                        await bumpMetric('redirected');
                        continue;
                    }
                    trace.emit('filter_redirect', {
                        from_room: it.env.room,
                        to_room: 'int-dev-announce',
                        reason: skipReason,
                        event_type: it.env.extra?.event_type
                    });
                    redirectToAnnounce(it, skipReason);
                    stats.redirected = (stats.redirected || 0) + 1;
                    await bumpMetric('redirected');
                }
            }
            // Route specific repos to int-dev-announce. This is in addition
            // to (and takes precedence over) the filter-based redirect.
            // Supports wildcard patterns like "detain/*" to match all repos under
            // an org, or exact repo names like "owner/repo".
            const announceRepos = (process.env.NOTIF_ANNOUNCE_REPOS || '')
                .split(',')
                .map(s => s.trim())
                .filter(Boolean);
            const repo = it.env.extra?.repo || '';
            const matchesAnnounce = announceRepos.some(pattern => {
                if (pattern.endsWith('/*')) {
                    // Wildcard pattern: matches "owner/repo" if pattern is "owner/*"
                    const prefix = pattern.slice(0, -2);
                    return repo.startsWith(prefix + '/') || repo === prefix;
                }
                return repo === pattern;
            });
            if (matchesAnnounce) {
                const originalRoom = it.env.room || 'unknown';
                trace.emit('announce_redirect', {
                    from_room: originalRoom,
                    to_room: 'int-dev-announce',
                    repo: it.env.extra?.repo,
                    matched_patterns: announceRepos
                });
                it.env.room = 'int-dev-announce';
                if (!it.env.extra) it.env.extra = {};
                it.env.extra.announce_redirect = true;
                it.env.extra.original_room = originalRoom;
            }
            // Auto-group GitHub events by commit SHA: inject a dedup_key so
            // the first event creates the trackable message and subsequent
            // job statuses for the same SHA edit it in-place.
            const beforeKey = it.env.extra?.dedup_key;
            const beforeSha = it.env.extra?._commit_sha;
            normalizeGithubDedup(it);
            const afterKey = it.env.extra?.dedup_key;
            const afterSha = it.env.extra?._commit_sha;
            if (beforeKey !== afterKey || beforeSha !== afterSha) {
                trace.emit('normalize', {
                    room: it.env.room,
                    event_type: it.env.extra?.event_type,
                    before: { dedup_key: beforeKey, _commit_sha: beforeSha },
                    after: { dedup_key: afterKey, _commit_sha: afterSha }
                });
            }
            // PR-context attachment: reroute push/delete/issue_comment into
            // the PR's trackable when applicable, and seed the branch→PR
            // index when a pull_request event arrives. Also rewrites the
            // delete event's verbose message to say what was deleted.
            const beforePrKey = it.env.extra?.dedup_key;
            const beforePrSha = it.env.extra?._commit_sha;
            const prContext = await attachPrContext(it.env);
            if (prContext.rerouted) {
                trace.emit('pr_context_attached', {
                    room: it.env.room,
                    event_type: it.env.extra?.event_type,
                    reason: prContext.reason,
                    pr_number: prContext.pr_number,
                    before: { dedup_key: beforePrKey, _commit_sha: beforePrSha },
                    after: { dedup_key: it.env.extra?.dedup_key, _commit_sha: it.env.extra?._commit_sha }
                });
            }
            // Workflow-flavoured events also seed the per-repo active-workflow
            // index so a later bot/downstream push can look up its parent.
            await recordActiveWorkflow(it.env);
            if (WORKFLOW_EVENT_TYPES.has(it.env.extra?.event_type || '') && it.env.extra?._commit_sha) {
                trace.emit('wfactive_record', {
                    repo: it.env.extra.repo,
                    commit_sha: it.env.extra._commit_sha,
                    event_type: it.env.extra.event_type
                });
            }
            // Action-triggered child pushes (vhs commit, sync-sugarcraft
            // subtree pushes, etc.) get re-keyed onto the parent's dedup_key
            // so they nest under the existing trackable instead of spawning
            // a new message. PR-branch routing already won above for pushes
            // to PR head branches, so don't double-rewrite.
            const alreadyPrRouted = (it.env.extra?.dedup_key || '').startsWith('github:pr:');
            if (!alreadyPrRouted && isActionTriggeredPush(it.env)) {
                const parentSha = await findParentSha(it.env);
                if (parentSha && parentSha !== it.env.extra._commit_sha) {
                    const ownSha = it.env.extra._commit_sha;
                    it.env.extra._original_commit_sha = ownSha;
                    it.env.extra._commit_sha = parentSha;
                    it.env.extra.dedup_key = `github:commit:${ parentSha }`;
                    console.log(`[notif]   ↳ attributing ${ it.env.extra.repo }@${ ownSha } to parent ${ parentSha } (action-triggered)`);
                    trace.emit('action_triggered_attribution', {
                        repo: it.env.extra.repo,
                        own_sha: ownSha,
                        parent_sha: parentSha,
                        new_dedup_key: it.env.extra.dedup_key
                    });
                }
            }
            valid.push(it);
        }

        const byRoom = new Map();
        for (const it of valid) {
            const room = it.env.room || 'notifications';
            if (!byRoom.has(room)) byRoom.set(room, []);
            byRoom.get(room).push(it);
        }

        for (const [room, batch] of byRoom.entries()) {
            await processRoom(room, batch, stats);
        }
    } finally {
        tickInFlight = false;
        const ms = Date.now() - t0;
        if (stats.drained > 0) {
            console.log(`[notif] tick drained=${ stats.drained } sent=${ stats.sent } edited=${ stats.edited } coalesced=${ stats.coalesced } redirected=${ stats.redirected } fallback=${ stats.fallback } dead=${ stats.dead } expired=${ stats.expired } ms=${ ms }`);
            trace.emit('tick_end', { stats: { ...stats }, ms });
            scheduleNextTick(POLL_INTERVAL_FAST_MS);
        } else {
            scheduleNextTick(POLL_INTERVAL_MS);
        }
        lastTickStats = { ...stats, ms, at: new Date().toISOString() };
    }
}

async function recoverProcessing() {
    try {
        const stuck = await redis.lrange(k('processing'), 0, -1);
        if (stuck && stuck.length) {
            console.log(`[notif] recovering ${ stuck.length } stuck items from processing list`);
            // Push them to the head of the queue so they get picked up next.
            const pipe = redis.pipeline();
            for (const j of stuck) pipe.rpush(k('queue'), j);
            pipe.del(k('processing'));
            await pipe.exec();
        }
    } catch (err) {
        console.error('[notif] recovery failed:', err.message);
    }
}

async function drainBatch() {
    const out = [];
    for (let i = 0; i < MAX_PER_TICK; i++) {
        const j = await redis.rpoplpush(k('queue'), k('processing'));
        if (!j) break;
        let env;
        try {
            env = JSON.parse(j);
        } catch (err) {
            console.error('[notif] bad envelope JSON, dead-lettering:', err.message);
            await deadLetterRaw(j, 'json_parse_failed');
            continue;
        }
        out.push({ raw: j, env });
    }
    return out;
}

async function processRoom(room, batch, stats) {
    // Bucket: items with dedup_key are "trackable" (each goes out as its own
    // activity so we can edit them individually). Items without dedup_key
    // are coalescable.
    const trackable = [];
    const coalescable = [];
    for (const it of batch) {
        if (it.env.extra && it.env.extra.dedup_key) trackable.push(it);
        else coalescable.push(it);
    }
    trace.emit('route', {
        room,
        batch_size: batch.length,
        trackable_count: trackable.length,
        coalescable_count: coalescable.length,
        trackable_keys: trackable.map(it => it.env.extra?.dedup_key)
    });

    // 1. Trackable items: try to edit existing recent activity, else new send.
    // Calls are awaited sequentially so when several events for the same
    // dedup_key arrive in one batch, the first call's handleSingleNew saves
    // to Redis before the next call's handleTrackable looks it up — no
    // grouping work needs to happen here.
    for (const it of trackable) {
        await handleTrackable(room, it, stats);
    }

    // 2. Coalescable items: combine within room.
    if (coalescable.length === 1) {
        await handleSingleNew(room, coalescable[0], stats);
    } else if (coalescable.length > 1) {
        await handleCoalesced(room, coalescable, stats);
    }
}

async function handleTrackable(room, it, stats) {
    let recent = null;
    let recentSource = null;
    const dedupKey = it.env.extra.dedup_key;

    // Primary lookup by dedup_key
    try {
        const raw = await redis.hget(recentKey(room), dedupKey);
        if (raw) { recent = JSON.parse(raw); recentSource = 'dedup_key'; }
    } catch (err) {
        console.warn('[notif] hget failed:', err.message);
    }

    // Secondary lookup by commit SHA: if we have a _commit_sha and the primary
    // lookup missed, try github:commit:{sha} as the key. This handles the case
    // where a push was saved with its PHP-set dedup_key before normalizeGithubDedup
    // changed it to the commit-based key, but a subsequent job is looking for the
    // commit-based key.
    if (!recent && it.env.extra._commit_sha) {
        try {
            const shaBasedKey = `github:commit:${ it.env.extra._commit_sha }`;
            const raw = await redis.hget(recentKey(room), shaBasedKey);
            if (raw) { recent = JSON.parse(raw); recentSource = 'commit_sha_fallback'; }
        } catch (err) {
            console.warn('[notif] hget sha-based lookup failed:', err.message);
        }
    }

    const ageMs = recent ? Date.now() - recent.ts : null;
    const inWindow = recent && recent.activityId && ageMs < EDIT_WINDOW_MS && recent.type === it.env.type;
    trace.emit('recent_lookup', {
        room,
        dedup_key: dedupKey,
        commit_sha: it.env.extra._commit_sha,
        event_type: it.env.extra.event_type,
        found: !!recent,
        source: recentSource,
        age_ms: ageMs,
        edit_window_ms: EDIT_WINDOW_MS,
        eligible_for_edit: !!inWindow,
        recent: trace.snapshotRecent(recent)
    });

    if (inWindow) {
        const ok = await tryEdit(room, it, recent, stats);
        if (ok) {
            await ackOne(it);
            return;
        }
        // Edit failed → fall through to new send (overwrite cache)
        trace.emit('edit_fell_through', { room, dedup_key: dedupKey });
    }

    await handleSingleNew(room, it, stats);
}

async function tryEdit(room, it, recent, stats) {
    let conversationRef = await loadConvRef(recent.conversationId);
    let usedConstructed = false;
    if (!conversationRef) {
        // Stored convref missing — happens when `handleSingleNew` sent the
        // initial activity via `tryConstructedConvRef` (the bot has never
        // received an inbound message in that channel), so no real
        // `convref:{conversationId}` was ever written to Redis. The very
        // case we need to support: a PR fires multiple events in the same
        // tick, the first lands via constructed ref, and every subsequent
        // event arrives here. Without this fallback, every edit silently
        // bails to `handleSingleNew` and spawns a new top-level message
        // — which is exactly what was producing un-grouped PR notifications.
        conversationRef = buildConstructedConvRef(room, recent.conversationId);
        usedConstructed = true;
        trace.emit('edit_using_constructed_convref', {
            room,
            conversation_id: recent.conversationId,
            dedup_key: it.env.extra.dedup_key
        });
    }
    console.log(`[notif] ✎ edit room=${ room } conv=${ shortConv(recent.conversationId) } activity=${ recent.activityId } dedup=${ it.env.extra.dedup_key }${ usedConstructed ? ' [constructed]' : '' } "${ preview(it.env.message) }"`);

    let newText, newCard;
    let newHeader = recent.header || null;
    let newItems = Array.isArray(recent.items) ? recent.items : null;
    let newHeaderIdentity = recent.header_identity || null;
    let mergeMode = null;
    if (it.env.type === 'msg') {
        const recentSha = recent.commit_sha || null;
        const currentSha = it.env.extra && it.env.extra._commit_sha ? it.env.extra._commit_sha : null;
        const isGithubAppend = recentSha && currentSha && recentSha === currentSha;

        if (isGithubAppend) {
            mergeMode = 'github_trackable';
            const merged = mergeGithubTrackable(recent, it.env);
            newText = merged.text;
            newHeader = merged.header;
            newItems = merged.items;
            newHeaderIdentity = merged.header_identity;
        } else {
            mergeMode = 'standard_append';
            // ── Standard append mode (non-GitHub or different commit) ──────
            const appended = (recent.appended_count || 0) + 1;
            if (appended <= APPEND_TRAIL_SUMMARY_AFTER) {
                const ts = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false });
                newText = `${ recent.text || '' }\n\n— update at ${ ts } (+${ appended }) —\n${ it.env.message || '' }`;
            } else {
                const ts = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false });
                const baseLine = (recent.text || '').split('\n')[0] || '';
                newText = `${ baseLine }\n\n— ${ appended } updates · last ${ ts } — most recent: ${ it.env.message || '' }`;
            }
        }
    } else {
        mergeMode = 'card_replace';
        newCard = Array.isArray(it.env.card) ? it.env.card : [it.env.card];
    }

    trace.emit('edit_merge', {
        room,
        activity_id: recent.activityId,
        conversation_id: recent.conversationId,
        dedup_key: it.env.extra.dedup_key,
        merge_mode: mergeMode,
        before: {
            text: recent.text,
            header: recent.header,
            items: recent.items,
            header_identity: recent.header_identity,
            commit_sha: recent.commit_sha
        },
        after: {
            text: newText,
            header: newHeader,
            items: newItems,
            header_identity: newHeaderIdentity
        },
        incoming: trace.snapshotEnvelope(it.env)
    });

    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(conversationRef, async (proactiveContext) => {
                const activity = newCard
                    ? buildCardActivity(newCard, recent.activityId)
                    : { type: 'message', id: recent.activityId, text: newText };
                await proactiveContext.updateActivity(activity);
            });
        }, {
            label: `notif edit ${ room }`,
            serviceUrl: conversationRef.serviceUrl,
            maxRetries: 3
        });
    } catch (err) {
        console.warn(`[notif] edit failed for ${ room }/${ it.env.extra.dedup_key }: ${ err.message }`);
        trace.emit('edit_failed', {
            room,
            activity_id: recent.activityId,
            dedup_key: it.env.extra.dedup_key,
            error: err.message
        });
        await bumpMetric('edit_failed');
        return false;
    }
    trace.emit('edit_ok', {
        room,
        activity_id: recent.activityId,
        dedup_key: it.env.extra.dedup_key,
        appended_count: (recent.appended_count || 0) + 1
    });

    const updated = {
        activityId: recent.activityId,
        ts: Date.now(),
        type: it.env.type,
        text: newText || recent.text,
        header: newHeader,
        items: newItems,
        header_identity: newHeaderIdentity,
        appended_count: (recent.appended_count || 0) + 1,
        conversationId: recent.conversationId,
        commit_sha: it.env.extra && it.env.extra._commit_sha ? it.env.extra._commit_sha : (recent.commit_sha || null)
    };
    await saveRecent(room, it.env.extra.dedup_key, updated);
    stats.edited++;
    await bumpMetric('edited');
    return true;
}

/**
 * Compute the initial `{ header, items, header_identity }` to persist
 * alongside a freshly sent envelope. For text messages this splits leading
 * `•` bullet lines (a push's commit list) out of the header so that any
 * subsequent merge for the same dedup_key sees the same structure that
 * `mergeGithubTrackable` would have produced — preventing the same push
 * from being appended as a verbatim bullet on re-delivery.
 */
function initialTrackableState(env) {
    if (env.type !== 'msg') {
        return { header: '', items: [], header_identity: null };
    }
    const message = String(env.message || '');
    if (!message) {
        return { header: '', items: [], header_identity: jobIdentity(env) };
    }
    const split = splitHeaderAndBullets(message);
    return {
        header: split.headerText,
        items: split.bulletItems,
        header_identity: jobIdentity(env)
    };
}

async function handleSingleNew(room, it, stats) {
    const conversationId = resolveChannel(room) || resolveChannel('notifications');
    if (!conversationId) {
        await fallbackSend(room, [it], stats, 'unknown_room');
        return;
    }
    const conversationRef = await loadConvRef(conversationId);

    let activity;
    if (it.env.type === 'card') {
        const cards = Array.isArray(it.env.card) ? it.env.card : [it.env.card];
        activity = buildCardActivity(cards, null);
    } else {
        activity = { type: 'message', text: it.env.message || '' };
    }

    // If no stored convref, try constructed one before falling back to webhook
    if (!conversationRef) {
        const constructedActivityId = await tryConstructedConvRef(room, conversationId, activity, stats);
        if (constructedActivityId) {
            if (constructedActivityId && it.env.extra && it.env.extra.dedup_key) {
                const initial = initialTrackableState(it.env);
                await saveRecent(room, it.env.extra.dedup_key, {
                    activityId: constructedActivityId,
                    ts: Date.now(),
                    type: it.env.type,
                    text: it.env.message,
                    ...initial,
                    appended_count: 0,
                    conversationId,
                    commit_sha: it.env.extra._commit_sha || null
                });
            }
            stats.sent++;
            await bumpMetric('sent');
            await ackOne(it);
            return;
        }
        // Still no working convref - fall back to webhook
        await fallbackSend(room, [it], stats, 'no_convref');
        return;
    }

    const previewText = it.env.type === 'card' ? `[card ×${ Array.isArray(it.env.card) ? it.env.card.length : 1 }]` : preview(it.env.message);
    console.log(`[notif] → send room=${ room } conv=${ shortConv(conversationId) } "${ previewText }"`);
    trace.emit('send_attempt', {
        room,
        conversation_id: conversationId,
        dedup_key: it.env.extra?.dedup_key,
        type: it.env.type,
        text: it.env.type === 'msg' ? it.env.message : undefined
    });
    let activityId = null;
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(conversationRef, async (proactiveContext) => {
                const sent = await proactiveContext.sendActivity(activity);
                activityId = sent && sent.id ? sent.id : null;
            });
        }, {
            label: `notif send ${ room }`,
            serviceUrl: conversationRef.serviceUrl,
            maxRetries: 3
        });
    } catch (err) {
        console.warn(`[notif] send failed for ${ room }: ${ err.message }`);
        trace.emit('send_failed', { room, conversation_id: conversationId, error: err.message });
        await fallbackSend(room, [it], stats, 'send_failed:' + err.message);
        return;
    }
    console.log(`[notif]   sent room=${ room } activity=${ activityId || '<none>' }`);
    trace.emit('send_ok', {
        room,
        conversation_id: conversationId,
        activity_id: activityId,
        dedup_key: it.env.extra?.dedup_key
    });

    if (activityId && it.env.extra && it.env.extra.dedup_key) {
        const initial = initialTrackableState(it.env);
        await saveRecent(room, it.env.extra.dedup_key, {
            activityId,
            ts: Date.now(),
            type: it.env.type,
            text: it.env.message,
            ...initial,
            appended_count: 0,
            conversationId,
            commit_sha: it.env.extra._commit_sha || null
        });
    }
    stats.sent++;
    await bumpMetric('sent');
    await ackOne(it);
}

async function handleCoalesced(room, items, stats) {
    const conversationId = resolveChannel(room) || resolveChannel('notifications');
    if (!conversationId) {
        await fallbackSend(room, items, stats, 'unknown_room');
        return;
    }
    const conversationRef = await loadConvRef(conversationId);

    const msgItems = items.filter(it => it.env.type === 'msg');
    const cardItems = items.filter(it => it.env.type === 'card');

    // If no stored convref, try constructed one before falling back to webhook
    if (!conversationRef) {
        // Build a combined activity to try with constructed convref
        let combinedActivity = null;
        if (msgItems.length > 0) {
            combinedActivity = { type: 'message', text: msgItems.map(it => it.env.message || '').join(APPEND_BREAK) };
        } else if (cardItems.length > 0) {
            const attachments = cardItems.flatMap(it => {
                const cards = Array.isArray(it.env.card) ? it.env.card : [it.env.card];
                return cards.map(c => ({ contentType: 'application/vnd.microsoft.card.adaptive', content: c }));
            });
            combinedActivity = { type: 'message', attachments };
        }

        if (combinedActivity) {
            const constructedActivityId = await tryConstructedConvRef(room, conversationId, combinedActivity, stats);
            if (constructedActivityId) {
                stats.sent += msgItems.length + cardItems.length;
                await bumpMetric('sent');
                for (const it of items) await ackOne(it);
                return;
            }
        }
        // Still no working convref - fall back to webhook
        await fallbackSend(room, items, stats, 'no_convref');
        return;
    }

    if (msgItems.length > 0) {
        await sendCombinedText(room, conversationRef, msgItems, stats);
    }
    if (cardItems.length > 0) {
        await sendCombinedCards(room, conversationRef, cardItems, stats);
    }
}

async function sendCombinedText(room, conversationRef, msgItems, stats) {
    let combined = '';
    let included = 0;
    let leftover = [];
    for (const it of msgItems) {
        const piece = it.env.message || '';
        const next = combined.length === 0 ? piece : combined + APPEND_BREAK + piece;
        if (next.length > COALESCE_MAX_CHARS && included > 0) {
            leftover = msgItems.slice(included);
            break;
        }
        combined = next;
        included++;
        if (included >= COALESCE_MAX_ITEMS) {
            leftover = msgItems.slice(included);
            break;
        }
    }
    console.log(`[notif] ⊕ coalesce(text) room=${ room } items=${ included }${ leftover.length ? ' leftover=' + leftover.length : '' } "${ preview(combined) }"`);
    trace.emit('coalesce_text_attempt', {
        room, items_in: included, leftover: leftover.length,
        bytes: combined.length, text: combined
    });
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(conversationRef, async (proactiveContext) => {
                await proactiveContext.sendActivity({ type: 'message', text: combined });
            });
        }, { label: `notif coalesce ${ room }`, serviceUrl: conversationRef.serviceUrl, maxRetries: 3 });
    } catch (err) {
        console.warn(`[notif] coalesced send failed for ${ room }: ${ err.message }`);
        trace.emit('coalesce_text_failed', { room, error: err.message });
        await fallbackSend(room, msgItems.slice(0, included), stats, 'send_failed:' + err.message);
        for (const it of leftover) await fallbackSend(room, [it], stats, 'leftover_after_failure');
        return;
    }
    console.log(`[notif]   sent (coalesced) room=${ room } items=${ included }`);
    trace.emit('coalesce_text_ok', { room, items_in: included });
    stats.sent++;
    stats.coalesced += included;
    await bumpMetric('coalesced');
    for (const it of msgItems.slice(0, included)) await ackOne(it);
    // Leftover gets re-queued at the head so it's picked up on the next tick.
    for (const it of leftover) {
        try {
            await redis.lrem(k('processing'), 1, it.raw);
            await redis.rpush(k('queue'), it.raw);
        } catch (_) { /* best effort */ }
    }
}

async function sendCombinedCards(room, conversationRef, cardItems, stats) {
    const max = Math.min(cardItems.length, COALESCE_MAX_ITEMS);
    const attachments = [];
    for (let i = 0; i < max; i++) {
        const cards = Array.isArray(cardItems[i].env.card) ? cardItems[i].env.card : [cardItems[i].env.card];
        for (const c of cards) {
            attachments.push({ contentType: 'application/vnd.microsoft.card.adaptive', content: c });
        }
    }
    console.log(`[notif] ⊕ coalesce(cards) room=${ room } items=${ max } attachments=${ attachments.length }`);
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(conversationRef, async (proactiveContext) => {
                await proactiveContext.sendActivity({ type: 'message', attachments });
            });
        }, { label: `notif card-coalesce ${ room }`, serviceUrl: conversationRef.serviceUrl, maxRetries: 3 });
    } catch (err) {
        console.warn(`[notif] card-coalesce send failed for ${ room }: ${ err.message }`);
        await fallbackSend(room, cardItems.slice(0, max), stats, 'send_failed:' + err.message);
        return;
    }
    console.log(`[notif]   sent (cards-coalesced) room=${ room }`);
    stats.sent++;
    stats.coalesced += max;
    for (const it of cardItems.slice(0, max)) await ackOne(it);
    for (const it of cardItems.slice(max)) {
        try {
            await redis.lrem(k('processing'), 1, it.raw);
            await redis.rpush(k('queue'), it.raw);
        } catch (_) { /* best effort */ }
    }
}

async function fallbackSend(room, items, stats, reason) {
    for (const it of items) {
        const url = it.env.fallback_webhook_url;
        if (!url) {
            console.warn(`[notif] ⤳ fallback abandoned room=${ room } reason=${ reason }_no_fallback`);
            trace.emit('fallback_abandoned', { room, reason: `${ reason }_no_fallback`, envelope: trace.snapshotEnvelope(it.env) });
            await deadLetter(it, `${ reason }_no_fallback`);
            stats.dead++;
            continue;
        }
        const urlHost = (() => { try { return new URL(url).host; } catch (_) { return 'unknown'; } })();
        const previewText = it.env.type === 'card' ? `[card ×${ Array.isArray(it.env.card) ? it.env.card.length : 1 }]` : preview(it.env.message);
        console.log(`[notif] ⤳ fallback room=${ room } host=${ urlHost } reason=${ reason } "${ previewText }"`);
        trace.emit('fallback_attempt', { room, url_host: urlHost, reason });
        try {
            const body = it.env.type === 'card'
                ? {
                    type: 'message',
                    attachments: (Array.isArray(it.env.card) ? it.env.card : [it.env.card]).map(c => ({
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: c
                    }))
                }
                : { type: 'message', message: it.env.message || '' };
            await axios.post(url, body, { timeout: 30000 });
            console.log(`[notif]   fallback OK room=${ room }`);
            trace.emit('fallback_ok', { room, url_host: urlHost });
            stats.fallback++;
            await bumpMetric('fallback');
            await ackOne(it);
        } catch (err) {
            console.error(`[notif fallback] webhook POST failed for ${ room }: ${ err.message }`);
            trace.emit('fallback_failed', { room, url_host: urlHost, error: err.message });
            await deadLetter(it, `fallback_failed:${ err.message }`);
            stats.dead++;
            await bumpMetric('fallback_failed');
        }
    }
}

async function ackOne(it) {
    try {
        await redis.lrem(k('processing'), 1, it.raw);
    } catch (err) {
        console.warn('[notif] ack lrem failed:', err.message);
    }
}

async function deadLetter(it, reason) {
    try {
        const env = { ...it.env, _dead_reason: reason, _dead_at: Date.now() };
        await redis.multi()
            .lpush(k('dead'), JSON.stringify(env))
            .ltrim(k('dead'), 0, 999)
            .lrem(k('processing'), 1, it.raw)
            .exec();
        trace.emit('dead_lettered', {
            room: it.env.room,
            reason,
            dedup_key: it.env.extra?.dedup_key,
            event_type: it.env.extra?.event_type
        });
        await bumpMetric('dead');
    } catch (err) {
        console.error('[notif] deadLetter failed:', err.message);
    }
}

async function deadLetterRaw(raw, reason) {
    try {
        await redis.multi()
            .lpush(k('dead'), JSON.stringify({ raw, _dead_reason: reason, _dead_at: Date.now() }))
            .ltrim(k('dead'), 0, 999)
            .lrem(k('processing'), 1, raw)
            .exec();
    } catch (err) {
        console.error('[notif] deadLetterRaw failed:', err.message);
    }
}

function redirectToAnnounce(it, reason) {
    const originalRoom = it.env.room || 'unknown';
    console.log(`[notif] ⇢ redirect ${ originalRoom } → int-dev-announce (${ reason })`);
    it.env.room = 'int-dev-announce';
    if (!it.env.extra) it.env.extra = {};
    it.env.extra.filtered = true;
    it.env.extra.filter_reason = reason;
    it.env.extra.original_room = originalRoom;
    // Annotate text envelopes with a small marker so the announce channel
    // is readable as "here's what got filtered, and why". Card envelopes
    // are forwarded unchanged — the marker would clutter Adaptive Cards.
    if (it.env.type === 'msg') {
        const head = `_filtered from ${ originalRoom }: ${ reason }_`;
        it.env.message = it.env.message ? `${ head }\n\n${ it.env.message }` : head;
    }
}

async function loadConvRef(conversationId) {
    try {
        // Conversation references are written by botActivityHandler on every
        // inbound message and live on the bot's primary Redis (Dragonfly),
        // not on the notification queue Redis.
        const client = redisBot || redis;
        const stored = await client.get(`convref:${ conversationId }`);
        if (!stored) {
            await bumpMetric(`convref_missing:${ conversationId }`);
            return null;
        }
        return JSON.parse(stored);
    } catch (err) {
        console.warn('[notif] loadConvRef failed:', err.message);
        return null;
    }
}

// Build a minimal ConversationReference from just a room name + conversation
// ID. Used both for first-send fallback (`tryConstructedConvRef`) and for
// edit fallback (`tryEdit`) when `loadConvRef` returns null because the bot
// has never observed an inbound activity in that channel.
//
// The key fields needed are `serviceUrl` and `conversation.id`. `aadObjectId`
// and `tenantId` are used for bot identity — read from env when set.
function buildConstructedConvRef(room, conversationId) {
    return {
        serviceUrl: TEAMS_SERVICE_URL,
        conversation: {
            id: conversationId,
            name: room,
            isGroup: true
        },
        aadObjectId: process.env.BOT_AAD_OBJECT_ID || 'unknown',
        tenantId: process.env.BOT_TENANT_ID || 'unknown',
        bot: {
            id: process.env.MicrosoftAppId,
            name: 'teams-chat-bot'
        },
        channelId: 'msteams',
        _constructed: true
    };
}

// Try to send via Bot Framework using a constructed ConversationReference.
// This is a fallback when loadConvRef returns null (e.g., bot was installed
// before onInstallationUpdateAdd started capturing convrefs).
// Returns the new activity id on success, null on failure.
async function tryConstructedConvRef(room, conversationId, activity, stats) {
    const constructedRef = buildConstructedConvRef(room, conversationId);
    console.log(`[notif] → trying constructed convref room=${ room } conv=${ shortConv(conversationId) } serviceUrl=${ TEAMS_SERVICE_URL }`);
    trace.emit('constructed_convref_attempt', { room, conversation_id: conversationId, op: 'send' });
    let activityId = null;
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(constructedRef, async (proactiveContext) => {
                const sent = await proactiveContext.sendActivity(activity);
                activityId = sent && sent.id ? sent.id : null;
            });
        }, {
            label: `notif constructed-ref ${ room }`,
            serviceUrl: TEAMS_SERVICE_URL,
            maxRetries: 2
        });
        console.log(`[notif]   constructed-convref success room=${ room } activity=${ activityId || '<none>' }`);
        trace.emit('constructed_convref_ok', { room, conversation_id: conversationId, op: 'send', activity_id: activityId });
        return activityId;
    } catch (err) {
        console.warn(`[notif] constructed-convref failed for ${ room }: ${ err.message }`);
        trace.emit('constructed_convref_failed', { room, conversation_id: conversationId, op: 'send', error: err.message });
        return null;
    }
}

async function saveRecent(room, dedupKey, value) {
    try {
        const pipe = redis.pipeline();
        pipe.hset(recentKey(room), dedupKey, JSON.stringify(value));
        pipe.expire(recentKey(room), Math.ceil(EDIT_WINDOW_MS / 1000));
        await pipe.exec();
        trace.emit('recent_saved', {
            room,
            dedup_key: dedupKey,
            activity_id: value.activityId,
            commit_sha: value.commit_sha,
            appended_count: value.appended_count,
            header_identity: value.header_identity,
            item_count: Array.isArray(value.items) ? value.items.length : 0
        });
    } catch (err) {
        console.warn('[notif] saveRecent failed:', err.message);
    }
}

async function bumpMetric(name) {
    try {
        await redis.incr(metric(name));
    } catch (_) { /* best effort */ }
}

function buildCardActivity(cards, replaceActivityId) {
    const attachments = cards.map(c => ({
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: c
    }));
    const activity = { type: 'message', attachments };
    if (replaceActivityId) activity.id = replaceActivityId;
    return activity;
}

async function getHealth() {
    if (!redis) {
        return { running: false };
    }
    try {
        const [queueDepth, processingDepth, deadDepth] = await Promise.all([
            redis.llen(k('queue')),
            redis.llen(k('processing')),
            redis.llen(k('dead'))
        ]);
        return {
            running: true,
            queue_depth: queueDepth,
            processing_depth: processingDepth,
            dead_depth: deadDepth,
            poll_interval_ms: POLL_INTERVAL_MS,
            edit_window_ms: EDIT_WINDOW_MS,
            max_per_tick: MAX_PER_TICK,
            last_tick: lastTickStats
        };
    } catch (err) {
        return { running: true, redis_error: err.message };
    }
}

// Test-only injection seam. Lets tests substitute mock redis / redisBot /
// adapter handles so the handleTrackable / tryEdit / handleSingleNew paths
// can be driven without a live Redis or Bot Framework adapter. Returns the
// prior values so a test can restore them in an `after` hook.
function _setInternalsForTest(overrides = {}) {
    const prev = { redis, redisBot, adapter };
    if ('redis' in overrides) redis = overrides.redis;
    if ('redisBot' in overrides) redisBot = overrides.redisBot;
    if ('adapter' in overrides) adapter = overrides.adapter;
    return prev;
}

module.exports = {
    startConsumer, stopConsumer, getHealth, runTick, getNotifRedis,
    mergeGithubTrackable,
    // exported for !notif wfactive and tests
    parseDownstreamMap, isActionTriggeredPush, findParentSha, recordActiveWorkflow,
    wfactiveKey, DOWNSTREAM_REPOS,
    EDIT_WINDOW_MS,
    // exported for tests only — do not call from production code
    _setInternalsForTest, handleTrackable, tryEdit, handleSingleNew,
    buildConstructedConvRef,
    attachPrContext, recordPrBranch, lookupPrByBranch,
    prBranchKey, prNumberFromIssue, branchFromRef
};
