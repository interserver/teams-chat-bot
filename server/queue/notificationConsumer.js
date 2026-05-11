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
function metric(name) { return k('metrics:') + name; }

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
        // commits[0].id is the "after" SHA for a push event
        sha = data.commits?.[0]?.id || data.after || null;
    } else if (eventType === 'commit_comment') {
        sha = data.comment?.commit_id || null;
    } else if (eventType === 'pull_request') {
        // Use the merge commit SHA (head SHA) if available
        sha = data.pull_request?.merge_commit_sha || data.pull_request?.head?.sha || null;
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
 * Format: [check_run|workflow] name | failures:N | conclusion:STATUS
 * Returns null if this envelope doesn't represent a runnable job.
 */
function buildGithubJobLine(env) {
    const extra = env.extra || {};
    const data = extra.data || {};
    const eventType = extra.event_type || '';

    let name = '';
    let failures = 0;
    let conclusion = '';
    let status = '';
    let htmlUrl = '';
    let startedAt = '';
    let completedAt = '';

    if (eventType === 'check_run') {
        const cr = data.check_run || {};
        name = cr.name || 'check_run';
        failures = cr.failures_count || 0;
        conclusion = cr.conclusion || '';
        status = cr.status || '';
        htmlUrl = cr.html_url || '';
        startedAt = cr.started_at || '';
        completedAt = cr.completed_at || '';
    } else if (eventType === 'workflow_job') {
        const wj = data.workflow_job || {};
        name = wj.name || wj.workflow_name || 'workflow_job';
        failures = wj.failures_count || 0;
        conclusion = wj.conclusion || '';
        status = wj.status || '';
        htmlUrl = wj.html_url || '';
        startedAt = wj.started_at || '';
        completedAt = wj.completed_at || '';
    } else {
        return null;
    }

    // Build the line components
    const label = eventType === 'check_run' ? 'check_run' : 'workflow';
    const parts = [`[${ label }] ${ name }`];

    // For check_run: also include commit short SHA and first line of commit msg
    // (pulled from the check_suite head_commit) so the line is self-contained.
    if (eventType === 'check_run') {
        const commit = data.check_run?.check_suite?.head_commit;
        if (commit) {
            const shortSha = (commit.id || '').slice(0, 7);
            const msgFirstLine = (commit.message || '').split('\n')[0];
            if (shortSha) parts.push(shortSha);
            if (msgFirstLine) parts.push(msgFirstLine.slice(0, 80));
        }
    } else if (eventType === 'workflow_job') {
        // workflow_job references a workflow_run — pull head SHA from there
        const wfRun = data.workflow_job?.workflow_run;
        if (wfRun) {
            const shortSha = (wfRun.head_sha || '').slice(0, 7);
            if (shortSha) parts.push(shortSha);
        }
    }

    if (status && !conclusion) {
        parts.push(`status: ${ status }`);
    }
    if (failures > 0) {
        parts.push(`failures: ${ failures }`);
    }
    if (conclusion) {
        parts.push(`conclusion: ${ conclusion }`);
    }

    // Show elapsed time if started but not completed
    if (startedAt && !completedAt) {
        const elapsed = elapsedMs(startedAt);
        if (elapsed !== null) parts.push(`elapsed: ${ elapsed }`);
    }

    let line = parts.join(' | ');
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
    if (extra.dedup_key) return; // already has one — preserve
    const sha = extractCommitSha(it.env);
    if (!sha) return; // no SHA — can't group
    extra.dedup_key = `github:commit:${ sha }`;
    extra._commit_sha = sha;
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
        for (const it of items) {
            const ev = (it.env.extra && it.env.extra.event_type) ? ` ev=${ it.env.extra.event_type }` : '';
            const dk = (it.env.extra && it.env.extra.dedup_key) ? ` dedup=${ it.env.extra.dedup_key }` : '';
            console.log(`[notif]   • room=${ it.env.room } type=${ it.env.type }${ ev }${ dk } msg="${ preview(it.env.message) }"`);
        }

        const valid = [];
        const nowSec = Math.floor(Date.now() / 1000);
        for (const it of items) {
            if (it.env.expires_at && it.env.expires_at < nowSec) {
                stats.expired++;
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
                        await ackOne(it);
                        stats.redirected = (stats.redirected || 0) + 1;
                        await bumpMetric('redirected');
                        continue;
                    }
                    redirectToAnnounce(it, skipReason);
                    stats.redirected = (stats.redirected || 0) + 1;
                    await bumpMetric('redirected');
                }
            }
            // Auto-group GitHub events by commit SHA: inject a dedup_key so
            // the first event creates the trackable message and subsequent
            // job statuses for the same SHA edit it in-place.
            normalizeGithubDedup(it);
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

    // 1. Trackable items: try to edit existing recent activity, else new send.
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
    try {
        const raw = await redis.hget(recentKey(room), it.env.extra.dedup_key);
        if (raw) recent = JSON.parse(raw);
    } catch (err) {
        console.warn('[notif] hget failed:', err.message);
    }

    if (recent && recent.activityId && (Date.now() - recent.ts < EDIT_WINDOW_MS) && recent.type === it.env.type) {
        const ok = await tryEdit(room, it, recent, stats);
        if (ok) {
            await ackOne(it);
            return;
        }
        // Edit failed → fall through to new send (overwrite cache)
    }

    await handleSingleNew(room, it, stats);
}

async function tryEdit(room, it, recent, stats) {
    const conversationRef = await loadConvRef(recent.conversationId);
    if (!conversationRef) return false;
    console.log(`[notif] ✎ edit room=${ room } conv=${ shortConv(recent.conversationId) } activity=${ recent.activityId } dedup=${ it.env.extra.dedup_key } "${ preview(it.env.message) }"`);

    let newText, newCard;
    if (it.env.type === 'msg') {
        const recentSha = recent.commit_sha || null;
        const currentSha = it.env.extra && it.env.extra._commit_sha ? it.env.extra._commit_sha : null;
        const isGithubAppend = recentSha && currentSha && recentSha === currentSha;

        if (isGithubAppend) {
            // ── GitHub commit append mode ──────────────────────────────────
            // All events for the same commit SHA land in one trackable message.
            // Parse existing text as: {header} + "— jobs —" + {existing job list}
            //
            // New job arrives → append to job list, update message
            // New commit (push) arrives → replace header only, keep job list
            //   (so the push replaces the placeholder job-only header with
            //    the real commit info while preserving CI status)
            const existingText = recent.text || '';
            const JOB_LIST_MARKER = '\n\n— jobs —\n';
            const COMMIT_SHA_RE = /^[0-9a-f]{7} [^\n]*/;
            let header = '';
            let jobList = '';

            const markerPos = existingText.indexOf(JOB_LIST_MARKER);
            if (markerPos !== -1) {
                header = existingText.slice(0, markerPos);
                jobList = existingText.slice(markerPos + JOB_LIST_MARKER.length);
            } else {
                // No marker yet.  If existing starts with a 7-char hex SHA it is
                // a real commit header; otherwise it is a job-only placeholder
                // (created when a job arrived before the push).  In that case the
                // incoming commit message becomes the new header.
                if (COMMIT_SHA_RE.test(existingText)) {
                    header = existingText;
                } else {
                    header = it.env.message || `[${ it.env.extra.event_type }]`;
                    jobList = existingText;
                }
            }

            // Determine the new job line to append/update
            const jobLine = buildGithubJobLine(it.env);
            const isCommitMessage = !jobLine && it.env.message && String(it.env.message).trim().length > 0;

            if (isCommitMessage) {
                // ── New commit/push message arrived — replace header, keep job list
                header = it.env.message || '';
                newText = header + (jobList ? JOB_LIST_MARKER + jobList : '');
            } else if (jobLine) {
                // ── New job status — always append as a new entry (don't update in-place)
                // This gives a chronological timeline: queued → in_progress → success/failure
                if (jobList.length === 0) {
                    jobList = jobLine;
                } else {
                    jobList = `${ jobList }\n${ jobLine }`;
                }
                newText = `${ header }${ JOB_LIST_MARKER }${ jobList }`;
            } else {
                // Fallback: just append
                newText = `${ header }\n\n${ it.env.message || '' }`;
            }
        } else {
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
        newCard = Array.isArray(it.env.card) ? it.env.card : [it.env.card];
    }

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
        await bumpMetric('edit_failed');
        return false;
    }

    const updated = {
        activityId: recent.activityId,
        ts: Date.now(),
        type: it.env.type,
        text: newText || recent.text,
        appended_count: (recent.appended_count || 0) + 1,
        conversationId: recent.conversationId,
        commit_sha: it.env.extra && it.env.extra._commit_sha ? it.env.extra._commit_sha : (recent.commit_sha || null)
    };
    await saveRecent(room, it.env.extra.dedup_key, updated);
    stats.edited++;
    await bumpMetric('edited');
    return true;
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
                await saveRecent(room, it.env.extra.dedup_key, {
                    activityId: constructedActivityId,
                    ts: Date.now(),
                    type: it.env.type,
                    text: it.env.message,
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
        await fallbackSend(room, [it], stats, 'send_failed:' + err.message);
        return;
    }
    console.log(`[notif]   sent room=${ room } activity=${ activityId || '<none>' }`);

    if (activityId && it.env.extra && it.env.extra.dedup_key) {
        await saveRecent(room, it.env.extra.dedup_key, {
            activityId,
            ts: Date.now(),
            type: it.env.type,
            text: it.env.message,
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
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(conversationRef, async (proactiveContext) => {
                await proactiveContext.sendActivity({ type: 'message', text: combined });
            });
        }, { label: `notif coalesce ${ room }`, serviceUrl: conversationRef.serviceUrl, maxRetries: 3 });
    } catch (err) {
        console.warn(`[notif] coalesced send failed for ${ room }: ${ err.message }`);
        await fallbackSend(room, msgItems.slice(0, included), stats, 'send_failed:' + err.message);
        for (const it of leftover) await fallbackSend(room, [it], stats, 'leftover_after_failure');
        return;
    }
    console.log(`[notif]   sent (coalesced) room=${ room } items=${ included }`);
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
            await deadLetter(it, `${ reason }_no_fallback`);
            stats.dead++;
            continue;
        }
        const urlHost = (() => { try { return new URL(url).host; } catch (_) { return 'unknown'; } })();
        const previewText = it.env.type === 'card' ? `[card ×${ Array.isArray(it.env.card) ? it.env.card.length : 1 }]` : preview(it.env.message);
        console.log(`[notif] ⤳ fallback room=${ room } host=${ urlHost } reason=${ reason } "${ previewText }"`);
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
            stats.fallback++;
            await bumpMetric('fallback');
            await ackOne(it);
        } catch (err) {
            console.error(`[notif fallback] webhook POST failed for ${ room }: ${ err.message }`);
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

// Try to send via Bot Framework using a constructed ConversationReference.
// This is a fallback when loadConvRef returns null (e.g., bot was installed
// before onInstallationUpdateAdd started capturing convrefs).
// Returns true if successful, false otherwise.
async function tryConstructedConvRef(room, conversationId, activity, stats) {
    // Use configurable service URL (supports GCC/GCC High via TEAMS_SERVICE_URL env var)
    const SERVICE_URL = TEAMS_SERVICE_URL;

    // Construct a minimal ConversationReference with what we know.
    // The key fields needed are serviceUrl and conversation.id.
    // aadObjectId and tenantId are used for bot identity - use env values if available.
    const constructedRef = {
        serviceUrl: SERVICE_URL,
        conversation: {
            id: conversationId,
            name: room,  // use the actual room name, not a hardcoded value
            isGroup: true
        },
        aadObjectId: process.env.BOT_AAD_OBJECT_ID || 'unknown',
        tenantId: process.env.BOT_TENANT_ID || 'unknown',
        bot: {
            id: process.env.MicrosoftAppId,
            name: 'teams-chat-bot'
        },
        channelId: 'msteams',
        _constructed: true  // marker for debugging
    };

    console.log(`[notif] → trying constructed convref room=${ room } conv=${ shortConv(conversationId) } serviceUrl=${ SERVICE_URL }`);
    let activityId = null;
    try {
        await runWithRetry(async () => {
            await adapter.continueConversation(constructedRef, async (proactiveContext) => {
                const sent = await proactiveContext.sendActivity(activity);
                activityId = sent && sent.id ? sent.id : null;
            });
        }, {
            label: `notif constructed-ref ${ room }`,
            serviceUrl: SERVICE_URL,
            maxRetries: 2
        });
        console.log(`[notif]   constructed-convref success room=${ room } activity=${ activityId || '<none>' }`);
        return activityId;
    } catch (err) {
        console.warn(`[notif] constructed-convref failed for ${ room }: ${ err.message }`);
        return null;
    }
}

async function saveRecent(room, dedupKey, value) {
    try {
        const pipe = redis.pipeline();
        pipe.hset(recentKey(room), dedupKey, JSON.stringify(value));
        pipe.expire(recentKey(room), Math.ceil(EDIT_WINDOW_MS / 1000));
        await pipe.exec();
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

module.exports = { startConsumer, stopConsumer, getHealth, runTick, getNotifRedis };
