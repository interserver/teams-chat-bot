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
        sha = data.pull_request?.merge_commit_sha || data.pull_request?.head?.sha || null;
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
    const eventType = extra.event_type || '';

    // For all GitHub events that carry a commit SHA: ensure we have _commit_sha set
    // even if dedup_key was already set by PHP. This enables secondary SHA-based lookup.
    const sha = extractCommitSha(it.env);
    if (!sha) return; // no SHA — can't group

    extra._commit_sha = sha;

    // For events that carry a commit SHA we always force the dedup key to
    // be commit-scoped so that pushes, workflow events, status events and
    // PR-review events for the same commit all merge into one trackable
    // message. PHP-set dedup keys are intentionally overridden.
    if (eventType === 'push' || eventType === 'workflow_job' ||
        eventType === 'check_run' || eventType === 'check_suite' ||
        eventType === 'workflow_run' || eventType === 'status' ||
        eventType === 'pull_request' || eventType === 'pull_request_review' ||
        eventType === 'pull_request_review_comment' || eventType === 'commit_comment') {
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
    return null;
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
            for (const s of stripped) {
                lines.push(bulletPrefix(level + 1) + leafOnly(s));
            }
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
                for (const s of statusGroup) {
                    lines.push(bulletPrefix(level + 2) + leafOnly(s));
                }
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
    let result = lines.join('\n');

    // Teams natively renders 2 visual levels of Markdown nested bullets
    // (level 1 plus one level of indentation). When the tree we built
    // never goes deeper than that — no line starts with the level-3 NBSP
    // prefix (4 NBSPs + `-`) — swap the level-2 NBSPs back to regular
    // spaces so Teams falls back to its native nested-list rendering
    // which looks cleaner than literal indented text. As soon as we have
    // depth 3+ the NBSPs stay so the visible indent doesn't collapse
    // past Teams' single-level Markdown nesting limit.
    const L3_PREFIX = NBSP.repeat(4) + '-';
    const hasDepth3 = result.split('\n').some(l => l.startsWith(L3_PREFIX));
    if (!hasDepth3) {
        result = result.replace(new RegExp('^' + NBSP + '+', 'gm'), m => ' '.repeat(m.length));
    }
    return result;
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
    const bulletMessage = condensed || verboseMessage;
    const headerMessage = verboseMessage || condensed;
    const incomingIdentity = jobIdentity(env);

    let header = recentObj.header || recentObj.text || '';
    let items = Array.isArray(recentObj.items) ? recentObj.items.map(it => ({ ...it })) : [];
    let headerIdentity = recentObj.header_identity || null;

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
    const dedupKey = it.env.extra.dedup_key;

    // Primary lookup by dedup_key
    try {
        const raw = await redis.hget(recentKey(room), dedupKey);
        if (raw) recent = JSON.parse(raw);
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
            if (raw) recent = JSON.parse(raw);
        } catch (err) {
            console.warn('[notif] hget sha-based lookup failed:', err.message);
        }
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
    let newHeader = recent.header || null;
    let newItems = Array.isArray(recent.items) ? recent.items : null;
    let newHeaderIdentity = recent.header_identity || null;
    if (it.env.type === 'msg') {
        const recentSha = recent.commit_sha || null;
        const currentSha = it.env.extra && it.env.extra._commit_sha ? it.env.extra._commit_sha : null;
        const isGithubAppend = recentSha && currentSha && recentSha === currentSha;

        if (isGithubAppend) {
            const merged = mergeGithubTrackable(recent, it.env);
            newText = merged.text;
            newHeader = merged.header;
            newItems = merged.items;
            newHeaderIdentity = merged.header_identity;
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

module.exports = { startConsumer, stopConsumer, getHealth, runTick, getNotifRedis, mergeGithubTrackable };
