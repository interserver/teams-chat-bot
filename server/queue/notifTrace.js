// Replay-grade trace logger for the notification consumer.
//
// Emits one JSON object per line (JSONL) to `.logs/notif-trace-YYYY-MM-DD.jsonl`
// for every decision point in the queue pipeline: drain, normalize, filter,
// route, recent-lookup, edit, send, fallback, dead-letter, ack.
//
// The intent is "given just this file, you can reconstruct every action the
// bot took and the data it acted on" — so each line carries enough context
// (envelope, recent state, merged output) to replay the scenario offline.
//
// Disable with `NOTIF_TRACE_LOG=0`. Override path with `NOTIF_TRACE_LOG=/abs/path.jsonl`
// (auto-rotation by date is skipped when an explicit path is given).

const fs = require('fs');
const path = require('path');

const DEFAULT_DIR = path.resolve(__dirname, '..', '..', '.logs');
const RAW = process.env.NOTIF_TRACE_LOG;
const DISABLED = RAW === '0' || RAW === 'false';
const EXPLICIT_PATH = RAW && !DISABLED && RAW !== '1' && RAW !== 'true' ? RAW : null;

let writeQueue = Promise.resolve();
let currentTick = 0;
let seqWithinTick = 0;
let warnedOnce = false;

function todayPath() {
    if (EXPLICIT_PATH) return EXPLICIT_PATH;
    const d = new Date();
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return path.join(DEFAULT_DIR, `notif-trace-${ yyyy }-${ mm }-${ dd }.jsonl`);
}

function ensureDir(p) {
    try {
        fs.mkdirSync(path.dirname(p), { recursive: true });
    } catch (err) {
        if (!warnedOnce) {
            console.warn(`[notif-trace] cannot create log dir ${ path.dirname(p) }: ${ err.message }`);
            warnedOnce = true;
        }
    }
}

function newTick() {
    currentTick++;
    seqWithinTick = 0;
    return currentTick;
}

// Strip large/redundant fields from an envelope for logging. We keep the full
// `extra` (event_type, dedup_key, _commit_sha, data, etc.) because that's the
// raw input the consumer's logic branches on. `data` itself can be large for
// GitHub webhooks; we leave it intact deliberately — the whole point of this
// log is replay fidelity, and silently dropping payload fields would lie to
// future-you about what the consumer saw.
function snapshotEnvelope(env) {
    if (!env || typeof env !== 'object') return env;
    return {
        type: env.type,
        room: env.room,
        message: env.message,
        // `card` may be very large; preserve attachment count plus a small
        // hint instead of the full JSON so cards don't dominate the log
        card: env.card ? (Array.isArray(env.card)
            ? { _cards: env.card.length }
            : { _card: true }) : undefined,
        fallback_webhook_url: env.fallback_webhook_url,
        expires_at: env.expires_at,
        extra: env.extra
    };
}

// `recent` rows can contain a fully rendered Markdown body (`text`) and the
// parsed items list. Both are essential for replay so we keep them whole.
function snapshotRecent(recent) {
    if (!recent) return null;
    return {
        activityId: recent.activityId,
        ts: recent.ts,
        type: recent.type,
        text: recent.text,
        header: recent.header,
        items: recent.items,
        header_identity: recent.header_identity,
        appended_count: recent.appended_count,
        conversationId: recent.conversationId,
        commit_sha: recent.commit_sha
    };
}

function write(record) {
    if (DISABLED) return;
    const filePath = todayPath();
    ensureDir(filePath);
    const line = JSON.stringify(record) + '\n';
    // Serialise writes so a slow disk can't interleave lines from concurrent
    // emit() calls. fs.appendFile is async but we chain on the prior promise.
    writeQueue = writeQueue.then(() => new Promise(resolve => {
        fs.appendFile(filePath, line, err => {
            if (err && !warnedOnce) {
                console.warn(`[notif-trace] write failed for ${ filePath }: ${ err.message }`);
                warnedOnce = true;
            }
            resolve();
        });
    }));
}

/**
 * Emit one trace event. `kind` identifies the decision point; everything
 * else is free-form payload. `tick` and `seq` are auto-attached so a reader
 * can sort by them when timestamps tie at millisecond resolution.
 */
function emit(kind, payload = {}) {
    if (DISABLED) return;
    seqWithinTick++;
    const now = Date.now();
    write({
        t: now,
        iso: new Date(now).toISOString(),
        tick: currentTick,
        seq: seqWithinTick,
        kind,
        ...payload
    });
}

function isEnabled() { return !DISABLED; }
function currentPath() { return DISABLED ? null : todayPath(); }

// Force any pending appendFile calls to flush. Used by tests so the
// assertions can read the file synchronously after triggering emits.
async function flush() { await writeQueue; }

module.exports = {
    emit, newTick, snapshotEnvelope, snapshotRecent,
    isEnabled, currentPath, flush
};
