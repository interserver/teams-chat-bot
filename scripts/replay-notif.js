#!/usr/bin/env node
/* eslint-disable no-console */
// Replay / inspect the notification consumer trace log.
//
// The consumer writes JSONL events to .logs/notif-trace-YYYY-MM-DD.jsonl —
// one line per decision point (drain, normalize, route, recent_lookup,
// edit_merge, send, fallback, dead-letter, ack). This script reads those
// files and prints a filtered, grouped timeline so you can answer "what
// did the bot do for SHA X" or "why didn't events for room Y get grouped".
//
// Usage:
//   scripts/replay-notif.js [--file PATH]... [filters] [--mode timeline|grouped|raw]
//
// Filters (logical AND across all that are given):
//   --room NAME              one room name (e.g. int-dev)
//   --commit SHA             match envelope/recent _commit_sha (prefix-matched)
//   --dedup KEY              match dedup_key (exact)
//   --event TYPE             match GitHub event_type (e.g. push, check_run)
//   --kind KIND              match trace event kind (drained, edit_merge…)
//   --tick N                 match a single tick number
//   --since ISO              only events at/after this UTC timestamp
//   --until ISO              only events at/before this UTC timestamp
//   --activity ID            match activity_id (often only available on edits)
//
// Modes:
//   timeline (default)       one-line summary per event, chronological
//   grouped                  bucket by dedup_key/commit_sha, show what
//                            happened for each trackable thread
//   raw                      dump the full JSON line for each match
//
// Examples:
//   scripts/replay-notif.js --commit 4fd48e4
//   scripts/replay-notif.js --room int-dev --since 2026-05-13T00:00:00Z
//   scripts/replay-notif.js --dedup github:pr:detain/foo:42 --mode grouped
//   scripts/replay-notif.js --kind edit_merge --tick 17 --mode raw

const fs = require('fs');
const path = require('path');
const readline = require('readline');

function parseArgs(argv) {
    const args = {
        files: [],
        filters: {},
        mode: 'timeline'
    };
    for (let i = 2; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-h' || a === '--help') { args.help = true; continue; }
        if (a === '--mode') { args.mode = argv[++i]; continue; }
        if (a === '--file') { args.files.push(argv[++i]); continue; }
        if (a === '--room') { args.filters.room = argv[++i]; continue; }
        if (a === '--commit') { args.filters.commit = argv[++i]; continue; }
        if (a === '--dedup') { args.filters.dedup = argv[++i]; continue; }
        if (a === '--event') { args.filters.event = argv[++i]; continue; }
        if (a === '--kind') { args.filters.kind = argv[++i]; continue; }
        if (a === '--tick') { args.filters.tick = parseInt(argv[++i], 10); continue; }
        if (a === '--activity') { args.filters.activity = argv[++i]; continue; }
        if (a === '--since') { args.filters.since = Date.parse(argv[++i]); continue; }
        if (a === '--until') { args.filters.until = Date.parse(argv[++i]); continue; }
        console.error('unknown arg:', a);
        process.exit(2);
    }
    return args;
}

function defaultFiles() {
    const dir = path.resolve(__dirname, '..', '.logs');
    try {
        return fs.readdirSync(dir)
            .filter(f => /^notif-trace-\d{4}-\d{2}-\d{2}\.jsonl$/.test(f))
            .sort()
            .map(f => path.join(dir, f));
    } catch (err) {
        console.error(`cannot read ${ dir }: ${ err.message }`);
        return [];
    }
}

function matchEvent(ev, filters) {
    if (filters.tick !== undefined && ev.tick !== filters.tick) return false;
    if (filters.kind && ev.kind !== filters.kind) return false;
    if (filters.since && ev.t < filters.since) return false;
    if (filters.until && ev.t > filters.until) return false;

    if (filters.room) {
        const rooms = collectField(ev, 'room');
        if (!rooms.includes(filters.room)) return false;
    }
    if (filters.commit) {
        const shas = collectShas(ev);
        if (!shas.some(s => s && s.startsWith(filters.commit))) return false;
    }
    if (filters.dedup) {
        const keys = collectField(ev, 'dedup_key');
        if (!keys.includes(filters.dedup)) return false;
    }
    if (filters.event) {
        const types = collectField(ev, 'event_type');
        if (!types.includes(filters.event)) return false;
    }
    if (filters.activity) {
        const ids = collectField(ev, 'activity_id');
        if (!ids.includes(filters.activity)) return false;
    }
    return true;
}

// Walk the event object and collect every value at any nesting depth keyed
// by `field`. Lets a filter match both top-level fields and ones tucked
// inside nested snapshots (envelope.extra.event_type, recent.commit_sha, …).
function collectField(node, field, out = []) {
    if (!node || typeof node !== 'object') return out;
    if (Array.isArray(node)) {
        for (const v of node) collectField(v, field, out);
        return out;
    }
    for (const [k, v] of Object.entries(node)) {
        if (k === field && (typeof v === 'string' || typeof v === 'number')) out.push(v);
        if (v && typeof v === 'object') collectField(v, field, out);
    }
    return out;
}

function collectShas(ev) {
    const out = [];
    out.push(...collectField(ev, 'commit_sha'));
    out.push(...collectField(ev, '_commit_sha'));
    // dedup_key carries the SHA inline for the `github:commit:{sha}` form, so
    // events that only carry the dedup_key (send_ok, edit_merge, recent_saved)
    // still match a --commit filter against the embedded SHA.
    for (const k of collectField(ev, 'dedup_key')) {
        const m = String(k).match(/^github:commit:([a-f0-9]+)/);
        if (m) out.push(m[1]);
    }
    return out;
}

function summarize(ev) {
    const t = new Date(ev.t).toISOString().slice(11, 23);
    const head = `${ t } tick=${ ev.tick }.${ ev.seq } ${ ev.kind }`;
    switch (ev.kind) {
    case 'drained':
        return `${ head } count=${ ev.count }`;
    case 'item_drained': {
        const e = ev.envelope || {};
        const x = e.extra || {};
        return `${ head } room=${ e.room } type=${ e.type } ev=${ x.event_type || '-' } dedup=${ x.dedup_key || '-' } sha=${ x._commit_sha || '-' }`;
    }
    case 'normalize':
        return `${ head } room=${ ev.room } ev=${ ev.event_type } ${ JSON.stringify(ev.before) } → ${ JSON.stringify(ev.after) }`;
    case 'filter_silent':
        return `${ head } room=${ ev.room } silent_drop ev=${ ev.event_type }`;
    case 'filter_redirect':
        return `${ head } ${ ev.from_room } → ${ ev.to_room } reason=${ ev.reason } ev=${ ev.event_type }`;
    case 'announce_redirect':
        return `${ head } ${ ev.from_room } → ${ ev.to_room } repo=${ ev.repo }`;
    case 'action_triggered_attribution':
        return `${ head } repo=${ ev.repo } own=${ ev.own_sha } → parent=${ ev.parent_sha }`;
    case 'wfactive_record':
        return `${ head } repo=${ ev.repo } sha=${ ev.commit_sha } ev=${ ev.event_type }`;
    case 'route':
        return `${ head } room=${ ev.room } batch=${ ev.batch_size } trackable=${ ev.trackable_count } coalescable=${ ev.coalescable_count }`;
    case 'recent_lookup':
        return `${ head } room=${ ev.room } dedup=${ ev.dedup_key } found=${ ev.found } source=${ ev.source || '-' } age=${ ev.age_ms }ms editable=${ ev.eligible_for_edit }`;
    case 'edit_merge': {
        const beforeItems = ev.before?.items?.length || 0;
        const afterItems = ev.after?.items?.length || 0;
        return `${ head } room=${ ev.room } dedup=${ ev.dedup_key } mode=${ ev.merge_mode } items: ${ beforeItems } → ${ afterItems } activity=${ ev.activity_id }`;
    }
    case 'edit_ok':
        return `${ head } room=${ ev.room } activity=${ ev.activity_id } appended=${ ev.appended_count }`;
    case 'edit_failed':
        return `${ head } room=${ ev.room } activity=${ ev.activity_id } error=${ ev.error }`;
    case 'edit_skipped_no_convref':
    case 'edit_fell_through':
        return `${ head } room=${ ev.room } dedup=${ ev.dedup_key }`;
    case 'send_attempt':
        return `${ head } room=${ ev.room } conv=${ ev.conversation_id?.slice(-12) } dedup=${ ev.dedup_key || '-' }`;
    case 'send_ok':
        return `${ head } room=${ ev.room } activity=${ ev.activity_id } dedup=${ ev.dedup_key || '-' }`;
    case 'send_failed':
        return `${ head } room=${ ev.room } error=${ ev.error }`;
    case 'coalesce_text_attempt':
        return `${ head } room=${ ev.room } items=${ ev.items_in } bytes=${ ev.bytes }`;
    case 'coalesce_text_ok':
        return `${ head } room=${ ev.room } items=${ ev.items_in }`;
    case 'coalesce_text_failed':
        return `${ head } room=${ ev.room } error=${ ev.error }`;
    case 'recent_saved':
        return `${ head } room=${ ev.room } dedup=${ ev.dedup_key } activity=${ ev.activity_id } items=${ ev.item_count } appended=${ ev.appended_count }`;
    case 'fallback_attempt':
    case 'fallback_ok':
    case 'fallback_failed':
    case 'fallback_abandoned':
        return `${ head } room=${ ev.room } host=${ ev.url_host || '-' } reason=${ ev.reason || '-' } ${ ev.error || '' }`;
    case 'dead_lettered':
        return `${ head } room=${ ev.room } reason=${ ev.reason } dedup=${ ev.dedup_key || '-' }`;
    case 'expired':
        return `${ head } room=${ ev.room } expires_at=${ ev.expires_at } now=${ ev.now_sec }`;
    case 'tick_end':
        return `${ head } ms=${ ev.ms } ${ JSON.stringify(ev.stats) }`;
    default:
        return `${ head } ${ JSON.stringify(ev).slice(0, 200) }`;
    }
}

async function readEvents(files, filters) {
    const matched = [];
    for (const file of files) {
        let stream;
        try { stream = fs.createReadStream(file, 'utf8'); } catch (err) {
            console.error(`skip ${ file }: ${ err.message }`);
            continue;
        }
        const rl = readline.createInterface({ input: stream, crlfDelay: Infinity });
        for await (const line of rl) {
            if (!line.trim()) continue;
            let ev;
            try { ev = JSON.parse(line); } catch (_) { continue; }
            if (matchEvent(ev, filters)) matched.push(ev);
        }
    }
    matched.sort((a, b) => a.t - b.t || a.tick - b.tick || a.seq - b.seq);
    return matched;
}

function timeline(events) {
    for (const ev of events) console.log(summarize(ev));
}

function grouped(events) {
    // Bucket by dedup_key (preferred) or commit_sha; events with neither go
    // into "unkeyed" so we still surface them.
    const buckets = new Map();
    for (const ev of events) {
        const keys = collectField(ev, 'dedup_key').filter(Boolean);
        const shas = collectShas(ev).filter(Boolean);
        const key = keys[0] || (shas[0] ? `sha:${ shas[0] }` : null) || 'unkeyed';
        if (!buckets.has(key)) buckets.set(key, []);
        buckets.get(key).push(ev);
    }
    for (const [key, list] of buckets) {
        console.log('');
        console.log(`── ${ key }  (${ list.length } events) ───`);
        for (const ev of list) console.log('  ' + summarize(ev));
    }
}

function raw(events) {
    for (const ev of events) console.log(JSON.stringify(ev));
}

(async () => {
    const args = parseArgs(process.argv);
    if (args.help) {
        // Print the leading `//` comment block as the help text.
        const self = fs.readFileSync(__filename, 'utf8');
        const lines = self.split('\n');
        const help = [];
        for (const line of lines) {
            if (line.startsWith('//')) help.push(line.replace(/^\/\/ ?/, ''));
            else if (line.startsWith('#!') || line.startsWith('/*')) continue;
            else if (help.length > 0) break;
        }
        process.stdout.write(help.join('\n') + '\n');
        return;
    }
    const files = args.files.length ? args.files : defaultFiles();
    if (!files.length) {
        console.error('no log files found (looked in .logs/notif-trace-*.jsonl). Pass --file PATH.');
        process.exit(1);
    }
    const events = await readEvents(files, args.filters);
    if (events.length === 0) {
        console.error('no events matched.');
        return;
    }
    switch (args.mode) {
    case 'timeline': timeline(events); break;
    case 'grouped': grouped(events); break;
    case 'raw': raw(events); break;
    default:
        console.error('unknown --mode:', args.mode);
        process.exit(2);
    }
})().catch(err => {
    console.error(err);
    process.exit(1);
});
