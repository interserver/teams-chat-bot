// Serialized-size guards for outbound Teams activities.
//
// Microsoft Teams rejects any activity whose serialized JSON exceeds
// ~25 KB with `RestError: Message size too large` — and because the SDK
// routes that error into `adapter.onTurnError` (which swallows it), the
// send looks successful to the caller while the message is silently lost.
// The only reliable defense is to measure UTF-8 byte length of the
// serialized activity BEFORE handing it to the adapter. Every send/update
// path should funnel through a helper here.
//
// JS string `.length` is UTF-16 code units, not bytes — emoji cost 4 UTF-8
// bytes for 2 code units and JSON escaping inflates further — so all
// measurements in this module are `Buffer.byteLength(..., 'utf8')`.

const ADAPTIVE_CARD_CONTENT_TYPE = 'application/vnd.microsoft.card.adaptive';
// Headroom under Teams' ~25 KB limit. The limit applies to the FULL wire
// envelope: continueConversation stamps on from/conversation/recipient
// fields (~300-700 B) that activityBytes never sees, so this guard only
// measures the caller-built shape. Do not tune above ~24.8K or those
// invisible fields can push a "fitting" activity back over the real cap.
const DEFAULT_MAX_ACTIVITY_BYTES = 24_000;
const TRUNCATION_MARKER = '\n\n✂ …truncated';

function parseMaxActivityBytes(raw) {
    const parsed = parseInt(raw, 10);
    if (!Number.isFinite(parsed) || parsed <= 0) return DEFAULT_MAX_ACTIVITY_BYTES;
    return parsed;
}

const MAX_ACTIVITY_BYTES = parseMaxActivityBytes(process.env.TEAMS_MAX_ACTIVITY_BYTES);

// ---------------------------------------------------------------------------
// Measurement
// ---------------------------------------------------------------------------

// Byte length of the activity once serialized to JSON — exactly what
// Bot Framework puts on the wire and what Teams measures.
function activityBytes(activity) {
    let json;
    try {
        json = JSON.stringify(activity);
    } catch (err) {
        throw new Error(`activityBytes: activity is not JSON-serializable (${ err.message })`, { cause: err });
    }
    if (typeof json !== 'string') {
        throw new Error('activityBytes: activity serialized to nothing — refusing to guess a size.');
    }
    return Buffer.byteLength(json, 'utf8');
}

// Fixed cost of the activity envelope around a text body, so callers can
// budget the text itself against (maxBytes - overhead). Includes the two
// quote characters around the text value, which `textJsonBytes` omits —
// `activityBytes({ type:'message', text }) === overhead + textJsonBytes(text)`.
function textActivityOverhead() {
    return activityBytes({ type: 'message', text: '' });
}

// ---------------------------------------------------------------------------
// Text truncation (byte-aware, surrogate-pair safe)
// ---------------------------------------------------------------------------

// Byte cost of the value once embedded in JSON, minus its surrounding
// quotes. JSON escaping is char-wise, so these costs are additive under
// concatenation — unlike raw Buffer.byteLength, which lies about \n, ", \,
// and lone surrogates (escaped to 6 ASCII bytes) and would let the
// "truncated" result still overflow the real activity budget.
function textJsonBytes(value) {
    return Buffer.byteLength(JSON.stringify(String(value ?? '')), 'utf8') - 2;
}

function isLoneHighSurrogateEnd(str, cut) {
    if (cut <= 0) return false;
    const prev = str.charCodeAt(cut - 1);
    return prev >= 0xD800 && prev <= 0xDBFF;
}

// First maxChars code units of str, backing off one unit when the cut
// would land between the halves of a surrogate pair.
function charCut(str, maxChars) {
    if (str.length <= maxChars) return str;
    const cut = isLoneHighSurrogateEnd(str, maxChars) ? maxChars - 1 : maxChars;
    return str.slice(0, cut);
}

// Largest code-unit index such that str.slice(0, index) serializes within
// maxBytes. Never splits a UTF-16 surrogate pair.
function safeCutIndex(str, maxBytes) {
    if (textJsonBytes(str) <= maxBytes) return str.length;
    let low = 0;
    let high = str.length;
    while (low < high) {
        const mid = Math.ceil((low + high) / 2);
        if (textJsonBytes(str.slice(0, mid)) <= maxBytes) low = mid;
        else high = mid - 1;
    }
    if (isLoneHighSurrogateEnd(str, low)) low -= 1;
    return low;
}

// Truncate `text` so its serialized size fits maxBytes, appending `marker`
// when there is room for it. Text that already fits is returned unchanged.
function truncateTextToBytes(text, maxBytes, marker = TRUNCATION_MARKER) {
    const value = String(text ?? '');
    if (textJsonBytes(value) <= maxBytes) return value;
    if (maxBytes <= 0) return '';
    const markerBytes = textJsonBytes(marker);
    const roomForMarker = markerBytes < maxBytes;
    const budget = roomForMarker ? maxBytes - markerBytes : maxBytes;
    const cut = safeCutIndex(value, budget);
    return value.slice(0, cut) + (roomForMarker ? marker : '');
}

// Fit `text` into a { type:'message', text } activity of at most maxBytes
// serialized. Returns the (possibly truncated) text.
function fitTextActivity(text, maxBytes = MAX_ACTIVITY_BYTES) {
    return truncateTextToBytes(text, maxBytes - textActivityOverhead());
}

// ---------------------------------------------------------------------------
// Adaptive-card trimming
// ---------------------------------------------------------------------------

// Recursive in-place trim of a parsed Adaptive Card: shorten long text and
// speak fields, drop base64 image payloads, remove empty arrays. Ported from
// dailyRecapController's original limiter; bounded recursion depth keeps it
// safe on pathological payloads.
function trimCardInPlace(node, depth) {
    if (depth > 20) return;
    if (!node || typeof node !== 'object') return;

    if (Array.isArray(node)) {
        for (const item of node) trimCardInPlace(item, depth + 1);
        return;
    }

    if (typeof node.text === 'string' && node.text.length > 2000) {
        node.text = `${ charCut(node.text, 2000) }…`;
    }
    if (typeof node.speak === 'string' && node.speak.length > 1000) {
        node.speak = `${ charCut(node.speak, 1000) }…`;
    }
    if (typeof node.url === 'string' && node.url.startsWith('data:')) node.url = '';
    if (typeof node.image === 'string' && node.image.startsWith('data:')) node.image = '';

    for (const [key, val] of Object.entries(node)) {
        if (Array.isArray(val) && val.length === 0) delete node[key];
        else trimCardInPlace(val, depth + 1);
    }
}

// Return a trimmed clone of the card (input is never mutated).
function trimCard(card) {
    const clone = JSON.parse(JSON.stringify(card));
    trimCardInPlace(clone, 0);
    return clone;
}

function buildCardAttachmentsActivity(cards, text) {
    const activity = {
        type: 'message',
        attachments: cards.map(card => ({ contentType: ADAPTIVE_CARD_CONTENT_TYPE, content: card }))
    };
    if (text) activity.text = String(text);
    return activity;
}

// Trim a card once and measure the resulting attachment activity —
// `fits` reports whether that single pass was enough; it does not shrink
// iteratively. Returns { card, fits, bytes }.
function fitCardActivity(card, maxBytes = MAX_ACTIVITY_BYTES) {
    const trimmed = trimCard(card);
    const bytes = activityBytes(buildCardAttachmentsActivity([trimmed], ''));
    return { card: trimmed, fits: bytes <= maxBytes, bytes };
}

// Greedy attachment fit: keep attachments from the front while the whole
// activity (with its optional text) stays under maxBytes, trimming each
// adaptive card first. Anything that no longer fits is dropped and counted
// in a human-readable note the caller should surface to the channel.
// Measurements reserve headroom so prepending the note cannot re-overflow.
const NOTE_RESERVE_BYTES = 160;

function fitAttachmentsActivity(attachments, text = '', maxBytes = MAX_ACTIVITY_BYTES) {
    const list = Array.isArray(attachments) ? attachments : [];
    const kept = [];
    let dropped = 0;

    for (const attachment of list) {
        const candidate = isAdaptiveAttachment(attachment)
            ? { ...attachment, content: trimCard(attachment.content) }
            : attachment;
        const bytes = activityBytes({ type: 'message', text: String(text ?? ''), attachments: [...kept, candidate] });
        if (bytes + NOTE_RESERVE_BYTES <= maxBytes) {
            kept.push(candidate);
            continue;
        }
        dropped += 1;
    }

    if (dropped === 0) return { attachments: kept, note: null };
    const note = `⚠️ ${ dropped } oversized card(s) were dropped to fit Teams' message size limit.`;
    return { attachments: kept, note };
}

function isAdaptiveAttachment(attachment) {
    return Boolean(attachment) &&
        attachment.contentType === ADAPTIVE_CARD_CONTENT_TYPE &&
        attachment.content && typeof attachment.content === 'object';
}

// ---------------------------------------------------------------------------
// Chunking (for paths that may send N messages instead of one)
// ---------------------------------------------------------------------------

// Split combined notification text into ordered chunks, each small enough
// to fit a { type:'message', text } activity under maxBytes. Breaks prefer
// `separator`, then paragraph breaks, then hard byte cuts. Content is
// never lost: concatenating the chunks reproduces the original text.
// Throws RangeError when maxBytes cannot even hold an empty text activity.
function splitTextIntoChunks(text, maxBytes = MAX_ACTIVITY_BYTES, separator = '\n\n') {
    const value = String(text ?? '');
    const budget = maxBytes - textActivityOverhead();
    if (budget <= 0) {
        throw new RangeError(`splitTextIntoChunks: maxBytes ${ maxBytes } is smaller than the ${ textActivityOverhead() } B message envelope — no chunk could ever fit`);
    }
    if (textJsonBytes(value) <= budget) return [value];
    const chunks = packSegments(attachJoiner(value.split(separator), separator), budget);
    return chunks.length > 0 ? chunks : [''];
}

// Re-attach each joiner to the segment before it so any pack/split walk
// downstream can concatenate its inputs losslessly.
function attachJoiner(parts, joiner) {
    return parts.map((part, i) => (i < parts.length - 1 ? `${ part }${ joiner }` : part));
}

// Greedy bin-packing of pre-joined segments into ≤budget-byte chunks.
// Oversized segments recurse into finer splits instead of being dropped.
function packSegments(segments, budget) {
    const chunks = [];
    let current = '';
    for (const segment of segments) {
        if (textJsonBytes(segment) > budget) {
            if (current.length > 0) chunks.push(current);
            current = '';
            chunks.push(...splitOversizedSegment(segment, budget));
            continue;
        }
        if (current.length === 0) {
            current = segment;
            continue;
        }
        if (textJsonBytes(current + segment) <= budget) {
            current += segment;
            continue;
        }
        chunks.push(current);
        current = segment;
    }
    if (current.length > 0) chunks.push(current);
    return chunks;
}

// A single segment too large for any chunk: split on paragraphs first, then
// lines, then hard byte cuts — always surrogate-safe, never lossy. Each
// recursive split MUST produce strictly shorter parts; a segment ending in
// its own separator would otherwise re-attach to reproduce itself and ping-
// pong between packSegments/splitOversizedSegment forever (C1 crash-loop).
function splitOversizedSegment(segment, budget) {
    for (const inner of ['\n\n', '\n']) {
        if (!segment.includes(inner)) continue;
        const parts = attachJoiner(segment.split(inner), inner);
        if (parts.every(part => part.length < segment.length)) return packSegments(parts, budget);
        return hardSplit(segment, budget); // trailing-separator degenerate split makes no progress
    }
    return hardSplit(segment, budget);
}

// Hard byte split as the last resort: each chunk is the largest surrogate-
// safe prefix that fits, so re-joining reproduces the input exactly.
function hardSplit(value, budget) {
    const out = [];
    let rest = value;
    while (textJsonBytes(rest) > budget) {
        let cut = safeCutIndex(rest, budget);
        if (cut === 0) cut = 1; // guarantee progress when one char alone exceeds the budget
        out.push(rest.slice(0, cut));
        rest = rest.slice(cut);
    }
    if (rest.length > 0) out.push(rest);
    return out;
}

module.exports = {
    MAX_ACTIVITY_BYTES,
    TRUNCATION_MARKER,
    ADAPTIVE_CARD_CONTENT_TYPE,
    activityBytes,
    textActivityOverhead,
    textJsonBytes,
    charCut,
    truncateTextToBytes,
    fitTextActivity,
    trimCard,
    trimCardInPlace,
    buildCardAttachmentsActivity,
    fitCardActivity,
    fitAttachmentsActivity,
    isAdaptiveAttachment,
    splitTextIntoChunks,
    hardSplit
};
