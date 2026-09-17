// Unit tests for the shared Teams activity-size guards.
// Teams measures the SERIALIZED activity in UTF-8 bytes; these tests pin
// byte math, surrogate-pair safety, lossless chunking, and card trimming.

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const {
    MAX_ACTIVITY_BYTES,
    TRUNCATION_MARKER,
    activityBytes,
    truncateTextToBytes,
    fitTextActivity,
    trimCard,
    fitCardActivity,
    fitAttachmentsActivity,
    splitTextIntoChunks
} = require('../server/lib/messageSize');

const LONE_HIGH_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])/;

describe('activityBytes', () => {
    it('counts UTF-8 bytes, not JS chars (emoji are multi-byte)', () => {
        const activity = { type: 'message', text: '🚀😀' };
        const json = JSON.stringify(activity);
        assert.equal(activityBytes(activity), Buffer.byteLength(json, 'utf8'));
        assert.ok(activityBytes(activity) > json.length, 'bytes must exceed chars for emoji text');
    });

    it('throws a clear error on unserializable activity', () => {
        const circular = { type: 'message' };
        circular.self = circular;
        assert.throws(() => activityBytes(circular), /not JSON-serializable/);
    });
});

describe('truncateTextToBytes', () => {
    it('returns text unchanged when it already fits', () => {
        assert.equal(truncateTextToBytes('hello', 100), 'hello');
    });

    it('appends the marker and respects the byte budget', () => {
        const out = truncateTextToBytes('x'.repeat(1000), 200);
        assert.ok(Buffer.byteLength(out, 'utf8') <= 200);
        assert.ok(out.endsWith(TRUNCATION_MARKER));
    });

    it('never splits a surrogate pair at the cut boundary', () => {
        // Budget chosen so the naive byte cut would land on the high
        // surrogate of an emoji; the result must back off, not leak a lone
        // surrogate (which would mangle Teams rendering and byte counts).
        const text = 'a😀b';
        const out = truncateTextToBytes(text, 4, '');
        assert.equal(out, 'a');
        const emojiRun = '🚀'.repeat(50);
        for (let max = 1; max < 40; max++) {
            const cut = truncateTextToBytes(emojiRun, max, '');
            assert.ok(Buffer.byteLength(cut, 'utf8') <= max, `fits at max=${ max }`);
            assert.ok(!LONE_HIGH_SURROGATE.test(cut), `no lone high surrogate at max=${ max }`);
        }
    });

    it('falls back to a plain byte cut when the marker itself does not fit', () => {
        const out = truncateTextToBytes('x'.repeat(100), 5);
        assert.ok(!out.includes('truncated'));
        assert.ok(Buffer.byteLength(out, 'utf8') <= 5);
    });
});

describe('fitTextActivity', () => {
    it('produces text whose wrapping activity fits the budget', () => {
        const out = fitTextActivity('🚀 '.repeat(20000), MAX_ACTIVITY_BYTES);
        assert.ok(activityBytes({ type: 'message', text: out }) <= MAX_ACTIVITY_BYTES);
    });
});

describe('splitTextIntoChunks', () => {
    const SEP = '\n\n— · —\n\n';

    it('returns a single chunk when the text already fits', () => {
        assert.deepEqual(splitTextIntoChunks('short', MAX_ACTIVITY_BYTES, SEP), ['short']);
    });

    it('prefers the separator and loses no content', () => {
        const pieces = Array.from({ length: 6 }, (_, i) => `🚀 ${ i } ` + '🚀'.repeat(4000));
        const text = pieces.join(SEP);
        const chunks = splitTextIntoChunks(text, MAX_ACTIVITY_BYTES, SEP);
        assert.ok(chunks.length > 1, 'must split an oversized join');
        assert.equal(chunks.join(''), text);
        for (const chunk of chunks) {
            assert.ok(activityBytes({ type: 'message', text: chunk }) <= MAX_ACTIVITY_BYTES);
        }
    });

    it('hard-splits unsplittable overflow without loss', () => {
        const text = 'a'.repeat(60000) + '😀'.repeat(10);
        const chunks = splitTextIntoChunks(text, MAX_ACTIVITY_BYTES, SEP);
        assert.equal(chunks.join(''), text);
        for (const chunk of chunks) {
            assert.ok(activityBytes({ type: 'message', text: chunk }) <= MAX_ACTIVITY_BYTES);
            assert.ok(!LONE_HIGH_SURROGATE.test(chunk));
        }
    });

    it('always returns at least one chunk', () => {
        assert.equal(splitTextIntoChunks('', MAX_ACTIVITY_BYTES, SEP).length, 1);
    });

    // C1 regressions: an oversized segment ENDING in its own separator made
    // attachJoiner re-produce the identical segment, ping-ponging
    // packSegments ↔ splitOversizedSegment until RangeError (stack overflow)
    // killed the consumer tick mid-sendCombinedText.
    it('terminates on an oversized piece that ends with the separator', () => {
        const text = 'q'.repeat(500) + SEP;
        const chunks = splitTextIntoChunks(text, 200, SEP);
        assert.ok(chunks.length >= 1);
        assert.equal(chunks.join(''), text, 'lossless');
        for (const chunk of chunks) {
            assert.ok(activityBytes({ type: 'message', text: chunk }) <= 200);
        }
    });

    it('terminates on an oversized piece containing no newline at all', () => {
        const text = `tiny${ SEP }${ 'w'.repeat(5000) }`;
        const chunks = splitTextIntoChunks(text, 300, SEP);
        assert.equal(chunks.join(''), text, 'lossless');
        for (const chunk of chunks) {
            assert.ok(activityBytes({ type: 'message', text: chunk }) <= 300);
            assert.ok(!LONE_HIGH_SURROGATE.test(chunk));
        }
    });

    it('terminates on a pathological mix of trailing separators and giant runs', () => {
        const text = 'a'.repeat(700) + SEP + 'tail line\n' + '😀'.repeat(400) + SEP + SEP + 'c'.repeat(700);
        const chunks = splitTextIntoChunks(text, 240, SEP);
        assert.ok(chunks.length > 2);
        assert.equal(chunks.join(''), text, 'lossless');
        for (const chunk of chunks) {
            assert.ok(activityBytes({ type: 'message', text: chunk }) <= 240);
            assert.ok(!LONE_HIGH_SURROGATE.test(chunk));
        }
    });

    it('throws RangeError when maxBytes cannot hold the envelope', () => {
        assert.throws(() => splitTextIntoChunks('hi', 10, SEP), RangeError);
    });
});

describe('trimCard', () => {
    it('truncates long text/speak fields, strips data: URLs and empty arrays', () => {
        const card = {
            type: 'AdaptiveCard',
            body: [
                { type: 'TextBlock', text: 'x'.repeat(5000) },
                { type: 'Image', url: 'data:image/png;base64,AAAA' },
                { type: 'Container', items: [] }
            ],
            speak: 'y'.repeat(3000)
        };
        const trimmed = trimCard(card);
        assert.equal(trimmed.body[0].text.length, 2001);
        assert.ok(trimmed.body[0].text.endsWith('…'));
        assert.equal(trimmed.body[1].url, '');
        assert.ok(!('items' in trimmed.body[2]), 'empty arrays are deleted');
        assert.equal(trimmed.speak.length, 1001);
        // input untouched (atomic predictability)
        assert.equal(card.body[0].text.length, 5000);
    });

    it('never leaves a lone high surrogate when the cap lands mid-emoji', () => {
        const text = 'a'.repeat(1999) + '😀'.repeat(50);
        const trimmed = trimCard({ type: 'AdaptiveCard', body: [{ type: 'TextBlock', text }] });
        assert.ok(!LONE_HIGH_SURROGATE.test(trimmed.body[0].text));
        assert.equal(trimmed.body[0].text.length, 2000); // 1999 kept + '…'
    });
});

describe('fitCardActivity', () => {
    it('reports fit against the serialized attachment activity', () => {
        const card = { type: 'AdaptiveCard', body: [{ type: 'TextBlock', text: 'small' }] };
        const { fits, bytes } = fitCardActivity(card, MAX_ACTIVITY_BYTES);
        assert.ok(fits);
        assert.ok(bytes <= MAX_ACTIVITY_BYTES);
    });
});

describe('fitAttachmentsActivity', () => {
    const adaptive = content => ({ contentType: 'application/vnd.microsoft.card.adaptive', content });
    const untrimmableHugeCard = () => ({
        type: 'AdaptiveCard',
        body: Array.from({ length: 300 }, (_, i) => ({ type: 'TextBlock', text: `block ${ i } ` + 'x'.repeat(1900) }))
    });

    it('keeps everything that fits and reports no note', () => {
        const { attachments, note } = fitAttachmentsActivity(
            [adaptive({ type: 'AdaptiveCard', body: [] }), adaptive({ type: 'AdaptiveCard', body: [] })],
            'hello', MAX_ACTIVITY_BYTES
        );
        assert.equal(attachments.length, 2);
        assert.equal(note, null);
    });

    it('drops trailing oversized cards and returns a note', () => {
        const { attachments, note } = fitAttachmentsActivity(
            [adaptive({ type: 'AdaptiveCard', body: [] }), adaptive(untrimmableHugeCard())],
            'hello', MAX_ACTIVITY_BYTES
        );
        assert.equal(attachments.length, 1);
        assert.ok(note && note.includes('1 oversized card(s) were dropped'));
    });

    it('drops a leading oversized card but keeps later ones that fit', () => {
        const { attachments, note } = fitAttachmentsActivity(
            [adaptive(untrimmableHugeCard()), adaptive({ type: 'AdaptiveCard', body: [] })],
            '', MAX_ACTIVITY_BYTES
        );
        assert.equal(attachments.length, 1);
        assert.ok(note.includes('1 oversized'));
    });
});
