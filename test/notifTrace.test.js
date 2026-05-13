const { describe, it, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

describe('notifTrace', () => {
    const tmpFile = path.join(os.tmpdir(), `notif-trace-test-${ process.pid }.jsonl`);
    let trace;

    before(() => {
        try { fs.unlinkSync(tmpFile); } catch (_) { /* ok */ }
        process.env.NOTIF_TRACE_LOG = tmpFile;
        // Module path must be required fresh so it picks up the env var.
        const modPath = require.resolve('../server/queue/notifTrace');
        delete require.cache[modPath];
        trace = require('../server/queue/notifTrace');
    });

    after(() => {
        try { fs.unlinkSync(tmpFile); } catch (_) { /* ok */ }
        delete process.env.NOTIF_TRACE_LOG;
    });

    it('writes one JSON line per emit with tick + seq + kind', async () => {
        trace.newTick();
        trace.emit('drained', { count: 2 });
        trace.emit('item_drained', { room: 'int-dev' });
        await trace.flush();

        const lines = fs.readFileSync(tmpFile, 'utf8').trim().split('\n');
        assert.equal(lines.length, 2);
        const a = JSON.parse(lines[0]);
        const b = JSON.parse(lines[1]);
        assert.equal(a.kind, 'drained');
        assert.equal(a.count, 2);
        assert.equal(a.tick, 1);
        assert.equal(a.seq, 1);
        assert.equal(b.kind, 'item_drained');
        assert.equal(b.room, 'int-dev');
        assert.equal(b.tick, 1);
        assert.equal(b.seq, 2);
        assert.ok(a.iso);
        assert.ok(typeof a.t === 'number');
    });

    it('snapshotEnvelope preserves extra payload for replay', () => {
        const env = {
            type: 'msg',
            room: 'int-dev',
            message: 'hi',
            extra: {
                event_type: 'push',
                dedup_key: 'github:commit:abc1234',
                _commit_sha: 'abc1234',
                data: { commits: [{ id: 'abc1234', message: 'fix' }] }
            }
        };
        const snap = trace.snapshotEnvelope(env);
        assert.equal(snap.type, 'msg');
        assert.equal(snap.room, 'int-dev');
        assert.deepEqual(snap.extra.data.commits, env.extra.data.commits);
    });

    it('snapshotRecent preserves header / items / commit_sha', () => {
        const recent = {
            activityId: 'act-1',
            ts: 123,
            type: 'msg',
            text: 'rendered',
            header: 'header',
            items: [{ identity: 'x', text: 'bullet' }],
            header_identity: 'h-id',
            appended_count: 1,
            conversationId: 'conv-1',
            commit_sha: 'abc1234'
        };
        const snap = trace.snapshotRecent(recent);
        assert.deepEqual(snap, recent);
        assert.equal(trace.snapshotRecent(null), null);
    });

    it('disabled trace produces no file output', async () => {
        const disabledFile = tmpFile + '.disabled';
        try { fs.unlinkSync(disabledFile); } catch (_) { /* ok */ }
        process.env.NOTIF_TRACE_LOG = '0';
        const modPath = require.resolve('../server/queue/notifTrace');
        delete require.cache[modPath];
        const traceOff = require('../server/queue/notifTrace');
        traceOff.newTick();
        traceOff.emit('drained', { count: 1 });
        await traceOff.flush();
        assert.equal(traceOff.isEnabled(), false);
        assert.equal(traceOff.currentPath(), null);
        // Re-enable for following tests
        process.env.NOTIF_TRACE_LOG = tmpFile;
        delete require.cache[modPath];
        trace = require('../server/queue/notifTrace');
    });
});
