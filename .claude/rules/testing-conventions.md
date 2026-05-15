---
paths:
  - test/**
---

# Test Conventions

- `node:test` only — `const { describe, it, before, beforeEach, after } = require('node:test');` + `const assert = require('node:assert/strict');`. No Jest, no Mocha.
- `NOTIF_TRACE_LOG = '0'` MUST be set at the top of any test that requires `server/queue/notificationConsumer.js` to silence trace file writes.
- Consumer/adapter mock seam: `consumer._setInternalsForTest({ redis, redisBot, adapter })` returns the originals; restore in `after()`. Canonical mock shape: `test/notifBatchMerge.test.js`, `test/notifConvrefFallback.test.js`.
- Mock Redis MUST support `get`/`set`/`hget`/`hset`/`hgetall`/`incr`/`pipeline`/`expire`/`lrem` for the batch/recent path. Seed `_lists.set('notif:processing', [...rawValues])` and assert empty after the tick.
- Adapter mock records `_calls[]` with `{ conversationRef, sentActivities, updatedActivities }`. Each `continueConversation(ref, cb)` builds a fake context that pushes into the record.
- Add a `match()` assertion in `test/commands.test.js` for every new command regex (one `describe` per module, multiple `it` per branch).
- ⚠️ `test/botActivityHandler.test.js` hangs at end of run (unclosed Mongo/MySQL/Redis). When running the full suite, list other files explicitly rather than relying on `test/*.test.js` glob.
- Run a single file in foreground for reliable output: `node --test test/<file>.test.js 2>&1 | tail -25`.
- Test data env vars: `beforeEach` saves `process.env[KEY]`, overrides; `afterEach` restores (see `test/validateEnv.test.js`).
