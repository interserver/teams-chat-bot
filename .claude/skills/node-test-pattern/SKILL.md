---
name: node-test-pattern
description: Writes node:test tests for the Teams Chat Bot. Use when user says 'add test', 'write a node:test', 'mock the consumer', 'test the queue', or creates files in test/*.test.js. Provides the canonical mock setup: `_setInternalsForTest({ redis, redisBot, adapter })` seam, hash-map-backed Redis mock with hget/hset/incr/pipeline/expire/lrem, adapter mock recording _calls[] with sentActivities/updatedActivities, NOTIF_TRACE_LOG=0 silencing, and beforeEach/after restoration. Do NOT use for Jest, Mocha, or Codeception — this project is strictly `node:test`.
paths:
  - test/*.test.js
  - test/**/*.test.js
---
# node:test Pattern

## Critical

- The project uses `node:test` ONLY. NEVER `require('mocha')`, `require('jest')`, or import `chai`. Use `const { describe, it, beforeEach, after } = require('node:test')` and `const assert = require('node:assert/strict')`.
- Tests live in `test/*.test.js` and are run via `npm test` (`node --test 'test/*.test.js'`). Do NOT create `*.spec.js`, `__tests__/`, or nested test dirs — the npm script glob will not pick them up.
- BEFORE requiring `server/queue/notificationConsumer.js` in any test, set `process.env.NOTIF_TRACE_LOG = '0'` at module top. Without it, the consumer's startup code writes JSONL trace files to `.logs/` and pollutes the workspace.
- When swapping internals on the notification consumer, ALWAYS capture the return value of `_setInternalsForTest(...)` and restore it in `after(() => _setInternalsForTest(originals || {}))`. Forgetting this leaks state into the next test file.
- ESLint (flat config, `eslint.config.js`) enforces 4-space indent, semicolons, single quotes, and `template-curly-spacing: always` (`${ var }` with spaces). Match this exactly or `npm run lint` will fail.

## Instructions

1. **Create the test file** at `test/<feature>.test.js` (kebab-case-ish; existing files use camelCase like `notifBatchMerge.test.js` — match the neighbouring style). The npm test glob is `test/*.test.js`, so it must end in `.test.js` and live directly under `test/`.

2. **Header boilerplate** — every test file starts with:
   ```js
   const { describe, it, beforeEach, after } = require('node:test');
   const assert = require('node:assert/strict');

   process.env.NOTIF_TRACE_LOG = '0';
   ```
   Only import the `node:test` symbols you actually use. Add `before`/`afterEach` only when needed.

3. **Pure-function tests come first, before any mock setup.** Group them with `describe(name, () => { it(case, () => { ... }) })`. Pure tests should require the function directly from its module — no env vars, no mocks. Example pattern from `test/notifBatchMerge.test.js`:
   ```js
   const { groupTrackableByDedup, canBatchMergeGroup } = require('../server/queue/notificationConsumer');

   describe('groupTrackableByDedup', () => {
       it('groups items by dedup_key preserving order', () => {
           const groups = groupTrackableByDedup([
               { env: { extra: { dedup_key: 'a' } } },
               { env: { extra: { dedup_key: 'b' } } }
           ]);
           assert.equal(groups.size, 2);
       });
   });
   ```
   Verify: pure tests do not call `_setInternalsForTest` and have no `beforeEach`.

4. **For integration tests that touch the notification consumer**, define three mock factories at module scope (copy verbatim from `test/notifBatchMerge.test.js` lines 22–104):

   ```js
   function makeRedisMock() {
       const hashes = new Map();
       const lists = new Map();
       return {
           _hashes: hashes,
           _lists: lists,
           async hget(key, field) {
               const h = hashes.get(key);
               return h ? (h[field] || null) : null;
           },
           async hset(key, field, value) {
               const h = hashes.get(key) || {};
               h[field] = value;
               hashes.set(key, h);
               return 1;
           },
           async incr() { return 1; },
           async get() { return null; },
           async expire() { return 1; },
           async lrem(key, count, value) {
               const arr = lists.get(key) || [];
               const idx = arr.indexOf(value);
               if (idx >= 0) arr.splice(idx, 1);
               lists.set(key, arr);
               return idx >= 0 ? 1 : 0;
           },
           pipeline() {
               const ops = [];
               const self = this;
               const chain = {
                   hset(key, field, value) { ops.push({ op: 'hset', key, field, value }); return chain; },
                   expire() { return chain; },
                   zadd() { return chain; },
                   zremrangebyscore() { return chain; },
                   async exec() {
                       for (const o of ops) {
                           if (o.op === 'hset') {
                               const h = self._hashes.get(o.key) || {};
                               h[o.field] = o.value;
                               self._hashes.set(o.key, h);
                           }
                       }
                       return [];
                   }
               };
               return chain;
           }
       };
   }
   ```
   Expose `_hashes` and `_lists` so assertions can read them directly. Verify: `redisMock._hashes` is a `Map` and `redisMock.pipeline()` returns a chainable object with `exec()`.

5. **Bot Redis mock** (the convref store) returns JSON-stringified convrefs from `convref:*` keys:
   ```js
   function makeBotRedisMock(convrefs = {}) {
       return {
           async get(key) {
               const m = key.match(/^convref:(.+)$/);
               if (!m) return null;
               const v = convrefs[m[1]];
               return v == null ? null : JSON.stringify(v);
           },
           async incr() { return 1; }
       };
   }
   ```
   Pass `{ 'int-dev-private': { conversation: { id: '19:...' }, serviceUrl: 'https://smba...' } }` to seed convrefs. Pass `{}` to force the constructed-convref fallback path.

6. **Adapter mock** records every `continueConversation` call and the activities sent/updated within its callback:
   ```js
   function makeAdapterMock() {
       const calls = [];
       return {
           _calls: calls,
           async continueConversation(conversationRef, callback) {
               const record = { conversationRef, sentActivities: [], updatedActivities: [] };
               calls.push(record);
               const proactiveContext = {
                   async sendActivity(activity) {
                       record.sentActivities.push(activity);
                       return { id: 'new-activity-' + calls.length };
                   },
                   async updateActivity(activity) {
                       record.updatedActivities.push(activity);
                       return null;
                   }
               };
               await callback(proactiveContext);
           }
       };
   }
   ```
   To simulate edit failure, reassign `adapterMock.continueConversation` in the specific test with a version that throws on `updateActivity` (see `test/notifBatchMerge.test.js` line 314–327).

7. **Wire mocks via `_setInternalsForTest`** inside `describe()`:
   ```js
   let originals;
   let redisMock, botRedisMock, adapterMock;

   beforeEach(() => {
       redisMock = makeRedisMock();
       botRedisMock = makeBotRedisMock({});
       adapterMock = makeAdapterMock();
       redisMock._lists.set('notif:processing', []);
       originals = _setInternalsForTest({
           redis: redisMock,
           redisBot: botRedisMock,
           adapter: adapterMock
       });
   });

   after(() => {
       _setInternalsForTest(originals || {});
   });
   ```
   `_setInternalsForTest` is destructured from the consumer: `const { _setInternalsForTest, handleTrackableBatch } = require('../server/queue/notificationConsumer');`. The keys it accepts are `redis`, `redisBot`, `adapter`. It returns the prior internals so they can be restored — capture them.

8. **Seed state via `redisMock._hashes` / `_lists` directly**, not via the public hset (faster, exact). To pre-populate a recent activity:
   ```js
   const key = `notif:recent:${ ROOM }`;
   redisMock._hashes.set(key, { [DEDUP]: JSON.stringify({ activityId: 'act-1', ts: Date.now() - 1000, /* ... */ }) });
   ```
   To make `ackOne()` succeed, push the raw envelope strings into `notif:processing`:
   ```js
   for (const it of group) redisMock._lists.get('notif:processing').push(it.raw);
   ```

9. **Assert against `adapterMock._calls`** — never against internal consumer state. The standard assertions are:
   ```js
   assert.equal(adapterMock._calls.length, 1, 'exactly one adapter call');
   assert.equal(adapterMock._calls[0].sentActivities.length, 1);
   assert.equal(adapterMock._calls[0].updatedActivities.length, 0);
   assert.equal(adapterMock._calls[0].updatedActivities[0].id, 'act-1');
   ```
   Also assert the stats object passed into the handler reflects the right outcome:
   ```js
   const stats = { sent: 0, edited: 0, coalesced: 0, fallback: 0, dead: 0, expired: 0, redirected: 0 };
   await handleTrackableBatch(ROOM, group, stats);
   assert.equal(stats.edited, 1);
   ```
   Read persisted state back from `redisMock._hashes` and `JSON.parse` it for value assertions.

10. **Use helper factories for envelopes** at the top of the file (after the mocks), one per event_type. Example from `test/notificationConsumer.test.js`:
    ```js
    function checkRunEnv(name, status, conclusion = '', message = CHECK_RUN_MSG, htmlUrl = 'https://example/cr') {
        return {
            type: 'msg',
            message,
            extra: {
                event_type: 'check_run',
                _commit_sha: '4fd48e4',
                dedup_key: 'github:commit:4fd48e4',
                data: { check_run: { name, status, conclusion, html_url: htmlUrl } }
            }
        };
    }
    ```
    Keep these pure (no closures over mocks) so they are reusable across describe blocks.

11. **Run the test suite** with `npm test`. To run a single file: `node --test test/<file>.test.js`. To run with `--test-only` filtering, mark with `it.only(...)`. Verify before claiming success: the final line should show `# pass <N>` and `# fail 0`.

12. **Lint the test file** with `npm run lint` before committing. ESLint flat config enforces 4-space indent, semicolons, single quotes, and `template-curly-spacing: always` — write `${ var }` not `${var}`.

## Examples

**User says:** "Add a test verifying that when the convref is missing, the consumer falls back to the constructed reference and still sends the activity."

**Actions taken:**
1. Create `test/notifConvrefFallback.test.js`.
2. Add the header: `node:test` imports, `assert/strict`, `process.env.NOTIF_TRACE_LOG = '0'`.
3. Require `handleTrackableBatch` and `_setInternalsForTest` from `../server/queue/notificationConsumer`.
4. Copy the three mock factories verbatim. Construct `makeBotRedisMock({})` (empty — forces fallback).
5. In a `describe('constructed convref fallback', () => { ... })`, set up `beforeEach`/`after` per Step 7.
6. Inside `it(...)`, build a trackable envelope, push its raw string into `notif:processing`, call `handleTrackableBatch`.
7. Assert: `adapterMock._calls.length === 1`, the call's `conversationRef.serviceUrl` matches `TEAMS_SERVICE_URL`, and `stats.sent === 1`.
8. Run `npm test` and confirm `# pass`.

**Result:** A new file shaped exactly like `notifBatchMerge.test.js`, that runs under `npm test` and validates the fallback path without hitting any real Redis/Teams endpoints.

## Common Issues

- **`Error: Cannot find module 'node:test'`** — Node version is below 18. Run `node --version`; this project requires ≥16.14.2 per Readme but `node:test` needs ≥18. Upgrade Node.
- **Trace files appearing in `.logs/` during tests** — You forgot `process.env.NOTIF_TRACE_LOG = '0'` before requiring the consumer. Move that line above any `require('../server/queue/notificationConsumer')`.
- **Test passes locally, fails in CI with `TypeError: redis.pipeline(...).zadd is not a function`** — Your `makeRedisMock` is out of date. Re-copy the full `pipeline()` chain from `test/notifBatchMerge.test.js` lines 48–67 — it must include `zadd`, `zremrangebyscore`, and `expire` as chainable no-ops.
- **Subsequent test files break with `Cannot read properties of undefined`** — `_setInternalsForTest` was not restored. Ensure your `describe` block has `after(() => _setInternalsForTest(originals || {}))`.
- **`assert.equal(adapterMock._calls.length, 1)` fails with `0`** — The handler short-circuited before reaching the adapter. Likely causes: the envelope is missing `extra.dedup_key`, the `room` value does not resolve through `CHANNELS` (use a known room like `int-dev-announce`), or the raw entry was not pushed into `notif:processing` so `ackOne` blew up.
- **`npm test` reports `0 tests`** — File is not in `test/` or does not end in `.test.js`. The npm script glob is `'test/*.test.js'`; nested directories like `test/queue/foo.test.js` are NOT picked up.
- **ESLint errors `Expected indentation of 4 spaces but found 2`** — Match the project's flat config in `eslint.config.js`: 4-space indent, single quotes, semicolons required, `${ var }` with internal spaces. Run `npm run lint` before committing.
- **`updateActivity` simulated failure does not register as a fallback** — `stats.edited` only increments on successful `updateActivity`. Assert `stats.edited === 0` and `stats.sent === 1` (the fresh send) when simulating edit failure; check `adapterMock._calls.reduce((n, c) => n + c.sentActivities.length, 0)` for the actual send count.