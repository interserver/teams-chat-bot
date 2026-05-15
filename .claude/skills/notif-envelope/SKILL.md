---
name: notif-envelope
description: Builds and enqueues notification envelopes onto Redis `notif:queue` in the canonical v=1 shape `{ v, id, ts, expires_at, room, type:'msg', message|card, extra:{ dedup_key, event_type, repo, _commit_sha, data }, fallback_webhook_url? }`. Covers channel resolution via `CHANNELS`, GitHub `dedup_key` auto-injection (`github:commit:{sha7}` / `github:pr:{repo}:{n}` / `github:issue:{repo}:{n}`), edit-window coalescing (`NOTIF_EDIT_WINDOW_MS`, 30 min), trackable batch-merge via `handleTrackableBatch`, announce-redirect via `decideAnnounceRedirect`, and dead-letter on `notif:dead`. Use when user says 'enqueue notification', 'queue a notification', 'add webhook event', 'send to notif queue', or modifies `server/queue/notificationConsumer.js`, `server/queue/filters.js`, `server/queue/channels.js`, or an external producer. Do NOT use for synchronous proactive sends via `POST /api/message` (those bypass the queue) or for daily-recap cards (`POST /api/dailyrecap` has its own pipeline).
paths:
  - server/queue/**/*.js
  - test/notif*.test.js
---
# Notification Envelope Skill

Produce and enqueue notification envelopes onto Redis `notif:queue` exactly as the canonical PHP producers (`MyAdmin\Notifications\Queue`, `NotificationQueue`) and `server/commands/notifAdmin.js` do today. The consumer in `server/queue/notificationConsumer.js` decides edit/coalesce/new — your job is to produce a correct envelope and `LPUSH` it.

## Critical

- `room` MUST resolve via `CHANNELS` in `server/queue/channels.js` — never write `19:*@thread.v2` IDs into envelopes. Validate with `resolve(room)` before enqueueing; unknown rooms → bail.
- Redis key prefix is `process.env.NOTIF_KEY_PREFIX || 'notif:'`. Always go through `k(name)` (i.e. `KEY_PREFIX + name`); never hardcode `'notif:queue'` strings.
- Push with `redis.lpush(k('queue'), JSON.stringify(envelope))`. Producers LPUSH; the consumer `RPOPLPUSH`es into `notif:processing`.
- The producer Redis is the **InterServer Redis** (`createNotifRedis()` / `getNotifRedis()` from `server/queue/notificationConsumer.js`), NOT the bot's primary Redis (Dragonfly, used for `convref:*`). Don't mix them.
- `extra.dedup_key` makes an envelope **trackable** — it edits the most recent matching activity within `NOTIF_EDIT_WINDOW_MS` (default 30 min). Without it, the envelope is coalescable (2–8 items merged with `\n\n— · —\n\n`).
- For GitHub events the consumer's `normalizeGithubDedup()` overwrites `dedup_key` based on `event_type` + `_commit_sha`. If you set a custom GitHub `dedup_key` it WILL be replaced — let the consumer derive it.
- Never emit `check_suite` or `workflow_run` events expecting them to render — `server/queue/filters.js` silently drops them as aggregates. Use individual `check_run` / `workflow_job` events.
- Sanitize `message` to ≤4000 chars (Teams message limit). Adaptive Cards must stay under 25 KB or the consumer trims/rejects.
- Failures dead-letter to `notif:dead` with `_dead_reason` + `_dead_at`. Drain via `!notif drain-dead`.

## Instructions

### Step 1 — Resolve the channel

Import `resolve` and use it as a guard:

```js
const { resolve: resolveChannel } = require('../queue/channels');
if (!resolveChannel(room)) throw new Error(`unknown room "${ room }"`);
```

Known rooms (see `server/queue/channels.js`): `notifications`, `bot-testing`, `int-dev`, `int-dev-private`, `int-development`, `int-hw`, `hardware`, `general`, `development`, `int-dev-announce`, `interserver.net`. **Add new rooms to `CHANNELS` first** — do not invent room names in producers.

Verify: `resolveChannel(room)` returns a non-null `19:*@thread.v2` ID before proceeding to Step 2.

### Step 2 — Get the notif Redis client

From inside the bot process (e.g. a command module):

```js
const { getNotifRedis } = require('../queue/notificationConsumer');
const notifRedis = getNotifRedis();
if (!notifRedis) {
    await context.sendActivity(MessageFactory.text('notif consumer is not running'));
    return;
}
```

From an external script/producer, instantiate directly:

```js
const { createNotifRedis } = require('../lib/redis');
const notifRedis = createNotifRedis();
```

Never reuse `deps.redis` (bot Redis) — that holds `convref:*` and lives on a different host (`REDIS_HOST_MY`).

### Step 3 — Build the envelope

Copy this shape verbatim from `server/commands/notifAdmin.js#testCmd`:

```js
const envelope = {
    v: 1,
    id: 'src-' + Date.now(),                     // unique id, your prefix
    ts: Math.floor(Date.now() / 1000),
    expires_at: Math.floor(Date.now() / 1000) + 300,   // seconds-since-epoch
    room: room.trim(),                            // friendly name, NOT conv ID
    type: 'msg',                                  // 'msg' | 'card'
    message: String(text).slice(0, 4000),         // markdown for type='msg'
    card: null,                                   // Adaptive Card JSON for type='card'
    extra: {
        dedup_key: 'src:scope:key',               // omit to make coalescable
        level: 'info',                            // info|warn|error (advisory)
        source: 'commands/yourModule'
    },
    fallback_webhook_url: null                    // optional Power Automate URL
};
```

Rules:

- `type: 'card'` → set `card` to the Adaptive Card JSON, leave `message` null.
- `extra.dedup_key` namespaces: use `admin:*` (admin probes), `monitor:{check}:{host}`, `ticket:{id}`, `pr:{repo}:{n}`. The `github:*` namespace is reserved — the consumer owns it.
- `extra.event_type`, `extra.repo`, `extra.data` — only set on GitHub envelopes; the consumer reads them for normalization and PR-context attachment (`attachPrContext` in `server/queue/notificationConsumer.js`).
- `fallback_webhook_url`: if the Bot Framework `continueConversation` fails after retries, the consumer POSTs `{ text }` to this URL. Useful for Power Automate flows.
- `_commit_sha`, `_original_commit_sha`, `original_room`, `filtered`, `_dead_reason`, `_dead_at` are **internal fields** — do not set them in producers.

Verify: the envelope serializes via `JSON.stringify(envelope)` without throwing, and `JSON.parse(JSON.stringify(envelope)).room` resolves through `CHANNELS`.

### Step 4 — GitHub envelopes: let the consumer derive `dedup_key`

For GitHub events, set `extra.event_type`, `extra.repo`, and `extra.data` (the raw webhook payload). Do NOT set `dedup_key` for these event types — `normalizeGithubDedup()` in `notificationConsumer.js:548` will overwrite it:

| Event type | Auto-injected `dedup_key` |
|---|---|
| `push`, `check_run`, `workflow_job`, `check_suite`, `workflow_run`, `status`, `commit_comment`, `deployment`, `deployment_status` | `github:commit:{sha7}` |
| `pull_request`, `pull_request_review`, `pull_request_review_comment` (with PR #) | `github:pr:{repo}:{n}` |
| `issue_comment` on a PR (via `issue.html_url /pull/{n}` or `issue.pull_request`) | `github:pr:{repo}:{n}` |
| `issue_comment` on a non-PR issue | `github:issue:{repo}:{n}` |
| `push`/`delete` whose branch is in `notif:prbranch:{repo}:{branch}` | `github:pr:{repo}:{n}` (rerouted by `attachPrContext`) |

The SHA is extracted by `extractCommitSha()` (`notificationConsumer.js:404`). For pushes that prefers `data.after` → `data.head_commit.id` → last commit in `commits[]`. Make sure your `extra.data` matches the raw GitHub webhook payload so SHA extraction works.

Don't emit envelopes for events `server/queue/filters.js` silently drops: `star`, `watch`, `fork`, `ping`, `check_suite`, `workflow_run`, or `check_run`/`workflow_job` whose name contains `${{` (unexpanded matrix template).

Verify with one of the tests in `test/notifAnnounceRedirect.test.js` or by enqueueing through `!notif test <room> <msg>` first.

### Step 5 — Enqueue

```js
await notifRedis.lpush(k('queue'), JSON.stringify(envelope));
```

Where `k` is defined locally as:

```js
const KEY_PREFIX = process.env.NOTIF_KEY_PREFIX || 'notif:';
function k(name) { return KEY_PREFIX + name; }
```

The consumer ticks every `NOTIF_POLL_MS` (5000, fast 1000 after activity). Bump `notif:metrics:enqueued` if you want producer-side visibility:

```js
await notifRedis.incr(k('metrics:enqueued')).catch(() => {});
```

Verify: `await notifRedis.llen(k('queue'))` increases by 1, and the envelope appears in `!notif status` within one tick.

### Step 6 — (Optional) Pre-decide announce-redirect locally

If you want to know whether a GitHub envelope will end up in `int-dev-announce`, call `decideAnnounceRedirect(repo, process.env.NOTIF_ANNOUNCE_REPOS, process.env.NOTIF_ANNOUNCE_REPOS_EXCLUDE)` from `notificationConsumer.js`. Returns `{ redirect, matched, excluded?, excludedBy? }`. The consumer does this for you on every tick — only call it for logging or routing previews.

### Step 7 — Wire tests

Follow `test/notifAnnounceRedirect.test.js` / `test/notifBatchMerge.test.js`:

- Use `node:test` (`describe`, `it`, `assert/strict`).
- Set `process.env.NOTIF_TRACE_LOG = '0'` at the top of the file to silence the JSONL trace.
- For consumer paths, use `_setInternalsForTest({ redis, redisBot, adapter })` to inject mocks; restore in `after()`.
- For producer paths, mock `notifRedis.lpush` and assert on the JSON it received.

Run: `npm test -- --test-name-pattern=notif`.

## Examples

### Example 1 — Admin command enqueues a coalescable probe

User says: "add a `!ping-room <room>` admin command that enqueues a single coalescable 'pong' to the named room."

Actions: new file `server/commands/pingRoom.js`, registered in `server/commands/index.js`:

```js
const { MessageFactory } = require('botbuilder');
const { resolve: resolveChannel } = require('../queue/channels');
const { getNotifRedis } = require('../queue/notificationConsumer');

const KEY_PREFIX = process.env.NOTIF_KEY_PREFIX || 'notif:';
function k(name) { return KEY_PREFIX + name; }

module.exports = {
    match(text, lcText, deps) {
        if (deps.ima !== 'admin') return null;
        const m = text.match(/^!ping-room\s+(\S+)$/i);
        return m ? { room: m[1] } : null;
    },
    async execute({ room }, { context }) {
        if (!resolveChannel(room)) {
            await context.sendActivity(MessageFactory.text(`unknown room "${ room }"`));
            return;
        }
        const notifRedis = getNotifRedis();
        if (!notifRedis) {
            await context.sendActivity(MessageFactory.text('notif consumer not running'));
            return;
        }
        const envelope = {
            v: 1,
            id: 'ping-' + Date.now(),
            ts: Math.floor(Date.now() / 1000),
            expires_at: Math.floor(Date.now() / 1000) + 300,
            room,
            type: 'msg',
            message: 'pong',
            card: null,
            extra: { level: 'info', source: 'commands/pingRoom' },
            fallback_webhook_url: null
        };
        await notifRedis.lpush(k('queue'), JSON.stringify(envelope));
        await context.sendActivity(MessageFactory.text(`✅ enqueued to \`${ room }\``));
    }
};
```

Result: no `dedup_key` → coalescable. If two pings fire within the same tick they merge into one activity separated by `— · —`.

### Example 2 — External producer for a GitHub `push`

User says: "enqueue a push event from our webhook handler."

Actions:

```js
const { createNotifRedis } = require('../lib/redis');
const { resolve: resolveChannel } = require('../queue/channels');

const KEY_PREFIX = process.env.NOTIF_KEY_PREFIX || 'notif:';
const k = (name) => KEY_PREFIX + name;

async function enqueuePush(githubPayload, room) {
    if (!resolveChannel(room)) throw new Error(`unknown room: ${ room }`);
    const headCommit = githubPayload.head_commit || {};
    const repo = githubPayload.repository?.full_name || '';
    const envelope = {
        v: 1,
        id: 'gh-push-' + githubPayload.after,
        ts: Math.floor(Date.now() / 1000),
        expires_at: Math.floor(Date.now() / 1000) + 1800,
        room,
        type: 'msg',
        message: `📦 ${ headCommit.author?.name || 'someone' } pushed to [${ repo }](${ githubPayload.repository?.html_url })\n${ headCommit.message?.split('\n')[0] || '' }`,
        card: null,
        extra: {
            event_type: 'push',
            repo,
            data: githubPayload,
            level: 'info',
            source: 'webhooks/github'
        },
        fallback_webhook_url: null
    };
    const redis = createNotifRedis();
    await redis.lpush(k('queue'), JSON.stringify(envelope));
    await redis.incr(k('metrics:enqueued')).catch(() => {});
}
```

Result: the consumer's `normalizeGithubDedup()` rewrites `extra.dedup_key` to `github:commit:{sha7}`. Subsequent `check_run` / `workflow_job` envelopes for the same commit edit this activity in-place, appending job-status lines until the 3-update summary kicks in.

## Common Issues

- **"unknown room"** at enqueue time → the room name is not in `CHANNELS` (`server/queue/channels.js`). Add it there before producing; aliasing multiple producer names to one conversationId is the canonical pattern (see `hardware` ↔ `int-hw`).
- **Envelope shows up in `notif:dead` with `_dead_reason: 'expired'`** → `expires_at` was in the past. Use `Math.floor(Date.now() / 1000) + N` (seconds), not milliseconds.
- **Envelope shows up in `notif:dead` with `_dead_reason: 'json_parse_failed'`** → you LPUSHed a JS object instead of `JSON.stringify(...)`. Always serialize.
- **Envelope queued but never delivered, no errors** → check `server/queue/filters.js#shouldSkip` — `low-signal github event {type}` and `__SILENT__` skips are silent. Run `!notif status` to see the `redirected` counter increase.
- **Multiple events for the same commit produce N messages instead of 1 edited message** → either `extra._commit_sha` got set in the producer (don't — `extractCommitSha` will derive it) OR `EDIT_WINDOW_MS` expired between events (`NOTIF_COMMIT_GROUP_WINDOW_MS` defaults to 3 min). Confirm with `!notif wfactive`.
- **PR pushes spawn their own trackable instead of nesting under the PR** → the `pull_request` opened/edited event must arrive **before** the push so `notif:prbranch:{repo}:{branch}` is populated. Check ordering in your producer.
- **`getNotifRedis()` returns null** → consumer hasn't called `startConsumer()` yet (still in bot startup). Either wait, or in scripts use `createNotifRedis()` directly from `server/lib/redis.js`.
- **Card envelope rejected as "too large"** → Adaptive Card serialised JSON exceeds 25 KB. Trim text fields, remove base64 inline images, or split the card. See `enforceCardSizeLimit()` in `server/api/dailyRecapController.js` for the trimming algorithm.
- **`!notif test` works but real producer doesn't** → the producer is hitting the wrong Redis. `notif:*` keys must be on the InterServer Redis (`REDIS_HOST_MY` / `REDIS_PORT_MY`). Verify with `redis-cli -h $REDIS_HOST_MY llen notif:queue`.