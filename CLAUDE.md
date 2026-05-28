# Teams Chat Bot

Microsoft Teams bot using RSC permissions to receive all channel messages without @mentions. Built on Bot Framework v4.23 (`botbuilder`, `botbuilder-dialogs`, `botframework-connector`), Express.js, Node.js. Backed by Redis (`ioredis`), MongoDB (`mongodb`), MySQL (`mysql2`). Card binding via `adaptivecards-templating`. Webhook/GitHub via `@octokit/rest`. Build via `esbuild` (`build.js`).

## Commands

```bash
npm start                  # node server/index.js on port 3978
npm run dev                # nodemon --inspect=9239 server/index.js
npm run build              # esbuild bundle via build.js → dist/index.js
npm run lint               # eslint . (eslint.config.js, flat config)
npm test                   # node --test 'test/*.test.js'
npm run server             # npm install && node server/index.js
npm run manifest           # Windows: zip appManifest/* → appManifest.zip
```

```bash
ngrok http 3978 --host-header="localhost:3978"
devtunnel host -p 3978 --allow-anonymous
```

## Architecture

- **Entry** `server/index.js`: `dotenv` loads `../.env`, `validateEnv()` runs, mounts `/api` with `apiLimiter` (60/min/IP), exposes `/health` + `/health/queue`, calls `startConsumer()` from `server/queue/notificationConsumer.js`, optionally calls `botActivityHandler.syncConversationReferences()` when `CHANNEL_SYNC_ENABLED=1`. SIGTERM/SIGINT → `stopConsumer()` + 10s graceful shutdown
- **Routes** (`server/api/index.js`): `POST /api/messages` → `botController.js` · `POST /api/message` → `msgController.js` (sendLimiter) · `POST /api/dailyrecap` → `dailyRecapController.js` (sendLimiter)
- **Bot** (`server/bot/`): `botActivityHandler.js` extends `TeamsActivityHandler` — owns MySQL pool, Mongo `usersCollection` (from `zone-mta`.`users`), Redis client, master-table introspection (`backup_masters`, `website_masters`, `vps_masters`, `qs_masters`), `onMessage` → `dispatch()`, `setMaster()`, `probeChannel()`, `syncConversationReferences()`. `dialogBot.js` · `teamsBot.js` host the OAuth waterfall
- **Commands** (`server/commands/`): 22 modules dispatched in order by `index.js` → first match wins. Admin gate: `if (ima !== 'admin') return null;`. Modules: `ima.js`, `ping.js`, `joke.js`, `setMaster.js`, `ticketCard.js`, `ticketSubmit.js`, `ticketPost.js`, `ticketQuick.js`, `mailbabyUser.js`, `ipLookup.js`, `blockEmail.js`, `blockDomain.js`, `blockHelp.js`, `githubIssues.js`, `githubLabels.js`, `assetSearch.js`, `hypervStatus.js`, `processingStatus.js`, `globalVar.js`, `dailyRecap.js`, `notifAdmin.js`, `help.js`
- **Queue** (`server/queue/`): `notificationConsumer.js` polls `notif:queue` → trackable / coalesced / single sends, dead-letters to `notif:dead`. `channels.js` exports `CHANNELS` + `resolve(roomName)` + `knownRooms()`. `filters.js` drops low-signal GitHub events (`star`, `watch`, `fork`, `ping`, successful `check_run`/`workflow_job`, `${{ matrix.* }}` placeholders). `notifTrace.js` writes JSONL events to `.logs/notif-trace-YYYY-MM-DD.jsonl` for replay
- **Shared libs** (`server/lib/`): `adapter.js` (`getAdapter()` / `setAdapter()` singleton) · `redis.js` (`createBotRedis()`, `createNotifRedis()`, `TEAMS_SERVICE_URL`) · `retry.js` (`runWithRetry(fn, opts)` exponential backoff + jitter, classifies `transient`/`auth`/`rateLimit`)
- **Dialogs** (`server/dialogs/`): `mainDialog.js` (OAuth waterfall: `promptStep` → `loginStep` → `displayTokenPhase1` → `displayTokenPhase2`) extends `logoutDialog.js` extends `ComponentDialog`
- **Cards** (`cards/`): `add_ticket.json` — Adaptive Card v1.5; `Action.Submit` carries `data.msteams.type` (`addTicketSubmit`/`addTicketCancel`) for routing; `activityId` injected at send time via `bot.updateActionSubmitData(element, sentActivity)`
- **Scripts**: `scripts/replay-notif.js` — JSONL trace replayer (`--mode timeline|grouped|raw`, `--dedup`, `--commit`, `--room`, `--tick`, `--event`, `--kind`, `--since/--until`, `--activity`)
- **Infra**: `infra/azure.bicep` + `infra/azure.parameters.json` · manifest `appManifest/manifest.json` (RSC: `ChannelMessage.Read.Group`, `ChatMessage.Read.Chat`, botId `6fa7ed27-9923-4d5a-9f2d-2c9b81cfdd2d`, validDomains `mybot.interserver.net` + `token.botframework.com`)
- **Deploy**: `m365agents.yml` / `m365agents.local.yml` (M365 Agents Toolkit)
- **Tests** (`test/`, `node:test`): `commands.test.js` (all command `match()` regexes), `botActivityHandler.test.js` (`isValidIP`/`isValidHostname`/`updateActionSubmitData` — ⚠️ hangs at end of run), `validateEnv.test.js` (REQUIRED env list), `filters.test.js`, `notifAnnounceRedirect.test.js`, `notifBatchMerge.test.js`, `notifConvrefFallback.test.js`, `notifPrContext.test.js`, `notifTrace.test.js`, `notificationConsumer.test.js`

## Environment (`.env`)

Required (`server/validateEnv.js`): `MicrosoftAppId`, `MicrosoftAppPassword`, `MYSQL_HOST`, `MYSQL_USER`, `MYSQL_PASS`, `MYSQL_DB`, `ZONEMTA_USERNAME`, `ZONEMTA_PASSWORD`, `ZONEMTA_HOST`.

Optional: `PORT` (3978), `connectionName` (OAuth dialog), `TEAMS_SERVICE_URL`, `REDIS_USER`/`REDIS_PASSWORD`, `REDIS_HOST_MY`/`REDIS_PORT_MY`, `DAILY_RECAP_URL`/`DAILY_RECAP_TOKEN`, `GITHUB_OWNER`/`GITHUB_REPO`, `TICKET_API_URL`/`TICKET_POST_API_URL`, `BOT_AAD_OBJECT_ID`/`BOT_TENANT_ID`, `CHANNEL_SYNC_ENABLED`, `NOTIF_POLL_MS`/`NOTIF_POLL_FAST_MS`/`NOTIF_MAX_PER_TICK`/`NOTIF_EDIT_WINDOW_MS`/`NOTIF_KEY_PREFIX`/`NOTIF_FILTER_ENABLED`/`NOTIF_CONSUMER_ENABLED`/`NOTIF_HEARTBEAT_MS`/`NOTIF_COMMIT_GROUP_WINDOW_MS`/`NOTIF_DOWNSTREAM_REPOS`/`NOTIF_ANNOUNCE_REPOS`/`NOTIF_ANNOUNCE_REPOS_EXCLUDE`/`NOTIF_TRACE_LOG`, `RATE_LIMIT_WINDOW_MS`/`RATE_LIMIT_MAX`/`RATE_LIMIT_SEND_WINDOW_MS`/`RATE_LIMIT_SEND_MAX`. Authoritative reference: `.env.example`.

## Key Patterns

### Command modules (`server/commands/*.js`)
Every module: `match(text, lcText, deps)` returns truthy match data or `null`; `execute(matchResult, deps)` performs the work. `deps` = `{ context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }`. Admin gate: `if (ima !== 'admin') return null;` as the FIRST line of `match`. Reply via `context.sendActivity(MessageFactory.text(...))`. New modules MUST be appended to the `commands` array in `server/commands/index.js` — order matters. Add a `match()` test in `test/commands.test.js`. See `server/commands/ima.js` (minimal) and `server/commands/githubIssues.js` (multi-action).

### Shared adapter & retry
Never `new BotFrameworkAdapter(...)` outside `server/lib/adapter.js` — always `getAdapter()`. Wrap every outbound `continueConversation`/`processActivity` in `runWithRetry(fn, { label, serviceUrl, maxRetries })` from `server/lib/retry.js`. Auto-trusts `serviceUrl` on auth failures, classifies `transient`/`auth`/`rateLimit`. See `server/api/botController.js` for `onTurnError` pattern.

### Notification queue envelopes
`server/queue/notificationConsumer.js` reads JSON envelopes from Redis `notif:queue`. Shape: `{ channel, text|attachments, extra?: { dedup_key, event_type, _commit_sha, data }, fallback_webhook_url? }`. `channel` resolves via `CHANNELS`. `extra.dedup_key` enables trackable edits within `NOTIF_EDIT_WINDOW_MS` (30 min). GitHub commit events auto-get `dedup_key=github:commit:{sha7}`; PR-related branches/comments rewrite to `github:pr:{repo}:{n}` via `attachPrContext`. Action-triggered pushes (bot pusher OR `NOTIF_DOWNSTREAM_REPOS` glob) are rewritten to the parent SHA's dedup_key. Within a tick, multiple items sharing a `dedup_key` MUST be folded via `handleTrackableBatch` (1 API call), not iterated.

### Channel routing
Always `resolve(roomName)` from `server/queue/channels.js`. Never hardcode `19:*@thread.v2` IDs. Unknown channels → 400. `SKIP_CHANNELS` array in `msgController.js` short-circuits disabled rooms. `NOTIF_ANNOUNCE_REPOS` (comma-list of `owner/*` or `owner/repo`) redirects to `int-dev-announce`; `NOTIF_ANNOUNCE_REPOS_EXCLUDE` exempts. Exclude wins on conflict.

### Proactive messaging
`server/api/msgController.js` reads `convref:{conversationId}` from Redis (`createBotRedis()`) and calls `sendProactiveMessage()`. `botActivityHandler.onMessage` + `onInstallationUpdateAdd` populate `convref:*`. Missing convref → `buildConstructedConvRef(room, conversationId)` fallback (`_constructed: true` marker) using `TEAMS_SERVICE_URL`. If even that fails → envelope's `fallback_webhook_url` (Power Automate).

### Daily recap cards
`server/api/dailyRecapController.js` requires `X-Daily-Recap-Token` header. Uses `adaptivecards-templating` server-side (Teams cannot expand `${...}`). `enforceCardSizeLimit()` trims cards over 25 KB; `scheduleAutoDelete()` removes after 5 minutes. Cached at `dailyrecap:{activityId}` (2-day TTL). `server/commands/dailyRecap.js` is the admin bot command (`daily recap`); compact mode forces `show_details/show_pie/show_month=false` + empties `orders` arrays.

### Adaptive cards
Cards in `cards/` follow Adaptive Card v1.5. Every `Action.Submit` MUST include `data.msteams.type` for routing in `botActivityHandler` and is paired with `activityId` injected at send time by `bot.updateActionSubmitData(element, sentActivity)` (see `server/commands/ticketCard.js`).

### RSC permissions
`appManifest/manifest.json` declares `ChannelMessage.Read.Group` + `ChatMessage.Read.Chat` under `authorization.permissions.resourceSpecific`. Bot receives all channel messages without @mention.

### Rate limiting
`express-rate-limit` in `server/index.js` and `server/api/index.js`. `/api`: 60/min/IP. `/api/message` + `/api/dailyrecap`: 30/min keyed by `${ip}:${channel}`. `validateXForwardedForHeader: false` because no trusted proxy.

### Trace + replay debugging
`server/queue/notifTrace.js` emits JSONL events to `.logs/notif-trace-YYYY-MM-DD.jsonl` (disable with `NOTIF_TRACE_LOG=0`). Replay with `node scripts/replay-notif.js --dedup <key> --mode timeline` (or `--mode grouped`, `--commit <sha>`, `--room`, `--event <type>`, `--kind <trace_kind>`). Signature kinds for grouping bugs: `recent_lookup`, `edit_skipped_no_convref`, `edit_fell_through`, multiple `recent_saved` with different activity ids.

## ESLint Rules (`eslint.config.js`)

- Flat config (ESLint 9), `ecmaVersion: 2022`, `sourceType: 'commonjs'`, `globals.node`
- `semi: [2, 'always']` · `indent: [2, 4]` · `template-curly-spacing: [2, 'always']` (`${ var }`)
- `space-before-function-paren`: `named: 'never'`, `anonymous: 'never'`, `asyncArrow: 'always'`
- `no-unused-vars: [1, { argsIgnorePattern: '^_' }]`

## Test Gotchas

- `test/botActivityHandler.test.js` hangs at end of run (unclosed Mongo/MySQL/Redis client). To run the suite verifiably, list other files explicitly: `node --test test/notificationConsumer.test.js test/notifPrContext.test.js test/notifTrace.test.js test/notifConvrefFallback.test.js test/filters.test.js test/commands.test.js test/validateEnv.test.js test/notifAnnounceRedirect.test.js test/notifBatchMerge.test.js`.
- Consumer internals are swappable via `_setInternalsForTest({ redis, redisBot, adapter })` — restore originals in `after()`.
- `NOTIF_TRACE_LOG=0` in tests to silence file writes.

## Deployment

Azure resources via `infra/azure.bicep` + `infra/azure.parameters.json`. Edit `appManifest/manifest.json` `botId` + `validDomains` before zipping. Linux: `cd appManifest && zip -r appManifest.zip manifest.json outline.png color.png`.

@./Readme.md
@./CALIBER_LEARNINGS.md

## Before Committing

Run `caliber refresh` before creating git commits. Stage modified docs: `caliber refresh && git add CLAUDE.md .claude/ .cursor/ .github/copilot-instructions.md AGENTS.md CALIBER_LEARNINGS.md 2>/dev/null`.

## Personal Preferences

- Do NOT run multiple Agent/sub-agent tasks in parallel — sequence them one at a time (rate-limit + file-collision risk).
- When executing a multi-slice plan, do NOT pause for approval between slices — proceed straight to the next slice after merge.
- For multi-library/multi-phase audit fixes, split into one PR per library/phase, not one bundled PR.

<!-- caliber:managed:learnings -->
## Session Learnings

Read `CALIBER_LEARNINGS.md` for patterns and anti-patterns learned from previous sessions.
These are auto-extracted from real tool usage — treat them as project-specific rules.
<!-- /caliber:managed:learnings -->

<!-- caliber:managed:model-config -->
## Model Configuration

Recommended default: `claude-sonnet-4-6` with high effort (stronger reasoning; higher cost and latency than smaller models).
Smaller/faster models trade quality for speed and cost — pick what fits the task.
Pin your choice (`/model` in Claude Code, or `CALIBER_MODEL` when using Caliber with an API provider) so upstream default changes do not silently change behavior.

<!-- /caliber:managed:model-config -->

<!-- caliber:managed:sync -->
## Context Sync

This project uses [Caliber](https://github.com/caliber-ai-org/ai-setup) to keep AI agent configs in sync across Claude Code, Cursor, Copilot, and Codex.
Configs update automatically before each commit via `caliber refresh`.
If the pre-commit hook is not set up, run `/setup-caliber` to configure everything automatically.
<!-- /caliber:managed:sync -->
