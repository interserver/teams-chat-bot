# Teams Chat Bot

Microsoft Teams bot using RSC permissions to receive all channel messages without @mentions. Built on Bot Framework v4.23, Express.js, Node.js. Backed by Redis (`ioredis`), MongoDB (`mongodb`), MySQL (`mysql2`).

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

- **Entry**: `server/index.js` — `dotenv` loads `../.env`, `validateEnv()` runs, mounts `/api` with `apiLimiter` (60/min/IP), exposes `/health` + `/health/queue`, starts `notificationConsumer` and optional `syncConversationReferences()`
- **Routes** (`server/api/index.js`): `POST /api/messages` → `botController.js` · `POST /api/message` → `msgController.js` (sendLimiter) · `POST /api/dailyrecap` → `dailyRecapController.js` (sendLimiter)
- **Bot** (`server/bot/`): `botActivityHandler.js` extends `TeamsActivityHandler` — owns MySQL pool, Mongo `usersCollection`, Redis client, master-table introspection, `onMessage` → `dispatch()`, `setMaster()`, `probeChannel()`, `syncConversationReferences()`. `dialogBot.js` · `teamsBot.js` for OAuth waterfall hosting
- **Commands** (`server/commands/`): 21 modules — each exports `{ match(text, lcText, deps), execute(matchResult, deps) }`. Dispatched in order by `index.js` → first match wins. Admin gate: `if (ima !== 'admin') return null;`
- **Queue** (`server/queue/`): `notificationConsumer.js` polls `notif:queue` → routes to trackable / coalesced / single sends, writes to `notif:dead` on failure. `channels.js` exports `CHANNELS` map + `resolve(roomName)`. `filters.js` drops low-signal GitHub events (`star`, `watch`, `fork`, `ping`, successful `check_run`/`workflow_job`)
- **Shared libs** (`server/lib/`): `adapter.js` (`getAdapter()` singleton `BotFrameworkAdapter`) · `redis.js` (`createBotRedis()`, `createNotifRedis()`, `TEAMS_SERVICE_URL`) · `retry.js` (`runWithRetry(fn, opts)` with exponential backoff + jitter, classifies transient/auth/rate-limit)
- **Dialogs** (`server/dialogs/`): `mainDialog.js` (OAuth waterfall: `promptStep` → `loginStep` → `displayTokenPhase1` → `displayTokenPhase2`) extends `logoutDialog.js` extends `ComponentDialog`
- **Cards** (`cards/`): `add_ticket.json` — Adaptive Card v1.5; `Action.Submit` carries `data.msteams.type` (`addTicketSubmit`/`addTicketCancel`) for routing
- **Infra**: `infra/azure.bicep` + `infra/azure.parameters.json` · manifest `appManifest/manifest.json` (RSC: `ChannelMessage.Read.Group`, `ChatMessage.Read.Chat`)
- **Deploy**: `m365agents.yml` / `m365agents.local.yml` (M365 Agents Toolkit)
- **Tests** (`test/`): `node:test` — `commands.test.js` covers all command `match()` regexes, `botActivityHandler.test.js` covers `isValidIP`/`isValidHostname`/`updateActionSubmitData`, `validateEnv.test.js` covers REQUIRED env list

## Environment (`.env`)

Required (see `server/validateEnv.js`): `MicrosoftAppId`, `MicrosoftAppPassword`, `MYSQL_HOST`, `MYSQL_USER`, `MYSQL_PASS`, `MYSQL_DB`, `ZONEMTA_USERNAME`, `ZONEMTA_PASSWORD`, `ZONEMTA_HOST`.

Optional: `PORT` (3978), `connectionName` (OAuth), `TEAMS_SERVICE_URL`, `REDIS_USER`/`REDIS_PASSWORD`, `REDIS_HOST_MY`/`REDIS_PORT_MY`, `DAILY_RECAP_URL`/`DAILY_RECAP_TOKEN`, `GITHUB_OWNER`/`GITHUB_REPO`, `TICKET_API_URL`, `BOT_AAD_OBJECT_ID`/`BOT_TENANT_ID`, `CHANNEL_SYNC_ENABLED`, `NOTIF_*` (poll, edit window, key prefix, filter, heartbeat — see Readme.md changelog), `RATE_LIMIT_*`.

## Key Patterns

### Command modules (`server/commands/*.js`)
Every module: `match(text, lcText, deps)` returns truthy match data or `null`; `execute(matchResult, deps)` performs the work. `deps` includes `{ context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }`. Admin-only commands gate on `ima !== 'admin'`. New commands MUST be added to the array in `server/commands/index.js`. See `server/commands/ima.js` for the minimal shape and `server/commands/githubIssues.js` for multi-action dispatch.

### Shared adapter & retry
Never instantiate `new BotFrameworkAdapter` directly — call `getAdapter()` from `server/lib/adapter.js`. Wrap every outbound `continueConversation`/`processActivity` in `runWithRetry(fn, { label, serviceUrl, maxRetries })` from `server/lib/retry.js`. The retry helper auto-trusts `serviceUrl` on auth failures and classifies errors as `transient`/`auth`/`rateLimit`.

### Notification queue envelopes
`notificationConsumer.js` reads JSON envelopes from Redis `notif:queue`. Envelopes can carry `extra.dedup_key` to enable trackable edits within `NOTIF_EDIT_WINDOW_MS` (30 min default), `fallback_webhook_url` for Power Automate fallback, and a `channel` field resolved through `CHANNELS`. GitHub commit events get `dedup_key=github:commit:{sha7}` auto-injected so commit + job events edit one message.

### Channel routing
All outbound proactive sends MUST resolve room names via `resolve(roomName)` from `server/queue/channels.js`. Never hardcode `19:*@thread.v2` IDs in callers. Unknown channels → 400. `SKIP_CHANNELS` array in `msgController.js` short-circuits disabled rooms.

### Proactive messaging
`server/api/msgController.js` reads `convref:{conversationId}` from Redis and calls `sendProactiveMessage()`. The bot stores convrefs in `botActivityHandler.onMessage` and `onInstallationUpdateAdd`. Missing convref → `tryConstructedConvRef` fallback using `TEAMS_SERVICE_URL` + known `conversationId`.

### Daily recap cards
`server/api/dailyRecapController.js` requires `X-Daily-Recap-Token` header. Uses `adaptivecards-templating` server-side (Teams cannot expand `${...}`). `enforceCardSizeLimit()` trims cards over 25 KB; `scheduleAutoDelete()` removes the card after 5 minutes. Template+data cached at `dailyrecap:{activityId}` (2-day TTL).

### Adaptive cards
Cards in `cards/` follow Adaptive Card v1.5. Every `Action.Submit` MUST include `data.msteams.type` for routing in `botActivityHandler` and is paired with `activityId` injected at send time by `bot.updateActionSubmitData(element, sentActivity)` (see `server/commands/ticketCard.js`).

### RSC permissions
`appManifest/manifest.json` declares `ChannelMessage.Read.Group` + `ChatMessage.Read.Chat` under `authorization.permissions.resourceSpecific`. The bot receives all channel messages without @mention.

### Rate limiting
`express-rate-limit` configured in `server/index.js` and `server/api/index.js`. General `/api`: 60 req/min/IP. `/api/message` + `/api/dailyrecap`: 30 req/min keyed by `${ip}:${channel}`.

## ESLint Rules (`eslint.config.js`)

- Flat config (ESLint 9), `ecmaVersion: 2022`, `sourceType: 'commonjs'`
- `semi: [2, 'always']` · `indent: [2, 4]` · `template-curly-spacing: [2, 'always']` (`${ var }`)
- `space-before-function-paren`: `named: 'never'`, `anonymous: 'never'`, `asyncArrow: 'always'`
- `no-unused-vars: [1, { argsIgnorePattern: '^_' }]`

## Deployment

Azure resources via `infra/azure.bicep` + `infra/azure.parameters.json`. Edit `appManifest/manifest.json` to replace `botId` and `validDomains` (`mybot.interserver.net`, `token.botframework.com`).

```bash
npm run manifest                                    # Windows
cd appManifest && zip -r appManifest.zip manifest.json outline.png color.png
```

@./Readme.md

<!-- caliber:managed:pre-commit -->
## Before Committing

Run `caliber refresh` before creating git commits to keep docs in sync with code changes.
After it completes, stage any modified doc files before committing:

```bash
caliber refresh && git add CLAUDE.md .claude/ .cursor/ .github/copilot-instructions.md AGENTS.md CALIBER_LEARNINGS.md 2>/dev/null
```
<!-- /caliber:managed:pre-commit -->

<!-- caliber:managed:learnings -->
## Session Learnings

Read `CALIBER_LEARNINGS.md` for patterns and anti-patterns learned from previous sessions.
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
