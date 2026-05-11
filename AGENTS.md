# Teams Chat Bot

Microsoft Teams bot using RSC permissions to receive all channel messages without @mentions. Built on Bot Framework v4.23, Express.js, Node.js. Backed by Redis (`ioredis`), MongoDB (`mongodb`), MySQL (`mysql2`).

## Commands

```bash
npm start                  # node server/index.js on port 3978
npm run dev                # nodemon --inspect=9239 server/index.js
npm run build              # esbuild via build.js → dist/index.js
npm run lint               # eslint . (eslint.config.js, flat config)
npm test                   # node --test 'test/*.test.js'
npm run server             # npm install && node server/index.js
```

```bash
ngrok http 3978 --host-header="localhost:3978"
devtunnel host -p 3978 --allow-anonymous
```

## Architecture

- **Entry**: `server/index.js` — `dotenv` loads `../.env`, `validateEnv()` runs, mounts `/api` with `apiLimiter` (60/min/IP), exposes `/health` + `/health/queue`, starts `notificationConsumer` and optional `syncConversationReferences()`
- **Routes** (`server/api/index.js`): `POST /api/messages` → `botController.js` · `POST /api/message` → `msgController.js` (sendLimiter) · `POST /api/dailyrecap` → `dailyRecapController.js`
- **Bot** (`server/bot/`): `botActivityHandler.js` extends `TeamsActivityHandler` — owns MySQL pool, Mongo `usersCollection`, Redis client, `onMessage` → `dispatch()`, `setMaster()`, `probeChannel()`, `syncConversationReferences()`. `dialogBot.js` · `teamsBot.js` host the OAuth waterfall
- **Commands** (`server/commands/`): 21 modules — each exports `{ match(text, lcText, deps), execute(matchResult, deps) }`. Dispatched in order by `index.js` → first match wins. Admin gate: `if (ima !== 'admin') return null;`. See `ima.js`, `ping.js`, `githubIssues.js`, `notifAdmin.js`, `dailyRecap.js`, `ticketCard.js`/`ticketSubmit.js`/`ticketQuick.js`
- **Queue** (`server/queue/`): `notificationConsumer.js` polls `notif:queue` → trackable / coalesced / single sends, dead-letters to `notif:dead`. `channels.js` exports `CHANNELS` + `resolve(roomName)`. `filters.js` drops low-signal GitHub events (`star`, `watch`, `fork`, `ping`, successful `check_run`/`workflow_job`)
- **Shared libs** (`server/lib/`): `adapter.js` (`getAdapter()` singleton) · `redis.js` (`createBotRedis()`, `createNotifRedis()`, `TEAMS_SERVICE_URL`) · `retry.js` (`runWithRetry(fn, opts)` exponential backoff + jitter, classifies transient/auth/rate-limit)
- **Dialogs** (`server/dialogs/`): `mainDialog.js` (OAuth waterfall: `promptStep` → `loginStep` → `displayTokenPhase1` → `displayTokenPhase2`) extends `logoutDialog.js` extends `ComponentDialog`
- **Cards** (`cards/`): `add_ticket.json` Adaptive Card v1.5; `Action.Submit` carries `data.msteams.type` for routing (`addTicketSubmit`/`addTicketCancel`) and `activityId` is injected at send time by `bot.updateActionSubmitData()`
- **Infra**: `infra/azure.bicep` + `infra/azure.parameters.json` · manifest `appManifest/manifest.json` (RSC: `ChannelMessage.Read.Group`, `ChatMessage.Read.Chat`)
- **Tests** (`test/`): `node:test` — `commands.test.js`, `botActivityHandler.test.js`, `validateEnv.test.js`

## Environment (`.env`)

Required (`server/validateEnv.js`): `MicrosoftAppId`, `MicrosoftAppPassword`, `MYSQL_HOST`, `MYSQL_USER`, `MYSQL_PASS`, `MYSQL_DB`, `ZONEMTA_USERNAME`, `ZONEMTA_PASSWORD`, `ZONEMTA_HOST`.

Optional: `PORT` (3978), `connectionName`, `TEAMS_SERVICE_URL`, `REDIS_USER`/`REDIS_PASSWORD`, `REDIS_HOST_MY`/`REDIS_PORT_MY`, `DAILY_RECAP_URL`/`DAILY_RECAP_TOKEN`, `GITHUB_OWNER`/`GITHUB_REPO`, `TICKET_API_URL`, `BOT_AAD_OBJECT_ID`/`BOT_TENANT_ID`, `CHANNEL_SYNC_ENABLED`, `NOTIF_*`, `RATE_LIMIT_*`. Full table in `Readme.md` v2 changelog.

## Key Patterns

### Command modules
Every `server/commands/*.js` exports `{ match, execute }`. `deps` = `{ context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }`. Admin gate first. Register in the array in `server/commands/index.js` (order matters). Add a `match()` test in `test/commands.test.js`.

### Shared adapter & retry
Never `new BotFrameworkAdapter(...)` outside `server/lib/adapter.js`. Wrap every outbound Bot Framework call in `runWithRetry(fn, { label, serviceUrl, maxRetries })`.

### Notification envelopes
`notif:queue` JSON envelopes: `{ channel, text|attachments, extra?: { dedup_key }, fallback_webhook_url? }`. `extra.dedup_key` enables edit-window coalescing (`NOTIF_EDIT_WINDOW_MS`, 30 min). GitHub commit events auto-get `dedup_key=github:commit:{sha7}`.

### Channel routing
Always `resolve(roomName)` from `server/queue/channels.js`. Never hardcode `19:*@thread.v2` IDs.

### Proactive messaging
`server/api/msgController.js` reads `convref:{conversationId}` from Redis. `botActivityHandler.onMessage` and `onInstallationUpdateAdd` populate `convref:*`. Missing convref → constructed-reference fallback using `TEAMS_SERVICE_URL`.

### Daily recap
`server/api/dailyRecapController.js` requires `X-Daily-Recap-Token`. `adaptivecards-templating` binds template+data server-side. `enforceCardSizeLimit()` trims cards over 25 KB; `scheduleAutoDelete()` removes after 5 minutes. Cached at `dailyrecap:{activityId}` (2-day TTL).

### Adaptive cards
v1.5 schema. Every `Action.Submit` MUST include `data.msteams.type` and is paired with `activityId` injected by `bot.updateActionSubmitData(element, sentActivity)`. See `server/commands/ticketCard.js`.

### RSC permissions
`appManifest/manifest.json` declares `ChannelMessage.Read.Group` + `ChatMessage.Read.Chat` under `authorization.permissions.resourceSpecific`.

### Rate limiting
`express-rate-limit` in `server/index.js` and `server/api/index.js`. `/api`: 60/min/IP. `/api/message` + `/api/dailyrecap`: 30/min keyed by `${ip}:${channel}`.

## ESLint (`eslint.config.js`)

- Flat config (ESLint 9), `ecmaVersion: 2022`, `sourceType: 'commonjs'`
- `semi: [2, 'always']` · `indent: [2, 4]` · `template-curly-spacing: [2, 'always']` (`${ var }`)
- `space-before-function-paren`: `named: 'never'`, `anonymous: 'never'`, `asyncArrow: 'always'`

## Deployment

`infra/azure.bicep` + `infra/azure.parameters.json`. Edit `appManifest/manifest.json` `botId` + `validDomains` before zipping.

## Before Committing

Run `caliber refresh` to keep docs in sync, then stage updated docs.

<!-- caliber:managed:learnings -->
## Session Learnings

Read `CALIBER_LEARNINGS.md` for patterns and anti-patterns learned from previous sessions.
These are auto-extracted from real tool usage — treat them as project-specific rules.
<!-- /caliber:managed:learnings -->
