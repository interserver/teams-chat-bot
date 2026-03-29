# Teams Chat Bot

Microsoft Teams bot using RSC permissions to receive all channel messages without @mentions. Built on Bot Framework v4.23, Express.js, Node.js.

## Commands

```bash
npm start                  # production: node server/index.js on port 3978
npm run dev                # nodemon --inspect=9239 server/index.js
npm run build              # esbuild bundle via build.js
npm run lint               # eslint .
npm run server             # npm install && node server/index.js
```

```bash
# Local tunnel (set endpoint https://<tunnel>/api/messages in Azure Bot)
ngrok http 3978 --host-header="localhost:3978"
devtunnel host -p 3978 --allow-anonymous
```

## Architecture

- **Entry**: `server/index.js` → mounts `/api` router on port `3978`
- **Routes**: `server/api/index.js` · `POST /api/messages` → `server/api/botController.js` · `POST /api/message` → `server/api/msgController.js`
- **Bot**: `server/bot/botActivityHandler.js` · `server/bot/teamsBot.js` · `server/bot/dialogBot.js`
- **Dialogs**: `server/dialogs/mainDialog.js` (OAuth waterfall) · `server/dialogs/logoutDialog.js`
- **Cards**: `cards/add_ticket.json` (Adaptive Card v1.5 with Teams submit actions)
- **Infra**: `infra/azure.bicep` + `infra/azure.parameters.json` · manifest: `appManifest/manifest.json`
- **Deploy**: `m365agents.yml` / `m365agents.local.yml` (Teams Toolkit / M365 Agents Toolkit)
- **Storage**: Redis (`ioredis`) for conversation references · `mongodb` · `mysql2`

## Environment (`.env`)

- `MicrosoftAppId` — Azure AD app client ID (`AAD_APP_CLIENT_ID`)
- `MicrosoftAppPassword` — Azure AD client secret
- `PORT` — defaults to `3978`
- `connectionName` — OAuth connection name used by `server/dialogs/mainDialog.js`

## Key Patterns

### Error handling in controllers
Both `server/api/botController.js` and `server/api/msgController.js` define:
- `TRANSIENT_RE` — `/ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i` → skip reply, log only
- `AUTH_ERROR_RE` — `/authorization has been denied|401|unauthorized/i` → log and return
- `adapter.onTurnError` — sends `sendTraceActivity` + user-facing error message
- `runWithRetry(context, handler)` — max `MAX_RETRIES=2` with `RETRY_DELAY_MS=1000`

### Proactive messaging
`server/api/msgController.js` reads key `convref:{conversationId}` from Redis and calls `sendProactiveMessage()`. Bot must store `TurnContext.activity.conversation` as a conversation reference when it joins.

### Adaptive cards
Cards in `cards/` follow Adaptive Card schema v1.5. Submit actions must include `data.msteams.type` for routing (e.g., `addTicketSubmit`, `addTicketCancel`). Handle in `server/bot/botActivityHandler.js` via `handleTeamsCardActionInvoke` or message activity type checks.

### RSC permissions
`appManifest/manifest.json` declares `ChannelMessage.Read.Group` and `ChatMessage.Read.Chat` under `authorization.permissions.resourceSpecific`. Bot receives all channel messages without @mention.

### Dialog structure
`server/dialogs/mainDialog.js` extends `LogoutDialog` → extends `ComponentDialog`. Waterfall steps: `promptStep` → `loginStep` → `displayTokenPhase1` → `displayTokenPhase2`. Uses `OAuthPrompt` with `process.env.connectionName`.

## ESLint Rules (`.eslintrc.js`)
- Semicolons required: `"semi": [2, "always"]`
- 4-space indent: `"indent": [2, 4]`
- Template literals: `${ variable }` (spaces inside braces)
- No space before named/anonymous function parens; space before async arrow

## Deployment

Azure resources provisioned via `infra/azure.bicep` using params from `infra/azure.parameters.json`. Teams app packaged from `appManifest/` — edit `appManifest/manifest.json` to replace `botId` and `validDomains` before zipping.

```bash
# Zip manifest (Windows)
npm run manifest
# Manual zip
cd appManifest && zip -r appManifest.zip manifest.json outline.png color.png
```

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
These are auto-extracted from real tool usage — treat them as project-specific rules.
<!-- /caliber:managed:learnings -->
