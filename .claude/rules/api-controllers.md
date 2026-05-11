---
paths:
  - server/api/**
---

# Controller Boilerplate

- NEVER `new BotFrameworkAdapter(...)` here. Import `getAdapter` from `../lib/adapter`.
- NEVER `new Redis(...)` here. Import `createBotRedis` (convref) or `getNotifRedis` (queue) from `../lib/redis`.
- Wrap every `adapter.continueConversation` / `adapter.processActivity` in `runWithRetry(fn, { label, serviceUrl, maxRetries })` from `../lib/retry` — handles transient/auth/rate-limit classification + `MicrosoftAppCredentials.trustServiceUrl()` on auth failures.
- Resolve room names through `resolve(roomName)` from `../queue/channels` — never hardcode `19:*@thread.v2` IDs.
- Validate request body shape and return `400` with `{ message }` or `{ ok: false, error }` (match the surrounding controller's envelope style).
- Auth-token endpoints: read `X-Daily-Recap-Token` (or similar) header and compare to env var; return `401` on mismatch. See `server/api/dailyRecapController.js`.
- Register the route in `server/api/index.js` and apply `sendLimiter` to any proactive-send endpoint.
