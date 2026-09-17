---
paths:
  - server/api/**
---

# Controller Boilerplate

- NEVER `new BotFrameworkAdapter(...)` here. Import `getAdapter` from `../lib/adapter`.
- NEVER `new Redis(...)` here. Import `createBotRedis` (convref) or `getNotifRedis` (queue) from `../lib/redis`.
- Wrap every `adapter.continueConversation` / `adapter.processActivity` in `runWithRetry(fn, { label, serviceUrl, maxRetries })` from `../lib/retry` — handles transient/auth/rate-limit classification + `MicrosoftAppCredentials.trustServiceUrl()` on auth failures.
- Size-guard every outbound activity with `../lib/messageSize` (`activityBytes`, `fitTextActivity`, `trimCard`, `MAX_ACTIVITY_BYTES`) BEFORE the adapter call — Teams rejects activities over ~25 KB serialized and `onTurnError` swallows the `RestError`, so an oversized send returns a false success. Clamp text rather than reject (`/api/message` responds `truncated: true`); return `400` only when a card cannot be trimmed to fit (`server/api/dailyRecapController.js`).
- Measure sizes with `Buffer.byteLength(JSON.stringify(...), 'utf8')`, never JS `.length` — emoji and JSON escaping make code-unit counts lie.
- Resolve room names through `resolve(roomName)` from `../queue/channels` — never hardcode `19:*@thread.v2` IDs.
- Validate request body shape and return `400` with `{ message }` or `{ ok: false, error }` (match the surrounding controller's envelope style).
- Auth-token endpoints: read `X-Daily-Recap-Token` (or similar) header and compare to env var; return `401` on mismatch. See `server/api/dailyRecapController.js`.
- Register the route in `server/api/index.js` and apply `sendLimiter` to any proactive-send endpoint.
- `onTurnError` handler MUST classify `TRANSIENT_RE` / `AUTH_ERROR_RE` and skip user-visible replies on transient/auth failures, and short-circuit size errors (413 / `message size too large`) without attempting a reply (see `server/api/botController.js`).
