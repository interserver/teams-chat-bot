---
name: api-route
description: Adds a new Express route and controller module to `server/api/`, following the exact patterns in `botController.js`/`msgController.js`: `BotFrameworkAdapter` setup, `TRANSIENT_RE`/`AUTH_ERROR_RE` regexes, `adapter.onTurnError`, `sleep()`, and `runWithRetry()`. Trigger on: 'add endpoint', 'new route', 'create API handler', 'add controller'. Do NOT use for modifying `POST /api/messages` (the bot message pipeline) or `POST /api/message` (proactive messaging).
---
# api-route

## Critical

- Never register a new route on the messages or message paths — those are reserved for the existing bot and message controllers in `server/api/`.
- Every controller that touches the Bot Framework adapter MUST include `TRANSIENT_RE`, `AUTH_ERROR_RE`, and `adapter.onTurnError` — skipping these causes silent cascading failures on transient network errors.
- `MicrosoftAppId` and `MicrosoftAppPassword` must be loaded from environment variables — never hardcode credentials.

## Instructions

1. **Create a new controller file in `server/api/`** using this exact boilerplate:

```js
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { BotFrameworkAdapter } = require('botbuilder');
const { MicrosoftAppCredentials } = require('botframework-connector');

const TRANSIENT_RE = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i;
const AUTH_ERROR_RE = /authorization has been denied|401|unauthorized/i;
const MAX_RETRIES = 2;
const RETRY_DELAY_MS = 1000;

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

async function runWithRetry(context, handler) {
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
            await handler(context);
            return;
        } catch (err) {
            const msg = err.message || '';
            const isTransient = TRANSIENT_RE.test(msg);
            const isAuth = AUTH_ERROR_RE.test(msg);
            if ((isTransient || isAuth) && attempt < MAX_RETRIES) {
                console.warn(`[retry] Attempt ${attempt} failed (${msg}), retrying in ${RETRY_DELAY_MS}ms...`);
                if (isAuth) MicrosoftAppCredentials.trustServiceUrl(context.activity.serviceUrl);
                await sleep(RETRY_DELAY_MS * attempt);
                continue;
            }
            throw err;
        }
    }
}

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

adapter.onTurnError = async (context, error) => {
    const errorMsg = error.message || 'Oops. Something went wrong!';
    console.error(`\n [onTurnError] unhandled error: ${ error }`);
    const isTransient = TRANSIENT_RE.test(errorMsg);
    const isAuthError = AUTH_ERROR_RE.test(errorMsg);
    if (isTransient) { console.error('[onTurnError] Transient network error, skipping reply to user.'); return; }
    if (isAuthError) { console.error('[onTurnError] Authorization error — check MicrosoftAppId/MicrosoftAppPassword and bot registration.'); return; }
    try {
        await context.sendTraceActivity('OnTurnError Trace', `${ error }`, 'https://www.botframework.com/schemas/error', 'TurnError');
        await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${ errorMsg }`);
    } catch (sendError) {
        console.error(`[onTurnError] Failed to send error message to user: ${ sendError.message }`);
    }
};

const <name>Handler = async (req, res) => {
    try {
        // TODO: implement handler logic
        res.json({ message: 'ok' });
    } catch (err) {
        console.error('[<name>Handler] Error:', err.message);
        res.status(500).json({ message: err.message });
    }
};

module.exports = <name>Handler;
```

   Verify the file exists in `server/api/` before proceeding.

2. **Register the route in `server/api/index.js`** — open the file and add one line after the existing `router.post` calls:

```js
router.post('/<path>', require('./<name>Controller'));
```

   Use `router.post` for write operations and `router.get` for read-only queries. Verify the router file still exports `module.exports = router;` after editing.

3. **If the route does NOT use the Bot Framework adapter** (e.g., a plain JSON endpoint), omit `runWithRetry` and `adapter.onTurnError`. Keep `TRANSIENT_RE`, `AUTH_ERROR_RE`, and the try/catch in the handler — return `res.status(500).json({ message: err.message })` on failure.

4. **Smoke-test the route** with:
```bash
npm run dev
curl -X POST http://localhost:3978/api/<path> -H 'Content-Type: application/json' -d '{}'
```
   Expect a JSON response. A 404 means the route was not registered in `server/api/index.js`.

## Examples

**User says:** "Add a `/api/status` endpoint that returns the Redis connection state."

**Actions taken:**
1. Create `server/api/statusController.js` with the boilerplate above (no adapter needed — plain JSON route).
2. Implement handler: call `redis.ping()`, return `{ status: 'ok' }` or `{ status: 'error', message: err.message }`.
3. Add `router.get('/status', require('./statusController'));` to `server/api/index.js`.

**Result:** `GET /api/status` returns `{ "status": "ok" }` when Redis is reachable.

## Common Issues

- **404 on the new route**: The `require()` path in `server/api/index.js` is wrong or the line was added outside the router block. Verify `server/api/index.js` contains `router.post('/<path>', require('./<name>Controller'));` and that the file path matches exactly.
- **`Cannot find module '../bot/botActivityHandler'`**: Only `server/api/botController.js` imports the activity handler. Do not copy that import into a new controller unless the route runs bot turns.
- **`Error: Missing MicrosoftAppId`**: `.env` is not loaded. Confirm `server/index.js` calls `require('dotenv').config()` before any controller is required, or that the env vars are exported in the shell.
- **`authorization has been denied` on every request**: `MicrosoftAppId`/`MicrosoftAppPassword` in `.env` do not match the Azure Bot registration. The `AUTH_ERROR_RE` handler will swallow the error silently — check the console for `[onTurnError] Authorization error` logs.
- **Unhandled promise rejection in handler**: The top-level `async (req, res)` try/catch is missing. Every handler exported from a controller must wrap its body in `try { ... } catch (err) { res.status(500).json({ message: err.message }); }`.
