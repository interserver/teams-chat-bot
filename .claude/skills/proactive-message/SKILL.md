---
name: proactive-message
description: Implements proactive messaging to a Teams conversation using the Redis conversation-reference pattern in `server/api/msgController.js`. Covers storing `convref:{conversationId}` keys in Redis via `ioredis`, retrieving them, and calling `sendProactiveMessage()` via `BotFrameworkAdapter.continueConversation`. Use when user says 'send proactive message', 'notify channel', 'push message to Teams', or modifies `server/api/msgController.js`. Do NOT use for reactive bot replies (those belong in `server/bot/botActivityHandler.js`).
---
# Proactive Message

## Critical

- The bot MUST have already received a message from the target conversation before a proactive message can be sent. `convref:{conversationId}` is written to Redis in `server/bot/botActivityHandler.js` inside `this.onMessage` — if this key is missing, the lookup returns `null` and you get a 404.
- Never skip `MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl)` on auth-error retries — without it the next attempt will also fail.
- Redis key format is always `convref:{conversationId}` (e.g. `convref:19:0c93975aae904b7db892891da3065c33@thread.v2`).

## Instructions

1. **Store the conversation reference on every incoming message** in `server/bot/botActivityHandler.js` inside `this.onMessage`:
   ```js
   const conversationReference = TurnContext.getConversationReference(context.activity);
   await this.redis.set(`convref:${conversationReference.conversation.id}`, JSON.stringify(conversationReference));
   ```
   Verify: run `redis-cli get 'convref:<id>'` on `dragonfly.mailbaby.net:6379` and confirm JSON is stored.

2. **Define `sendProactiveMessage`** in `server/api/msgController.js` using `BotFrameworkAdapter.continueConversation` with retry logic (mirrors the existing pattern):
   ```js
   const { BotFrameworkAdapter } = require('botbuilder');
   const { MicrosoftAppCredentials } = require('botframework-connector');
   const Redis = require('ioredis');

   const TRANSIENT_RE = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i;
   const AUTH_ERROR_RE = /authorization has been denied|401|unauthorized/i;
   const MAX_RETRIES = 2;
   const RETRY_DELAY_MS = 1000;

   function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

   const adapter = new BotFrameworkAdapter({
       appId: process.env.MicrosoftAppId,
       appPassword: process.env.MicrosoftAppPassword
   });

   async function sendProactiveMessage(conversationReference, messageText) {
       for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
           try {
               await adapter.continueConversation(conversationReference, async (proactiveContext) => {
                   await proactiveContext.sendActivity(messageText);
               });
               return;
           } catch (err) {
               const msg = err.message || '';
               const isTransient = TRANSIENT_RE.test(msg);
               const isAuth = AUTH_ERROR_RE.test(err);
               if ((isTransient || isAuth) && attempt < MAX_RETRIES) {
                   console.warn(`[proactive retry] Attempt ${attempt} failed (${msg}), retrying in ${RETRY_DELAY_MS}ms...`);
                   if (isAuth) MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl);
                   await sleep(RETRY_DELAY_MS * attempt);
                   continue;
               }
               throw err;
           }
       }
   }
   ```
   Verify: function signature and retry constants match `server/api/botController.js` exactly.

3. **Wire up `adapter.onTurnError`** immediately after creating the adapter (same shape as `server/api/botController.js`):
   ```js
   adapter.onTurnError = async (context, error) => {
       const errorMsg = error.message || 'Oops. Something went wrong!';
       console.error(`\n [onTurnError] unhandled error: ${ error }`);
       if (TRANSIENT_RE.test(errorMsg)) { console.error('[onTurnError] Transient network error, skipping reply to user.'); return; }
       if (AUTH_ERROR_RE.test(errorMsg)) { console.error('[onTurnError] Authorization error — check MicrosoftAppId/MicrosoftAppPassword and bot registration.'); return; }
       try {
           await context.sendTraceActivity('OnTurnError Trace', `${ error }`, 'https://www.botframework.com/schemas/error', 'TurnError');
           await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${ errorMsg }`);
       } catch (sendError) { console.error(`[onTurnError] Failed to send error message to user: ${ sendError.message }`); }
   };
   ```

4. **Create the Redis client and implement `msgHandler`**:
   ```js
   const redis = new Redis({ host: 'dragonfly.mailbaby.net', port: 6379 });
   redis.on('connect', () => console.log("✅ Connected to Redis"));
   redis.on('error', (err) => console.error("❌ Redis error:", err));

   const msgHandler = async (req, res) => {
       const targetConversationId = '19:<your-thread-id>@thread.v2';
       try {
           const stored = await redis.get(`convref:${targetConversationId}`);
           if (stored) {
               const conversationReference = JSON.parse(stored);
               await sendProactiveMessage(conversationReference, req.body.message);
               res.json({ message: "sent" });
           } else {
               res.status(404).json({ message: "no conversation reference found" });
           }
       } catch (err) {
           console.error('[msgHandler] Error sending proactive message:', err.message);
           res.status(500).json({ message: err.message });
       }
   };

   module.exports = msgHandler;
   ```
   Verify: `POST /api/message` is already registered in `server/api/index.js` — do not add another route.

5. **Test end-to-end**:
   ```bash
   curl -X POST http://localhost:3978/api/message \
     -H 'Content-Type: application/json' \
     -d '{"message": "Hello from proactive!"}'
   # Expected: {"message":"sent"}
   # If 404: the bot hasn't received a message in that channel yet — send any message to the bot first.
   ```

## Examples

**User says:** "Add a proactive endpoint that pings the `int-dev-private` channel when a deploy completes."

**Actions taken:**
1. Confirm the conversation reference is stored by sending a test message in the channel and verifying the Redis key exists.
2. In `server/api/msgController.js`, set `targetConversationId` to the `int-dev-private` channel's conversation ID.
3. Call `sendProactiveMessage(conversationReference, req.body.message)` inside `msgHandler`.
4. `POST /api/message` with `{ "message": "Deploy complete ✅" }` → bot posts in channel.

**Result:** Channel receives the message without any user @mention trigger.

## Common Issues

- **`404 no conversation reference found`**: The bot has not yet received a message in that conversation. Send any message to the bot in the target channel first, then retry.
- **`401 / authorization has been denied`**: `MicrosoftAppId` or `MicrosoftAppPassword` in `.env` is wrong, or the service URL is not trusted. The retry loop calls `MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl)` automatically — if it persists, verify the Azure Bot registration credentials.
- **`ECONNRESET` / `socket hang up`**: Transient network error. The `sendProactiveMessage` retry loop (max 2 attempts, 1 s delay) handles this automatically. If it persists, check connectivity to `botframework.com`.
- **Redis `ECONNREFUSED`**: Confirm the Redis host `dragonfly.mailbaby.net:6379` is reachable: `redis-cli -h dragonfly.mailbaby.net -p 6379 ping` should return `PONG`.
- **`Cannot read properties of null (reading 'serviceUrl')`**: `JSON.parse(stored)` returned a malformed reference. Delete the key with `redis-cli del 'convref:<id>'` and let the bot re-store it on the next incoming message.
