---
name: bot-activity-handler
description: Extends `server/bot/botActivityHandler.js` (TeamsActivityHandler subclass) with new activity handlers — message reactions, invoke actions, card submit handling, member events. Use when user says 'handle activity', 'respond to card submit', 'add bot event handler', or modifies files in `server/bot/`. Do NOT use for dialog logic (OAuthPrompt, WaterfallDialog) — use the bot-dialog skill instead.
---
# bot-activity-handler

## Critical

- Every handler registered in the constructor **must** call `await next()` as its final statement — omitting it silently swallows all subsequent middleware.
- Card submit actions arrive as **message activities**, not invoke activities. Check the `msteams.type` field within `context.activity.value` inside `onMessage`, not `onInvokeActivity`.
- After sending an Adaptive Card, call `updateActionSubmitData()` to inject `activityId` into every `Action.Submit` so the bot can later call `context.updateActivity()` to replace the card.
- Never import from `botbuilder-dialogs` in this file — that belongs in `server/dialogs/`.

## Instructions

1. **Identify the target handler type.** Determine which Bot Framework lifecycle hook applies:
   - Incoming text/card-submit → `this.onMessage`
   - Member joins/leaves → `this.onMembersAdded` / `this.onMembersRemoved`
   - Reaction added/removed → `this.onReactionsAdded` / `this.onReactionsRemoved`
   - Message deleted/updated → `this.onMessageDelete` / `this.onMessageUpdate`
   - Conversation update → `this.onConversationUpdate`
   - Token response → `this.onTokenResponseEvent`
   Verify the hook name exists in `BotActivityHandler` constructor in `server/bot/botActivityHandler.js` before proceeding.

2. **Add the handler inside the constructor** of `BotActivityHandler` in `server/bot/botActivityHandler.js`. Follow this exact skeleton:
   ```js
   this.onReactionsAdded(async (context, next) => {
       console.log('got onReactionsAdded event');
       console.log(context.activity);
       // your logic here
       await next();
   });
   ```
   Verify `await next()` is the last line before the closing `}`.

3. **For card submit actions**, add a branch inside the existing `this.onMessage` handler. Card action type is found in `context.activity.value` as the `msteams.type` field. Pattern:
   ```js
   } else if (context.activity.value?.msteams?.type === 'myActionSubmit') {
       const { field1, field2 } = context.activity.value;
       try {
           // process submission
           await context.updateActivity({
               type: 'message',
               id: context.activity.value.activityId,
               conversation: context.activity.conversation,
               text: 'Submission received'
           });
       } catch (error) {
           await context.sendActivity(MessageFactory.text(error.message));
           console.log('❌ Request failed:', error.message);
       }
   }
   ```
   The `activityId` is injected by `updateActionSubmitData()` — verify the card JSON's `Action.Submit` nodes include `data.activityId` after the send.

4. **For sending a new Adaptive Card**, load from `cards/`, send once to get the activity ID, patch the card, then update:
   ```js
   const cardPath = path.join(__dirname, '../../cards/my_card.json');
   let cardContents = JSON.parse(fs.readFileSync(cardPath, 'utf8'));
   const sentActivity = await context.sendActivity({
       attachments: [CardFactory.adaptiveCard(cardContents)]
   });
   cardContents.body.forEach(el => this.updateActionSubmitData(el, sentActivity));
   await context.updateActivity({
       type: 'message',
       id: sentActivity.id,
       conversation: context.activity.conversation,
       attachments: [CardFactory.adaptiveCard(cardContents)]
   });
   ```
   Verify the card file exists in `cards/` and its `Action.Submit` nodes have a `msteams.type` field matching your handler branch.

5. **For helper methods**, add them as class methods outside the constructor:
   ```js
   myHelper(param) {
       // pure logic, no async DB calls unless truly needed
   }
   ```
   Verify the method is outside the constructor closing `}` but inside the class closing `}`.

6. **Run and verify**: `npm run dev`, trigger the activity in Teams or via Bot Framework Emulator, and confirm the `console.log` output matches the handler name.

## Examples

**User says**: "Add a handler that logs when a message is deleted and sends a DM to the deleter."

**Actions taken**:
1. Locate the existing `this.onMessageDelete` stub in `server/bot/botActivityHandler.js:531`.
2. Replace the stub body with logic to extract `context.activity.from.id` and call `context.sendActivity`.
3. Keep `await next()` as last line.

**Result**:
```js
this.onMessageDelete(async (context, next) => {
    console.log('got onMessageDelete event');
    const deleterId = context.activity.from.id;
    await context.sendActivity(MessageFactory.text(`Message deleted by ${deleterId}`));
    await next();
});
```

## Common Issues

- **Card submit never fires**: The card's `Action.Submit` is missing `"data": { "msteams": { "type": "yourActionName" } }`. Check `cards/your_card.json` and ensure that field is present.
- **`context.updateActivity` throws `Activity not found`**: `activityId` was not injected into the card before it was sent. Verify `updateActionSubmitData()` is called after `context.sendActivity` and before `context.updateActivity`.
- **Handler body runs but `await next()` missing → downstream middleware broken**: Add `await next()` as the last line inside every `this.onXxx` callback.
- **`TypeError: Cannot read properties of undefined (reading 'type')` when reading the msteams type**: Use optional chaining: `context.activity.value?.msteams?.type`.
- **Reaction events not received**: RSC permission `ChannelMessage.Read.Group` must be declared in `appManifest/manifest.json` under `authorization.permissions.resourceSpecific`. Verify with `cat appManifest/manifest.json | grep -A5 resourceSpecific`.
