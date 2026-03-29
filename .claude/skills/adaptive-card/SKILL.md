---
name: adaptive-card
description: Creates a new Adaptive Card JSON file in `cards/` following the v1.5 schema used in `cards/add_ticket.json`. Handles ColumnSet layouts, Input.ChoiceSet, Input.Text, Action.Submit with `data.msteams.type` routing, and Action.ToggleVisibility. Use when user says 'add card', 'create adaptive card', 'new card', or adds files to `cards/`. Do NOT use for modifying existing card handler logic in `server/bot/`.
---
# adaptive-card

## Critical

- **Every `Action.Submit` must include a `msteams.type` field in its `data` object** — this is the routing key checked in `server/bot/botActivityHandler.js`. The handler reads this value to dispatch the correct branch. Missing it means the bot will never handle the submission.
- **Every `Action.Submit` must also carry `data.activityId`** — the bot's `updateActionSubmitData()` helper injects this at send time to enable `context.updateActivity()`. Do not hardcode it; leave `activityId` absent from the JSON file itself.
- Schema version must be `"version": "1.5"` — the existing card and Teams client both target v1.5.
- Card files live in the `cards/` directory. The bot loads them with `fs.readFileSync` and parses them as JSON, so the file must be valid JSON (no comments, no trailing commas).

## Instructions

1. **Name the file** using kebab-case in the `cards/` directory (e.g., `cards/close_ticket.json`). Verify `cards/` exists with `ls cards/` before creating.

2. **Scaffold the root envelope** — always this exact shape:
   ```json
   {
     "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
     "type": "AdaptiveCard",
     "version": "1.5",
     "body": []
   }
   ```

3. **Add a header row** using a nested `ColumnSet` with an `Icon` + `TextBlock` pair, matching the pattern in `cards/add_ticket.json` lines 3–50:
   ```json
   {
     "type": "ColumnSet",
     "spacing": "None",
     "columns": [{
       "type": "Column", "width": "stretch",
       "items": [{
         "type": "ColumnSet", "spacing": "ExtraSmall",
         "columns": [
           { "type": "Column", "width": "auto", "verticalContentAlignment": "Center",
             "items": [{ "type": "Icon", "name": "<FluentIconName>", "color": "Accent", "size": "Small" }] },
           { "type": "Column", "width": "stretch", "spacing": "ExtraSmall", "verticalContentAlignment": "Center",
             "items": [{ "type": "TextBlock", "text": "<Title>", "wrap": true, "weight": "Bolder", "color": "Accent", "size": "Large" }] }
         ]
       }]
     }]
   }
   ```

4. **Add `Input.ChoiceSet` fields** inside `ColumnSet` columns for side-by-side dropdowns. Required fields: `id`, `label`, `placeholder`, `choices` (array of `{title, value}`), `isRequired: true`, `errorMessage`, `value` (default), `spacing: "None"`.

5. **Add `Input.Text` fields** at the body level for single-line or multiline text. Required fields: `id`, `label`, `placeholder`, `errorMessage`, `spacing: "None"`. Add `"isMultiline": true` for multi-line.

6. **Add optional toggle sections** using `Action.ToggleVisibility` on a `TextRun` inside a `RichTextBlock`. The toggled element must have a matching `id` and `"isVisible": false` in the JSON. Include paired `chevronDown`/`chevronUp` `Icon` elements with `isVisible: false` on the up-chevron.

7. **Add the `ActionSet`** inside a right-aligned `ColumnSet` column:
   ```json
   {
     "type": "ActionSet",
     "separator": true,
     "spacing": "None",
     "horizontalAlignment": "Right",
     "actions": [
       { "type": "Action.Submit", "title": "Submit", "style": "positive",
         "data": { "msteams": { "type": "<cardNameSubmit>" } } },
       { "type": "Action.Submit", "title": "Cancel", "style": "destructive",
         "data": { "msteams": { "type": "<cardNameCancel>" } } }
     ]
   }
   ```
   Use camelCase for `msteams.type` values (e.g., `addTicketSubmit`, `addTicketCancel`).

8. **Add bot handler branches** in `server/bot/botActivityHandler.js` inside the `this.onMessage` handler. Check the `msteams.type` field in `context.activity.value` and read field values from the same object. Use `context.updateActivity()` to replace the card on success (see lines 201–206 and 227–231 for the exact pattern).

9. **Add the trigger command** — add an `else if (lcText === '<trigger phrase>')` branch that reads the card via `fs.readFileSync`, sends it, then calls `this.updateActionSubmitData(element, sentActivity)` recursively to inject `activityId` before updating the activity. Follow lines 248–263 exactly.

   Verify the new card file parses cleanly: `node -e "JSON.parse(require('fs').readFileSync('cards/add_ticket.json','utf8'))"` (substitute `add_ticket.json` with your card's filename) — must exit 0.

## Examples

**User says:** "add a card for closing a ticket"

**Actions taken:**
1. Create `cards/close_ticket.json` with root envelope, header (`Icon: "CheckMark"`, title `"Close Ticket"`), an `Input.ChoiceSet` for `reason` (id: `reason`), an `Input.Text` for `note` (id: `note`, isMultiline: true), and an `ActionSet` with `msteams.type: "closeTicketSubmit"` / `"closeTicketCancel"`.
2. In `server/bot/botActivityHandler.js`, add handler for `msteams.type === "closeTicketSubmit"` reading `context.activity.value.reason` and `context.activity.value.note`.
3. Add `else if (lcText === 'close ticket')` branch to send the card.

**Result:** User types "close ticket" in Teams → card renders → submit fires → `closeTicketSubmit` branch processes fields.

## Common Issues

- **Card renders but Submit does nothing / bot never enters the handler**: The `msteams.type` field is missing or misspelled in the card JSON's `Action.Submit` data. Confirm the value exactly matches the string checked in `server/bot/botActivityHandler.js`.
- **`context.updateActivity()` throws "Activity not found"**: `activityId` was not injected. Ensure `updateActionSubmitData()` is called after the initial `sendActivity` and before `updateActivity`, as in lines 256–263.
- **`SyntaxError: Unexpected token` on startup**: Trailing comma or JS comment in the JSON file. Run `node -e "JSON.parse(require('fs').readFileSync('cards/add_ticket.json','utf8'))"` (substitute your card's filename) to locate the line.
- **Dropdown shows but value is not submitted**: The `Input.ChoiceSet` is missing an `id` field. Every input element needs a unique `id` that becomes the key in the activity value object.
- **Toggle section never shows**: The `targetElements` array in `Action.ToggleVisibility` must exactly match the `id` strings of the target elements. Typo in one id will silently fail.
