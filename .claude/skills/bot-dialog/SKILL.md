---
name: bot-dialog
description: Creates a new waterfall dialog following the server/dialogs/mainDialog.js pattern. Scaffolds a ComponentDialog subclass with WaterfallDialog, adds it to the dialog set, and wires up run(context, accessor). Use when user says 'add dialog', 'new dialog flow', 'create waterfall', or adds files to server/dialogs/. Do NOT use for modifying the OAuth/login flow in mainDialog.js directly.
---
# bot-dialog

## Critical

- New dialogs MUST extend `ComponentDialog` (or `LogoutDialog` if logout-interrupt is needed), never `TeamsActivityHandler`.
- `this.initialDialogId` MUST be set to the `WaterfallDialog` ID in the constructor or the dialog will never start.
- The `run(context, accessor)` method is required — `DialogBot` calls `this.dialog.run(context, this.dialogState)` directly.
- Dialog ID constants MUST be module-level `const` strings, not inline literals.
- Do NOT modify `server/dialogs/mainDialog.js` or `server/dialogs/logoutDialog.js` for new flows.

## Instructions

1. **Create a new dialog file in `server/dialogs/`** using this exact skeleton:
   ```js
   // Copyright (c) Microsoft Corporation. All rights reserved.
   // Licensed under the MIT License.

   const { ComponentDialog, DialogSet, DialogTurnStatus, WaterfallDialog } = require('botbuilder-dialogs');

   const <NAME>_DIALOG = '<Name>Dialog';
   const <NAME>_WATERFALL_DIALOG = '<Name>WaterfallDialog';

   class <Name>Dialog extends ComponentDialog {
       constructor() {
           super(<NAME>_DIALOG);

           this.addDialog(new WaterfallDialog(<NAME>_WATERFALL_DIALOG, [
               this.stepOne.bind(this),
               this.stepTwo.bind(this)
           ]));

           this.initialDialogId = <NAME>_WATERFALL_DIALOG;
       }

       async run(context, accessor) {
           const dialogSet = new DialogSet(accessor);
           dialogSet.add(this);
           const dialogContext = await dialogSet.createContext(context);
           const results = await dialogContext.continueDialog();
           if (results.status === DialogTurnStatus.empty) {
               await dialogContext.beginDialog(this.id);
           }
       }

       async stepOne(stepContext) {
           await stepContext.context.sendActivity('Step one.');
           return await stepContext.next();
       }

       async stepTwo(stepContext) {
           await stepContext.context.sendActivity('Done.');
           return await stepContext.endDialog();
       }
   }

   module.exports.<Name>Dialog = <Name>Dialog;
   ```
   Verify the file is saved in `server/dialogs/` with the correct name, and `module.exports.<Name>Dialog` matches the class name.

2. **Add prompts** if user input is needed. Import from `botbuilder-dialogs` and register in the constructor before the `WaterfallDialog`:
   ```js
   const { TextPrompt, WaterfallDialog, ... } = require('botbuilder-dialogs');
   const TEXT_PROMPT = 'TextPrompt';
   // inside constructor:
   this.addDialog(new TextPrompt(TEXT_PROMPT));
   ```
   Call via `stepContext.prompt(TEXT_PROMPT, 'Your question?')` in the step before the one that reads `stepContext.result`.
   Verify: every prompt constant is declared at module level and registered with `this.addDialog`.

3. **Wire the dialog into `DialogBot`** — in the file that instantiates `DialogBot` (typically `server/bot/teamsBot.js` or `server/index.js`), import and pass the new dialog:
   ```js
   const { TicketDialog } = require('../dialogs/ticketDialog');
   const dialog = new TicketDialog();
   const bot = new DialogBot(conversationState, userState, dialog);
   ```
   `DialogBot` will call `dialog.run(context, this.dialogState)` on every message.
   Verify: `this.dialog` in `DialogBot` points to the new instance.

4. **Run the linter** to catch unused variables and missing requires:
   ```bash
   npm run lint
   ```
   Fix any reported issues before testing.

5. **Smoke-test locally**:
   ```bash
   npm run dev
   ngrok http 3978 --host-header="localhost:3978"
   ```
   Update the Azure Bot messaging endpoint to `https://<tunnel>/api/messages`, then send a message in Teams to trigger the dialog.

## Examples

**User says**: "Add a ticket-collection dialog that asks for title and description"

**Actions**:
- Create `server/dialogs/ticketDialog.js`
- Constants: `TICKET_DIALOG`, `TICKET_WATERFALL_DIALOG`, `TEXT_PROMPT`
- Constructor: `addDialog(new TextPrompt(TEXT_PROMPT))`, then `WaterfallDialog` with `[askTitle, askDescription, confirm]`
- `askTitle`: `stepContext.prompt(TEXT_PROMPT, 'Enter ticket title:')`
- `askDescription`: reads `stepContext.result` as title, prompts for description
- `confirm`: reads description, sends summary, calls `stepContext.endDialog()`
- Export: `module.exports.TicketDialog = TicketDialog`
- Wire: `const { TicketDialog } = require('../dialogs/ticketDialog'); new DialogBot(cs, us, new TicketDialog())`

**Result**: Bot asks for title, then description, then confirms — all state managed by the waterfall step context.

## Common Issues

- **Dialog never starts / silent no-op**: `this.initialDialogId` is missing or doesn't match the string passed to `new WaterfallDialog(...)`. Verify both are identical.
- **`TypeError: this.dialog.run is not a function`**: The new dialog class is missing the `async run(context, accessor)` method. Copy it verbatim from the skeleton in Step 1.
- **`Error: DialogSet.add(): A dialog with an id of '<Name>Dialog' already exists`**: Two dialogs share the same ID string. Rename one of the module-level ID constants.
- **Step reads `stepContext.result` as `undefined`**: The previous step returned `stepContext.next()` instead of `stepContext.prompt(...)`. Ensure the prompting step returns the prompt call, and the reading step is the one immediately after.
- **Lint error `'DialogTurnStatus' is defined but never used`**: Only import what you use. If there's no `results.status` check, remove `DialogTurnStatus` from the destructure.
