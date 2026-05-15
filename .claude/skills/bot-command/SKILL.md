---
name: bot-command
description: Creates a new command module in server/commands/ following the { match(text, lcText, deps), execute(matchResult, deps) } shape used by all 21 existing modules. Use when user says 'add command', 'new bot command', 'add !something', 'create handler for X message', or when adding files to server/commands/. Handles admin-gate (ima !== 'admin'), deps destructuring, MessageFactory.text replies, registration in server/commands/index.js, and the required match() test in test/commands.test.js. Do NOT use for OAuth dialog logic (use bot-dialog skill), adaptive-card JSON authoring (use adaptive-card skill), card-submit handler routing changes inside TeamsActivityHandler (use bot-activity-handler skill), or editing the dispatcher in server/commands/index.js itself.
paths:
  - server/commands/*.js
  - test/commands.test.js
---
# Bot Command

## Critical

- Every command module MUST export exactly `{ match(text, lcText, deps), execute(matchResult, deps) }`. No other shape is dispatched by `server/commands/index.js`.
- `match()` MUST be synchronous, MUST NOT throw, and MUST return either a truthy object (the parsed data) or `null`.
- Admin-only commands MUST have `if (deps.ima !== 'admin') return null;` as the FIRST line inside `match()` — never inside `execute()`. Gating in execute leaks the command's existence to non-admins via dispatch ordering.
- Order in `server/commands/index.js` matters — first match wins. Place narrow regexes BEFORE broad ones (e.g. `blockEmail` before `blockHelp`).
- Every new module MUST be added to the `commands` array in `server/commands/index.js` AND have a `match()` test in `test/commands.test.js`. A module without registration is dead code; a module without a test breaks the lint/CI guarantee that regexes don't regress.
- Reply ONLY via `context.sendActivity(MessageFactory.text(...))` imported from `botbuilder`. Do not call `new BotFrameworkAdapter` — outbound proactive sends go through `server/lib/adapter.js`, not commands.
- `deps` is exactly `{ context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }`. Destructure only what you use. Do not add new fields to `deps` without updating `server/bot/botActivityHandler.js` where `dispatch()` is called.
- Use the project's ESLint style: 4-space indent, single quotes, semicolons, `${ var }` with spaces inside template curlies, `space-before-function-paren` named: never.

## Instructions

1. **Create the module file** at `server/commands/<camelCaseName>.js`. Match the existing naming: `ima.js`, `ping.js`, `blockEmail.js`, `githubIssues.js`. Use this skeleton:

   ```js
   const { MessageFactory } = require('botbuilder');

   module.exports = {
       match(text, lcText, { ima }) {
           if (ima !== 'admin') return null; // remove this line if command is public
           const m = lcText.match(/^yourcmd\s+(.+)$/i);
           return m ? { arg: m[1].trim() } : null;
       },
       async execute({ arg }, { context, db, redis, bot }) {
           await context.sendActivity(MessageFactory.text(`Result: ${ arg }`));
       }
   };
   ```

   Verify before proceeding: file path is exactly `server/commands/<name>.js`, exports the two-method object, and admin-gate (if applicable) is the first line of `match`.

2. **Choose match style based on existing precedent**:
   - Exact string (case-insensitive): `return lcText === 'ima' ? {} : null;` — see `server/commands/ima.js`.
   - One regex with captures: `const m = text.match(/^ping\s+(.+)$/i); return m ? { target: m[1].trim() } : null;` — see `server/commands/ping.js`.
   - Multiple sub-actions: return `{ action: 'list' | 'add' | 'remove', ... }` — see `server/commands/blockEmail.js`.
   - Card-submit handler (no text): inspect `deps.context.activity.value.msteams.type` and return `{ action }` — see `server/commands/ticketSubmit.js`.

   Verify before proceeding: regexes anchor with `^` and `$` to avoid matching mid-sentence (unless intentional like `ipLookup.js`), and `lcText` is preferred over `text` for case-insensitive matches.

3. **In `execute()`, destructure only the deps you need** from the second arg. Examples from the codebase:
   - DB query: `async execute({ ip }, { context, db }) { const [rows] = await db.query('SELECT ...', [ip]); }` — `server/commands/ipLookup.js`. Always use parameterised queries.
   - Redis: `async execute({ action, email }, { context, redis }) { await redis.sadd('blocked_emails', email); }` — `server/commands/blockEmail.js`.
   - Shell: `const { stdout } = await execFileAsync('ping', args, { timeout: 15000 });` — `server/commands/ping.js`. Use `execFileAsync` not `execAsync`; pass args as an array (never string interpolation) to avoid shell injection.
   - HTTP: `const axios = require('axios'); const res = await axios.post(url, new URLSearchParams(params), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });`
   - Bot helpers: `bot.isValidIP(target)`, `bot.isValidHostname(target)`, `bot.updateActionSubmitData(element, sentActivity)` — methods on the `TeamsActivityHandler` subclass.

   Verify before proceeding: every `await` is inside try/catch where I/O can fail (DB, shell, HTTP). Error replies use the format `⚠️ Error: ${ err.message }` or `❌ Invalid ...` — see `server/commands/ping.js`.

4. **Register the module** by adding `require('./<name>')` to the `commands` array in `server/commands/index.js`. Place narrow exact-string matchers (e.g. `ima`, `ping`) at the top; place catch-all/help matchers (`help`) at the very bottom. If the new command's regex could match strings that an existing command already handles, place yours BEFORE that command.

   Verify before proceeding: `npm test` still passes — first-match-wins ordering hasn't broken an existing test.

5. **Add a `match()` test** in `test/commands.test.js`. Import the module at the top of the file and add a `describe(...)` block with positive, negative, and (if admin-gated) non-admin cases. Pattern:

   ```js
   const yourCmd = require('../server/commands/yourCmd');

   describe('yourCmd command', () => {
       const deps = { ima: 'admin' };
       it('matches "yourcmd foo"', () => {
           const m = yourCmd.match('yourcmd foo', 'yourcmd foo', deps);
           assert.ok(m);
           assert.equal(m.arg, 'foo');
       });
       it('does not match unrelated text', () => {
           assert.equal(yourCmd.match('hello', 'hello', deps), null);
       });
       it('does not match for non-admin', () => {
           assert.equal(yourCmd.match('yourcmd foo', 'yourcmd foo', { ima: 'user' }), null);
       });
   });
   ```

   Verify with `npm test` — every described `it` must pass.

6. **Lint and run the full suite**:
   ```bash
   npm run lint
   npm test
   ```
   Verify before claiming done: zero lint errors, all `node:test` describes pass. Fix any indent/quote violations to match `eslint.config.js` (4 spaces, single quotes, `template-curly-spacing: always`).

## Examples

**User says:** "Add a `!uptime <host>` command that runs `uptime` on a remote host via ssh and returns the load average."

**Actions taken:**

1. Create `server/commands/uptime.js`:
   ```js
   const { MessageFactory } = require('botbuilder');

   module.exports = {
       match(text, lcText, { ima }) {
           if (ima !== 'admin') return null;
           const m = lcText.match(/^!?uptime\s+(\S+)$/i);
           return m ? { host: m[1] } : null;
       },
       async execute({ host }, { context, bot, execFileAsync }) {
           if (!bot.isValidHostname(host) && !bot.isValidIP(host)) {
               await context.sendActivity(MessageFactory.text(`❌ Invalid host: \`${ host }\``));
               return;
           }
           try {
               const { stdout } = await execFileAsync('ssh', ['-o', 'BatchMode=yes', '-o', 'ConnectTimeout=5', host, 'uptime'], { timeout: 10000 });
               await context.sendActivity(MessageFactory.text('```\n' + stdout.trim() + '\n```'));
           } catch (err) {
               await context.sendActivity(MessageFactory.text(`⚠️ Error: ${ err.stderr || err.message }`));
           }
       }
   };
   ```
2. Add `require('./uptime'),` to the `commands` array in `server/commands/index.js`, before `help` and after `ping` (similar shell-exec command).
3. Add to `test/commands.test.js`:
   ```js
   const uptime = require('../server/commands/uptime');
   describe('uptime command', () => {
       const deps = { ima: 'admin' };
       it('matches "uptime host1"', () => {
           const m = uptime.match('uptime host1', 'uptime host1', deps);
           assert.ok(m);
           assert.equal(m.host, 'host1');
       });
       it('does not match for non-admin', () => {
           assert.equal(uptime.match('uptime host1', 'uptime host1', { ima: 'user' }), null);
       });
   });
   ```
4. Run `npm run lint && npm test`.

**Result:** Admin users typing `uptime example.com` get the SSH `uptime` output in a code block. Non-admins get no response (the dispatcher falls through). Tests cover positive, negative-admin cases.

## Common Issues

- **"Command never fires for me but fires for admins"**: You forgot to remove the `if (ima !== 'admin') return null;` line for a public command, or `deps.ima` isn't `'admin'` for your account. Check `usersCollection.findOne({ email })` in `server/bot/botActivityHandler.js` — `ima` is read from the Mongo `users` doc.

- **"Command never matches even though regex looks right"**: `lcText` is already lowercased; matching `/Foo/` against `lcText` will fail. Either use `lcText.match(/foo/i)` or match against `text` for case-sensitive needs. Also confirm an earlier command in the array isn't swallowing the input — first match wins, so reorder in `server/commands/index.js`.

- **"`ReferenceError: MessageFactory is not defined`"**: Add `const { MessageFactory } = require('botbuilder');` at the top of the file. Do not destructure from `botbuilder-dialogs`.

- **"`TypeError: Cannot read properties of undefined (reading 'sendActivity')`"**: You destructured `{ context }` but the dispatcher passed `deps` positionally. Confirm `execute(matchResult, deps)` signature — `matchResult` is FIRST, `deps` is SECOND.

- **"ESLint fails with `Expected indentation of 8 spaces but found 4`"**: Block indent inside `execute` should be 8 spaces (4 for the function body inside the module-level object, plus 4 for the inner block). Match the precise indentation in `server/commands/blockEmail.js`.

- **"Test passes locally but `match()` returns truthy for both admin and non-admin"**: You wrote `deps.ima === 'admin'` in a positive check but never destructured `ima` from the third arg. The third arg is `deps` — write `match(text, lcText, { ima })` or `match(text, lcText, deps) { if (deps.ima !== 'admin') ...`.

- **"`db.query` throws `ER_NO_SUCH_TABLE`"**: The MySQL pool connects to `MYSQL_DB` from `.env`. Master tables (`backup_masters`, `website_masters`, `vps_masters`, `qs_masters`) live in that DB; ticket/billing tables may live elsewhere. Either fully-qualify the table (`zone-mta.users`) or confirm the schema name in `validateEnv.js`.

- **"`redis` calls hang forever in tests"**: `match()` tests should never call into `redis`/`db`/`execFileAsync`. If they do, the design is wrong — move I/O into `execute()`. Tests in `test/commands.test.js` only exercise `match()`.