---
paths:
  - server/commands/**
---

# Command Module Pattern

- Export `{ match(text, lcText, deps), execute(matchResult, deps) }` — no other shape is dispatched by `server/commands/index.js`.
- `match` returns the parsed data object (truthy) or `null`. Never throw from `match`.
- Admin-only: `if (deps.ima !== 'admin') return null;` as the FIRST line of `match`.
- `deps` is `{ context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }` — destructure only what you use.
- Reply via `context.sendActivity(MessageFactory.text(...))` (import from `botbuilder`).
- Card-submit handlers read `context.activity.value.msteams.type` — see `server/commands/ticketSubmit.js`.
- Update card edits via `context.updateActivity({ type: 'message', id: context.activity.value.activityId, conversation: context.activity.conversation, text/attachments })`.
- Register every new module in the `commands` array in `server/commands/index.js`. Order matters — first match wins.
- Add a `match()` test for every regex in `test/commands.test.js`.
- Use `axios` (already a dep) for HTTP. URL-encoded form posts: `new URLSearchParams(params)` + `Content-Type: application/x-www-form-urlencoded` (see `server/commands/ticketQuick.js`, `ticketSubmit.js`).
- DB queries via `deps.db.query(sql, params)` with `?`/`??` placeholders (mysql2 promise pool). Mongo via `deps.usersCollection.findOne`/`insertOne`/`deleteOne`. Redis via `deps.redis.get`/`set`/`sadd`/`srem`/`smembers`.
- Multi-action commands (`githubIssues.js`, `githubLabels.js`, `notifAdmin.js`) use a `{ action: '...', ...payload }` match shape + `switch` in `execute`.
