---
paths:
  - server/lib/**
---

# Shared Lib Conventions

- `adapter.js` owns the single `BotFrameworkAdapter` (`getAdapter()` / `setAdapter()`); `redis.js` owns `createBotRedis()` / `createNotifRedis()` / `TEAMS_SERVICE_URL`; `retry.js` owns `runWithRetry(fn, { label, serviceUrl, maxRetries })`. Never construct those clients outside these modules.
- `messageSize.js` owns every outbound size guard. Teams rejects activities whose serialized JSON exceeds ~25 KB and the SDK swallows that `RestError` in `onTurnError`, so size MUST be measured before the send — not recovered from after it.
- Measure with `activityBytes()` (`Buffer.byteLength(JSON.stringify(...), 'utf8')`). NEVER use JS `.length`: it counts UTF-16 code units and ignores JSON escaping.
- Budget is `MAX_ACTIVITY_BYTES` (default 24000, override via `TEAMS_MAX_ACTIVITY_BYTES`). Keep headroom under 25 KB — `continueConversation` stamps on from/conversation/recipient fields the caller-built shape never sees.
- Cut text through `truncateTextToBytes` / `fitTextActivity` / `safeCutIndex` / `charCut` — never raw `slice()`, which can emit a lone surrogate. Budget against `textJsonBytes`, not raw byte length.
- Cards: `trimCard()` returns a trimmed clone (never mutate the input); `fitCardActivity()` measures one pass; `fitAttachmentsActivity()` drops what will not fit and returns a `note` for the caller to surface.
- `splitTextIntoChunks()` is lossless — concatenating the chunks MUST reproduce the input, since callers map delivered chunks back onto queue items. Every recursive split must make strict progress (no ping-pong on trailing separators).
- Add new helpers to `module.exports` and cover them in `test/messageSize.test.js`; consumer-level behaviour goes in `test/notifOversizedGuard.test.js`.
