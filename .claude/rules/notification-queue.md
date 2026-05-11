---
paths:
  - server/queue/**
---

# Notification Queue Rules

- Redis keys use prefix `process.env.NOTIF_KEY_PREFIX || 'notif:'`. Key helper: `k(name) => KEY_PREFIX + name`.
- Queue: `notif:queue` (LPUSH/BRPOPLPUSH). In-flight: `notif:processing`. Failed: `notif:dead` (with `_dead_reason`, `_dead_at`).
- Envelope JSON: `{ channel, text|attachments, extra?: { dedup_key }, fallback_webhook_url? }`. `channel` MUST resolve via `CHANNELS`.
- Trackable mode: envelopes with `extra.dedup_key` edit the most recent activity within `NOTIF_EDIT_WINDOW_MS` (30 min). After 3 appends, summarize as `N updates · last HH:MM`.
- GitHub commit grouping: events carrying a commit SHA auto-get `dedup_key=github:commit:{sha7}` if no key present. Window: `NOTIF_COMMIT_GROUP_WINDOW_MS` (3 min).
- Filters (`server/queue/filters.js`): drop `star`/`watch`/`fork`/`ping`; drop successful `check_run`/`workflow_job`; redirect via `LOW_SIGNAL_GITHUB_EVENTS`/`SUCCESSFUL_*_CONCLUSIONS` sets.
- Metrics counters: increment `notif:metrics:{enqueued|sent|edited|coalesced|redirected|fallback|dead}` on the corresponding event.
- Recovery: on startup move `notif:processing` items back to `notif:queue`.
