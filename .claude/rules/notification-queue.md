---
paths:
  - server/queue/**
---

# Notification Queue Rules

- Redis keys use prefix `process.env.NOTIF_KEY_PREFIX || 'notif:'`. Key helper: `k(name) => KEY_PREFIX + name`.
- Queue: `notif:queue` (LPUSH/BRPOPLPUSH). In-flight: `notif:processing`. Failed: `notif:dead` (with `_dead_reason`, `_dead_at`).
- Envelope JSON: `{ channel, text|attachments, extra?: { dedup_key, event_type, _commit_sha, data }, fallback_webhook_url? }`. `channel` MUST resolve via `CHANNELS`.
- Trackable mode: envelopes with `extra.dedup_key` edit the most recent activity within `NOTIF_EDIT_WINDOW_MS` (30 min). After 3 appends, summarize as `N updates · last HH:MM`.
- GitHub commit grouping: events carrying a commit SHA auto-get `dedup_key=github:commit:{sha7}`. Window: `NOTIF_COMMIT_GROUP_WINDOW_MS` (3 min).
- PR-context attachment (`attachPrContext`): rewrites `dedup_key` to `github:pr:{repo}:{n}` for pushes/deletes/creates on a branch in `notif:prbranch:{repo}:{branch}` AND for `issue_comment` whose `issue.pull_request` is set OR `issue.html_url` matches `/pull/{n}`. Runs AFTER `normalizeGithubDedup` and BEFORE `isActionTriggeredPush`.
- Within a single tick, items sharing a `dedup_key` MUST be folded via `groupTrackableByDedup` → `canBatchMergeGroup` → `handleTrackableBatch` (1 API call) — NOT iterated sequentially.
- Filters (`server/queue/filters.js`): drop `star`/`watch`/`fork`/`ping`; drop successful `check_run`/`workflow_job`; drop `${{ matrix.* }}` placeholders. Use `LOW_SIGNAL_GITHUB_EVENTS`/`SUCCESSFUL_CHECK_CONCLUSIONS`/`SUCCESSFUL_WORKFLOW_CONCLUSIONS` sets.
- Announce redirect: `NOTIF_ANNOUNCE_REPOS` (comma-list of `owner/*` or `owner/repo`) → `int-dev-announce`; `NOTIF_ANNOUNCE_REPOS_EXCLUDE` exempts. Exclude wins. Use `decideAnnounceRedirect(repo, listRaw, excludeRaw)`.
- Metrics counters: increment `notif:metrics:{enqueued|sent|edited|coalesced|redirected|fallback|dead}` on the corresponding event.
- Recovery on startup: move `notif:processing` items back to `notif:queue`.
- Trace every decision via `trace.emit(kind, payload)` from `notifTrace.js` — match positive and negative branches (e.g. `announce_redirect` + `announce_excluded`).
- Test mock seam: `_setInternalsForTest({ redis, redisBot, adapter })` swaps module handles. Restore in `after()`.
