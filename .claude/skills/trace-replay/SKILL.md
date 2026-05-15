---
name: trace-replay
description: Debugs notification-grouping bugs via `scripts/replay-notif.js` reading `.logs/notif-trace-YYYY-MM-DD.jsonl`. Covers `--mode timeline|grouped|raw`, filters (`--dedup`, `--commit`, `--room`, `--tick`, `--event`, `--kind`, `--since/--until`, `--activity`), and signature kinds (`recent_lookup`, `edit_skipped_no_convref`, `edit_fell_through`, `recent_saved`, `announce_redirect`, `batch_merge`, `batch_skipped`). Use when user reports 'commits not grouping', 'new message instead of edit', 'trace this dedup key', 'why did this fall through', or asks to inspect `notif:queue` historical behavior. Do NOT use for live-queue admin (use `!notif status`/`drain-dead`) or for non-queue bug investigation.
paths:
  - .logs/notif-trace-*.jsonl
  - scripts/replay-notif.js
  - server/queue/notifTrace.js
  - server/queue/notificationConsumer.js
---
# Trace Replay (Notification Queue)

Diagnose why notifications grouped, edited, redirected, or fell through to a new message by replaying recorded JSONL traces from `notificationConsumer`.

## Critical

- The trace files live in `.logs/notif-trace-YYYY-MM-DD.jsonl`. Each line is one JSON event written by `server/queue/notifTrace.js`. `.logs/` is gitignored — files only exist on the host that ran the consumer.
- Never edit the trace files. They are append-only forensic evidence; mutating them invalidates `--mode grouped` reconstruction.
- Trace events do NOT contain message bodies in full — only signatures, dedup_keys, room names, commit SHAs, activityIds, and the `kind` of decision the consumer made. If the user expects to see full message text, tell them traces redact bodies.
- This skill investigates **past** behavior. For *current* queue depth, dead-letter contents, or rooms with cached convrefs, use the `!notif` admin commands (`!notif status`, `!notif rooms`, `!notif drain-dead`, `!notif wfactive`) — NOT this skill.
- All times in trace files are ms epoch (`ts` field). When reporting findings to the user, convert to local time so the timeline reads naturally.

## Instructions

1. **Confirm a trace file exists for the window in question.** Ask the user for the approximate date/time of the bad grouping behavior, then check:
   ```bash
   ls -lah .logs/notif-trace-*.jsonl
   ```
   If no file exists for the relevant date, stop — there is nothing to replay. Suggest enabling tracing going forward (the consumer writes traces automatically when `notifTrace.js` is wired; verify by reading `server/queue/notificationConsumer.js` for `trace(` calls). **Verify file exists before proceeding.**

2. **Run the broadest sensible query first** to see the shape of the data. Use `--mode timeline` with the narrowest known anchor (a `dedup_key`, commit SHA, room, or activityId):
   ```bash
   node scripts/replay-notif.js --mode timeline --dedup 'github:commit:abc1234'
   node scripts/replay-notif.js --mode timeline --commit abc1234de56
   node scripts/replay-notif.js --mode timeline --room int-dev-private --since '2026-05-15T14:00' --until '2026-05-15T15:00'
   node scripts/replay-notif.js --mode timeline --activity 1715787600123
   ```
   `--mode timeline` prints events in chronological order with `ts`, `kind`, `dedup_key`, `room`, `activityId`. **Verify the events for the reported incident appear before proceeding.** If nothing prints, broaden filters (drop `--since/--until`, try `--event push` or `--event check_run`).

3. **Identify the decision kind that explains the bug.** Read the `kind` column for each event and map it to behavior:

   | `kind` | Meaning | Look for |
   |---|---|---|
   | `recent_lookup` | Consumer queried Redis for a recent activityId on this `dedup_key` | `hit:true` (edit path) vs `hit:false` (new message path) |
   | `recent_saved` | Consumer stored a new `activityId` against the `dedup_key` for future edits | Confirms a baseline message was sent |
   | `edit_skipped_no_convref` | Edit path aborted because no `convref:{conversationId}` was cached | Bot never observed inbound activity in that room — fix by triggering any inbound message or running `syncConversationReferences()` |
   | `edit_fell_through` | Edit attempt failed at `continueConversation` time; consumer sent a new message instead | Inspect `err` field — auth/transient/rateLimit classification from `server/lib/retry.js` |
   | `announce_redirect` | Message was redirected to `int-dev-announce` by `filters.js` (low-signal GitHub event) | Confirms expected filtering; the user may want the filter disabled or the event upgraded |
   | `batch_merge` | Multiple envelopes merged into one coalesced send within the same tick | `count` field shows how many; check `NOTIF_COALESCE_MAX_ITEMS`/`MAX_CHARS` if unexpected |
   | `batch_skipped` | A candidate envelope was excluded from a batch (e.g. different room, dedup conflict) | Cross-check the skipped item's `dedup_key` against the batch's |

   **Verify which kind explains the user's report before proposing a fix.** If the user reports "commits not grouping" and the timeline shows `recent_lookup hit:false` followed by `recent_saved` for the same `dedup_key`, the edit window expired — check `NOTIF_EDIT_WINDOW_MS` (default 1800000) and `NOTIF_COMMIT_GROUP_WINDOW_MS` (default 180000).

4. **Switch to `--mode grouped` to reconstruct the activity-level story.** Once you know the offending `dedup_key` or `activityId`, run:
   ```bash
   node scripts/replay-notif.js --mode grouped --dedup 'github:commit:abc1234'
   node scripts/replay-notif.js --mode grouped --activity 1715787600123
   ```
   This collapses per-tick events under each `activityId` so you can see the lifecycle: initial send → N edits → expiry. **Verify the lifecycle matches what the user saw in Teams.** A common mismatch: trace shows successful edits but Teams shows N separate messages — the user is looking at a different `dedup_key` (e.g. a workflow_job for a different commit). Loop back to Step 3 with the correct anchor.

5. **Use `--mode raw` only when timeline/grouped omit a field you need.** Raw prints the full JSON object per line — useful for inspecting `extra`, `err.code`, classification, or fallback_webhook_url. Pipe to `jq` for filtering:
   ```bash
   node scripts/replay-notif.js --mode raw --dedup 'github:commit:abc1234' | jq 'select(.kind == "edit_fell_through")'
   ```

6. **Combine filters to narrow noisy days.** Filters AND together:
   ```bash
   node scripts/replay-notif.js --mode timeline \
     --room int-dev-private \
     --event workflow_job \
     --kind edit_fell_through \
     --since '2026-05-15T00:00' --until '2026-05-15T23:59'
   ```
   `--tick N` isolates a single consumer tick (one poll cycle) — useful when batch_merge is suspect. `--event` matches the originating GitHub event type stored in the envelope's `extra.event`.

7. **Cross-reference findings with code.** After identifying the `kind`, open the matching code path in `server/queue/notificationConsumer.js` (search for the literal `kind` string in `trace(` calls) and explain to the user *why* the consumer made that decision. **Quote file path + line numbers** in your reply so the user can jump straight there.

8. **Report findings to the user as a short timeline.** Format:
   - `14:02:11` — push event `github:commit:abc1234` → `recent_lookup hit:false` → `recent_saved activity=...`
   - `14:02:14` — workflow_job → `recent_lookup hit:true` → edit OK
   - `14:05:33` — workflow_job → `recent_lookup hit:false` (window expired) → new message
   Then state the root cause in one sentence and link to the line in `notificationConsumer.js` that produced the final `kind`.

## Examples

**Example 1 — User says: "Three commits to detain/sugarcraft just posted as 3 separate messages instead of grouping"**

Actions:
1. `ls .logs/notif-trace-2026-05-15.jsonl` → file exists.
2. `node scripts/replay-notif.js --mode timeline --room int-dev-private --event push --since '2026-05-15T13:50' --until '2026-05-15T14:10'` → shows three `push` events, each with a different `dedup_key` (`github:commit:aaa1111`, `bbb2222`, `ccc3333`) and `kind=recent_saved` for each.
3. Diagnosis: each push had a unique SHA, so the consumer correctly treated them as three independent trackables. The user expected commit grouping, but commit grouping only collapses **the same SHA's** push + workflow events into one message — not three separate commits into one.
4. Reply: "Trace shows 3 distinct `dedup_key`s (one per commit SHA). The grouping logic at `server/queue/notificationConsumer.js` injects `github:commit:{sha7}` per commit, so different SHAs always produce different messages by design. If you want N commits in one message, that needs a new coalescing rule keyed on push-batch, not per-commit."

Result: user gets a clear answer with the line of code that decided the behavior, and a concrete next step.

**Example 2 — User says: "The workflow_job for commit abc1234 made a new message instead of editing the existing one"**

Actions:
1. `node scripts/replay-notif.js --mode grouped --dedup 'github:commit:abc1234'` → shows activity `1715787600123` created at 13:30 with `recent_saved`, then at 14:05 a workflow_job event hit `recent_lookup hit:false`.
2. Compute elapsed: 14:05 − 13:30 = 35 min > `NOTIF_EDIT_WINDOW_MS` (default 30 min).
3. Reply: "Edit window expired (35 min elapsed, default `NOTIF_EDIT_WINDOW_MS=1800000` = 30 min). The workflow_job correctly started a fresh message. Bump `NOTIF_EDIT_WINDOW_MS` in `.env` if you want longer-lived edit windows, but note Teams may garbage-collect editable activities after ~30 min anyway."

Result: root cause identified by clock math against an env var, with the exact knob to turn.

## Common Issues

- **"No such file `.logs/notif-trace-...jsonl`"** — Either the consumer was not running on that date, or tracing is not wired. Check `server/queue/notificationConsumer.js` for `require('./notifTrace')` and `trace(` calls. If absent, traces were never written; no replay is possible.
- **"Cannot find module 'scripts/replay-notif.js'"** — Confirm CWD is the repo root. The script is invoked with `node scripts/replay-notif.js`, not via npm script. If the file is missing, the trace-replay tooling has not been deployed to this checkout; check `git log -- scripts/replay-notif.js`.
- **Timeline prints nothing for a known incident** — Filters AND together; one over-narrow filter silently zeroes the result. Drop filters one at a time starting with `--since/--until`, then `--event`, then `--kind`. If still empty, the event may have been filtered out entirely by `server/queue/filters.js` before reaching the consumer — check `filters.js` for the event type.
- **`edit_skipped_no_convref` for a room that *does* receive bot messages** — The `convref:{conversationId}` key is per-conversation, not per-room name. A renamed channel or migrated team produces a new `conversationId`. Resolve via the room → `conversationId` mapping in `server/queue/channels.js` and check `redis-cli GET convref:{conversationId}` for the current ID.
- **`batch_merge count:1`** — Not actually a batch; the consumer logs `batch_merge` even for single-item ticks in some code paths. Confirm by checking the `kind` of the immediately following event — a real batch will show `recent_saved` once with a coalesced activityId, while a singleton will look indistinguishable from a normal single send.
- **Timestamps in trace don't match Teams** — `ts` in traces is server epoch ms (UTC); Teams shows the viewer's local time. Convert with `date -d @$((TS/1000))` or in `jq`: `select(.ts) | .ts |= (./1000 | strftime("%Y-%m-%d %H:%M:%S"))`.
- **`--commit abc1234` returns nothing but `--dedup github:commit:abc1234` works** — `--commit` matches the SHA stored in `extra.commit_sha` on the envelope; some event types (e.g. `check_suite`) don't carry it. Always fall back to the `--dedup` form which matches the canonical key.