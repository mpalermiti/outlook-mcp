# redo-redo-todo-dates-use-configured-timezone
**What it does:** Nothing new ships. The To Do time-zone fix this task would rebuild is already rebuilt on the latest code, in the open change that also carries its live tests, so a second copy would only conflict with it.
**Pictures:** not needed — no code changes; this closes a duplicate task.
## Behaviours
1. Every offline behaviour of the time-zone fix is already in the open live-check change, which sits on the current main — proof: the text (the table under "Checked" below)
2. The three pieces of the old fix that are missing there were removed on purpose, because main now refuses to move a repeating task's due date; rebuilding them would bring back the six failures the checker found — proof: the text (the `todo.py` diff summary below)
3. No product code, test or doc changes in this session — proof: `git diff` of this commit touches only `agent/`
## Not covered
- The open findings on the live-check change (`checker-checker-todo-dates-live-check`) and the parked plain-task question stay where they are.

## Checked (2026-11-15)
- `construct/redo-todo-dates-live-check` (f3a0319) has merge-base `62c3fc0` = `origin/main` HEAD. VERIFIED (`git merge-base`).
- All 24 proof tests named in `redo-todo-dates-use-configured-timezone.md` were grepped by `def` name on the kept branch (`construct-kept/redo-todo-dates-use-configured-timezone-20261114-094156`, 1b0a8a4) and on the live-check branch. 21 exist on both. The 3 absent from live-check:
  - `test_updating_a_repeating_tasks_due_keeps_its_stored_zone` (offline) and `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` (live): retired by `checker-todo-dates-live-check`, since #242 refuses that move; replaced by main's `test_moving_a_repeating_tasks_due_is_refused_and_leaves_one_task`.
  - `test_a_repeating_task_keeps_the_zone_on_its_rule` (live): renamed `test_a_repeating_task_keeps_the_zone_sent_or_utc` after Michael's 2026-11-14 probe showed his account stores UTC.
- `git diff` kept → live-check on `src/outlook_mcp/tools/todo.py`: the only code difference is that `_due_for_update` (the second read that broke #242's tests, line 40 of the queue) is gone; `update_task` labels a plain task's due with the config zone and lets #242 refuse a repeating one. Everything else in the fix (`_due_in_zone`, `_recurrence_in_zone`, `due_time_zone` on read, relative-due resolution, legacy-zone fallback) is the same.
- Smoke on this checkout: 3138 passed, 4 skipped, 131 deselected; 1 failed + 25 errors all `tests/test_http_transport.py` (sandbox EPERM binding 127.0.0.1), same as previous sessions. Ruff check + format clean.
- Consequence: `checker-redo-todo-dates-use-configured-timezone` (queue line 40) reviews the kept branch this task would have replaced; its three findings are the ones the live-check branch already fixed. It can be retired with this task once the live-check change merges. Not edited here (not this task's line).
