# todo-dates-live-check
**What it does:** Acts on Michael's live probe of the To Do time-zone fix: his account stores a repeating task's rule in UTC whatever zone is sent, so the live tests now accept that, the docs stop promising the zone is kept, and the same run now also shows whether moving a repeating task's due date leaves a duplicate behind. This change carries the fix from its own pull request (#239) plus these amendments, so it can replace that pull request.
**Pictures:** not needed — nothing on screen changes; these are tests, docs and comments.
## Behaviours
1. The plain-task fix is unchanged: a plain task's due still goes out in the configured zone, and its live test still expects midnight on the written day there — proof: `TestTodoDatesUseTheConfiguredZone` (offline, unchanged) and `test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone` (live, unchanged, not run).
2. A repeating task still sends the configured zone on its rule; nothing in what is sent changes — proof: `TestTodoDatesUseTheConfiguredZone` (offline, unchanged).
3. The live test for creating a repeating task passes when Microsoft keeps the zone sent or stores UTC instead, and fails on a 400 or any third zone; the due and the rule must agree with each other — proof: `test_a_repeating_task_keeps_the_zone_sent_or_utc` (live, not run); its accept/refuse logic is `_zone_kept_or_coerced`, tested offline by `test_live_zone_rule_accepts_the_zone_sent_or_utc`.
4. The live test for moving a repeating task's due date passes when the task keeps whatever zone it was stored in (the zone sent, or UTC) and fails if the update relabels it with the server's zone — proof: `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` (live, not run).
5. That same test fails if moving the due date leaves more than one task with the test's title, and deletes any copy it finds — proof: `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` (live, not run).
6. A new live test sends exactly what `main` sends when a repeating task's due date moves (the due labelled UTC) and fails if that leaves a duplicate, cleaning up any copy — proof: `test_mains_due_patch_on_a_repeating_task_leaves_one_task` (live, not run).
7. README, SKILL, CHANGELOG, the config description and the code comments say the zone is *sent* on a repeating task's rule and that some accounts store it as UTC; none says Microsoft keeps it — proof: `test_docs_do_not_promise_a_repeating_tasks_zone_is_kept` (offline grep over those files).
8. Offline suite no worse than at the start of the session — proof: the suite run recorded in PROGRESS.
## Not covered
- Running the live tests: only Michael can. The command is in `agent/findings/todo-dates-live-check.md`.
- Fixing a duplicate if the run finds one: that would be a separate bug in `update_task`, queued after the run.
- The plain-task keep-or-drop question; Michael's probe answered it in practice ("keep the plain-task fix"), but the parked card is his to close.
