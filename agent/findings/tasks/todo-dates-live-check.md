# todo-dates-live-check
**What it does:** Acts on Michael's live probe of the To Do time-zone fix: his account stores a repeating task's rule in UTC whatever zone is sent, so the live tests now accept that and the docs stop promising the zone is kept. Moving a repeating task's due date is refused, as main's earlier fix decided, because Microsoft leaves a second copy behind. This change carries the fix from its own pull request (#239) plus these amendments, so it can replace that pull request.
**Pictures:** not needed — nothing on screen changes; these are tests, docs and comments.
## Behaviours
1. The plain-task fix is unchanged: a plain task's due still goes out in the configured zone, and its live test still expects midnight on the written day there — proof: `TestTodoDatesUseTheConfiguredZone` (offline, unchanged) and `test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone` (live, unchanged, not run).
2. A repeating task still sends the configured zone on its rule; nothing in what is sent changes — proof: `TestTodoDatesUseTheConfiguredZone` (offline, unchanged).
3. The live test for creating a repeating task passes when Microsoft keeps the zone sent or stores UTC instead, and fails on a 400 or any third zone; the due and the rule must agree with each other — proof: `test_a_repeating_task_keeps_the_zone_sent_or_utc` (live, not run); its accept/refuse logic is `_zone_kept_or_coerced`, tested offline by `test_live_zone_rule_accepts_the_zone_sent_or_utc`.
4. Moving a repeating task's due date is refused with only the repeating-task message, whatever zone the task is stored in and however the due is written, after one read and with nothing sent — proof: `test_a_repeating_task_gets_only_the_repeating_refusal` (offline)
5. On a real repeating task the refused move leaves exactly one task with the test's title, still on its written day; the test runs in UTC so the runner's zone cannot shift that day — proof: `test_moving_a_repeating_tasks_due_is_refused_and_leaves_one_task` (live, main's #242 test, not run)
6. No live test from this change replays a due move on a repeating task: the two that did were removed, since main never sends that move — proof: the text (neither removed test name appears under `tests/`)
7. README, SKILL, CHANGELOG, the config description and the code comments say the zone is *sent* on a repeating task's rule and that some accounts store it as UTC; none says Microsoft keeps it — proof: `test_docs_do_not_promise_a_repeating_tasks_zone_is_kept` (offline grep over those files).
8. Offline suite no worse than at the start of the session — proof: the suite run recorded in PROGRESS.
## Not covered
- Running the live tests: only Michael can; `RELEASING.md`'s `live_write` tier runs them before tagging.
- The plain-task keep-or-drop question; Michael's probe answered it in practice ("keep the plain-task fix"), but the parked card is his to close.
