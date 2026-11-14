# checker-checker-todo-dates-use-configured-timezone
**What it does:** Moving the due date of a repeating To Do task now keeps the time zone that task is already stored in, and reading a task now says which zone its due date is in. The docs stop promising an hour To Do does not keep, and the live test is rebuilt so that a pass really shows Microsoft used the zone.
**Pictures:** not needed — this changes what the server sends and returns, and the docs; nothing on a screen in this project changes.
## Behaviours
1. Reading a task returns the zone of its due date beside the date, so a time in a named zone can't be mistaken for UTC — proof: test_list_tasks_reports_the_zone_beside_due, test_get_task_reports_the_zone_beside_due
2. A task with no due date returns no zone, and a due date with no zone attached returns none rather than a guess — proof: test_a_task_without_a_due_reports_no_zone, test_a_due_without_a_zone_reports_no_zone
3. Changing the due date of a repeating task keeps the zone it is stored in, even when your settings name another — proof: test_updating_a_repeating_tasks_due_keeps_its_stored_zone
4. A moment with its own offset, on a repeating task, is converted to local time in the task's stored zone — proof: test_an_instant_on_a_repeating_task_becomes_wall_clock_in_its_stored_zone
5. A repeating task stored under a Windows zone name keeps that name when the new date has no zone — proof: test_a_windows_stored_zone_is_echoed_for_a_zone_less_due
6. A moment with its own offset on a task stored under a Windows zone name is refused before anything is changed, with a plain way forward, because this server cannot convert into that zone — proof: test_an_instant_on_a_windows_stored_zone_is_refused
7. A repeating task stored under a custom zone Microsoft won't accept back is refused before anything is changed — proof: test_a_custom_stored_zone_is_refused
8. A task that is not repeating, or repeats with no due date yet, still uses your configured zone — proof: test_updating_a_plain_tasks_due_uses_the_configured_zone, test_a_repeating_task_with_no_due_uses_the_configured_zone
9. If Microsoft returns no task when the server looks it up, the change is refused and nothing is sent — proof: test_a_due_update_on_a_task_graph_cannot_return_is_refused
10. A bad due date is refused without any network call, and an update that doesn't touch the due date makes no extra lookup — proof: test_an_invalid_due_is_refused_before_any_read, test_an_update_without_a_due_reads_nothing
11. The docs say To Do keeps only the date of a non-repeating task, and no longer say every due date is anchored or that a task due at 9:00 means 9:00 — proof: the text of README.md, SKILL.md, CHANGELOG.md and the `timezone` field in config.py
12. The live test checks the stored hour on a plain task, reads the zone back from the repeat rule on a separate repeating task, and adds a case that moves a repeating task's due date; a date alone no longer counts as a pass — proof: the text of tests/test_live_todo_write.py (written, not run) and agent/findings/todo-dates-live-check.md
## Not covered
- Whether a plain task should carry your zone at all. That is parked as a question for Michael, with both payloads shown.
- Running the live test. Only Michael can; the offline tests cannot see what Microsoft does with the zone.
- Converting a due date into your zone when reading it. The response names the zone and leaves the date as Microsoft stored it.
