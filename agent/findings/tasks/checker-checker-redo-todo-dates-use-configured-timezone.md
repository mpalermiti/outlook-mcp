# checker-checker-redo-todo-dates-use-configured-timezone
**What it does:** Removes the test the last change added, because the fix you shipped in 1.33.0 already has a broader test that checks the same thing, and corrects the notes that said a gap was open.
**Pictures:** not needed — only a duplicate test is removed; nothing on screen changes.

## Behaviours
1. The duplicate `test_update_task_a_repeating_task_only_gets_the_repeating_refusal` is gone from the To Do tests, so the change adds no test of its own — proof: `git diff 21f7151 -- tests/test_todo.py` is empty
2. A repeating task given any due shape still gets only the repeating-task refusal, after one read and with nothing sent, on Windows, custom, IANA, UTC and missing stored zones — proof: `test_write_payloads_reach_the_wire.py::TestTodoDatesUseTheConfiguredZone::test_a_repeating_task_gets_only_the_repeating_refusal` (20 cases, on main since #252, green)
3. The notes no longer say that check had no named test — proof: the text (`agent/PROGRESS.md` newest entry, `agent/findings/tasks/checker-redo-todo-dates-use-configured-timezone.md` Behaviour 2)

## Not covered
- Nothing in `src/` changes; the four original findings were settled on main by #252.
