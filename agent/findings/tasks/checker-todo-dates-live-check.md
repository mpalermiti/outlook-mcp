# checker-todo-dates-live-check
**What it does:** Makes the To Do time-zone change agree with the fix already on main that refuses to move a repeating task's due date. Changing a due date now reads the task once, refuses a repeating task straight away with the real reason, and the docs and live tests stop promising otherwise.
**Pictures:** not needed — nothing renders; the change is in tool behaviour, tests and docs.
## Behaviours
1. Changing the due date of a task reads it exactly once: main's two To Do tests that failed on "awaited 2 times" pass again
   — proof: `test_update_task_refuses_due_on_a_repeating_task`, `test_update_task_due_on_a_plain_task_reads_then_patches`
2. A repeating task given any `due` (plain time, `Z`, offset, `+7d`), whatever zone it is stored in (IANA, Windows name, custom), gets only the "repeating task is refused" message, after one read, and nothing is sent
   — proof: `test_a_repeating_task_gets_only_the_repeating_refusal`
3. A task that doesn't repeat still has its new due sent in the configured zone, and a bad `due` is refused before any read
   — proof: `test_updating_a_plain_tasks_due_uses_the_configured_zone`, `test_update_task_labels_a_zone_less_due_with_the_configured_zone`, `test_an_invalid_due_is_refused_before_any_read`
4. The stored-zone code for moving a repeating task, and the four payload tests that exercised it, are gone — proof: the text (`_due_for_update` absent from `src/`)
5. The README, SKILL.md and changelog no longer say a repeating task's due can be changed while keeping its zone — proof: `test_docs_do_not_promise_moving_a_repeating_tasks_due`
6. The two live tests that could only fail (moving a repeating task's due, and replaying main's old PATCH) are deleted, and the findings file no longer asks Michael to run them — proof: the text (`tests/test_live_todo_write.py`, `agent/findings/todo-dates-live-check.md`)
7. Full offline suite green apart from the sandbox-blocked HTTP transport tests, with every toolset too; ruff clean — proof: the text (PROGRESS entry)
## Not covered
- The parked live-run card for the duplicate check is left as it is on the queue (workers don't edit other lines); its run is no longer needed and this file and the queue note say so.
- The overlapping checker card on the redo change is a separate PR; this fix resolves its findings 1–3 on this branch only.
