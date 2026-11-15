# redo-todo-dates-live-check
## Ready to start
1. The cause is known: yes — the reviewed change no longer applies to main only because the 1.32.0 release moved the changelog's Unreleased entries; `git apply --check -3` of its delta conflicts in CHANGELOG.md alone, the other seven files apply cleanly.
2. One finish line: yes — the Behaviours below.
3. One part of the code: yes — the To Do module (`tools/todo.py`), its tests, and the docs that describe it.
4. No accounts, money or user data: yes — nothing signs in, pays or stores data; the live tests are written, not run.
5. Fits one session: yes — a mechanical re-apply plus one hand-merged changelog section.

**What it does:** The To Do time-zone fix, as last reviewed (with main's refusal to move a repeating task's due date folded in), carried onto the latest main. Nothing in it changes except where its changelog notes sit: 1.32.0 has since been released, so they go back under Unreleased.
**Pictures:** not needed — this changes what is sent to Microsoft and the docs; nothing on a screen in this project changes.
## Behaviours
1. The change sits on the latest main (1.32.0, the azure-identity cap, the preflight fix), and outside CHANGELOG.md every file's change is byte-identical to the reviewed version — proof: the text (`git diff` of each file's delta against the kept branch's delta, recorded in PROGRESS)
2. The fix's changelog notes sit under Unreleased, above 1.32.0, and none of them is duplicated inside an already-released version — proof: the text of CHANGELOG.md
3. A due date with no zone on a new task, or on an update to a task that does not repeat, goes out labelled with the configured zone; an instant (`Z`, offset, `+7d`) keeps its moment as local time there — proof: `TestTodoDatesUseTheConfiguredZone`, `test_updating_a_plain_tasks_due_uses_the_configured_zone`
4. A repeating task given any `due` gets only main's "repeating task" refusal, after a single read, and nothing is sent — proof: `test_a_repeating_task_gets_only_the_repeating_refusal`, `test_update_task_refuses_due_on_a_repeating_task`
5. The docs say the zone is sent on a repeating task's rule, never that Microsoft keeps it, and never that a repeating task's due can be moved — proof: `test_docs_do_not_promise_a_repeating_tasks_zone_is_kept`, `test_docs_do_not_promise_moving_a_repeating_tasks_due`
6. The live tests for a plain task (west and east of UTC) and a repeating task's rule (zone kept or UTC) are present and not run — proof: the text of `tests/test_live_todo_write.py`
7. Offline suite no worse than at the start of the session, with every toolset too; ruff clean — proof: the text (PROGRESS entry)
## Not covered
- The open checker findings on the reviewed version (main's #242 live test now needs a UTC pin; the old task file's behaviours 4–6) are queued separately and not fixed here.
- The separate redo of the earlier fix is now covered by this one, which carries the whole fix; that queue line is left for the construct.
- Running the live tests: only Michael can.
