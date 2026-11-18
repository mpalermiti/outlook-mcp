# checker-redo-todo-dates-use-configured-timezone

## Ready to start
1. **The cause is known**: yes — the kept branch (1b0a8a4) merged main's #242 refusal beside its own stored-zone read; all four findings trace to that merge, seen in the queue card and in `git diff 1b0a8a4 HEAD -- src/outlook_mcp/tools/todo.py` (-56 lines: the stored-zone branch is gone on main).
2. **One finish line**: yes — the Behaviours below.
3. **One part of the code**: yes — To Do `update_task` (`src/outlook_mcp/tools/todo.py`) and its tests; no product code changes.
4. **No accounts, money or user data**: yes — offline tests and a doc grep only.
5. **Fits one session**: yes — the fix already shipped on main as #252 (1.33.0); what is left is checking each finding against main.

**What it does:** Checks that the four problems the checker found on the old, superseded To Do time-zone branch are all gone in the fix that shipped (#252), and confirms main's own test already stops the "wrong refusal message" problem coming back.
**Pictures:** not needed — no screen or on-screen text changes.

## Behaviours
1. Changing a task's due date reads the task once, not twice; a repeating task is refused and a plain one is patched after that single read — proof: `test_update_task_refuses_due_on_a_repeating_task`, `test_update_task_due_on_a_plain_task_reads_then_patches` (both `get.assert_awaited_once()`, green at HEAD)
2. A repeating task given a `Z`, offset, `+7d` or zone-less due, stored in a Windows or custom zone, on a non-UTC config, only ever gets #242's "repeating task" refusal — never a time-zone retry hint — and nothing is sent — proof: `test_write_payloads_reach_the_wire.py::TestTodoDatesUseTheConfiguredZone::test_a_repeating_task_gets_only_the_repeating_refusal` (on main since #252; a duplicate added here was dropped, see checker-checker-redo-todo-dates-use-configured-timezone)
3. README, SKILL.md and CHANGELOG no longer say a repeating task's due can be moved keeping its zone; they describe the refusal — proof: the text (README.md:70, :459; SKILL.md:152; CHANGELOG.md 1.32.0 refusal, Unreleased #72 entry has no zone-keeping bullet)
4. The live repeating-task test accepts UTC as well as the zone sent, and the docs say the zone is *sent*, with some accounts storing UTC — proof: `tests/test_live_todo_write.py` `_zone_kept_or_coerced` / `test_a_repeating_task_keeps_the_zone_sent_or_utc` (unrun: live tier); the text (README.md:325, :623; config.py:303-308; CHANGELOG.md Unreleased #72 entry)

## Not covered
- The live tier is not run here (mailbox is off-limits); Michael ran it before #252 merged.
