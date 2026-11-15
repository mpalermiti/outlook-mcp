# checker-redo-todo-dates-live-check
**What it does:** The live check that a repeating task's date can't be moved now runs in UTC, like the two checks beside it, so it no longer fails for someone whose computer is set far east of London when nothing is actually broken.
**Pictures:** not needed — a test file changes; nothing a person sees in the server's replies changes.

## Ready to start
1. The cause is known: yes — `tests/test_live_todo_write.py:253` creates the repeating task with the runner's own zone, and the To Do fix now sends the due as wall-clock time in that zone.
2. One finish line: yes — the Behaviours below.
3. One part of the code: yes — the To Do live test file, plus one offline guard in `tests/test_todo.py`.
4. No accounts, money or user data: yes — test code only; nothing runs against a mailbox here.
5. Fits one session: yes — a one-line config pin and a short guard.

## Behaviours
1. The live test that a repeating task's due move is refused creates and updates its task under a UTC copy of the runner's config, so the written day it asserts is the day sent whatever zone the runner uses — proof: `test_moving_a_repeating_tasks_due_is_refused_and_leaves_one_task` (live, not run here) and the text of the change.
2. An offline guard fails if any live To Do test creates a task with a `due` under the runner's unpinned config, so the next live test added in that file can't bring the same zone dependence back — proof: `test_live_tests_pin_the_zone_of_every_due_they_write` (red before the pin, green after).

## Not covered
- The live test is not run: the mailbox is off-limits to workers. Michael's account stores UTC, so his run passed before and will pass after.
- Finding 2 of the older `checker-checker-todo-dates-live-check` card (stale behaviours in an old task file and a stale parked card) is a separate task.
