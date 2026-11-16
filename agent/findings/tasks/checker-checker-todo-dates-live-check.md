# checker-checker-todo-dates-live-check
## Ready to start
1. The cause is known: yes — finding 1's unpinned live test and finding 2's stale behaviours 4–6 and parked card, both seen at the lines the checker names.
2. One finish line: yes — the Behaviours below.
3. One part of the code: yes — the To Do live-test notes under `agent/` (finding 1 is already on main via #252).
4. No accounts, money or user data: yes — notes only; no live test is run.
5. Fits one session: yes — one guard check and three note edits.

**What it does:** Brings the notes for the To Do time-zone fix in line with what was merged: they stop promising two live checks that were removed, and the request asking Michael to run those checks is withdrawn. The live test the checker flagged was already fixed on main, so no code changes.
**Pictures:** not needed — nothing on screen changes; only the agent's notes change.
## Behaviours
1. Main's live test for the repeating-task refusal creates and updates its task under a UTC copy of the config, not the runner's own zone — proof: `tests/test_todo.py::TestRepeatingTaskZoneIsPerAccount::test_live_tests_pin_the_zone_of_every_due_they_write` (offline, already on main)
2. The old task notes describe the refusal that shipped, naming only live tests that exist, and no longer promise a duplicate-run check — proof: the text of `agent/findings/tasks/todo-dates-live-check.md`, and a grep for the two deleted test names under `agent/findings/tasks/` finding none outside "removed" wording
3. The parked request asking Michael to run the five live tests is retired with a note saying why — proof: the text of the `todo-repeating-due-duplicate-live-run` line in `agent/QUEUE.md`
4. Offline suite no worse than at the start of the session — proof: the suite run recorded in PROGRESS
## Not covered
- Running any live test: only Michael can, and none is needed for this.
- The other parked To Do question (keep the zone on plain tasks): it is Michael's to close; noted in suggestions.
