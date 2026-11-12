# todo-dates-live-check
**What it does:** Gets the one live check of the To Do time-zone fix ready for Michael to run: a command that works as written, an honest count of what it creates, and a table that says what each result means. No worker can run it; the mailbox is off-limits here.
**Pictures:** not needed — nothing on screen changes; this is a test run and its notes.
## Behaviours
1. The command Michael is given finds the tests: it switches to the fix's branch first, because the tests exist only there and on `main` the command would find nothing — proof: the text (`agent/findings/todo-dates-live-check.md` § The test to run), checked by `git show` of both branches.
2. The notes say how many throwaway tasks the run makes (four, about a month out, each deleted) instead of the stale "one" — proof: the text.
3. A pass on the plain-task test can only mean Microsoft used the zone: simulated readbacks for "zone honoured", "zone ignored" and "hour kept" pass, fail and fail, in both Los Angeles and Tokyo — proof: the text (§ Checked offline, the rehearsal output).
4. The functions the live tests call exist on the fix branch with the arguments the tests pass — proof: the text (§ Checked offline).
5. Offline suite no worse than at the start of the session — proof: the suite run recorded in PROGRESS.
## Not covered
- Running the live tests: only Michael can, and the result goes on the card or the PR.
- The plain-task keep-or-drop question; it stays parked separately and this run feeds it.
