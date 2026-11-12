# todo-dates-live-check

The fix for To Do dates is proven up to the point the request leaves this server. What Microsoft does with it can only be seen on a real mailbox.

## What is now sent

Produced by the real code path (`create_task`, serialized the way the Graph adapter does) for a task due 23:30, repeating every two days, with a reminder, on a server set to `America/Los_Angeles`:

```json
{
  "dueDateTime": {"dateTime": "2026-12-08T23:30:00", "timeZone": "America/Los_Angeles"},
  "isReminderOn": true,
  "recurrence": {
    "pattern": {"interval": 2, "type": "daily"},
    "range": {"recurrenceTimeZone": "America/Los_Angeles", "startDate": "2026-12-08", "type": "noEnd"}
  },
  "reminderDateTime": {"dateTime": "2026-12-08T23:30:00", "timeZone": "America/Los_Angeles"},
  "title": "Example"
}
```

Before the fix the same call labelled both dates `UTC` and sent no `recurrenceTimeZone`.

## The test to run

Run it on the fix's branch, before the PR merges. The test class and the fix exist only there; on `main` the command finds no tests and proves nothing (checked 2026-11-11: `TestDueIsAnchoredInTheConfiguredZone` is in `construct/redo-todo-dates-use-configured-timezone`'s `tests/test_live_todo_write.py` and not in `main`'s).

```
git fetch origin && git switch --detach origin/construct/redo-todo-dates-use-configured-timezone
uv sync
OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write "tests/test_live_todo_write.py::TestDueIsAnchoredInTheConfiguredZone" -v
git switch -
```

Four tests (the plain-task one runs once in Los Angeles and once in Tokyo), each creating one throwaway task about a month out and deleting it. Each sets its own zone, so the machine it runs on makes no difference. Not run by any worker.

| Test | Pass means | A failure means |
|---|---|---|
| `test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone[America/Los_Angeles]` and `[Asia/Tokyo]` | Microsoft used the zone to pick the day on a plain task, and stored midnight there, as the #72 probe saw. West of UTC that reads back as the written day; in Tokyo it reads back as 15:00 UTC the day before, which is what option A on the plain-task question means east of UTC. | Back at midnight UTC: the zone was ignored. Back with the hour kept: the docs' "keeps only the date" is wrong. Paste the value. |
| `test_a_repeating_task_keeps_the_zone_on_its_rule` | Microsoft accepts `recurrenceTimeZone` on a task and stores it. | A 400 here: non-UTC servers can't create repeating tasks until that field is dropped. |
| `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` | Changing the due date of a series anchored in another zone is accepted and leaves its zone alone. | A 400 here, or a changed zone: the stored-zone rule needs revisiting. Paste the error. |

No date is asserted on the repeating tasks: the #72 probe saw a repeating task's due move to midnight the next day, a shift whose rule is still unknown. A date alone can't show the zone was used.

The two older round-trip tests in the same file (`test_every_create_field_persists`, `test_every_update_field_persists`) now pin `timezone: "UTC"`, so the full live tier the release checklist asks for after a `todo.py` change no longer depends on the runner's configured zone.

## Checked offline (2026-11-11, todo-dates-live-check)

- The plain-task assertion has teeth. Its own arithmetic, fed simulated readbacks for a task written 30 days from 2026-11-11 (due 2026-12-11T23:30 in the zone):

  | Zone | Zone honoured (midnight there) | Zone ignored (midnight UTC) | Hour kept |
  |---|---|---|---|
  | America/Los_Angeles | `12-11T08:00` PASS | `12-11T00:00` FAIL | `12-12T07:30` FAIL |
  | Asia/Tokyo | `12-10T15:00` PASS | `12-11T00:00` FAIL | `12-11T14:30` FAIL |

  So a pass can only mean Microsoft used the zone to pick the day.
- Every function the tests call (`create_task`, `update_task`, `delete_task`, `list_tasks`, `_resolve_list_id`) exists on the fix branch with the keyword arguments the tests pass; `_format_task` there returns `due` and `due_time_zone`, which the plain-task test reads. Read from the branch with `git show`, not run.
- Not checked, and only the run can: anything Microsoft does.

## Prerequisites

- VERIFIED: the payload above, by running the real code path offline with a mock client.
- VERIFIED: the new test is well-formed, by the plain suite importing it (127 deselected, up from 126).
- ASSUMED: Microsoft honours `timeZone` on a task's `dueDateTime`. The #72 probe (2026-09-25, a Windows zone name) points that way; an IANA name has not been seen.
- ASSUMED: a repeating task reads back labelled with its own zone, which is what the update path relies on. Seen once in the #72 probe.
- ASSUMED: Microsoft accepts `recurrenceTimeZone` on a task. No earlier live run sent it.
