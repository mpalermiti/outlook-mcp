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

```
OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write "tests/test_live_todo_write.py::TestDueIsAnchoredInTheConfiguredZone" -v
```

Three tests, each creating one throwaway task about a month out and deleting it. Not run by any worker.

| Test | Pass means | A failure means |
|---|---|---|
| `test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone` | Microsoft used the zone to pick the day on a plain task, and stored midnight there, as the #72 probe saw. | Back at midnight UTC: the zone was ignored. Back with the hour kept: the docs' "keeps only the date" is wrong. Paste the value. |
| `test_a_repeating_task_keeps_the_zone_on_its_rule` | Microsoft accepts `recurrenceTimeZone` on a task and stores it. | A 400 here: non-UTC servers can't create repeating tasks until that field is dropped. |
| `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` | Changing the due date of a series anchored in another zone is accepted and leaves its zone alone. | A 400 here, or a changed zone: the stored-zone rule needs revisiting. Paste the error. |

No date is asserted on the repeating tasks: the #72 probe saw a repeating task's due move to midnight the next day, a shift whose rule is still unknown. A date alone can't show the zone was used.

## Prerequisites

- VERIFIED: the payload above, by running the real code path offline with a mock client.
- VERIFIED: the new test is well-formed, by the plain suite importing it (127 deselected, up from 126).
- ASSUMED: Microsoft honours `timeZone` on a task's `dueDateTime`. The #72 probe (2026-09-25, a Windows zone name) points that way; an IANA name has not been seen.
- ASSUMED: a repeating task reads back labelled with its own zone, which is what the update path relies on. Seen once in the #72 probe.
- ASSUMED: Microsoft accepts `recurrenceTimeZone` on a task. No earlier live run sent it.
