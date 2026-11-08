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

Creates one throwaway task about a month out and deletes it. Not run by any worker.

| Result | Meaning |
|---|---|
| Pass | Microsoft honours the zone on a task's due date and accepts it on a repeat rule. |
| Due date assertion fails, on the written date | Microsoft ignored the zone on the due date: the fix does not take effect there. |
| Due date assertion fails, some other value | Microsoft rewrites a task's due date (for example to the day only); the test's expectation needs revisiting with that value. |
| 400 from Microsoft | It refuses the zone on a task, most likely on the repeat rule. Non-UTC servers can't create repeating tasks until that field is dropped. |

## Prerequisites

- VERIFIED: the payload above, by running the real code path offline with a mock client.
- VERIFIED: the new test is well-formed, by the plain suite importing it (127 deselected, up from 126).
- ASSUMED: Microsoft honours `timeZone` on a task's `dueDateTime` as it does on an event. Read in Graph's model, not confirmed on a task.
- ASSUMED: Microsoft accepts `recurrenceTimeZone` on a task. No earlier live run sent it.
