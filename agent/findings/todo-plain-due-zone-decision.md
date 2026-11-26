# todo-plain-due-zone-decision

Should a task that does not repeat carry your time zone when it is sent to Microsoft? Repeating tasks need it either way; this is only about plain ones.

## What is sent today (A: keep the zone)

Real code path (`create_task`, serialized as the Graph adapter does) on a server set to `America/Los_Angeles`:

| Caller's `due` | `dueDateTime` sent |
|---|---|
| `2026-12-08T23:30:00` | `{"dateTime": "2026-12-08T23:30:00", "timeZone": "America/Los_Angeles"}` |
| `2026-12-09T07:30:00Z` (the same moment) | `{"dateTime": "2026-12-08T23:30:00", "timeZone": "America/Los_Angeles"}` |

## What B would send (plain tasks back to the old UTC label)

The same code with the zone set to UTC, which is what a plain task's payload would become:

| Caller's `due` | `dueDateTime` sent |
|---|---|
| `2026-12-08T23:30:00` | `{"dateTime": "2026-12-08T23:30:00", "timeZone": "UTC"}` |
| `2026-12-09T07:30:00Z` (the same moment) | `{"dateTime": "2026-12-09T07:30:00", "timeZone": "UTC"}` |

The second row is the difference. "11:30 tonight" given as a moment, or "in 7 days" worked out late in the evening, lands on the 9th instead of the 8th. A date typed plainly lands on the same day either way.

## What it costs

- **A:** read back on a server east of UTC, a plain task's date is the day before the one written. The #72 probe sent 10-15 as China Standard Time and got 16:00 UTC on 10-14. The value is right as a moment and `due_time_zone` says UTC, but an agent that reads off the date gets the 14th.
- **B:** the second row above, and your earlier comments on #72 (2026-09-25 and 2026-11-06, as the checker reports them; workers can't read the thread) said plain due dates could stay as they were. The 2026-11-08 card asked for due dates to be covered, and that is what shipped on the PR.

## Prerequisites

- VERIFIED: both tables, by running the real code path offline with a mock client.
- VERIFIED by the checker's reading of the thread, not by this worker: what the #72 probe returned, and what the comments said.
- ASSUMED: the probe's result holds for an IANA zone name. The probe used a Windows name. The live check tests this.
