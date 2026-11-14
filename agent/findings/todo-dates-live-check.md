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

## Michael's probe (2026-11-14)

One repeating task shaped like #239's payload, with the Windows name: due `2026-12-14T09:00:00` "Pacific Standard Time" and `recurrenceTimeZone` "Pacific Standard Time" sent; due `2026-12-14T08:00:00` UTC (midnight Pacific, right day) and `recurrenceTimeZone` UTC came back. The #72 probe (2026-09-25, another account) kept its zone. So a repeating task's stored zone is per account, like `is_online` (#70). The plain-task tests had already passed live, west and east of UTC.

## What changed after it (2026-11-14, todo-dates-live-check)

- The fix (#239's change, c5af221) is carried onto this branch unchanged, plus the items below, so this PR can replace #239.
- `test_a_repeating_task_keeps_the_zone_on_its_rule` is now `test_a_repeating_task_keeps_the_zone_sent_or_utc`: passes when the due and the rule are both the zone sent or both UTC (`_zone_kept_or_coerced`, tested offline in `tests/test_todo.py::TestRepeatingTaskZoneIsPerAccount`).
- `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` reads the stored zones before the update and requires the same ones after, whichever they were; it also fails if a second task with its (unique) title exists afterwards, deleting the copy.
- New `test_mains_due_patch_on_a_repeating_task_leaves_one_task`: creates a repeating task with `main`'s payload (UTC config, which sends exactly what `main` sends) and patches its due through the SDK the way `main`'s `update_task` does (due labelled `UTC`), then counts tasks with its title. This is Michael's item 3, "does changing a repeating task's due make a duplicate on main", answered without depending on this branch's `update_task`.
- Docs (README, CHANGELOG, config description, code comments) say the zone is sent and some accounts store UTC; nothing says it is kept. Guarded by `test_docs_do_not_promise_a_repeating_tasks_zone_is_kept`.
- Duplicate check rehearsed offline with a faked listing: one task passes; a second with the same title fails and is deleted.

## Reconciled with main's #242 (2026-11-14, checker-todo-dates-live-check)

Main's #242 fix (b0acf56) refuses every due change on a repeating task, because Microsoft leaves a second copy; #243's commit records that PATCH leaving a duplicate on a personal account on 2026-11-14. That answers Michael's item 3, so the duplicate run is no longer needed:

- `update_task` reads the task once (main's read), refuses a repeating task with #242's message before any zone handling, and sends a plain task's due in the config zone without a read for it. `_due_for_update` and its stored-zone branch are gone.
- Deleted live tests: `test_moving_a_repeating_tasks_due_keeps_its_stored_zone` (head refuses the call on every account, and main's `test_moving_a_repeating_tasks_due_is_refused_and_leaves_one_task` asserts the opposite) and `test_mains_due_patch_on_a_repeating_task_leaves_one_task` (replays a PATCH main no longer sends; #243 already measured it). Main's #242 live test guards the duplicate.
- The parked live-run card for those tests is withdrawn: nothing in it needs running now. What stays live-only on this branch is the three create tests in `TestDueIsAnchoredInTheConfiguredZone`, which `RELEASING.md`'s `live_write` tier runs before tagging.

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
- VERIFIED by Michael's run (reported 2026-11-14): Microsoft honours `timeZone` on a plain task's `dueDateTime`, west and east of UTC.
- VERIFIED by Michael's probe (2026-11-14): Microsoft accepts `recurrenceTimeZone` on a task; his account stored UTC on both the due and the rule. The #72 probe's account kept the zone.
- VERIFIED by #243's commit message (b0acf56): changing a repeating task's due leaves a duplicate on a personal account, so `update_task` refuses it; no stored-zone update path remains.
