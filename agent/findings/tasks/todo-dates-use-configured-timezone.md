# todo-dates-use-configured-timezone
**What it does:** A To Do due date, reminder or repeat rule now uses the time zone in your settings instead of always UTC. To Do keeps only the date of a non-repeating task, so there the zone decides which day is stored.
**Pictures:** not needed — this changes what is sent to Microsoft; nothing on a screen in this project changes.
## Behaviours
1. A due date with no zone, on a new or updated task, goes out as written and labelled with your configured zone — proof: test_a_zone_less_due_is_labelled_with_the_configured_zone, test_a_bare_date_is_labelled_with_the_configured_zone, test_update_task_labels_a_zone_less_due_with_the_configured_zone
2. A due date that names its own moment (a `Z` or an offset) keeps that exact moment, shown as local time in your zone — proof: test_a_z_due_becomes_wall_clock_in_the_configured_zone, test_an_offset_due_becomes_wall_clock_in_the_configured_zone, test_update_task_converts_a_z_due_to_the_configured_zone
3. A relative due date such as "in 7 days" is worked out before sending, instead of the literal text going to Microsoft — proof: test_a_relative_due_is_resolved_before_it_is_sent
4. A reminder uses the same time and zone as the due date — proof: test_a_reminder_matches_the_due_time_and_zone
5. A repeat rule carries your configured zone, unless the caller already named one — proof: test_recurrence_carries_the_configured_zone, test_recurrence_without_a_due_carries_the_configured_zone, test_a_recurrence_zone_the_caller_named_is_kept
6. A server left on the default UTC setting sends the same repeat rule as before — proof: test_a_utc_config_leaves_the_recurrence_as_it_was
7. An old zone name Microsoft refuses (such as EST) falls back to UTC, as calendar events do; a misspelt zone is refused before anything is sent — proof: test_a_legacy_config_zone_falls_back_to_utc, test_a_config_zone_that_does_not_exist_is_refused_before_sending
8. The settings docs say the zone applies to To Do tasks as well as events — proof: the text of README.md and the `timezone` field in config.py
9. Whether Microsoft honours the zone on a real task is left as a live test for Michael to run — proof: test_a_late_evening_due_lands_on_the_next_utc_day (written, not run)
## Not covered
- Completing a task: the completion time is the present moment, already correct in UTC.
- A repeat rule's start date is still the caller's; it is not cross-checked against the due date as calendar events are.
- Converting on read: a task's `due` comes back as Microsoft stored it, with its zone beside it.
