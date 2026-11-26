# pr-todo-dates-use-configured-timezone-s1794187944-806779
**What it does:** Tidies the To Do time-zone fix after review: the changelog no longer says people on UTC see no change at all, two older live tests stop depending on the zone of the machine that runs them, the list of promises matches the code, and the live check gains a Tokyo case so the run shows what happens east of UTC.
**Pictures:** not needed — this changes tests and notes; nothing on a screen in this project changes.
## Behaviours
1. The changelog says that on a UTC server a due with no zone and a repeat rule go out as before, while a due with a `Z` or an offset is now sent as the UTC time of that moment, which can be a different day from the one written — proof: test_a_utc_config_sends_an_offset_due_as_utc_wall_clock, test_a_utc_config_sends_a_zone_less_due_as_written, and the text of CHANGELOG.md
2. The two older live tests that check every field of a created and an updated task run in UTC whatever zone the person running them has configured — proof: the text of tests/test_live_todo_write.py (`test_every_create_field_persists`, `test_every_update_field_persists`)
3. The live test for a plain task's due runs twice, once in Los Angeles and once in Tokyo, each setting its own zone and expecting midnight on the written day in that zone — proof: test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone[America/Los_Angeles] and [Asia/Tokyo] (written, not run)
4. The promises list of the original fix names the live tests that exist and says a repeating task keeps its stored zone when its due date changes — proof: the text of agent/findings/tasks/todo-dates-use-configured-timezone.md
5. The live-check notes say the run now creates four throwaway tasks and what the Tokyo case shows — proof: the text of agent/findings/todo-dates-live-check.md
## Not covered
- Running any live test: Michael runs them; the keep-or-drop question on plain tasks stays parked until he has.
- Any change to what the server sends.
