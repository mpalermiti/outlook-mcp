# checker-todo-dates-use-configured-timezone
**What it does:** Stops one new To Do date test from failing by itself around every US clock change. The test now runs on a fixed date instead of today's, so it gives the same answer all year.
**Pictures:** not needed — only a test changes; nothing on a screen in this project changes.
## Behaviours
1. The test for a relative due date such as "in 7 days" gives the same result on any day of the year, because it sets the clock itself — proof: test_a_relative_due_is_resolved_before_it_is_sent
2. It checks the exact local time sent when no clock change falls inside the seven days — proof: test_a_relative_due_is_resolved_before_it_is_sent[no-clock-change]
3. It checks the exact local time sent when the clocks go forward or back inside the seven days — proof: test_a_relative_due_is_resolved_before_it_is_sent[clocks-go-forward], test_a_relative_due_is_resolved_before_it_is_sent[clocks-go-back]
4. It checks the exact local time when the due moment lands in the hour that happens twice each autumn — proof: test_a_relative_due_is_resolved_before_it_is_sent[the-repeated-hour]
5. A code change that gets the local time wrong still fails the test — proof: the text of the progress entry, which records the test failing against a deliberately broken conversion
## Not covered
- The product code: what is sent to Microsoft does not change.
- The checker's suggested comparison against the current moment. It still fails for one hour each autumn, so the test sets the clock instead.
