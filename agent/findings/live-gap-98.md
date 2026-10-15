# live-gap-98 — the live test that shows #98

Issue #98: `outlook_update_event` moving a series master's date in the same zone, with no
`recurrence`, answers `updated` and leaves every occurrence where it was. This session does not
fix it. It writes the test that would catch it, so the bug can be confirmed and, once fixed,
kept fixed.

Re-checked 2026-10-14 before starting: #98 open, unassigned, no comments. PR search hits
#103, #83 and #63 do not touch it (#103's body says "#98 is not folded in"; #83 filed it).

## Where

`tests/test_live_calendar_write.py::TestMovingAMastersDateMovesItsOccurrences`, two tests:

| Test | Series created | Update sent |
|---|---|---|
| `test_a_timed_weekly_series_moved_a_day_lands_on_the_new_day` | weekly Monday 09:00–09:30 `America/New_York`, 3 occurrences | `start`/`end` on the Tuesday, 09:00–09:30, no `timezone`, no `recurrence` |
| `test_an_all_day_weekly_series_moved_a_day_lands_on_the_new_day` | weekly all-day Monday, 3 occurrences | `start`/`end` Tuesday 00:00 → Wednesday 00:00, nothing else |

The timed case is the issue's own reproduction. The all-day case is its second one, but
`numbered` only: this tier's rules (`tests/conftest.py`) never write a `noEnd` range.
Both follow the tier's rules: no attendees, a bounded range, a subject carrying the
`[outlook-mcp live-write guard] safe to delete` marker (plus ` issue98-timed` /
` issue98-all-day`), and the event deleted in a `finally`.

## What each test asserts

1. **Premise:** before the move, the series' occurrences are three consecutive Mondays.
   Read through `/me/events/{id}/instances`, in the event's own zone (`America/New_York`;
   UTC for the all-day series, which Graph anchors in UTC whatever it is sent).
2. The update answers `status: updated`.
3. **The claim:** after the move, the occurrences are three consecutive Tuesdays starting the
   day after the original Monday, at the original time of day.

The step that matters is 3. Every earlier `live_write` test asserts the master's start or type
after a date move, which is why #98 went unseen: the master can look moved while its
occurrences are not.

## How it stays out of the way

Both tests are `xfail(strict=True, raises=_SeriesDidNotMoveError)`.

- `_SeriesDidNotMoveError` is raised only for #98's exact signature: after an `updated`, the
  occurrences are the original three Mondays at the original time of day, untouched.
- Anything that moved the series somewhere other than the three Tuesdays (other Mondays, the
  same Mondays at another time, an all-day start that is not midnight) is a plain failure.
  Pinned offline by `tests/test_live_gap_98_classifier.py` (2026-10-14, after a review showed
  the first version let those through as XFAIL).
- Anything else (a failed premise, an error from Graph, occurrences on a third day, a changed
  time of day) is a plain failure and shows red.
- If the occurrences land on Tuesday, the test passes, and strict `xfail` turns that pass into a
  failure (`XPASS(strict)`). That means a fix landed or Graph changed; delete the marker and
  the test becomes the regression guard for the fix.

So a release-time run of the whole write tier (RELEASING.md step 1d) stays green while #98 is
open, and cannot quietly stay green after it is fixed.

Default runs never select it: the module carries `pytest.mark.live_write`, deselected by
`addopts`, and the tier skips without `OUTLOOK_MCP_LIVE_WRITE=1`. Default suite after this
change: 1365 passed, 87 deselected (was 85; the two new tests).

## How to run it

```bash
OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write \
  "tests/test_live_calendar_write.py::TestMovingAMastersDateMovesItsOccurrences" -v -rxX
```

About a minute. Creates and deletes two throwaway events about a month out. `-rxX` prints the
xfail reason beside each XFAIL. It does not print the occurrence dates: for an XFAIL pytest
shows only the marker's reason. A plain failure's message does list the start times.

## What a pass and a fail look like

| You see | Meaning | Next |
|---|---|---|
| `XFAIL` on both, reason starting `#98:` | The bug is real on this build, and the test catches it | Nothing to do now; the test is ready for whoever fixes #98 |
| `FAILED [XPASS(strict)]` | The occurrences moved: fixed, or Graph changed | Remove the `xfail` marker; the issue's premise needs a re-check |
| `FAILED` with an `AssertionError` | Something else: the premise or the landing day is wrong | Read the message; it lists the occurrence start times |
| `XFAIL` on one, `XPASS(strict)` on the other | Timed and all-day series behave differently | Worth a note on #98; the fix may need two paths |
| `SKIPPED` | `OUTLOOK_MCP_LIVE_WRITE` not set, no token, or a `read_only` config | Set it and re-run |

## Verification done here (offline only)

- VERIFIED: default suite green, ruff check and format clean; the two tests are collected and
  deselected (deselected count 85 → 87).
- VERIFIED: pytest's `xfail(strict=True, raises=Sub)` semantics, on a scratch file: raising the
  subclass gives XFAIL, a plain `AssertionError` gives FAILED, a pass gives FAILED
  (`XPASS(strict)`).
- VERIFIED: both test bodies dry-run against fakes for `create_event`, `update_event`,
  `delete_event` and `_instances` (scratch file, not committed): the #98 shape raises
  `_SeriesDidNotMoveError`; three Tuesdays passes; three Wednesdays raises a plain
  `AssertionError`; the event is deleted in every case; the update carries no `timezone`,
  `recurrence` or attendees.
- ASSUMED (from the issue's live transcript, not re-run): Graph answers this patch `updated`
  and leaves the series on Monday, timed and all-day.
- ASSUMED: Graph returns an all-day occurrence's start as midnight when asked for UTC. If not,
  the all-day test fails red on the time-of-day check with the start times in the message,
  rather than reporting a wrong XFAIL.
- NOT RUN: either live test. Workers never touch a mailbox.
