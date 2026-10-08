"""Tests for the shared recurrence helpers.

``build_patterned_recurrence`` is the strict dict→SDK converter that To Do has
used since 1.5.0 (moved here from ``tools/todo.py``); its behavior is covered
by ``test_todo.py`` and must not drift. The new surface is
``build_event_recurrence`` — the calendar entry point that additionally accepts
JSON strings and the documented shorthands, and reconciles ``range.startDate``
with the event's own start.
"""

from datetime import date

import pytest
from msgraph.generated.models.day_of_week import DayOfWeek
from msgraph.generated.models.patterned_recurrence import PatternedRecurrence
from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType
from msgraph.generated.models.recurrence_range_type import RecurrenceRangeType

from outlook_mcp.tools._recurrence import (
    build_event_recurrence,
    build_patterned_recurrence,
    serialize_recurrence,
)

_START = "2026-09-07T12:30:00Z"  # a Monday
_FULL = {
    "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday", "friday"]},
    "range": {"type": "noEnd", "startDate": "2026-09-07"},
}


class TestBuildEventRecurrenceDict:
    def test_dict_payload_builds_typed_recurrence(self):
        """The issue's exact payload converts to a typed PatternedRecurrence."""
        pr = build_event_recurrence(_FULL, start=_START)

        assert isinstance(pr, PatternedRecurrence)
        assert pr.pattern.type is RecurrencePatternType.Weekly
        assert pr.pattern.interval == 1
        assert pr.pattern.days_of_week == [DayOfWeek.Monday, DayOfWeek.Friday]
        assert pr.range.type is RecurrenceRangeType.NoEnd
        assert pr.range.start_date == date(2026, 9, 7)

    def test_json_string_payload_is_parsed(self):
        """A JSON-encoded dict is accepted — some client bridges stringify nested args."""
        import json

        pr = build_event_recurrence(json.dumps(_FULL), start=_START)

        assert pr.pattern.type is RecurrencePatternType.Weekly
        assert pr.pattern.days_of_week == [DayOfWeek.Monday, DayOfWeek.Friday]

    def test_missing_start_date_defaults_to_event_start(self):
        """Graph requires range.startDate to match the event start; default it when absent."""
        payload = {"pattern": {"type": "daily", "interval": 1}, "range": {"type": "noEnd"}}

        pr = build_event_recurrence(payload, start="2026-04-15T09:00:00Z")

        assert pr.range.start_date == date(2026, 4, 15)

    def test_conflicting_start_date_is_rejected(self):
        """A startDate that disagrees with the event start is a 400 from Graph — catch it here."""
        payload = {
            "pattern": {"type": "daily", "interval": 1},
            "range": {"type": "noEnd", "startDate": "2026-10-01"},
        }

        with pytest.raises(ValueError, match="startDate"):
            build_event_recurrence(payload, start="2026-04-15T09:00:00Z")

    def test_missing_range_is_defaulted_not_rejected(self):
        """Callers routinely send only a pattern; supply an open-ended range."""
        pr = build_event_recurrence({"pattern": {"type": "daily", "interval": 1}}, start=_START)

        assert pr.range.type is RecurrenceRangeType.NoEnd
        assert pr.range.start_date == date(2026, 9, 7)

    def test_offset_start_uses_the_date_as_written(self):
        """+02:00 12:30 is the 7th to the caller; don't shift the series into UTC's day."""
        pr = build_event_recurrence(
            {"pattern": {"type": "daily", "interval": 1}},
            start="2026-09-07T00:30:00+02:00",
        )

        assert pr.range.start_date == date(2026, 9, 7)


class TestARangeThatEndsBeforeItBegins:
    """#86: ``startDate`` is re-derived from the start and ``endDate`` is the caller's.

    So a start past the series end used to build a range ending before it
    began, which Graph refuses with ``400 ErrorInvalidParameter: StartDateV2
    should be earlier or equal to EndDateV2``. The refusal is ours now, and it
    never moves ``endDate``: extending a series is not what the caller asked for.
    """

    _START = "2027-04-05T09:00:00"

    @staticmethod
    def _ending(end: str, range_type: str = "endDate", **extra) -> dict:
        return {
            "pattern": {"type": "daily", "interval": 1},
            "range": {"type": range_type, "endDate": end, **extra},
        }

    def test_an_end_before_the_start_is_refused_naming_both_dates(self):
        """The issue's exact repro."""
        with pytest.raises(ValueError) as excinfo:
            build_event_recurrence(self._ending("2027-03-26"), start=self._START)

        message = str(excinfo.value)
        assert "range.endDate (2027-03-26)" in message
        assert "first occurrence (2027-04-05)" in message
        assert "on or after 2027-04-05" in message

    def test_it_is_refused_when_the_caller_supplied_a_matching_start_date_too(self):
        """The verified-``startDate`` branch reaches the check, not only the defaulted one."""
        with pytest.raises(ValueError, match="range.endDate"):
            build_event_recurrence(
                self._ending("2027-03-26", startDate="2027-04-05"), start=self._START
            )

    def test_an_end_on_the_start_day_is_accepted(self):
        """Graph allows equal dates — a one-day range — so the boundary must pass."""
        pr = build_event_recurrence(self._ending("2027-04-05"), start=self._START)

        assert pr.range.start_date == date(2027, 4, 5)
        assert pr.range.end_date == date(2027, 4, 5)

    def test_a_numbered_series_as_graph_reads_it_back_is_accepted(self):
        """The reason the check reads ``endDate`` only on an ``endDate`` range.

        Graph reads a ``numbered`` series back with ``endDate: "0001-01-01"``,
        and accepts any ``endDate`` on one and ignores it (verified live
        2026-10-05). ``serialize_recurrence`` passes the field through, so this
        is the shape every echo of a numbered series carries — the
        ``outlook_get_event`` round trip and the zone re-anchor's re-send alike.
        """
        pr = build_event_recurrence(
            self._ending("0001-01-01", "numbered", numberOfOccurrences=2), start=self._START
        )

        assert pr.range.type is RecurrenceRangeType.Numbered
        assert pr.range.number_of_occurrences == 2

    def test_a_pascal_case_range_type_is_checked_too(self):
        """The converter accepts the SDK's member name, so the check has to see it as well."""
        with pytest.raises(ValueError, match="range.endDate"):
            build_event_recurrence(self._ending("2027-03-26", "EndDate"), start=self._START)

    def test_a_malformed_end_gets_the_converters_refusal(self):
        """One wording for a bad date, whichever check reaches it first."""
        with pytest.raises(ValueError, match="range.endDate must be YYYY-MM-DD"):
            build_event_recurrence(self._ending("2027-13-01"), start=self._START)


class TestShorthands:
    def test_weekly_uses_the_start_weekday(self):
        """'weekly' on a Monday start means every Monday."""
        pr = build_event_recurrence("weekly", start=_START)

        assert pr.pattern.type is RecurrencePatternType.Weekly
        assert pr.pattern.interval == 1
        assert pr.pattern.days_of_week == [DayOfWeek.Monday]
        assert pr.range.type is RecurrenceRangeType.NoEnd
        assert pr.range.start_date == date(2026, 9, 7)

    def test_daily(self):
        pr = build_event_recurrence("daily", start=_START)

        assert pr.pattern.type is RecurrencePatternType.Daily
        assert pr.pattern.interval == 1

    def test_weekdays_is_monday_through_friday(self):
        pr = build_event_recurrence("weekdays", start=_START)

        assert pr.pattern.type is RecurrencePatternType.Weekly
        assert pr.pattern.days_of_week == [
            DayOfWeek.Monday,
            DayOfWeek.Tuesday,
            DayOfWeek.Wednesday,
            DayOfWeek.Thursday,
            DayOfWeek.Friday,
        ]

    def test_monthly_is_absolute_on_the_start_day(self):
        pr = build_event_recurrence("monthly", start=_START)

        assert pr.pattern.type is RecurrencePatternType.AbsoluteMonthly
        assert pr.pattern.day_of_month == 7

    def test_yearly_is_absolute_on_the_start_month_and_day(self):
        pr = build_event_recurrence("yearly", start=_START)

        assert pr.pattern.type is RecurrencePatternType.AbsoluteYearly
        assert pr.pattern.month == 9
        assert pr.pattern.day_of_month == 7

    def test_unknown_shorthand_is_rejected_with_the_valid_set(self):
        with pytest.raises(ValueError, match="daily"):
            build_event_recurrence("biweekly", start=_START)

    def test_shorthand_is_case_insensitive(self):
        pr = build_event_recurrence("Weekly", start=_START)

        assert pr.pattern.type is RecurrencePatternType.Weekly

    def test_non_mapping_json_is_rejected(self):
        """A JSON array parses but isn't a recurrence — don't fall through to shorthand."""
        with pytest.raises(ValueError):
            build_event_recurrence("[1, 2]", start=_START)


class TestStrictDictConverter:
    """``build_patterned_recurrence`` keeps To Do's 1.5.0 contract exactly."""

    def test_rejects_a_string(self):
        with pytest.raises(ValueError, match="dict"):
            build_patterned_recurrence("weekly")  # type: ignore[arg-type]

    def test_requires_both_pattern_and_range(self):
        with pytest.raises(ValueError, match="pattern"):
            build_patterned_recurrence({"pattern": {"type": "daily"}})

    def test_rejects_an_invalid_enum_with_the_valid_set(self):
        with pytest.raises(ValueError, match="Invalid pattern.type"):
            build_patterned_recurrence(
                {"pattern": {"type": "fortnightly"}, "range": {"type": "noEnd"}}
            )


class TestSerializeRecurrence:
    def test_round_trips_to_the_documented_json_shape(self):
        """What create accepts, get returns — same camelCase Graph shape."""
        pr = build_event_recurrence(_FULL, start=_START)

        assert serialize_recurrence(pr) == {
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday", "friday"]},
            "range": {"type": "noEnd", "startDate": "2026-09-07"},
        }

    def test_none_stays_none(self):
        assert serialize_recurrence(None) is None

    def test_unset_fields_are_omitted(self):
        """Don't pad every event detail with a dozen nulls."""
        pr = build_event_recurrence("daily", start=_START)

        assert serialize_recurrence(pr) == {
            "pattern": {"type": "daily", "interval": 1},
            "range": {"type": "noEnd", "startDate": "2026-09-07"},
        }

    def test_full_range_fields_survive(self):
        pr = build_event_recurrence(
            {
                "pattern": {"type": "absoluteMonthly", "interval": 2, "dayOfMonth": 7},
                "range": {
                    "type": "numbered",
                    "startDate": "2026-09-07",
                    "numberOfOccurrences": 4,
                    "recurrenceTimeZone": "UTC",
                },
            },
            start=_START,
        )

        assert serialize_recurrence(pr) == {
            "pattern": {"type": "absoluteMonthly", "interval": 2, "dayOfMonth": 7},
            "range": {
                "type": "numbered",
                "startDate": "2026-09-07",
                "numberOfOccurrences": 4,
                "recurrenceTimeZone": "UTC",
            },
        }


class TestEventStartDate:
    """``event_start_date`` also parses what Graph hands *back*, not just user input."""

    def test_parses_graph_seven_digit_fractional_seconds(self):
        """Graph returns "…T09:00:00.0000000"; datetime.fromisoformat rejects that on 3.10."""
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-09-07T09:00:00.0000000") == date(2026, 9, 7)

    def test_parses_ordinary_shapes(self):
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-09-07T09:00:00Z") == date(2026, 9, 7)
        assert event_start_date("2026-09-07T09:00:00+02:00") == date(2026, 9, 7)
        assert event_start_date("2026-09-07") == date(2026, 9, 7)

    def test_rejects_garbage(self):
        from outlook_mcp.tools._recurrence import event_start_date

        with pytest.raises(ValueError, match="Invalid event start"):
            event_start_date("not-a-date")

    def test_an_instant_takes_its_date_in_the_anchor_zone(self):
        """01:00Z on Thursday is 18:00 on Wednesday in Los Angeles.

        Both dates are correct answers to different questions; the one Graph
        needs is the event's date in the zone the series is expanded against.
        """
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-10-29T01:00:00Z") == date(2026, 10, 29)
        assert event_start_date("2026-10-29T01:00:00Z", "America/Los_Angeles") == date(2026, 10, 28)

    def test_a_naive_start_is_never_shifted(self):
        """It is already wall-clock time in the zone; converting would move it."""
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-10-29T01:00:00", "America/Los_Angeles") == date(2026, 10, 29)
        assert event_start_date("2026-10-29", "Asia/Tokyo") == date(2026, 10, 29)

    def test_an_unresolvable_zone_falls_back_to_the_written_date(self):
        """Graph hands back Windows zone names and Python maps none of them.

        Falling back to the text is the behaviour from before anchoring
        existed: wrong only when the offset and the zone disagree, and never
        worse than not trying. Refusing instead would break `update_event` for
        every event Graph reports with a Windows name.
        """
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-10-29T01:00:00Z", "Pacific Standard Time") == date(
            2026, 10, 29
        )


class TestMovePattern:
    """A series' pattern names days, and those days belong to its start date.

    When the start's local date moves — a re-anchor across midnight, or a series
    moved to another day — the pattern has to move with it, or Graph schedules
    every occurrence on the old day beside a master on the new one.
    """

    @staticmethod
    def _move(pattern: dict, old: date, new: date) -> dict:
        from outlook_mcp.tools._recurrence import move_pattern

        return move_pattern({"pattern": pattern, "range": {"type": "noEnd"}}, old=old, new=new)[
            "pattern"
        ]

    def test_weekly_days_move_together_with_the_week_boundary(self):
        """A fortnightly Sunday-and-Monday block stays one block a day earlier.

        Shifting the days and leaving `firstDayOfWeek` on Sunday would split
        Saturday and Sunday across two weeks of the fortnight.
        """
        moved = self._move(
            {
                "type": "weekly",
                "interval": 2,
                "daysOfWeek": ["sunday", "monday"],
                "firstDayOfWeek": "sunday",
            },
            old=date(2026, 11, 8),
            new=date(2026, 11, 7),
        )
        assert moved["daysOfWeek"] == ["saturday", "sunday"]
        assert moved["firstDayOfWeek"] == "saturday"

    def test_a_missing_week_boundary_is_graphs_sunday_and_moves_too(self):
        """A hand-written fortnightly pattern that omits `firstDayOfWeek`.

        Graph reads the omission as Sunday. Shifting the days and leaving the
        boundary unset kept it on Sunday, so `[saturday, sunday]` straddled it
        and the Sundays landed a week late. Graph's own read-back always carries
        the field, so only a pattern a caller wrote can reach this.
        """
        moved = self._move(
            {"type": "weekly", "interval": 2, "daysOfWeek": ["sunday", "monday"]},
            old=date(2026, 11, 8),
            new=date(2026, 11, 7),
        )
        assert moved["daysOfWeek"] == ["saturday", "sunday"]
        assert moved["firstDayOfWeek"] == "saturday"

    @pytest.mark.parametrize("boundary", [{"firstDayOfWeek": "sunday"}, {}], ids=["set", "omitted"])
    def test_every_week_the_boundary_is_left_as_it_was(self, boundary):
        """Every week, the boundary schedules nothing — only the week start shown.

        `same_pattern` ignores it at `interval == 1` for the same reason, so
        moving it here would change a field the comparison calls irrelevant.
        """
        moved = self._move(
            {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"], **boundary},
            old=date(2026, 11, 5),
            new=date(2026, 11, 4),
        )
        assert moved["daysOfWeek"] == ["wednesday"]
        assert moved.get("firstDayOfWeek") == boundary.get("firstDayOfWeek")

    def test_weekly_wraps_forward_across_the_week(self):
        moved = self._move(
            {"type": "weekly", "interval": 1, "daysOfWeek": ["saturday"]},
            old=date(2026, 11, 7),
            new=date(2026, 11, 8),
        )
        assert moved["daysOfWeek"] == ["sunday"]

    def test_an_unmoved_date_changes_nothing(self):
        pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 1}
        assert self._move(pattern, old=date(2026, 11, 1), new=date(2026, 11, 1)) == pattern

    def test_daily_has_no_days_to_move(self):
        pattern = {"type": "daily", "interval": 1}
        assert self._move(pattern, old=date(2026, 11, 5), new=date(2026, 11, 4)) == pattern

    def test_monthly_takes_the_new_day_within_the_month(self):
        moved = self._move(
            {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15},
            old=date(2026, 11, 15),
            new=date(2026, 11, 14),
        )
        assert moved["dayOfMonth"] == 14

    def test_monthly_pushed_into_another_month_is_refused(self):
        """The day before "the 1st" is not any single dayOfMonth."""
        with pytest.raises(ValueError, match="different month"):
            self._move(
                {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 1},
                old=date(2026, 12, 1),
                new=date(2026, 11, 30),
            )

    def test_yearly_takes_the_new_month_and_day(self):
        moved = self._move(
            {"type": "absoluteYearly", "interval": 1, "month": 1, "dayOfMonth": 1},
            old=date(2027, 1, 1),
            new=date(2026, 12, 31),
        )
        assert (moved["month"], moved["dayOfMonth"]) == (12, 31)

    def test_an_absolute_pattern_not_anchored_on_its_start_is_refused(self):
        with pytest.raises(ValueError, match="does not match that start date"):
            self._move(
                {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 20},
                old=date(2026, 11, 15),
                new=date(2026, 11, 14),
            )

    def test_a_relative_pattern_whose_day_moves_is_refused(self):
        """The day before the first Thursday is not always the first Wednesday."""
        with pytest.raises(ValueError, match="relative pattern"):
            self._move(
                {
                    "type": "relativeMonthly",
                    "interval": 1,
                    "daysOfWeek": ["thursday"],
                    "index": "first",
                },
                old=date(2026, 11, 5),
                new=date(2026, 11, 4),
            )


class TestSamePattern:
    def test_graph_defaults_do_not_make_a_pattern_different(self):
        """Graph fills in `month: 0`, `dayOfMonth: 0` and `index` on a weekly read.

        A hand-written pattern describing the same days is the same pattern;
        comparing whole dicts would call it an edit and skip moving it.
        """
        from outlook_mcp.tools._recurrence import same_pattern

        read_back = {
            "type": "weekly",
            "interval": 1,
            "month": 0,
            "dayOfMonth": 0,
            "daysOfWeek": ["thursday"],
            "firstDayOfWeek": "sunday",
            "index": "first",
        }
        assert same_pattern({"type": "weekly", "daysOfWeek": ["Thursday"]}, read_back)

    def test_a_different_day_is_a_different_pattern(self):
        from outlook_mcp.tools._recurrence import same_pattern

        assert not same_pattern(
            {"type": "weekly", "daysOfWeek": ["monday"]},
            {"type": "weekly", "daysOfWeek": ["thursday"]},
        )

    def test_nothing_is_the_same_as_nothing(self):
        from outlook_mcp.tools._recurrence import same_pattern

        assert not same_pattern(None, None)

    def test_a_moved_week_boundary_is_a_different_fortnightly_pattern(self):
        """Every other week, `firstDayOfWeek` decides which days share a week.

        A caller who changed only the boundary has written a new schedule, so it
        must not be taken for the stored one and shifted a second time.
        """
        from outlook_mcp.tools._recurrence import same_pattern

        base = {"type": "weekly", "interval": 2, "daysOfWeek": ["sunday", "monday"]}
        assert not same_pattern(
            {**base, "firstDayOfWeek": "monday"}, {**base, "firstDayOfWeek": "sunday"}
        )
        # Graph's default is Sunday, so omitting it is the same boundary.
        assert same_pattern(base, {**base, "firstDayOfWeek": "sunday"})

    def test_the_week_boundary_does_not_matter_every_week(self):
        from outlook_mcp.tools._recurrence import same_pattern

        base = {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]}
        assert same_pattern({**base, "firstDayOfWeek": "monday"}, base)


_WRONG_TYPES = [
    ({"pattern": {"type": "daily"}, "range": "invalid"}, "range must be an object"),
    ({"pattern": {"type": "daily"}, "range": ["noEnd"]}, "range must be an object"),
    ({"pattern": "daily", "range": {"type": "noEnd"}}, "pattern must be an object"),
    (
        {"pattern": {"type": "daily", "interval": None}, "range": {"type": "noEnd"}},
        "pattern.interval must be a whole number",
    ),
    (
        {"pattern": {"type": "daily", "interval": 2.5}, "range": {"type": "noEnd"}},
        "pattern.interval must be a whole number",
    ),
    (
        {"pattern": {"type": "daily", "interval": True}, "range": {"type": "noEnd"}},
        "pattern.interval must be a whole number",
    ),
    (
        {"pattern": {"type": "weekly", "daysOfWeek": "monday"}, "range": {"type": "noEnd"}},
        "daysOfWeek must be a list",
    ),
    (
        {"pattern": {"type": "weekly", "daysOfWeek": [1]}, "range": {"type": "noEnd"}},
        "Invalid pattern.daysOfWeek 1",
    ),
    ({"pattern": {"type": 7}, "range": {"type": "noEnd"}}, "Invalid pattern.type 7"),
    (
        {"pattern": {"type": "daily"}, "range": {"type": "numbered", "numberOfOccurrences": None}},
        "numberOfOccurrences must be a whole number",
    ),
    (
        {"pattern": {"type": "daily"}, "range": {"type": "endDate", "endDate": 20270101}},
        "endDate must be a YYYY-MM-DD string",
    ),
    (
        {"pattern": {"type": "daily"}, "range": {"type": "noEnd", "recurrenceTimeZone": 5}},
        "recurrenceTimeZone must be a zone name string",
    ),
]


class TestWrongJsonTypesAreRefusedByName:
    """A value of the wrong JSON type is caller input, so it gets an input error.

    These used to escape as `TypeError` / `AttributeError`, which the server
    reports as a crash with the message withheld — or, for a string where the
    pattern object belongs, were accepted: `"type" in "daily"` is a substring
    test, so the converter built an empty pattern without complaint.
    """

    @pytest.mark.parametrize("recurrence,message", _WRONG_TYPES)
    def test_the_converter_names_the_field(self, recurrence, message):
        with pytest.raises(ValueError, match=message):
            build_event_recurrence(recurrence, start=_START)

    @pytest.mark.parametrize("recurrence,message", _WRONG_TYPES)
    def test_the_up_front_shape_check_names_it_too(self, recurrence, message):
        from outlook_mcp.tools._recurrence import check_recurrence_shape

        with pytest.raises(ValueError, match=message):
            check_recurrence_shape(recurrence)


class TestPatternTypeSpelling:
    """The converter accepts `weekly` and the SDK's `Weekly`; comparison and moves must too.

    Seen case-sensitively, a `Weekly` pattern skipped every type-specific field,
    so an edited Monday pattern compared equal to a stored Thursday one and was
    shifted as if it were the series' own.
    """

    def test_a_pascal_case_pattern_on_a_different_day_is_a_different_pattern(self):
        from outlook_mcp.tools._recurrence import same_pattern

        assert not same_pattern(
            {"type": "Weekly", "daysOfWeek": ["monday"]},
            {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]},
        )

    def test_a_pascal_case_echo_is_the_same_pattern(self):
        from outlook_mcp.tools._recurrence import same_pattern

        assert same_pattern(
            {"type": "Weekly", "daysOfWeek": ["thursday"]},
            {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]},
        )

    def test_a_pascal_case_pattern_moves_like_any_other(self):
        from outlook_mcp.tools._recurrence import move_pattern

        moved = move_pattern(
            {"pattern": {"type": "Weekly", "daysOfWeek": ["thursday"]}, "range": {}},
            old=date(2026, 11, 5),
            new=date(2026, 11, 4),
        )
        assert moved["pattern"]["daysOfWeek"] == ["wednesday"]


def test_a_whole_number_in_another_spelling_is_still_accepted():
    """The control: a digit string and an integral float were accepted before, and still are."""
    built = build_event_recurrence(
        {
            "pattern": {"type": "daily", "interval": "2"},
            "range": {"type": "numbered", "numberOfOccurrences": 3.0},
        },
        start=_START,
    )
    assert built.pattern.interval == 2
    assert built.range.number_of_occurrences == 3


class TestNumbersInAnotherSpelling:
    """The converter reads `"15"` as 15, so comparison and moves have to as well.

    Compared raw, an echoed monthly pattern carrying `"15"` was taken for an edit
    of the stored 15 and sent unmoved across a date boundary, and a move of it
    was refused as not anchored on its own start.
    """

    def test_a_string_day_is_the_same_day(self):
        from outlook_mcp.tools._recurrence import same_pattern

        assert same_pattern(
            {"type": "absoluteMonthly", "interval": "1", "dayOfMonth": "15"},
            {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15},
        )

    def test_a_string_day_moves(self):
        from outlook_mcp.tools._recurrence import move_pattern

        moved = move_pattern(
            {"pattern": {"type": "absoluteMonthly", "dayOfMonth": "15"}, "range": {}},
            old=date(2026, 11, 15),
            new=date(2026, 11, 14),
        )
        assert moved["pattern"]["dayOfMonth"] == 14
