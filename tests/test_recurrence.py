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
        assert event_start_date("2026-10-29T01:00:00Z", "America/Los_Angeles") == date(
            2026, 10, 28
        )

    def test_a_naive_start_is_never_shifted(self):
        """It is already wall-clock time in the zone; converting would move it."""
        from outlook_mcp.tools._recurrence import event_start_date

        assert event_start_date("2026-10-29T01:00:00", "America/Los_Angeles") == date(
            2026, 10, 29
        )
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
