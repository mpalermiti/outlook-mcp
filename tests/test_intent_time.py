"""Typed time windows for the intent tools (contracts §3): ``validation.resolve_window``.

The intent tools take a window as a closed set of typed forms and parse no
natural language. A window is a half-open ``[start, end)`` pair of UTC instants,
computed in one zone: the zone that both reads the input and shows the times.
The cases that go wrong silently are the ones pinned here: day lengths across a
daylight-saving change, a midnight that does not exist, and the sign of an offset.
"""

from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pytest

from outlook_mcp.validation import WINDOW_FORMS, resolve_window

UTC = timezone.utc
# A Thursday.
NOW = datetime(2026, 11, 26, 15, 30, tzinfo=UTC)


def _utc(*args) -> datetime:
    return datetime(*args, tzinfo=UTC)


def _local(zone: str, *args) -> datetime:
    return datetime(*args, tzinfo=ZoneInfo(zone)).astimezone(UTC)


# ── Keywords ──────────────────────────────────────────────────────────────


@pytest.mark.parametrize("zone", ["UTC", "America/Los_Angeles", "Asia/Kolkata"])
def test_today_and_tomorrow_are_local_days(zone):
    local_today = NOW.astimezone(ZoneInfo(zone)).date()
    start, end = resolve_window("today", zone, now=NOW)
    assert start == _local(zone, local_today.year, local_today.month, local_today.day)
    assert end == start + timedelta(days=1)
    tomorrow = resolve_window("tomorrow", zone, now=NOW)
    assert tomorrow == (end, end + timedelta(days=1))


@pytest.mark.parametrize("zone", ["UTC", "America/Los_Angeles", "Asia/Kolkata"])
def test_weeks_start_on_monday(zone):
    start, end = resolve_window("this_week", zone, now=NOW)
    # 2026-11-26 is a Thursday everywhere at NOW: the week began Monday the 23rd.
    assert start == _local(zone, 2026, 11, 23)
    assert end == _local(zone, 2026, 11, 30)
    assert resolve_window("next_week", zone, now=NOW) == (end, _local(zone, 2026, 12, 7))


@pytest.mark.parametrize("zone", ["UTC", "America/Los_Angeles", "Asia/Kolkata"])
def test_sunday_2359_is_still_this_week_and_monday_0000_starts_the_next(zone):
    sunday_late = _local(zone, 2026, 11, 29, 23, 59)
    start, end = resolve_window("this_week", zone, now=sunday_late)
    assert (start, end) == (_local(zone, 2026, 11, 23), _local(zone, 2026, 11, 30))

    monday = _local(zone, 2026, 11, 30, 0, 0)
    start, _ = resolve_window("this_week", zone, now=monday)
    assert start == monday
    assert resolve_window("today", zone, now=monday)[0] == monday


def test_keywords_ignore_case():
    assert resolve_window("Today", "UTC", now=NOW) == resolve_window("today", "UTC", now=NOW)


# ── Day lengths and the missing midnight ──────────────────────────────────


def test_a_spring_forward_day_is_23_hours_and_a_fall_back_day_25():
    start, end = resolve_window("2026-03-08", "America/Los_Angeles", now=NOW)
    assert end - start == timedelta(hours=23)
    start, end = resolve_window("2026-11-01", "America/Los_Angeles", now=NOW)
    assert end - start == timedelta(hours=25)


def test_a_midnight_inside_a_daylight_saving_gap_starts_at_the_first_instant_after_it():
    # Santiago skips 00:00-01:00 on 2026-09-06: the day starts at 01:00 local.
    start, end = resolve_window("2026-09-06", "America/Santiago", now=NOW)
    assert start == _utc(2026, 9, 6, 4, 0)
    assert start.astimezone(ZoneInfo("America/Santiago")).hour == 1
    assert end == _local("America/Santiago", 2026, 9, 7)


def test_every_local_midnight_of_a_year_round_trips_or_starts_after_its_gap():
    zone = ZoneInfo("America/Santiago")
    day = date(2026, 1, 1)
    while day.year == 2026:
        start, end = resolve_window(day.isoformat(), "America/Santiago", now=NOW)
        local = start.astimezone(zone)
        assert local.date() == day
        assert local.hour in (0, 1)
        assert end > start
        day += timedelta(days=1)


# ── Offsets and intervals ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("+7d", (NOW, NOW + timedelta(days=7))),
        ("7d", (NOW - timedelta(days=7), NOW)),
        ("-7d", (NOW - timedelta(days=7), NOW)),
        ("+24h", (NOW, NOW + timedelta(hours=24))),
        ("24h", (NOW - timedelta(hours=24), NOW)),
        ("+2w", (NOW, NOW + timedelta(weeks=2))),
        ("90m", (NOW - timedelta(minutes=90), NOW)),
    ],
)
def test_a_signed_offset_looks_forward_and_a_bare_one_looks_back(value, expected):
    assert resolve_window(value, "UTC", now=NOW) == expected


def test_an_interval_takes_an_instant_on_each_side():
    assert resolve_window("now/+7d", "UTC", now=NOW) == (NOW, NOW + timedelta(days=7))
    start, end = resolve_window(
        "2026-11-27T09:00:00/2026-11-27T17:00:00", "America/Los_Angeles", now=NOW
    )
    assert start == _local("America/Los_Angeles", 2026, 11, 27, 9)
    assert end == _local("America/Los_Angeles", 2026, 11, 27, 17)


def test_an_interval_in_the_wrong_order_is_refused():
    with pytest.raises(ValueError, match="later"):
        resolve_window("+7d/now", "UTC", now=NOW)
    with pytest.raises(ValueError, match="later"):
        resolve_window("now/now", "UTC", now=NOW)


def test_a_span_over_the_tools_maximum_is_refused():
    assert resolve_window("+400d", "UTC", now=NOW, max_days=400)
    with pytest.raises(ValueError, match="400 days"):
        resolve_window("+401d", "UTC", now=NOW, max_days=400)
    with pytest.raises(ValueError, match="31 days"):
        resolve_window("2026-01-01/2026-03-01", "UTC", now=NOW, max_days=31)


# ── The zone ──────────────────────────────────────────────────────────────


def test_the_zone_given_reads_the_input():
    utc_day = resolve_window("2026-11-27", "UTC", now=NOW)
    kolkata_day = resolve_window("2026-11-27", "Asia/Kolkata", now=NOW)
    assert kolkata_day[0] == utc_day[0] - timedelta(hours=5, minutes=30)


def test_an_abbreviation_is_refused_by_the_zone_check():
    with pytest.raises(ValueError, match="abbreviation"):
        resolve_window("today", "PST", now=NOW)


# ── Refusals ──────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "value",
    ["tomorrow at 3", "next tuesday", "", "  ", "today/tomorrow", "+7x", "2026-13-01", "7d/"],
)
def test_anything_outside_the_grammar_is_refused_with_every_form(value):
    with pytest.raises(ValueError) as refused:
        resolve_window(value, "UTC", now=NOW)
    text = str(refused.value)
    for form in ("today", "tomorrow", "this_week", "next_week", "+7d", "7d", "now/+7d"):
        assert form in text
    assert text.count(WINDOW_FORMS) == 1


def test_a_phrase_is_answered_with_the_typed_time_to_pass_instead():
    with pytest.raises(ValueError, match=r"2026-11-27T15:00"):
        resolve_window("tomorrow at 3", "UTC", now=NOW)


def test_control_characters_are_refused():
    with pytest.raises(ValueError):
        resolve_window("today\x00", "UTC", now=NOW)


# ── Clock times (§3.1 `clock`): agenda's day_start and day_end ─────────────


@pytest.mark.parametrize(
    ("value", "minutes"), [("00:00", 0), ("09:00", 540), ("18:30", 1110), ("23:59", 1439)]
)
def test_a_clock_is_minutes_after_midnight(value, minutes):
    from outlook_mcp.validation import validate_clock

    assert validate_clock(value) == minutes


def test_twenty_four_hundred_is_only_a_days_end():
    from outlook_mcp.validation import validate_clock

    assert validate_clock("24:00", end_of_day=True) == 1440
    with pytest.raises(ValueError, match="24:00"):
        validate_clock("24:00")


@pytest.mark.parametrize("value", ["9:00", "09:60", "25:00", "24:01", "0900", "noon", "", " 09:00"])
def test_anything_else_is_refused(value):
    from outlook_mcp.validation import validate_clock

    with pytest.raises(ValueError, match="HH:MM"):
        validate_clock(value, end_of_day=True)
