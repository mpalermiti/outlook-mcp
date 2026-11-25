"""Offline guard for the #98 live test's verdict.

`TestMovingAMastersDateMovesItsOccurrences` was a strict `xfail` while #98 was
open. With the fix it is the regression guard, so `_assert_moved` decides
whether a live run shows green or red: only three consecutive Tuesdays at the
original time of day may pass. #98's own signature (the original Mondays,
untouched) and every other landing must fail red. The live test itself never
runs offline; this checks the assertion it ends on, with fake occurrences.
"""

from __future__ import annotations

from datetime import date, timedelta
from types import SimpleNamespace

import pytest

import tests.test_live_calendar_write as live

MONDAY = date(2026, 11, 16)
TUESDAY = MONDAY + timedelta(days=1)


def _occurrences(first: date, clock: str):
    return [
        SimpleNamespace(
            start=SimpleNamespace(date_time=f"{(first + timedelta(weeks=n)).isoformat()}T{clock}")
        )
        for n in range(3)
    ]


def _verdict(first: date, clock: str, time_of_day: str) -> str:
    assert_moved = live.TestMovingAMastersDateMovesItsOccurrences._assert_moved
    try:
        assert_moved(_occurrences(first, clock), MONDAY, time_of_day=time_of_day)
    except AssertionError as e:
        return f"red: {e}"
    return "pass"


@pytest.mark.parametrize(
    ("time_of_day", "clock"),
    [("09:00:00", "09:00:00.0000000"), ("00:00:00", "00:00:00.0000000")],
)
def test_the_issue_98_signature_fails_red_and_names_the_issue(time_of_day, clock):
    assert _verdict(MONDAY, clock, time_of_day).startswith("red: #98 is back")


@pytest.mark.parametrize(
    ("first", "clock", "time_of_day"),
    [
        (MONDAY, "04:00:00.0000000", "09:00:00"),  # Mondays, but the time moved
        (MONDAY + timedelta(weeks=1), "09:00:00.0000000", "09:00:00"),  # a week late
        (MONDAY, "08:00:00.0000000", "00:00:00"),  # all-day not at midnight
        (TUESDAY + timedelta(days=1), "09:00:00.0000000", "09:00:00"),  # Wednesdays
        (TUESDAY, "10:00:00.0000000", "09:00:00"),  # Tuesdays, but the time moved
    ],
)
def test_any_other_landing_fails_red(first, clock, time_of_day):
    assert _verdict(first, clock, time_of_day).startswith("red")


def test_three_tuesdays_at_the_same_time_pass():
    assert _verdict(TUESDAY, "09:00:00.0000000", "09:00:00") == "pass"
