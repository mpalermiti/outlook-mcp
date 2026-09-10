"""Relative date offsets on every datetime parameter.

Two months of trajectories show the agent reaching for `after="7d"` fifty times
and being refused every time, then falling back to a hand-computed ISO timestamp.
It was guessing the convention every other CLI uses, and the old error —
"Invalid datetime format: 7d" — never told it what the accepted format was.

Direction follows the CLI convention the agent was reaching for: bare and `-`
mean *ago*, `+` means *from now*. That asymmetry is the one thing here that could
silently produce a wrong answer (a task due "7d" landing in the past), so it is
pinned in both directions and stated in every docstring that takes a date.
"""

from datetime import datetime, timedelta, timezone

import pytest

from outlook_mcp.validation import validate_datetime

NOW = datetime(2026, 9, 10, 12, 0, 0, tzinfo=timezone.utc)


def _expect(delta: timedelta) -> str:
    return (NOW + delta).strftime("%Y-%m-%dT%H:%M:%SZ")


@pytest.mark.parametrize(
    ("value", "delta"),
    [
        ("7d", timedelta(days=-7)),
        ("-7d", timedelta(days=-7)),
        ("+7d", timedelta(days=7)),
        ("24h", timedelta(hours=-24)),
        ("+2h", timedelta(hours=2)),
        ("30m", timedelta(minutes=-30)),
        ("+90m", timedelta(minutes=90)),
        ("2w", timedelta(weeks=-2)),
        ("+1w", timedelta(weeks=1)),
        ("0d", timedelta()),
        ("now", timedelta()),
    ],
)
def test_relative_offsets_resolve_against_now(value, delta):
    assert validate_datetime(value, now=NOW) == _expect(delta)


@pytest.mark.parametrize("value", ["7D", "24H", "2W", "NOW", "+1W"])
def test_units_are_case_insensitive(value):
    """An agent that shouts is still asking for the same thing."""
    assert validate_datetime(value, now=NOW).endswith("Z")


def test_bare_offset_means_ago_not_from_now():
    """The direction that matters: a wrong reading here backdates a task."""
    assert validate_datetime("7d", now=NOW) < NOW.strftime("%Y-%m-%dT%H:%M:%SZ")
    assert validate_datetime("+7d", now=NOW) > NOW.strftime("%Y-%m-%dT%H:%M:%SZ")


def test_relative_ignores_the_config_timezone():
    """An offset from now is an instant; no zone can change it."""
    assert validate_datetime("7d", "America/Los_Angeles", now=NOW) == validate_datetime(
        "7d", "UTC", now=NOW
    )


def test_absolute_iso_still_works():
    assert validate_datetime("2026-07-04T03:44:03Z") == "2026-07-04T03:44:03Z"


def test_bare_date_still_works():
    assert validate_datetime("2026-10-22") == "2026-10-22T00:00:00Z"


def test_zone_less_input_still_honours_the_config_timezone():
    """Guard the 1.15.0 behaviour: naive input is not resolved against the host clock."""
    assert validate_datetime("2026-10-22", "America/Los_Angeles") == "2026-10-22T07:00:00Z"


@pytest.mark.parametrize(
    "value",
    ["7y", "7", "d7", "abc", "7 d", "--7d", "+", "7dd", "1e3d", "99999999999d"],
)
def test_unsupported_shapes_are_rejected(value):
    with pytest.raises(ValueError):
        validate_datetime(value, now=NOW)


def test_error_names_the_accepted_formats():
    """The old message said only what was wrong. This one says what to write."""
    with pytest.raises(ValueError) as exc:
        validate_datetime("last tuesday", now=NOW)

    message = str(exc.value)
    assert "7d" in message  # a relative example
    assert "+7d" in message  # and the other direction
    assert "2026" in message or "ISO" in message  # the absolute form


def test_now_defaults_to_the_current_clock():
    """Without an injected clock it uses real time — the production path."""
    before = datetime.now(timezone.utc) - timedelta(seconds=5)
    resolved = datetime.strptime(validate_datetime("0d"), "%Y-%m-%dT%H:%M:%SZ").replace(
        tzinfo=timezone.utc
    )
    assert resolved >= before
