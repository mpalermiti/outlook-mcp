"""Tests for calendar read tools."""

from unittest.mock import AsyncMock, MagicMock
from zoneinfo import ZoneInfoNotFoundError

import pytest

from outlook_mcp.tools.calendar_read import (
    _compute_calendar_range,
    _format_event_summary,
    _resolve_timezone,
    get_event,
    list_events,
)


def _make_mock_event(**overrides):
    """Factory for mock Graph SDK event objects."""
    event = MagicMock()
    event.id = overrides.get("id", "AAMkAG123=")
    event.subject = overrides.get("subject", "Team Meeting")
    event.start = MagicMock(
        date_time=overrides.get("start_dt", "2026-04-15T10:00:00"),
        time_zone=overrides.get("start_tz", "UTC"),
    )
    event.end = MagicMock(
        date_time=overrides.get("end_dt", "2026-04-15T11:00:00"),
        time_zone=overrides.get("end_tz", "UTC"),
    )
    event.location = MagicMock(display_name=overrides.get("location", "Room 101"))
    event.is_all_day = overrides.get("is_all_day", False)
    event.organizer = MagicMock()
    event.organizer.email_address = MagicMock()
    event.organizer.email_address.name = overrides.get("organizer_name", "Boss")
    event.organizer.email_address.address = overrides.get("organizer_email", "boss@test.com")
    event.response_status = MagicMock(
        response=MagicMock(value=overrides.get("response_status", "accepted"))
    )
    event.is_online_meeting = overrides.get("is_online", False)
    event.categories = overrides.get("categories", [])
    # Detail fields
    event.body = MagicMock(content=overrides.get("body_content", "<p>Agenda here</p>"))
    event.online_meeting = MagicMock(
        join_url=overrides.get("join_url", None)
    )
    event.recurrence = overrides.get("recurrence", None)
    event.type = overrides.get("type", None)
    attendee_data = overrides.get("attendees", [])
    attendees = []
    for a in attendee_data:
        att = MagicMock()
        att.email_address = MagicMock()
        att.email_address.name = a.get("name", "")
        att.email_address.address = a.get("email", "")
        att.status = MagicMock(
            response=MagicMock(value=a.get("response", "none"))
        )
        attendees.append(att)
    event.attendees = attendees
    return event


class TestFormatEventSummary:
    def test_formats_basic_event(self):
        """_format_event_summary extracts all expected fields."""
        event = _make_mock_event()
        result = _format_event_summary(event)
        assert result["id"] == "AAMkAG123="
        assert result["subject"] == "Team Meeting"
        assert result["start"] == "2026-04-15T10:00:00 (UTC)"
        assert result["end"] == "2026-04-15T11:00:00 (UTC)"
        assert result["location"] == "Room 101"
        assert result["is_all_day"] is False
        assert result["organizer"] == "Boss"
        assert result["response_status"] == "accepted"
        assert result["is_online"] is False


class TestListEvents:
    async def test_list_events_returns_events(self):
        """list_events returns structured event list."""
        mock_event = _make_mock_event()
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC")
        assert result["count"] == 1
        assert result["events"][0]["subject"] == "Team Meeting"
        assert result["has_more"] is False

    async def test_list_events_days_computes_range(self):
        """days param computes start/end relative to now."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        result = await list_events(mock_client, days=3, timezone="UTC")
        assert result["count"] == 0
        # Verify calendarView was called
        mock_client.me.calendar_view.get.assert_called_once()

    async def test_list_events_with_explicit_dates(self):
        """list_events with after/before validates and uses them."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        result = await list_events(
            mock_client,
            after="2026-04-15T00:00:00Z",
            before="2026-04-22T00:00:00Z",
            timezone="UTC",
        )
        assert result["count"] == 0

    async def test_list_events_rejects_invalid_dates(self):
        """list_events rejects bad date strings."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError, match="Invalid datetime"):
            await list_events(mock_client, after="not-a-date", timezone="UTC")


class TestListEventsConcise:
    async def test_list_events_concise_drops_body_and_attendees(self):
        """concise=True returns is_organizer + attendees_count, no body/organizer/categories."""
        mock_event = _make_mock_event(
            attendees=[
                {"name": "Alice", "email": "alice@test.com", "response": "accepted"},
                {"name": "Bob", "email": "bob@test.com", "response": "none"},
            ],
            categories=["Blue Category"],
        )
        # is_organizer is read via getattr, set it explicitly on the mock.
        mock_event.is_organizer = True

        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC", concise=True)
        event = result["events"][0]
        assert event["id"] == "AAMkAG123="
        assert event["subject"] == "Team Meeting"
        # Concise-mode-only fields.
        assert event["is_organizer"] is True
        assert event["is_online_meeting"] is False
        assert event["attendees_count"] == 2
        # Fields that must be dropped.
        assert "organizer" not in event
        assert "response_status" not in event
        assert "categories" not in event
        assert "body" not in event
        assert "attendees" not in event

    async def test_list_events_default_keeps_full_shape(self):
        """concise=False (default) preserves the existing event shape."""
        mock_event = _make_mock_event()
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC")
        event = result["events"][0]
        # Existing fields.
        assert "organizer" in event
        assert "response_status" in event
        assert "is_online" in event
        # New concise-only fields must NOT bleed into default mode.
        assert "is_organizer" not in event
        assert "attendees_count" not in event


class TestGetEvent:
    async def test_get_event_validates_id(self):
        """get_event rejects invalid event IDs."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await get_event(mock_client, event_id="bad id with spaces!")

    async def test_get_event_returns_full_detail(self):
        """get_event returns full event detail including body and attendees."""
        mock_event = _make_mock_event(
            body_content="<p>Full agenda</p>",
            join_url="https://teams.microsoft.com/join/123",
            attendees=[
                {"name": "Alice", "email": "alice@test.com", "response": "accepted"},
                {"name": "Bob", "email": "bob@test.com", "response": "tentativelyAccepted"},
            ],
            categories=["Blue Category"],
        )
        builder = MagicMock()
        builder.get = AsyncMock(return_value=mock_event)
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await get_event(mock_client, event_id="AAMkAG123=")
        assert result["id"] == "AAMkAG123="
        assert result["body"] == "<p>Full agenda</p>"
        assert result["online_meeting_url"] == "https://teams.microsoft.com/join/123"
        assert len(result["attendees"]) == 2
        assert result["attendees"][0]["name"] == "Alice"
        assert result["attendees"][0]["response"] == "accepted"
        assert result["categories"] == ["Blue Category"]

class TestEventDetailRecurrence:
    """#41 follow-on: recurrence must come back as the same JSON shape create accepts."""

    def _detail(self, **overrides):
        from outlook_mcp.tools.calendar_read import _format_event_detail

        return _format_event_detail(_make_mock_event(**overrides))

    def test_no_recurrence_is_none(self):
        assert self._detail()["recurrence"] is None

    def test_recurrence_is_a_dict_not_a_repr(self):
        """str(PatternedRecurrence) leaked a ~500-char Python repr into the payload."""
        from outlook_mcp.tools._recurrence import build_event_recurrence

        pr = build_event_recurrence(
            {
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {"type": "noEnd", "startDate": "2026-09-07"},
            },
            start="2026-09-07T12:30:00Z",
        )

        assert self._detail(recurrence=pr)["recurrence"] == {
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
            "range": {"type": "noEnd", "startDate": "2026-09-07"},
        }

    def test_series_master_type_is_surfaced(self):
        """`type` is what tells a client the series actually took — see #41's repro."""
        assert self._detail(type=MagicMock(value="seriesMaster"))["type"] == "seriesMaster"

    def test_missing_type_is_empty_string(self):
        assert self._detail()["type"] == ""

    def test_summary_carries_type_so_listings_can_tell_series_apart(self):
        summary = _format_event_summary(_make_mock_event(type=MagicMock(value="seriesMaster")))

        assert summary["type"] == "seriesMaster"

    def test_summary_still_omits_the_recurrence_object(self):
        """`type` is one short string; the full pattern stays detail-only."""
        assert "recurrence" not in _format_event_summary(_make_mock_event())

    def test_concise_mode_omits_type(self):
        from outlook_mcp.tools.calendar_read import _format_event_concise

        concise = _format_event_concise(_make_mock_event(type=MagicMock(value="seriesMaster")))

        assert "type" not in concise


class TestTimezoneResolution:
    """A calendar range needs a time zone database; say so when there isn't one.

    The regression these guard: on a host with no IANA database (Windows, or a
    slim Linux image) every calendar tool failed with a bare `Error executing
    tool outlook_list_events` and no text, because `_wrap_tool_errors` holds an
    unexpected exception's message server-side.
    """

    def test_a_real_zone_resolves(self):
        assert _resolve_timezone("America/Los_Angeles").key == "America/Los_Angeles"

    def test_a_typo_names_the_zone_and_where_to_fix_it(self):
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("America/Los_Angelez")
        message = str(excinfo.value)
        assert "America/Los_Angelez" in message
        assert "config.json" in message

    def test_a_missing_database_blames_the_install_not_the_config(self, monkeypatch):
        monkeypatch.setattr(
            "outlook_mcp.tools.calendar_read.importlib.util.find_spec",
            lambda name: None,
        )
        monkeypatch.setattr(
            "outlook_mcp.tools.calendar_read.ZoneInfo",
            MagicMock(side_effect=ZoneInfoNotFoundError("no such key")),
        )
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("America/Los_Angeles")
        message = str(excinfo.value)
        assert "tzdata" in message
        assert "config.json" not in message

    def test_the_message_survives_the_wrapper_the_model_sees(self):
        """ValueError is the anticipated-failure channel, so the text reaches the caller."""
        with pytest.raises(ValueError) as excinfo:
            _compute_calendar_range(7, None, None, "Not/AZone")
        assert "Not/AZone" in str(excinfo.value)

    def test_a_long_zone_name_is_truncated_like_every_other_echo(self):
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("X" * 200)
        assert "X" * 51 not in str(excinfo.value)


class TestCalendarRangeWithoutBounds:
    def test_days_widens_the_window_from_now(self):
        start, end = _compute_calendar_range(7, None, None, "America/Los_Angeles")
        assert start < end
        assert start.endswith("Z") and end.endswith("Z")

    def test_explicit_bounds_win_over_days(self):
        start, end = _compute_calendar_range(
            7, "2026-09-11T00:00:00", "2026-09-12T00:00:00", "America/Los_Angeles"
        )
        # Midnight Pacific on a DST date is 07:00Z.
        assert start == "2026-09-11T07:00:00Z"
        assert end == "2026-09-12T07:00:00Z"
