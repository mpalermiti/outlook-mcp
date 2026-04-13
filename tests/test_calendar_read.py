"""Tests for calendar read tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools.calendar_read import (
    _format_event_summary,
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
