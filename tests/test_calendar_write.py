"""Tests for calendar write tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.calendar_write import (
    create_event,
    delete_event,
    rsvp,
    update_event,
)


def _make_event_builder():
    """Create a MagicMock event builder with async endpoints.

    by_event_id is a sync method on the real Graph SDK that returns a
    request builder. We use MagicMock so chained attribute access
    (e.g., .accept.post) works without producing coroutines.
    """
    builder = MagicMock()
    builder.get = AsyncMock()
    builder.patch = AsyncMock()
    builder.delete = AsyncMock()
    builder.accept.post = AsyncMock()
    builder.decline.post = AsyncMock()
    builder.tentatively_accept.post = AsyncMock()
    return builder


class TestCreateEvent:
    async def test_create_event_calls_post(self):
        """create_event builds Event and calls me.events.post()."""
        mock_client = AsyncMock()
        mock_response = MagicMock()
        mock_response.id = "AAMkAGnew="
        mock_response.subject = "Lunch"
        mock_client.me.events.post = AsyncMock(return_value=mock_response)

        result = await create_event(
            mock_client,
            subject="Lunch",
            start="2026-04-15T12:00:00Z",
            end="2026-04-15T13:00:00Z",
        )
        assert result["status"] == "created"
        assert result["event_id"] == "AAMkAGnew="
        mock_client.me.events.post.assert_called_once()

    async def test_create_event_raises_read_only(self):
        """create_event raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await create_event(
                mock_client,
                subject="Lunch",
                start="2026-04-15T12:00:00Z",
                end="2026-04-15T13:00:00Z",
                read_only=True,
            )

    async def test_create_event_with_attendees(self):
        """create_event passes attendees to Event object."""
        mock_client = AsyncMock()
        mock_response = MagicMock()
        mock_response.id = "AAMkAGnew="
        mock_response.subject = "Team Sync"
        mock_client.me.events.post = AsyncMock(return_value=mock_response)

        result = await create_event(
            mock_client,
            subject="Team Sync",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            attendees=["alice@test.com", "bob@test.com"],
            is_online=True,
        )
        assert result["status"] == "created"
        mock_client.me.events.post.assert_called_once()


class TestUpdateEvent:
    async def test_update_event_patches_fields(self):
        """update_event PATCHes changed fields on the event."""
        builder = _make_event_builder()
        mock_response = MagicMock()
        mock_response.id = "AAMkAG123="
        mock_response.subject = "Updated Meeting"
        builder.patch = AsyncMock(return_value=mock_response)

        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            subject="Updated Meeting",
            location="Room 202",
        )
        assert result["status"] == "updated"
        assert result["event_id"] == "AAMkAG123="
        builder.patch.assert_called_once()

    async def test_update_event_raises_read_only(self):
        """update_event raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                subject="Nope",
                read_only=True,
            )


class TestDeleteEvent:
    async def test_delete_event_calls_delete(self):
        """delete_event calls .delete() on the event."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await delete_event(mock_client, event_id="AAMkAG123=")
        assert result["status"] == "deleted"
        builder.delete.assert_called_once()

    async def test_delete_event_raises_read_only(self):
        """delete_event raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await delete_event(mock_client, event_id="AAMkAG123=", read_only=True)


class TestRsvp:
    async def test_rsvp_accept_calls_accept_post(self):
        """rsvp with response=accept calls accept.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(mock_client, event_id="AAMkAG123=", response="accept")
        assert result["status"] == "accepted"
        builder.accept.post.assert_called_once()

    async def test_rsvp_decline_calls_decline_post(self):
        """rsvp with response=decline calls decline.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(mock_client, event_id="AAMkAG123=", response="decline")
        assert result["status"] == "declined"
        builder.decline.post.assert_called_once()

    async def test_rsvp_tentative_calls_tentatively_accept_post(self):
        """rsvp with response=tentative calls tentatively_accept.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(mock_client, event_id="AAMkAG123=", response="tentative")
        assert result["status"] == "tentativelyAccepted"
        builder.tentatively_accept.post.assert_called_once()

    async def test_rsvp_raises_read_only(self):
        """rsvp raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await rsvp(
                mock_client,
                event_id="AAMkAG123=",
                response="accept",
                read_only=True,
            )
