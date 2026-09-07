"""Tests for calendar write tools."""

from datetime import date
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.calendar_write import (
    create_event,
    delete_event,
    rsvp,
    update_event,
)

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _created(subject: str, event_id: str = "AAMkAGnew="):
    """Mock Graph response for a successful event POST."""
    resp = MagicMock()
    resp.id = event_id
    resp.subject = subject
    return resp


def _posted(mock_client):
    """The Event object handed to me.events.post()."""
    return mock_client.me.events.post.call_args[0][0]


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
            config=_CFG,
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
                config=_CFG_RO,
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
            config=_CFG,
        )
        assert result["status"] == "created"
        mock_client.me.events.post.assert_called_once()

    async def test_create_event_ignores_no_recurrence(self):
        """No recurrence argument leaves the Event's recurrence unset (regression guard)."""
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Lunch"))

        await create_event(
            mock_client,
            subject="Lunch",
            start="2026-04-15T12:00:00Z",
            end="2026-04-15T13:00:00Z",
            config=_CFG,
        )

        assert _posted(mock_client).recurrence is None

    async def test_create_event_sets_recurrence_from_dict(self):
        """#41: a Graph-shape recurrence dict reaches the posted Event as PatternedRecurrence."""
        from msgraph.generated.models.day_of_week import DayOfWeek
        from msgraph.generated.models.patterned_recurrence import PatternedRecurrence
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Weekly sync"))

        await create_event(
            mock_client,
            subject="Weekly sync",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:30:00Z",
            recurrence={
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday", "friday"]},
                "range": {"type": "noEnd", "startDate": "2026-09-07"},
            },
            config=_CFG,
        )

        posted = _posted(mock_client)
        assert isinstance(posted.recurrence, PatternedRecurrence)
        assert posted.recurrence.pattern.type is RecurrencePatternType.Weekly
        assert posted.recurrence.pattern.days_of_week == [DayOfWeek.Monday, DayOfWeek.Friday]

    async def test_create_event_sets_recurrence_from_json_string(self):
        """The workaround shape from #41 — a JSON string — now produces a real series."""
        import json

        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Weekly sync"))

        await create_event(
            mock_client,
            subject="Weekly sync",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:30:00Z",
            recurrence=json.dumps(
                {
                    "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                    "range": {"type": "noEnd", "startDate": "2026-09-07"},
                }
            ),
            config=_CFG,
        )

        assert _posted(mock_client).recurrence.pattern.type is RecurrencePatternType.Weekly

    async def test_create_event_expands_the_documented_shorthand(self):
        """The docstring has advertised "weekly" since 1.0 — make it mean something."""
        from msgraph.generated.models.day_of_week import DayOfWeek
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        await create_event(
            mock_client,
            subject="Standup",
            start="2026-09-07T09:00:00Z",  # a Monday
            end="2026-09-07T09:15:00Z",
            recurrence="weekly",
            config=_CFG,
        )

        posted = _posted(mock_client)
        assert posted.recurrence.pattern.type is RecurrencePatternType.Weekly
        assert posted.recurrence.pattern.days_of_week == [DayOfWeek.Monday]

    async def test_create_event_rejects_an_unknown_shorthand(self):
        """Silently dropping an unparseable recurrence is what caused #41."""
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Nope"))

        with pytest.raises(ValueError, match="daily"):
            await create_event(
                mock_client,
                subject="Nope",
                start="2026-09-07T09:00:00Z",
                end="2026-09-07T09:15:00Z",
                recurrence="biweekly",
                config=_CFG,
            )

        mock_client.me.events.post.assert_not_called()

    async def test_create_event_checks_permission_before_building_recurrence(self):
        """Read-only mode must short-circuit before any recurrence parsing."""
        mock_client = AsyncMock()

        with pytest.raises(ReadOnlyError):
            await create_event(
                mock_client,
                subject="Weekly sync",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:30:00Z",
                recurrence="biweekly",
                config=_CFG_RO,
            )


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
            config=_CFG,
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
                config=_CFG_RO,
            )

    async def test_update_event_ignores_no_recurrence(self):
        """Omitting recurrence leaves it unset on the patch (regression guard)."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(mock_client, event_id="AAMkAG123=", subject="X", config=_CFG)

        assert builder.patch.call_args[0][0].recurrence is None
        builder.get.assert_not_called()

    async def test_update_event_sets_recurrence_with_explicit_start(self):
        """Patching start + recurrence together needs no extra round trip."""
        from msgraph.generated.models.day_of_week import DayOfWeek
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:30:00Z",
            recurrence="weekly",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.type is RecurrencePatternType.Weekly
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Monday]
        builder.get.assert_not_called()

    async def test_update_event_reads_current_start_when_none_given(self):
        """Recurrence alone: Graph needs the range anchored on the event's own start."""
        from msgraph.generated.models.day_of_week import DayOfWeek

        current = MagicMock()
        # Graph returns 7 fractional digits, which datetime.fromisoformat rejects on 3.10.
        current.start = MagicMock(date_time="2026-09-07T12:30:00.0000000", time_zone="UTC")

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG
        )

        builder.get.assert_awaited_once()
        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Monday]
        assert patched.recurrence.range.start_date == date(2026, 9, 7)

    async def test_update_event_accepts_a_graph_recurrence_object(self):
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-09-07T12:30:00Z",
            recurrence={
                "pattern": {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 7},
                "range": {"type": "numbered", "numberOfOccurrences": 3},
            },
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.type is RecurrencePatternType.AbsoluteMonthly
        assert patched.recurrence.range.number_of_occurrences == 3

    async def test_update_event_errors_when_start_cannot_be_resolved(self):
        """Better a named error than a Graph 400 on a series with no anchor."""
        current = MagicMock()
        current.start = None

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="start"):
            await update_event(
                mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG
            )

        builder.patch.assert_not_called()

    async def test_update_event_rejects_bad_recurrence_before_patching(self):
        builder = _make_event_builder()
        builder.patch = AsyncMock()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="daily"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-09-07T12:30:00Z",
                recurrence="biweekly",
                config=_CFG,
            )

        builder.patch.assert_not_called()


class TestDeleteEvent:
    async def test_delete_event_calls_delete(self):
        """delete_event calls .delete() on the event."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await delete_event(mock_client, event_id="AAMkAG123=", config=_CFG)
        assert result["status"] == "deleted"
        builder.delete.assert_called_once()

    async def test_delete_event_raises_read_only(self):
        """delete_event raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await delete_event(mock_client, event_id="AAMkAG123=", config=_CFG_RO)


class TestRsvp:
    async def test_rsvp_accept_calls_accept_post(self):
        """rsvp with response=accept calls accept.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(
            mock_client, event_id="AAMkAG123=", response="accept", config=_CFG,
        )
        assert result["status"] == "accepted"
        builder.accept.post.assert_called_once()

    async def test_rsvp_decline_calls_decline_post(self):
        """rsvp with response=decline calls decline.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(
            mock_client, event_id="AAMkAG123=", response="decline", config=_CFG,
        )
        assert result["status"] == "declined"
        builder.decline.post.assert_called_once()

    async def test_rsvp_tentative_calls_tentatively_accept_post(self):
        """rsvp with response=tentative calls tentatively_accept.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(
            mock_client, event_id="AAMkAG123=", response="tentative", config=_CFG,
        )
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
                config=_CFG_RO,
            )
