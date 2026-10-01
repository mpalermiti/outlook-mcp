"""Tests for calendar write tools."""

from datetime import date
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools._recurrence import build_event_recurrence, serialize_recurrence
from outlook_mcp.tools.calendar_write import (
    create_event,
    delete_event,
    rsvp,
    update_event,
)

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)
# A configured zone that is not UTC, so "defaulted from config" and "hardcoded
# UTC" cannot pass the same assertion.
_CFG_LA = Config(client_id="test", timezone="America/Los_Angeles")


def _created(subject: str, event_id: str = "AAMkAGnew="):
    """Mock Graph response for a successful event POST."""
    resp = MagicMock()
    resp.id = event_id
    resp.subject = subject
    return resp


def _posted(mock_client):
    """The Event object handed to me.events.post()."""
    return mock_client.me.events.post.call_args[0][0]


#: The change key Graph returns in `additional_data["@odata.etag"]`, in its real
#: shape. Tests that assert the recurrence re-send pins its patch compare
#: against this rather than re-typing it.
CURRENT_ETAG = 'W/"a0B0kEE/hUqXUZYaDD6+nQAJKU6nXA=="'


def _current_event(
    date_time: str = "2026-09-07T12:30:00.0000000",
    start_zone: str = "UTC",
    end_zone: str | None = None,
    event_type: str = "singleInstance",
    recurrence=None,
    is_all_day: bool = False,
    etag: str | None = CURRENT_ETAG,
    cancelled: list | None = None,
    edited: list | None = None,
):
    """An event as Graph returns it from a plain GET.

    Two details are the wire shape rather than the obvious one, and both have
    already cost a bug each: `start.time_zone` is **UTC** whatever the event is
    anchored in, because Graph projects it unless the request carries a
    `Prefer: outlook.timezone` header and this server never sends one; and the
    anchor lives in `original_start_time_zone` / `original_end_time_zone`,
    which are separate because Graph stores them separately.

    Graph also returns seven fractional digits, which `fromisoformat` rejects
    before 3.11 and the project floor is 3.10.
    """
    event = MagicMock(type=MagicMock(value=event_type))
    event.start = MagicMock(date_time=date_time, time_zone="UTC")
    event.original_start_time_zone = start_zone
    event.original_end_time_zone = end_zone if end_zone is not None else start_zone
    event.recurrence = recurrence
    # Explicit, not left to MagicMock: an auto-created attribute is truthy, and
    # `update_event` reads this to decide whether to anchor the patch in UTC.
    # Left implicit, every mocked event looks like an all-day one.
    event.is_all_day = is_all_day
    # Explicit for the same reason as `is_all_day`, and it bites harder: left to
    # MagicMock, `additional_data` is truthy and `.get("@odata.etag")` hands back
    # another MagicMock, so code reading the change key gets a value no Graph
    # response contains and every assertion about it passes vacuously.
    event.additional_data = {"@odata.etag": etag} if etag is not None else {}
    # Explicit for the same reason again, and here a truthy MagicMock would
    # refuse every series-master time patch as one with occurrences to lose.
    # A plain GET really returns None for both; they are set to the empty lists
    # the `$select`ed occurrence read returns for a clean series, because this
    # fixture answers both reads and only that one is looked at for them.
    event.cancelled_occurrences = cancelled if cancelled is not None else []
    event.exception_occurrences = edited if edited is not None else []
    return event


def _thursday_utc_series():
    """A weekly series stored the way pre-anchoring events are: Thursday 02:00Z.

    That is Wednesday 18:00 in Los Angeles, so re-anchoring it there moves the
    local date back a day — the case the start-date-only re-derivation missed.
    `firstDayOfWeek` is set because Graph always returns it.
    """
    return build_event_recurrence(
        {
            "pattern": {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": ["thursday"],
                "firstDayOfWeek": "sunday",
            },
            "range": {"type": "numbered", "numberOfOccurrences": 3},
        },
        start="2026-11-05T02:00:00Z",
        zone="UTC",
    )


def _series_master_builder(recurrence, current=None):
    """A builder whose event is a UTC-anchored series master, and a client for it."""
    if current is None:
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=recurrence,
        )
    builder = _make_event_builder()
    builder.get = AsyncMock(return_value=current)
    builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
    mock_client = MagicMock()
    mock_client.me.events.by_event_id = MagicMock(return_value=builder)
    return builder, mock_client


def _make_event_builder():
    """Create a MagicMock event builder with async endpoints.

    by_event_id is a sync method on the real Graph SDK that returns a
    request builder. We use MagicMock so chained attribute access
    (e.g., .accept.post) works without producing coroutines.
    """
    builder = MagicMock()
    # `get` answers with the shape a plain GET actually returns, not a bare
    # MagicMock. An auto-created attribute here is truthy and stringifies to
    # something no Graph response contains, so code reading the wrong field —
    # or reading a field it should not — sails through looking correct. A
    # test that wants a different event replaces this wholesale.
    builder.get = AsyncMock(return_value=_current_event())
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
        """start + recurrence together: one read, shared, for the zone.

        This path made no request at all before the timezone fix. It now reads
        the event once, because Graph rejects a start patch carrying no
        ``timeZone`` and the zone the event is already stored in is the only
        one that does not relocate it. What this pins is that the read is
        *shared* with the recurrence anchor rather than repeated — the count is
        the assertion, not the absence.
        """
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
        assert builder.get.await_count == 1

    async def test_update_event_reads_current_start_when_none_given(self):
        """Recurrence alone: Graph needs the range anchored on the event's own start."""
        from msgraph.generated.models.day_of_week import DayOfWeek

        current = MagicMock()
        # Graph returns 7 fractional digits, which datetime.fromisoformat rejects on 3.10.
        current.start = MagicMock(date_time="2026-09-07T12:30:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "UTC"
        current.original_end_time_zone = "UTC"

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

    async def test_update_event_leaves_new_fields_unset_when_omitted(self):
        """Critical: omitting them must not force False/empty onto the event."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(mock_client, event_id="AAMkAG123=", subject="X", config=_CFG)

        patched = builder.patch.call_args[0][0]
        assert patched.attendees is None
        assert patched.is_all_day is None
        assert patched.is_online_meeting is None
        assert patched.show_as is None

    async def test_update_event_replaces_attendees(self):
        """Graph replaces the whole collection — we send exactly what the caller gave."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            attendees=["alice@test.com", "bob@test.com"],
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert [a.email_address.address for a in patched.attendees] == [
            "alice@test.com",
            "bob@test.com",
        ]

    async def test_update_event_empty_attendee_list_clears_them(self):
        """[] is a real instruction ("no attendees"), distinct from None ("don't touch")."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(mock_client, event_id="AAMkAG123=", attendees=[], config=_CFG)

        assert builder.patch.call_args[0][0].attendees == []

    async def test_update_event_validates_attendee_emails(self):
        builder = _make_event_builder()
        builder.patch = AsyncMock()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                attendees=["alice@test.com", "not-an-email"],
                config=_CFG,
            )

        builder.patch.assert_not_called()

    async def test_update_event_sets_is_all_day_with_bounds(self):
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-22T00:00:00Z",
            end="2026-10-23T00:00:00Z",
            is_all_day=True,
            config=_CFG,
        )

        assert builder.patch.call_args[0][0].is_all_day is True

    async def test_update_event_rejects_all_day_without_bounds(self):
        """Live Graph returns "Missing parameters: Event.Start" — name it here instead."""
        builder = _make_event_builder()
        builder.patch = AsyncMock()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="start and end"):
            await update_event(
                mock_client, event_id="AAMkAG123=", is_all_day=True, config=_CFG
            )

        builder.patch.assert_not_called()

    async def test_update_event_can_turn_all_day_off(self):
        """False is an instruction too — must not be swallowed as "unset"."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-22T00:00:00Z",
            end="2026-10-23T00:00:00Z",
            is_all_day=False,
            config=_CFG,
        )

        assert builder.patch.call_args[0][0].is_all_day is False
    async def test_remove_recurrence_sends_an_explicit_null(self):
        """The SDK DROPS `event.recurrence = None` — it must go via additional_data.

        Asserting on the serialized JSON, not on the attribute: the whole failure
        mode here is a payload the SDK silently omits, and an attribute-level
        assertion would stay green through exactly that bug.
        """
        from kiota_serialization_json.json_serialization_writer import (
            JsonSerializationWriter,
        )

        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG
        )

        writer = JsonSerializationWriter()
        builder.patch.call_args[0][0].serialize(writer)
        assert '"recurrence": null' in writer.get_serialized_content().decode()

    async def test_setting_event_recurrence_none_is_unsendable(self):
        """Pins why additional_data is used.

        The previous version of this test serialized with a bare
        ``JsonSerializationWriter`` and asserted the field was omitted. The
        adapter does not use that writer — it goes through the backing-store
        proxy, which emits the null — so the assertion described a serializer
        this code never touches and the "if kiota ever changes" escape hatch
        could never fire. Through the real path the plain assignment does not
        silently drop the field; it makes the whole payload unserializable.

        If kiota ever starts emitting a usable top-level null, this fails and
        the additional_data workaround can be simplified.
        """
        from msgraph.generated.models.event import Event

        from tests.test_write_payloads_reach_the_wire import wire

        event = Event()
        event.subject = "x"
        event.recurrence = None

        with pytest.raises(AssertionError, match="top-level field"):
            wire(event)

    async def test_omitting_remove_recurrence_sends_no_recurrence_key(self):
        """A partial patch must not blank a series just because it edited the subject."""
        from kiota_serialization_json.json_serialization_writer import (
            JsonSerializationWriter,
        )

        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(mock_client, event_id="AAMkAG123=", subject="X", config=_CFG)

        writer = JsonSerializationWriter()
        builder.patch.call_args[0][0].serialize(writer)
        assert "recurrence" not in writer.get_serialized_content().decode()

    async def test_remove_recurrence_conflicts_with_setting_one(self):
        builder = _make_event_builder()
        builder.patch = AsyncMock()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="both"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-09-07T12:30:00Z",
                recurrence="weekly",
                remove_recurrence=True,
                config=_CFG,
            )

        builder.patch.assert_not_called()

    async def test_remove_recurrence_needs_no_extra_round_trip(self):
        """Unlike setting one, clearing needs no start anchor — don't GET the event."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG
        )

        builder.get.assert_not_called()


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


class TestEventTimezone:
    """The zone an event is anchored in, on both write paths.

    Every one of these is about a value that used to be the literal string
    ``"UTC"`` regardless of what anyone asked for. The instant was right; the
    anchor was not, and the anchor is what a recurring series is expanded
    against — so a 09:00 weekly meeting became 08:00 the week the clocks went
    back, reported as ``status: updated`` throughout.
    """

    async def test_create_anchors_in_the_configured_zone_by_default(self):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        await create_event(
            mock_client,
            subject="Standup",
            start="2026-10-28T09:00:00",
            end="2026-10-28T09:15:00",
            config=_CFG_LA,
        )

        event = _posted(mock_client)
        assert event.start.time_zone == "America/Los_Angeles"
        assert event.end.time_zone == "America/Los_Angeles"

    async def test_create_keeps_the_caller_written_datetime_verbatim(self):
        """The string is not normalized to UTC on the way through.

        Two things depend on that. A naive value has no instant until a zone
        resolves it, and resolving it against the *host* clock is the bug
        1.15.0 removed; and Graph requires an all-day event to begin on a
        midnight boundary, which local midnight converted to UTC is not.
        """
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Offsite"))

        await create_event(
            mock_client,
            subject="Offsite",
            start="2026-10-28T09:00:00",
            end="2026-10-29T00:00:00",
            config=_CFG_LA,
        )

        assert _posted(mock_client).start.date_time == "2026-10-28T09:00:00"

    async def test_create_takes_an_explicit_zone_over_the_configured_one(self):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Review"))

        await create_event(
            mock_client,
            subject="Review",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/New_York",
            config=_CFG_LA,
        )

        assert _posted(mock_client).start.time_zone == "America/New_York"

    async def test_an_abbreviation_is_refused_before_the_network(self):
        """PDT is what an agent sends when a user says "3pm Pacific".

        Graph answers it with `400 TimeZoneNotSupportedException`; refusing
        locally costs no round trip and says which name to use instead.
        """
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Sync"))

        with pytest.raises(ValueError) as excinfo:
            await create_event(
                mock_client,
                subject="Sync",
                start="2026-10-28T09:00:00",
                end="2026-10-28T10:00:00",
                timezone="PDT",
                config=_CFG_LA,
            )

        assert "America/Los_Angeles" in str(excinfo.value)
        mock_client.me.events.post.assert_not_called()

    async def test_update_keeps_the_zone_the_event_is_stored_in(self):
        """Patching a colleague's New York meeting must not move it here.

        The mock models what a plain GET actually returns, and the difference
        is the whole finding: Graph projects `start` into UTC unless the
        request carries `Prefer: outlook.timezone`, which this server never
        sends, so `start.time_zone` reads "UTC" for every event regardless of
        its anchor. An earlier version of this test set `start.time_zone` to
        the anchor — a shape the wire never produces — and passed against code
        that read the wrong field and relocated every event it patched.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000", start_zone="America/New_York"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T11:00:00",
            end="2026-10-28T12:00:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "America/New_York"
        assert patched.end.time_zone == "America/New_York"

    async def test_a_series_is_anchored_on_its_date_in_the_events_own_zone(self):
        """`2026-10-29T01:00:00Z` in Los Angeles is Wednesday the 28th, 18:00.

        The recurrence has to be built against that date, not the one in the
        text. While every event was anchored in UTC the two were the same day
        by construction; once the anchor is real they diverge, and the text
        date builds a Thursday pattern starting the 29th for an event that
        happens on Wednesday the 28th.

        Graph does not refuse that. Verified live 2026-09-21: it accepts the
        inconsistent master and schedules the whole series a day late —
        occurrences on Thursday 18:00 — answering `status: created`.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Evening sync"))

        await create_event(
            mock_client,
            subject="Evening sync",
            start="2026-10-29T01:00:00Z",
            end="2026-10-29T02:00:00Z",
            timezone="America/Los_Angeles",
            recurrence="weekly",
            config=_CFG_LA,
        )

        posted = _posted(mock_client)
        assert posted.recurrence.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert posted.recurrence.range.start_date == date(2026, 10, 28)

    async def test_a_naive_start_is_already_local_and_is_not_shifted(self):
        """The control. A zone-less start is wall-clock time in the zone.

        Converting it would be the 1.15.0 host-clock bug wearing a new hat, so
        the conversion has to apply only where the datetime names an instant.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Evening sync"))

        await create_event(
            mock_client,
            subject="Evening sync",
            start="2026-10-29T18:00:00",
            end="2026-10-29T19:00:00",
            timezone="America/Los_Angeles",
            recurrence="weekly",
            config=_CFG_LA,
        )

        posted = _posted(mock_client)
        assert posted.recurrence.pattern.days_of_week == [DayOfWeek.Thursday]
        assert posted.recurrence.range.start_date == date(2026, 10, 29)

    async def test_the_legacy_warning_only_claims_fixed_offset_where_it_is_true(
        self, caplog, monkeypatch
    ):
        """`CET` is Graph-refused *and* observes daylight saving.

        Both zones take the same UTC fallback, so the only thing separating
        them is the sentence the operator reads. Telling them CET "never
        observes daylight saving" is false diagnostic information in the one
        message they get, and the lookup tables being right does not help if
        the prose is wrong.
        """
        for zone, expect_fixed in (("EST", True), ("CET", False)):
            monkeypatch.setattr("outlook_mcp.validation._warned_legacy_config_zone", set())
            caplog.clear()
            mock_client = AsyncMock()
            mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

            with caplog.at_level("WARNING"):
                await create_event(
                    mock_client,
                    subject="Standup",
                    start="2026-07-01T09:00:00",
                    end="2026-07-01T09:30:00",
                    config=Config(client_id="test", timezone=zone),
                )

            assert _posted(mock_client).start.time_zone == "UTC", zone
            assert ("never observes daylight saving" in caplog.text) is expect_fixed, (
                f"{zone}: fixed-offset claim present={not expect_fixed and 'wrongly' or ''}"
            )

    async def test_an_explicit_zone_re_anchors_both_ends(self):
        """"Move this to Eastern" means the whole event, not one end of it.

        Preserving a split only makes sense while the caller has named no zone;
        once they have, applying it to one end and not the other leaves the
        event in a state nobody asked for.
        """
        current = _current_event(
            date_time="2026-11-10T13:00:00.0000000",
            start_zone="America/New_York",
            end_zone="America/Los_Angeles",
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-10T09:00:00",
            end="2026-11-10T10:00:00",
            timezone="Europe/London",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "Europe/London"
        assert patched.end.time_zone == "Europe/London"

    async def test_a_zone_on_an_all_day_event_is_refused_not_ignored(self):
        """Graph coerces an all-day anchor to UTC, so the zone cannot be applied.

        Accepting it and substituting UTC reports `updated` for work not done —
        the shape this project keeps four guards against, and the one the new
        argument was most likely to reintroduce. Refused both ways: when the
        caller says `is_all_day=True` outright, which costs no round trip, and
        when only the stored event knows, which unavoidably costs one.
        """
        # Stated by the caller — refused before any network call.
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-10-22T00:00:00",
                end="2026-10-23T00:00:00",
                is_all_day=True,
                timezone="America/New_York",
                config=_CFG_LA,
            )
        assert "all-day" in str(excinfo.value)
        assert builder.get.await_count == 0
        builder.patch.assert_not_called()

        # Known only from the stored event — refused after the one read.
        builder = _make_event_builder()
        builder.get = AsyncMock(
            return_value=_current_event(
                date_time="2026-10-22T00:00:00.0000000", is_all_day=True
            )
        )
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-10-22T00:00:00",
                end="2026-10-23T00:00:00",
                timezone="America/New_York",
                config=_CFG_LA,
            )
        assert "all-day" in str(excinfo.value)
        builder.patch.assert_not_called()

    async def test_a_zone_with_no_times_to_apply_it_to_is_refused(self):
        """A lone timezone would patch nothing and answer `updated`.

        Graph rejects a start patch carrying no timeZone, so the zone is never
        an independent edit; accepting it alone would be the silent no-op this
        project keeps four guards against. Refused before any network call,
        like every other input check here.
        """
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                timezone="America/New_York",
                config=_CFG_LA,
            )

        assert "requires start and end" in str(excinfo.value)
        assert builder.get.await_count == 0
        builder.patch.assert_not_called()

    async def test_a_series_master_zone_change_brings_its_recurrence_along(self):
        """Graph refuses the patch without it, and says nothing about zones.

        `400 ErrorPropertyValidationFailure: At least one property failed
        validation` is the whole of the message. Verified live: the identical
        patch succeeds when the recurrence travels with it. Without this, every
        series created before events carried a zone is unrepairable except by
        deleting and rebuilding it.
        """
        from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=build_event_recurrence("weekly", start="2026-10-28T16:00:00Z"),
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "America/Los_Angeles"
        assert patched.recurrence is not None
        assert patched.recurrence.pattern.type is RecurrencePatternType.Weekly
        # Re-anchored on the new start rather than echoing the stored range, and
        # the stale recurrenceTimeZone dropped so Graph re-derives it.
        assert patched.recurrence.range.start_date == date(2026, 10, 28)
        assert patched.recurrence.range.recurrence_time_zone is None

    async def test_an_echoed_recurrence_loses_its_stale_range_zone(self):
        """The read-modify-write round trip plus a new `timezone`.

        `outlook_get_event` returns `recurrence` including
        `range.recurrenceTimeZone`, and the docstring invites callers to hand it
        straight back. Combined with an explicit `timezone`, that sends the old
        zone beside the new anchor — which Graph refuses with
        `400 ErrorPropertyValidationFailure`, turning an ordinary edit into an
        opaque failure. The explicit argument is the more specific instruction,
        so the stale field goes.

        This is narrower than the general passthrough I declined to remove in
        the create path: there a conflicting zone is caller error, loudly
        refused. Here it is two arguments of the same call disagreeing, and the
        call is one the tool advertises.
        """
        echoed = {
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["wednesday"]},
            "range": {
                "type": "numbered",
                "numberOfOccurrences": 2,
                "recurrenceTimeZone": "Pacific Standard Time",
            },
        }

        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000",
            start_zone="America/Los_Angeles",
            event_type="seriesMaster",
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            recurrence=echoed,
            timezone="America/New_York",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "America/New_York"
        assert patched.recurrence.range.recurrence_time_zone is None
        assert patched.recurrence.range.number_of_occurrences == 2

    async def test_an_echoed_recurrence_keeps_its_zone_without_an_explicit_one(self):
        """The control, and the reason this is a flag rather than a rule.

        With no `timezone` argument the caller's `recurrenceTimeZone` is the only
        statement about the range's zone, and removing it would be the sanitizer
        mistake — discarding input for no stated reason. It survives.
        """
        echoed = {
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["wednesday"]},
            "range": {
                "type": "numbered",
                "numberOfOccurrences": 2,
                "recurrenceTimeZone": "Pacific Standard Time",
            },
        }

        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000",
            start_zone="America/Los_Angeles",
            event_type="seriesMaster",
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            recurrence=echoed,
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.range.recurrence_time_zone == "Pacific Standard Time"

    async def test_a_single_event_keeps_its_recurrence_untouched(self):
        """The re-send is a workaround for one Graph rule, not a habit.

        A single instance has no recurrence to carry, and sending one would turn
        a time edit into a series — the opposite of a partial patch.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000", start_zone="UTC", recurrence=None
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        assert builder.patch.call_args[0][0].recurrence is None

    async def test_a_resent_recurrence_pins_the_patch_to_the_version_it_read(self):
        """The re-send echoes back state it read, so it must not clobber a newer one.

        Every other field in this patch is something the caller supplied; the
        recurrence is the one thing read from Graph and handed straight back. A
        concurrent re-patterning between the GET and the PATCH would therefore be
        silently reverted to whatever this call read. `If-Match` turns that into
        `412 ErrorIrresolvableConflict` instead, which `errors._HINT_TABLE`
        answers with "re-read the item and send the update again".

        Verified live on a consumer mailbox: a stale change key is refused and
        the current one succeeds, on the SDK's own request builder.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=build_event_recurrence("weekly", start="2026-10-28T16:00:00Z"),
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        # The header, not merely "a request configuration was passed" — a config
        # carrying nothing, or the wrong change key, would pass that.
        config_sent = builder.patch.call_args.kwargs["request_configuration"]
        assert config_sent.headers.get("If-Match") == {CURRENT_ETAG}
        # And the recurrence really did travel, so this is the pin on that patch
        # rather than a header on an unrelated one.
        assert builder.patch.call_args[0][0].recurrence is not None

    async def test_a_patch_that_echoes_nothing_back_is_not_pinned(self):
        """The control. Pinning every patch would be a different tool.

        `update_event` is last-writer-wins by design everywhere else: the caller
        supplied those values and means them. Only a patch that reshapes a series
        master rests on state it read — the occurrence check, and the echoed
        recurrence — so only it conflicts. This one is a single event.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000", start_zone="America/Los_Angeles"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            subject="Renamed",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/New_York",
            config=_CFG_LA,
        )

        assert "request_configuration" not in builder.patch.call_args.kwargs

    @pytest.mark.parametrize(
        "change",
        [
            {"start": "2026-11-05T03:00:00Z", "end": "2026-11-05T03:30:00Z"},
            {"recurrence": "weekly"},
        ],
        ids=["same-zone-move", "recurrence-only"],
    )
    async def test_a_master_reshape_is_pinned_even_without_a_resend(self, change):
        """The occurrence check is a read the patch depends on, so the patch is pinned to it.

        Without the pin, an occurrence edited in another client after the check
        found none would be discarded by this patch, silently — the exact loss
        the check exists to prevent. Editing or deleting an occurrence moves the
        master's change key (verified live), so `If-Match` turns that race into
        a 412. The key is the *first* read's: the re-send echoes that one.
        """
        builder, mock_client = _series_master_builder(_thursday_utc_series())

        await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **change)

        config_sent = builder.patch.call_args.kwargs["request_configuration"]
        assert config_sent.headers.get("If-Match") == {CURRENT_ETAG}

    async def test_an_unchanged_zone_does_not_resend_the_recurrence(self):
        """The control. Same series, same zone, no re-send.

        Without this the test above passes for a version that re-sends the
        recurrence on every start patch, which would make each one a series
        rewrite.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000",
            start_zone="America/Los_Angeles",
            event_type="seriesMaster",
            recurrence=build_event_recurrence("weekly", start="2026-10-28T09:00:00"),
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T11:00:00",
            end="2026-10-28T12:00:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        assert builder.patch.call_args[0][0].recurrence is None

    async def test_a_resent_series_moves_its_weekday_when_the_local_date_moves(self):
        """The main repair case: a US evening series stored in UTC.

        Thursday 02:00Z is Wednesday 18:00 in Los Angeles. Re-deriving only
        `startDate` sent a Wednesday start and `startDate` beside
        `daysOfWeek: ["thursday"]`, and Graph scheduled every occurrence on
        Thursday — verified live, reported as `updated`. The days move by as
        many days as the start did. Every week, `firstDayOfWeek` schedules
        nothing, so it stays where it was rather than moving the week start
        Outlook displays.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        builder, mock_client = _series_master_builder(_thursday_utc_series())

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-04T18:00:00",
            end="2026-11-04T18:30:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        sent = builder.patch.call_args[0][0].recurrence
        assert sent.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert sent.pattern.first_day_of_week is DayOfWeek.Sunday
        assert sent.range.start_date == date(2026, 11, 4)
        assert sent.range.number_of_occurrences == 3

    async def test_an_echoed_series_crosses_midnight_with_a_new_zone(self):
        """`outlook_get_event`'s recurrence handed straight back, across a date boundary.

        It carries the stored `startDate`, which `_reconcile_range` refused as a
        mismatch — loud, but it meant the advertised round trip failed for
        exactly the events this argument exists to repair. A pattern equal to
        the stored one is the series' own, so it moves like the re-send does.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        stored = _thursday_utc_series()
        builder, mock_client = _series_master_builder(stored)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-04T18:00:00",
            end="2026-11-04T18:30:00",
            recurrence=serialize_recurrence(stored),
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        sent = builder.patch.call_args[0][0].recurrence
        assert sent.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert sent.range.start_date == date(2026, 11, 4)

    async def test_a_new_pattern_beside_a_stale_start_date_is_sent_as_given(self):
        """The control for the echo: a caller who changed the pattern meant it.

        Their Monday pattern was written for the Monday start they sent, so
        moving it by the start's shift would make it Friday. Only the stale
        `startDate` goes.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        stored = _thursday_utc_series()
        builder, mock_client = _series_master_builder(stored)
        edited = serialize_recurrence(stored)
        edited["pattern"]["daysOfWeek"] = ["monday"]

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-09T18:00:00",
            end="2026-11-09T18:30:00",
            recurrence=edited,
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        sent = builder.patch.call_args[0][0].recurrence
        assert sent.pattern.days_of_week == [DayOfWeek.Monday]
        assert sent.range.start_date == date(2026, 11, 9)

    async def test_a_master_without_a_readable_recurrence_is_refused(self):
        """Never `event.recurrence = None` — the #63/#64 unsendable-null shape.

        And never a patch without the recurrence either: Graph refuses a
        master's zone change unless it travels along, with a 400 that names
        neither. So the refusal is ours, with the remedy.
        """
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=None,
        )
        builder, mock_client = _series_master_builder(None, current=current)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-04T18:00:00",
                end="2026-11-04T18:30:00",
                timezone="America/Los_Angeles",
                config=_CFG_LA,
            )

        assert "Pass `recurrence` alongside `timezone`" in str(excinfo.value)
        builder.patch.assert_not_called()

    @pytest.mark.parametrize(
        "times",
        [
            {"start": "2026-11-04T18:00:00", "end": "2026-11-04T18:30:00",
             "timezone": "America/Los_Angeles"},
            {"start": "2026-11-05T03:00:00Z", "end": "2026-11-05T03:30:00Z"},
            {"end": "2026-11-05T02:45:00Z"},
            {"recurrence": {"pattern": {"type": "weekly", "daysOfWeek": ["thursday", "friday"]},
                            "range": {"type": "numbered", "numberOfOccurrences": 3}}},
        ],
        ids=["re-anchor", "same-zone-move", "end-only", "pattern-change"],
    )
    async def test_a_change_that_would_discard_occurrences_is_refused(self, times):
        """Graph restores every edited and deleted occurrence when a series is reshaped.

        Measured live, one edited and one deleted occurrence each time: all
        four of these shapes lost both, reported as success. A subject patch
        and a recurrence re-sent unchanged kept them — so the loss follows the
        change to the series' times or pattern, and the refusal is keyed on
        that rather than on the re-send. It names what would be lost, and
        nothing is patched.
        """
        moved = MagicMock(subject="Moved to 10", start=MagicMock(date_time="2026-11-12T03:00:00"))
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=_thursday_utc_series(),
            edited=[moved],
            cancelled=["OID.AAMkAG123=.2026-11-19"],
        )
        builder, mock_client = _series_master_builder(None, current=current)

        with pytest.raises(ValueError) as excinfo:
            await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **times)

        message = str(excinfo.value)
        assert "Moved to 10" in message and "2026-11-12T03:00:00" in message
        assert "deleted 2026-11-19" in message
        assert "Nothing was modified" in message
        builder.patch.assert_not_called()
        # The occurrence read names both collections — a plain GET omits them,
        # so a read that forgot to ask would see a series that looks clean — and
        # `subject` and `start`, because the nested exceptions take the same
        # projection: verified live, without them both come back None.
        selected = builder.get.call_args.kwargs["request_configuration"].query_parameters.select
        assert set(selected) == {"subject", "start", "cancelledOccurrences", "exceptionOccurrences"}

    async def test_an_occurrence_read_that_omits_them_is_refused(self):
        """Fail closed: a read that does not say is not a clean series."""
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000",
            start_zone="UTC",
            event_type="seriesMaster",
            recurrence=_thursday_utc_series(),
        )
        current.exception_occurrences = None
        builder, mock_client = _series_master_builder(None, current=current)

        with pytest.raises(ValueError, match="Could not read which occurrences"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-05T03:00:00Z",
                end="2026-11-05T03:30:00Z",
                config=_CFG_LA,
            )
        builder.patch.assert_not_called()

    async def test_update_rejects_bad_input_without_asking_graph_anything(self):
        """Every input check belongs above the first `await`.

        Before the anchor-zone read existed these rejected with zero network
        calls. Hoisting the read above them cost two things: a stale event id
        plus a malformed time surfaced as `404 ErrorItemNotFound` instead of
        the input error the caller could act on, and every rejected patch paid
        for a full-event round trip.

        `patch.assert_not_called()` cannot see this — the builder's `get`
        answers happily — so the count on `get` is the assertion.
        """
        cases = [
            ({"start": "not-a-date"}, "Invalid datetime"),
            ({"end": "also-not-a-date"}, "Invalid datetime"),
            ({"is_all_day": True}, "requires start and end"),
            ({"recurrence": "weekly", "remove_recurrence": True}, "not both"),
            # Checked before the reads a recurrence patch needs, so a bad shape
            # is named as one rather than surfacing as the occurrence refusal.
            ({"recurrence": "fortnightly"}, "Invalid recurrence"),
            (
                {"recurrence": {"pattern": {"type": "weekly", "daysOfWeek": ["funday"]},
                                "range": {"type": "noEnd"}}},
                "pattern.daysOfWeek",
            ),
            (
                {"recurrence": {"pattern": {"type": "daily"}, "range": "invalid"}},
                "range must be an object",
            ),
        ]
        for kwargs, expected in cases:
            builder = _make_event_builder()
            builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
            mock_client = MagicMock()
            mock_client.me.events.by_event_id = MagicMock(return_value=builder)

            with pytest.raises(ValueError) as excinfo:
                await update_event(
                    mock_client, event_id="AAMkAG123=", config=_CFG_LA, **kwargs
                )

            assert expected in str(excinfo.value), kwargs
            assert builder.get.await_count == 0, f"{kwargs} asked Graph before refusing"
            builder.patch.assert_not_called()

    async def test_an_all_day_event_is_anchored_in_utc(self):
        """Graph stores an all-day event in UTC whatever zone it is sent.

        Verified live 2026-09-21: a midnight `Z` and a naive midnight, both
        labelled America/Los_Angeles, came back `originalStartTimeZone: UTC`.
        So the zone buys nothing at Graph and costs correctness here — `00:00Z`
        labelled Los Angeles is 17:00 the previous day, which builds the
        recurrence for the wrong date and the wrong weekday. 2026-10-22 is a
        Thursday, and the shape below is the one this tool's own error text
        tells callers to use.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Offsite"))

        await create_event(
            mock_client,
            subject="Offsite",
            start="2026-10-22T00:00:00Z",
            end="2026-10-23T00:00:00Z",
            is_all_day=True,
            recurrence="weekly",
            config=_CFG_LA,
        )

        posted = _posted(mock_client)
        assert posted.start.time_zone == "UTC"
        assert posted.end.time_zone == "UTC"
        assert posted.recurrence.range.start_date == date(2026, 10, 22)
        assert posted.recurrence.pattern.days_of_week == [DayOfWeek.Thursday]

    async def test_an_all_day_patch_is_anchored_in_utc_too(self):
        """Including when the stored event is all-day and the caller says nothing.

        A patch that leaves `is_all_day` alone still has to agree with what the
        event actually is, or it re-anchors an all-day event into a zone Graph
        will discard — having first used that zone to derive the wrong date.
        """
        current = _current_event(
            date_time="2026-10-22T00:00:00.0000000",
            start_zone="America/New_York",
            is_all_day=True,
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-22T00:00:00",
            end="2026-10-23T00:00:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "UTC"
        assert patched.end.time_zone == "UTC"

    async def test_a_legacy_config_zone_keeps_working_and_says_so(self, caplog, monkeypatch):
        """An upgrade must not break a server whose config.json already says EST.

        `EST` resolves in Python, Graph refuses it, and this server never sent
        it anywhere before events carried a real zone — so an install holding
        one has been working fine and would, on upgrade, start failing every
        `outlook_create_event` that does not name a zone. Accept and warn is
        the rule for stored config; hard-error is for input no legacy file can
        carry.
        """
        # The latch is process-global and nothing resets it. This test passes
        # today only because it is the sole trigger in the run; a second one
        # anywhere would leave `caplog` empty and the assertion would pass for
        # the wrong reason.
        monkeypatch.setattr("outlook_mcp.validation._warned_legacy_config_zone", set())

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        with caplog.at_level("WARNING"):
            await create_event(
                mock_client,
                subject="Standup",
                start="2026-07-01T09:00:00",
                end="2026-07-01T09:30:00",
                config=Config(client_id="test", timezone="EST"),
            )

        # UTC, not America/New_York. The two are not the same zone: EST is a
        # fixed UTC-05:00 that never observes daylight saving, which is how
        # `resolve_timezone` — and so every calendar read — already reads this
        # config value. Anchoring writes in a DST-observing zone would make the
        # two halves of the server disagree about the same string all summer.
        # A July date is deliberate: it is when they differ.
        posted = _posted(mock_client)
        assert posted.start.time_zone == "UTC"
        assert posted.start.date_time == "2026-07-01T09:00:00"
        assert "EST" in caplog.text
        assert "America/New_York" in caplog.text, "the warning names the zone that works"
        assert "config.json" in caplog.text

    async def test_the_tolerance_does_not_extend_to_the_argument(self):
        """The asymmetry is the design, so it gets an assertion.

        A stored config value can predate the validation; a `timezone` passed
        on the call cannot. Tolerating it there would be inventing a zone the
        caller did not ask for, on input they wrote this second.
        """
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        with pytest.raises(ValueError, match="America/New_York"):
            await create_event(
                mock_client,
                subject="Standup",
                start="2026-10-28T09:00:00",
                end="2026-10-28T09:30:00",
                timezone="EST",
                config=Config(client_id="test", timezone="EST"),
            )

        mock_client.me.events.post.assert_not_called()

    async def test_a_typo_in_the_config_zone_is_still_an_error(self):
        """Tolerance is for zones that used to work, not for broken config.

        A misspelt `config.timezone` already fails every calendar *read*
        through `resolve_timezone`, so an install carrying one is visibly
        broken. Inventing a zone for its writes would make the two halves of
        the server disagree about the same config value.
        """
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        with pytest.raises(ValueError, match="Not/AZone"):
            await create_event(
                mock_client,
                subject="Standup",
                start="2026-10-28T09:00:00",
                end="2026-10-28T09:30:00",
                config=Config(client_id="test", timezone="Not/AZone"),
            )

    async def test_an_explicitly_empty_zone_is_refused_not_defaulted(self):
        """`timezone=""` is a caller who meant something and sent nothing.

        Treating it as "use the configured zone" anchors the event somewhere
        they never named and reports success, while `timezone="  "` — one
        keystroke away — is refused. Two spellings of the same mistake must not
        take different paths.
        """
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Sync"))

        for blank in ("", "   "):
            with pytest.raises(ValueError) as excinfo:
                await create_event(
                    mock_client,
                    subject="Sync",
                    start="2026-10-28T09:00:00",
                    end="2026-10-28T10:00:00",
                    timezone=blank,
                    config=_CFG_LA,
                )
            assert "empty" in str(excinfo.value), blank

        mock_client.me.events.post.assert_not_called()

    async def test_update_keeps_a_split_zone_event_split(self):
        """A flight leaves New York and lands in Los Angeles.

        Graph stores the two anchors separately — verified live 2026-09-21,
        08:00 `America/New_York` to 11:00 `America/Los_Angeles` comes back as
        13:00Z to 19:00Z with both intact. Deriving one zone from the start and
        stamping it on both ends relabels the landing time and moves it three
        hours, silently, in a patch that only meant to shift the departure.
        """
        current = _current_event(
            date_time="2026-11-10T13:00:00.0000000",
            start_zone="America/New_York",
            end_zone="America/Los_Angeles",
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-10T09:00:00",
            end="2026-11-10T12:00:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "America/New_York"
        assert patched.end.time_zone == "America/Los_Angeles"

    async def test_an_end_only_patch_uses_the_end_anchor(self):
        """The narrowest case, and the one that shows the two are independent."""
        current = _current_event(
            date_time="2026-11-10T13:00:00.0000000",
            start_zone="America/New_York",
            end_zone="America/Los_Angeles",
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            end="2026-11-10T12:30:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.end.time_zone == "America/Los_Angeles"
        assert patched.start is None

    async def test_a_recurrence_only_update_uses_the_events_local_date(self):
        """The stored start Graph returns is the UTC one, and its date can differ.

        `start.dateTime` comes back naive and already projected into UTC, so a
        18:00 Pacific event reads as 01:00 the next day. Building the pattern
        from that text puts the series on the wrong weekday. The instant has to
        be reconstructed and read in the event's own zone.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        current = _current_event(
            date_time="2026-10-29T01:00:00.0000000", start_zone="America/Los_Angeles"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            recurrence="weekly",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert patched.recurrence.range.start_date == date(2026, 10, 28)

    async def test_a_prefer_header_graph_ignored_is_refused(self):
        """A start projected into some other zone is UTC text posing as local time.

        Graph echoes the zone it was asked for in `start.timeZone` — verified
        live for `Pacific Standard Time` and `America/Los_Angeles` alike — so a
        second read that still says `UTC` means the header was not honoured.
        Reading `02:00` as Pacific wall clock would build the series a day late
        and report `updated`; one comparison turns that into a refusal before
        anything is patched.
        """
        # Every read answers the UTC projection, as if the header were ignored.
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000", start_zone="Pacific Standard Time"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                recurrence="weekly",
                config=_CFG_LA,
            )

        assert "'Pacific Standard Time'" in str(excinfo.value)
        assert "'UTC'" in str(excinfo.value)
        builder.patch.assert_not_called()

    async def test_a_second_read_without_a_start_is_refused(self):
        """The same wrong-day outcome as an ignored header, reached another way.

        Falling back to the date as written means the UTC date, which is the
        bug the second read exists to fix. So a reply with no start refuses too.
        """
        utc_view = _current_event(
            date_time="2026-11-05T02:00:00.0000000", start_zone="Pacific Standard Time"
        )
        startless = _current_event(start_zone="Pacific Standard Time")
        startless.start = None

        def _get(request_configuration=None):
            prefers = request_configuration and request_configuration.headers.get("Prefer")
            return startless if prefers else utc_view

        builder = _make_event_builder()
        builder.get = AsyncMock(side_effect=_get)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="no start at all"):
            await update_event(
                mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG_LA
            )
        builder.patch.assert_not_called()

    async def test_a_windows_zone_evening_event_takes_its_local_weekday(self):
        """An 18:00 Pacific event is next-day in UTC, and the series follows the local day.

        This replaces a test that pinned the *wrong* answer. Graph returns the
        stored start projected into UTC and names the event's zone in Windows
        terms, which Python cannot map — so `2026-11-05T02:00:00Z` for a
        Wednesday-evening event built a Thursday series, and Graph accepted it
        and scheduled the whole thing a day late.

        The mock answers the second read — the one carrying
        `Prefer: outlook.timezone` — with the local wall clock, because that is
        what Graph does. A `MagicMock` that ignores the header returns the UTC
        value to both reads and the test passes against the unfixed code, which
        is exactly how this went unnoticed once already.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        utc_view = _current_event(
            date_time="2026-11-05T02:00:00.0000000", start_zone="Pacific Standard Time"
        )
        local_view = _current_event(
            date_time="2026-11-04T18:00:00.0000000", start_zone="Pacific Standard Time"
        )
        # Graph names the projection it applied, echoing the requested zone.
        local_view.start.time_zone = "Pacific Standard Time"

        # Graph answers the local wall clock only when the request actually
        # carries `Prefer: outlook.timezone` naming the event's own zone, so the
        # mock keys on the header rather than on a request config being present
        # at all: a version that sent an empty config, a misspelled preference
        # or someone else's zone would otherwise pass this.
        prefers = []

        def _get(request_configuration=None):
            preference = None
            if request_configuration is not None:
                preference = request_configuration.headers.get("Prefer")
            prefers.append(preference)
            return local_view if preference else utc_view

        builder = _make_event_builder()
        builder.get = AsyncMock(side_effect=_get)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            recurrence="weekly",
            config=_CFG_LA,
        )

        assert prefers == [None, {'outlook.timezone="Pacific Standard Time"'}], (
            f"the second read did not ask for the event's own zone: {prefers}"
        )
        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert patched.recurrence.range.start_date == date(2026, 11, 4)

    async def test_a_resolvable_zone_needs_no_second_read(self):
        """The Prefer read is the fallback, not the default.

        An IANA anchor converts locally, so paying for a second GET on every
        recurrence-only update would be waste. The count is the assertion.
        """
        current = _current_event(
            date_time="2026-11-05T02:00:00.0000000", start_zone="America/Los_Angeles"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            recurrence="weekly",
            config=_CFG_LA,
        )

        assert builder.get.await_count == 1
        assert builder.patch.call_args[0][0].recurrence.range.start_date == date(2026, 11, 4)

    async def test_a_zone_less_start_is_accepted_with_an_unmappable_zone(self):
        """The control, and the remedy the refusal recommends.

        A zone-less start is already wall-clock time in the event's zone, so
        its date needs no conversion and the Windows name does not matter.
        A refusal here would make the advice in the error message wrong.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        current = _current_event(
            date_time="2026-10-29T01:00:00.0000000", start_zone="Pacific Standard Time"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T18:00:00",
            end="2026-10-28T19:00:00",
            recurrence="weekly",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Wednesday]
        assert patched.recurrence.range.start_date == date(2026, 10, 28)

    async def test_a_utc_anchored_event_needs_no_such_help(self):
        """The control: most events predating this change are anchored in UTC.

        There the stored date and the local date are the same day, so a
        recurrence-only update keeps working exactly as it did.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        current = _current_event(date_time="2026-10-29T01:00:00.0000000", start_zone="UTC")

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            recurrence="weekly",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Thursday]
        assert patched.recurrence.range.start_date == date(2026, 10, 29)

    async def test_update_refuses_rather_than_guess_an_unreadable_anchor(self):
        """An event in a custom zone reports `tzone://Microsoft/Custom`.

        Every available move is wrong: echoing the sentinel is a 400, and any
        substitute — the config zone, UTC — relocates the event while
        answering `updated`. So the call is refused, naming the argument that
        resolves it. Documented behaviour; I have not reproduced a custom-zone
        event live, which is why the guard is a refusal and not a translation.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000", start_zone="tzone://Microsoft/Custom"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-10-28T11:00:00",
                end="2026-10-28T12:00:00",
                config=_CFG_LA,
            )

        assert "custom time zone" in str(excinfo.value)
        builder.patch.assert_not_called()

    async def test_the_anchor_never_comes_from_start_time_zone(self):
        """The regression guard for the finding itself.

        `start.time_zone` is "UTC" on every plain GET, so code reading it
        cannot preserve anything. This pins that the two fields disagreeing
        resolves in favour of the anchor — the assertion that fails the moment
        anyone reaches for the obvious field again.
        """
        current = _current_event(
            date_time="2026-10-28T16:00:00.0000000", start_zone="Europe/London"
        )

        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-10-28T11:00:00",
            end="2026-10-28T12:00:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "Europe/London"


class TestShowAs:
    """``show_as`` is Graph's ``freeBusyStatus`` — Outlook's "Show as" field.

    Graph honours all six values on both POST and PATCH for a consumer mailbox;
    that claim belongs to the live tier and is pinned there, not here. These
    assert the half a mock can see: the argument becomes a typed enum member on
    the model, and omitting it leaves the field untouched rather than writing a
    default of our own.
    """

    @pytest.mark.parametrize(
        ("supplied", "expected"),
        [
            ("free", "free"),
            ("tentative", "tentative"),
            ("busy", "busy"),
            ("oof", "oof"),
            ("workingElsewhere", "workingElsewhere"),
            ("unknown", "unknown"),
            # Case and separator normalisation.
            ("Tentative", "tentative"),
            ("WORKINGELSEWHERE", "workingElsewhere"),
            # Aliases for the two values whose Graph names don't match the
            # labels Outlook shows, which is what an agent reads off the UI.
            ("out_of_office", "oof"),
            ("Out of office", "oof"),
            ("outofoffice", "oof"),
            ("working_elsewhere", "workingElsewhere"),
            ("working elsewhere", "workingElsewhere"),
            ("working-elsewhere", "workingElsewhere"),
        ],
    )
    async def test_create_event_maps_show_as_to_the_sdk_enum(self, supplied, expected):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        await create_event(
            mock_client,
            subject="Standup",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:00:00Z",
            show_as=supplied,
            config=_CFG,
        )

        # The enum member, not the string it came from: a raw string serializes
        # fine for the values that happen to match Graph's spelling and silently
        # does not for the ones that don't.
        assert _posted(mock_client).show_as.value == expected

    async def test_create_event_leaves_show_as_unset_when_omitted(self):
        """Graph's own default is `busy`; sending one would take that decision."""
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        await create_event(
            mock_client,
            subject="Standup",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:00:00Z",
            config=_CFG,
        )

        assert _posted(mock_client).show_as is None

    async def test_update_event_patches_show_as_alone(self):
        """Unlike is_all_day, showAs needs nothing resent alongside it."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client, event_id="AAMkAG123=", show_as="workingElsewhere", config=_CFG
        )

        patched = builder.patch.call_args[0][0]
        assert patched.show_as.value == "workingElsewhere"
        assert patched.start is None
        assert patched.end is None

    @pytest.mark.parametrize("bogus", ["maybe", "Out to lunch", "", "   ", "0", "oof!"])
    async def test_invalid_show_as_is_refused_naming_what_is_valid(self, bogus):
        """A bogus value must not reach Graph, and the refusal must be actionable.

        Graph answers a bad showAs with a 400 whose text does not list the
        alternatives. The caller here is a model choosing from a menu it cannot
        see, so the message carries the menu.

        The expected values are read off ``FreeBusyStatus`` rather than typed
        here, so this cannot pass while the message names a subset — which is
        how it previously missed that `unknown` went unlisted in `SKILL.md`.
        If the SDK ever gains a member, this fails until the refusal names it.
        """
        from msgraph.generated.models.free_busy_status import FreeBusyStatus

        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        with pytest.raises(ValueError) as excinfo:
            await create_event(
                mock_client,
                subject="Standup",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                show_as=bogus,
                config=_CFG,
            )

        message = str(excinfo.value)
        for member in FreeBusyStatus:
            assert member.value in message, f"refusal does not name {member.value!r}: {message}"
        mock_client.me.events.post.assert_not_called()

    async def test_invalid_show_as_on_update_never_reaches_graph(self):
        """The same gate on the write sibling — #1's lesson, checked not assumed."""
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="show_as"):
            await update_event(mock_client, event_id="AAMkAG123=", show_as="maybe", config=_CFG)

        builder.patch.assert_not_called()

    async def test_invalid_show_as_is_refused_before_any_graph_call(self):
        """Refuse bad input before the network, not after — `.get()` counts.

        `update_event(recurrence=…)` without `start` reads the event back to
        anchor the recurrence range. Resolving `show_as` after that block meant
        a rejected value still cost a Graph round trip: harmless to the mailbox,
        but it breaks the project's rule that input is validated before Graph is
        called, and it is the kind of ordering that stops being harmless the
        moment a write is added above it.
        """
        builder = _make_event_builder()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        # A usable anchor, so that without the fix the recurrence block runs to
        # completion and this test fails on `get.assert_not_called()` — the
        # thing it is about — rather than on an incidental TypeError from a
        # MagicMock datetime.
        builder.get = AsyncMock(
            return_value=MagicMock(start=MagicMock(date_time="2026-09-07T12:30:00"))
        )
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="show_as"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                recurrence="weekly",
                show_as="maybe",
                config=_CFG,
            )

        builder.get.assert_not_called()
        builder.patch.assert_not_called()
