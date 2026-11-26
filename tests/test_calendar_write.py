"""Tests for calendar write tools."""

from datetime import date, datetime, timedelta, timezone
from unittest.mock import AsyncMock, MagicMock

import httpx
import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools import calendar_write
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


def reads_no_times(events) -> MagicMock:
    """Make `events.by_event_id(id).get()` answer with an event that has no times.

    A converted create reads its event back that way (`_times_in_utc`). With no
    times to read, the check says nothing, which is what the tests using this
    need: they are about other arguments, and
    `TestAConvertedTimeIsCheckedAgainstGraph` covers the check. `by_event_id` is
    synchronous, as the SDK's is. Returns the request builder it hands out.
    """
    reader = MagicMock()
    reader.get = AsyncMock(return_value=MagicMock(start=None, end=None))
    events.by_event_id = MagicMock(return_value=reader)
    return reader


def async_sdk() -> AsyncMock:
    """An `AsyncMock` SDK client whose events read back without times (`reads_no_times`)."""
    client = AsyncMock()
    reads_no_times(client.me.events)
    return client


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
    # Explicit once more: under a policy without `mail_send`, `update_event`
    # reads this to decide whether a reword is mailed on, and a truthy
    # MagicMock would make every mocked event a meeting with guests.
    event.attendees = []
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


def _on_the_default_calendar():
    """`GET /me/events/{id}/calendar` for an event on the default calendar."""
    from msgraph.generated.models.calendar import Calendar
    from msgraph.generated.models.email_address import EmailAddress

    return Calendar(
        id="AAMkDEF=", is_default_calendar=True, owner=EmailAddress(address="me@outlook.com")
    )


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
    # Under a policy without `mail_send`, `update_event` asks which calendar the
    # event is on; this one is on the default calendar, the user's own. Events
    # on someone else's calendar: `test_writes_into_shared_calendars_and_lists`.
    builder.calendar.get = AsyncMock(return_value=_on_the_default_calendar())
    builder.patch = AsyncMock()
    builder.delete = AsyncMock()
    builder.accept.post = AsyncMock()
    builder.decline.post = AsyncMock()
    builder.tentatively_accept.post = AsyncMock()
    return builder


class TestCreateEvent:
    async def test_create_event_calls_post(self):
        """create_event builds Event and calls me.events.post()."""
        mock_client = async_sdk()
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
        mock_client = async_sdk()
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
        mock_client = async_sdk()
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

        await update_event(mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG)

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
            await update_event(mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG)

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
            await update_event(mock_client, event_id="AAMkAG123=", is_all_day=True, config=_CFG)

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

        await update_event(mock_client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG)

        writer = JsonSerializationWriter()
        builder.patch.call_args[0][0].serialize(writer)
        assert '"recurrence": null' in writer.get_serialized_content().decode()

    async def test_setting_event_recurrence_none_is_not_sendable_on_every_supported_kiota(self):
        """Pins why additional_data is used, on whichever kiota is installed (#233).

        Through the adapter's real path (the backing-store proxy, not a bare
        ``JsonSerializationWriter``) the plain assignment ``event.recurrence =
        None`` does one of two things, by kiota-serialization-json version:

        - before 1.15 the null is written beside the object body and the whole
          payload is refused (``ValueError("Invalid Json output")``): the
          request never leaves the process;
        - from 1.15 it is written inside the body under the Python attribute
          name. ``recurrence`` happens to be Graph's name too, so this one
          works; a multi-word field such as ``mobile_phone`` does not.

        The package accepts kiota from 1.13 (``msgraph-core>=1.5.2``), so the
        plain assignment would break ``remove_recurrence`` on some supported
        installs. ``additional_data`` is sent as written on every version. The
        workaround can only go once the floor guarantees 1.15 — not because
        this test fails, which it should not on either side.
        """
        import json
        from importlib.metadata import version

        from msgraph.generated.models.event import Event

        from tests.test_write_payloads_reach_the_wire import adapter_json

        event = Event()
        event.subject = "x"
        event.recurrence = None

        installed = tuple(
            int(n) for n in version("microsoft-kiota-serialization-json").split(".")[:2]
        )
        if installed < (1, 15):
            with pytest.raises(ValueError, match="Invalid Json output"):
                adapter_json(event)
        else:
            assert json.loads(adapter_json(event)) == {
                "@odata.type": "#microsoft.graph.event",
                "subject": "x",
                "recurrence": None,
            }

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

        await update_event(mock_client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG)

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
            mock_client,
            event_id="AAMkAG123=",
            response="accept",
            config=_CFG,
        )
        assert result["status"] == "accepted"
        builder.accept.post.assert_called_once()

    async def test_rsvp_decline_calls_decline_post(self):
        """rsvp with response=decline calls decline.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(
            mock_client,
            event_id="AAMkAG123=",
            response="decline",
            config=_CFG,
        )
        assert result["status"] == "declined"
        builder.decline.post.assert_called_once()

    async def test_rsvp_tentative_calls_tentatively_accept_post(self):
        """rsvp with response=tentative calls tentatively_accept.post()."""
        builder = _make_event_builder()
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await rsvp(
            mock_client,
            event_id="AAMkAG123=",
            response="tentative",
            config=_CFG,
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
        """ "Move this to Eastern" means the whole event, not one end of it.

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
            return_value=_current_event(date_time="2026-10-22T00:00:00.0000000", is_all_day=True)
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
        answers with "re-read the item and check before sending the update again".

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

        config_sent = builder.patch.call_args.kwargs["request_configuration"]
        assert not config_sent.headers.get("If-Match")

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
            {
                "start": "2026-11-04T18:00:00",
                "end": "2026-11-04T18:30:00",
                "timezone": "America/Los_Angeles",
            },
            {"start": "2026-11-05T03:00:00Z", "end": "2026-11-05T03:30:00Z"},
            {"end": "2026-11-05T02:45:00Z"},
            {
                "recurrence": {
                    "pattern": {"type": "weekly", "daysOfWeek": ["thursday", "friday"]},
                    "range": {"type": "numbered", "numberOfOccurrences": 3},
                }
            },
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
                {
                    "recurrence": {
                        "pattern": {"type": "weekly", "daysOfWeek": ["funday"]},
                        "range": {"type": "noEnd"},
                    }
                },
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
                await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **kwargs)

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

    @pytest.mark.parametrize(
        "zone",
        [
            pytest.param('Pacific Standard Time"; odata.maxpagesize=1', id="quote"),
            pytest.param("Pacific Standard Time\r\nX-Extra: 1", id="crlf"),
            pytest.param("Pacific Standard Time\n", id="lf"),
            pytest.param("Pacific\\Standard Time", id="backslash"),
            pytest.param("Pacific\x00Standard Time", id="nul"),
            pytest.param("Hora estándar del Pacífico", id="non-ascii"),
        ],
    )
    async def test_a_zone_a_header_cannot_carry_never_reaches_one(self, zone):
        """The second read names the event's stored zone inside `outlook.timezone="…"`.

        That name is Graph's, not the caller's, but it is written into a header
        as is, so one that would end the quoted value or the header line, or
        that a header cannot encode, is refused before the read is made.
        """
        current = _current_event(date_time="2026-11-05T02:00:00.0000000", start_zone=zone)
        prefers = []

        def _get(request_configuration=None):
            prefers.append(
                None
                if request_configuration is None
                else request_configuration.headers.get("Prefer")
            )
            return current

        builder = _make_event_builder()
        builder.get = AsyncMock(side_effect=_get)
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="Nothing was modified"):
            await update_event(
                mock_client, event_id="AAMkAG123=", recurrence="weekly", config=_CFG_LA
            )
        assert prefers == [None], f"a Prefer header was built from {zone!r}: {prefers}"
        builder.patch.assert_not_called()

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


def _one_day_series_ending_thursday():
    """A UTC series at Thursday 23:00Z whose ``endDate`` range ends that same Thursday.

    In Tokyo that start is Friday 08:00, so re-anchoring it there moves the
    first occurrence's local date past the end with nobody typing a date.
    """
    return build_event_recurrence(
        {
            "pattern": {"type": "daily", "interval": 1},
            "range": {"type": "endDate", "endDate": "2026-11-05"},
        },
        start="2026-11-05T23:00:00Z",
        zone="UTC",
    )


def _thursday_evening_master():
    stored = _one_day_series_ending_thursday()
    current = _current_event(
        date_time="2026-11-05T23:00:00.0000000",
        start_zone="UTC",
        event_type="seriesMaster",
        recurrence=stored,
    )
    return stored, *_series_master_builder(stored, current=current)


class TestARangeThatEndsBeforeItBegins:
    """#86 on every route into ``_reconcile_range``: refused, and nothing sent.

    ``endDate`` is never moved to make room — extending a series is not what
    the caller asked for — and the refusal is ours rather than Graph's
    ``400 StartDateV2 should be earlier or equal to EndDateV2``.
    """

    async def test_a_zone_re_anchor_that_crosses_midnight_is_refused(self):
        """The re-send: the caller typed neither a new date nor a range.

        Same instant, new zone — Thursday 23:00Z is Friday 08:00 in Tokyo, a
        day past the series' end. A zone test that never crosses midnight
        cannot see this.
        """
        _, builder, mock_client = _thursday_evening_master()

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-06T08:00:00",
                end="2026-11-06T08:30:00",
                timezone="Asia/Tokyo",
                config=_CFG,
            )

        assert "range.endDate (2026-11-05)" in str(excinfo.value)
        assert "first occurrence (2026-11-06)" in str(excinfo.value)
        builder.patch.assert_not_called()

    async def test_a_zone_re_anchor_that_keeps_the_date_is_sent(self):
        """The control: London in November is UTC+0, so the date holds and the range is valid.

        Without it, the refusal above could be the fixture's rather than the check's.
        """
        from msgraph.generated.models.recurrence_range_type import RecurrenceRangeType

        _, builder, mock_client = _thursday_evening_master()

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-05T23:00:00",
            end="2026-11-05T23:30:00",
            timezone="Europe/London",
            config=_CFG,
        )

        sent = builder.patch.call_args[0][0].recurrence.range
        assert sent.type is RecurrenceRangeType.EndDate
        assert sent.start_date == date(2026, 11, 5)
        assert sent.end_date == date(2026, 11, 5)

    async def test_an_echoed_recurrence_with_a_start_past_its_end_is_refused(self):
        """Update with the caller's own recurrence: ``outlook_get_event``'s, handed back."""
        stored, builder, mock_client = _thursday_evening_master()

        with pytest.raises(ValueError, match=r"range\.endDate \(2026-11-05\)"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-12T23:00:00",
                end="2026-11-12T23:30:00",
                recurrence=serialize_recurrence(stored),
                config=_CFG,
            )

        builder.patch.assert_not_called()

    async def test_create_with_a_range_ending_before_the_start_is_refused(self):
        """Create: no network at all."""
        mock_client = AsyncMock()

        with pytest.raises(ValueError, match=r"range\.endDate \(2027-03-26\)"):
            await create_event(
                mock_client,
                subject="Ends before it begins",
                start="2027-04-05T09:00:00Z",
                end="2027-04-05T09:30:00Z",
                recurrence={
                    "pattern": {"type": "daily", "interval": 1},
                    "range": {"type": "endDate", "endDate": "2027-03-26"},
                },
                config=_CFG,
            )

        mock_client.me.events.post.assert_not_called()


def _monday_series(zone: str = "America/New_York", *, all_day: bool = False, pattern=None):
    """#98's series: weekly from Monday 2026-11-02, three, 09:00 in ``zone``.

    All-day, it is midnight UTC, as Graph stores one. ``recurrenceTimeZone`` is
    set because Graph's read-back always carries it.
    """
    clock = "00:00:00" if all_day else "09:00:00"
    stored_zone = "UTC" if all_day else zone
    recurrence = build_event_recurrence(
        {
            "pattern": pattern
            or {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": ["monday"],
                "firstDayOfWeek": "sunday",
            },
            "range": {
                "type": "numbered",
                "numberOfOccurrences": 3,
                "recurrenceTimeZone": stored_zone,
            },
        },
        start=f"2026-11-02T{clock}",
        zone=stored_zone,
    )
    current = _current_event(
        date_time="2026-11-02T00:00:00.0000000" if all_day else "2026-11-02T14:00:00.0000000",
        start_zone=stored_zone,
        event_type="seriesMaster",
        recurrence=recurrence,
        is_all_day=all_day,
    )
    return _series_master_builder(recurrence, current=current)


class TestMovingASeriesToAnotherDay:
    """#98: a series master's start moved to another day, in its own zone, no `recurrence`.

    Graph answered `updated` and snapped the master back onto its pattern:
    every occurrence stayed on the old day. Verified live. A start sent with the
    series' own recurrence, its days and `startDate` moved with it, lands on the
    new day (``TestMovingASeriesKeepsItsDays`` in the live tier), so that is what
    goes out, exactly as if the caller had echoed ``outlook_get_event``'s
    recurrence beside the new start.
    """

    @pytest.mark.parametrize("zone", ["America/New_York", "Eastern Standard Time"])
    async def test_a_weekly_series_moved_a_day_takes_its_days_along(self, zone):
        """The issue's case. A Windows-named zone too: a zone-less start needs no conversion."""
        from msgraph.generated.models.day_of_week import DayOfWeek

        builder, mock_client = _monday_series(zone)

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T09:00:00",
            end="2026-11-03T09:30:00",
            config=_CFG_LA,
        )

        assert result["status"] == "updated"
        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == zone
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Tuesday]
        assert patched.recurrence.range.start_date == date(2026, 11, 3)
        assert patched.recurrence.range.number_of_occurrences == 3
        # The event's own zone is not changing, so Graph re-derives the same one.
        assert patched.recurrence.range.recurrence_time_zone is None
        # It echoes state it read, so it is pinned like every other master reshape.
        config_sent = builder.patch.call_args.kwargs["request_configuration"]
        assert config_sent.headers.get("If-Match") == {CURRENT_ETAG}

    async def test_an_all_day_series_moved_a_day_takes_its_days_along(self):
        """The issue's second case: Graph anchors it in UTC, so this compared UTC with UTC."""
        from msgraph.generated.models.day_of_week import DayOfWeek

        builder, mock_client = _monday_series(all_day=True)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T00:00:00",
            end="2026-11-04T00:00:00",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.time_zone == "UTC"
        assert patched.recurrence.pattern.days_of_week == [DayOfWeek.Tuesday]
        assert patched.recurrence.range.start_date == date(2026, 11, 3)

    @pytest.mark.parametrize(
        "times",
        [
            {"start": "2026-11-02T10:00:00", "end": "2026-11-02T10:30:00"},
            {"end": "2026-11-02T10:00:00"},
        ],
        ids=["same-day-time", "end-only"],
    )
    async def test_a_change_that_keeps_the_day_sends_no_recurrence(self, times):
        """The control: a time patch on the same day is not a series rewrite."""
        builder, mock_client = _monday_series()

        await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **times)

        assert builder.patch.call_args[0][0].recurrence is None

    @pytest.mark.parametrize(
        ("start", "end", "moved_to"),
        [
            # 03:00Z on the 3rd is 22:00 on Monday the 2nd in New York.
            ("2026-11-03T03:00:00Z", "2026-11-03T03:30:00Z", None),
            # 14:00Z on the 3rd is 09:00 on Tuesday the 3rd there.
            ("2026-11-03T14:00:00Z", "2026-11-03T14:30:00Z", date(2026, 11, 3)),
        ],
        ids=["utc-date-moves-local-does-not", "local-date-moves"],
    )
    async def test_the_events_own_day_decides_not_the_utc_one(self, start, end, moved_to):
        """A typed instant is judged on the day it falls on in the event's zone.

        Taken from the UTC text, the first would shift every occurrence a day
        late, reported as `updated`, and the second would look unmoved.
        """
        builder, mock_client = _monday_series()

        await update_event(mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG_LA)

        sent = builder.patch.call_args[0][0].recurrence
        assert (sent and sent.range.start_date) == moved_to

    async def test_an_instant_in_a_zone_this_server_cannot_read_is_refused(self):
        """A `Z` start on a series in a zone no table here knows: its local day is unknowable.

        Sent as written, a new day would leave every occurrence behind, and
        guessing the day from the UTC text can move them to the wrong one. The
        remedy is a zone-less time, which is read in the event's own zone.
        """
        builder, mock_client = _monday_series("Middle Earth Standard Time")

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T14:00:00Z",
                end="2026-11-03T14:30:00Z",
                config=_CFG_LA,
            )

        message = str(excinfo.value)
        assert "Middle Earth Standard Time" in message
        assert "without a zone or offset" in message
        assert "Nothing was modified" in message
        builder.patch.assert_not_called()

    async def test_a_master_without_a_readable_start_date_is_refused(self):
        """Fail closed: with no stored `startDate` there is no telling whether the day moves."""
        builder, mock_client = _monday_series()
        builder.get.return_value.recurrence = None

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                config=_CFG_LA,
            )

        assert "Pass `recurrence`" in str(excinfo.value)
        assert "Nothing was modified" in str(excinfo.value)
        builder.patch.assert_not_called()

    async def test_a_pattern_with_no_exact_move_is_refused_not_left_behind(self):
        """A "first Monday" pattern moved a day has no exact equivalent: refused, not `updated`."""
        builder, mock_client = _monday_series(
            pattern={"type": "relativeMonthly", "interval": 1, "daysOfWeek": ["monday"]}
        )

        with pytest.raises(ValueError, match="relative pattern cannot be moved"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                config=_CFG_LA,
            )

        builder.patch.assert_not_called()


def _weekdays_begun_on_a_saturday(zone: str):
    """A Monday-to-Friday series whose range starts Saturday 2026-10-31.

    The "weekdays" shorthand on a weekend start makes one. Its first occurrence
    is assumed to be Monday 2026-11-02 at 09:00 in ``zone`` — what #98 suggests
    Graph does with a master that is off its pattern — so `startDate` is not
    the day the series is on. The plain GET reports it projected into UTC.
    """
    recurrence = build_event_recurrence(
        {
            "pattern": {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                "firstDayOfWeek": "sunday",
            },
            "range": {"type": "numbered", "numberOfOccurrences": 10},
        },
        start="2026-10-31T09:00:00",
        zone=zone,
    )
    return _current_event(
        date_time="2026-11-02T14:00:00.0000000",
        start_zone=zone,
        event_type="seriesMaster",
        recurrence=recurrence,
    )


class TestASeriesWhosePatternSkipsItsFirstDay:
    """A same-day time change must not read as a move from the range's first day.

    Measured from `startDate` (Saturday), a Monday 09:00 to 10:00 change is a
    two-day move, and the re-send would shift every weekday by two, reported as
    `updated`. So the day the series is on comes off the master's own start.
    """

    async def test_a_same_day_time_change_sends_no_recurrence_and_reads_nothing_more(self):
        builder, mock_client = _series_master_builder(
            None, current=_weekdays_begun_on_a_saturday("America/New_York")
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-02T10:00:00",
            end="2026-11-02T10:30:00",
            config=_CFG_LA,
        )

        assert builder.patch.call_args[0][0].recurrence is None
        # The event read and the occurrence read; an IANA zone converts locally.
        assert builder.get.await_count == 2

    async def test_a_move_to_another_day_is_refused(self):
        """`_follow_new_start` measures from `startDate`, so the days would land two off."""
        builder, mock_client = _series_master_builder(
            None, current=_weekdays_begun_on_a_saturday("America/New_York")
        )

        with pytest.raises(ValueError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                config=_CFG_LA,
            )

        message = str(excinfo.value)
        assert "2026-10-31" in message and "2026-11-02" in message
        assert "Nothing was modified" in message
        builder.patch.assert_not_called()

    async def test_a_windows_named_zone_asks_graph_for_the_local_day(self):
        """The UTC start cannot be converted here, so Graph projects it, as for #77."""
        utc_view = _weekdays_begun_on_a_saturday("Eastern Standard Time")
        local_view = _weekdays_begun_on_a_saturday("Eastern Standard Time")
        local_view.start = MagicMock(
            date_time="2026-11-02T09:00:00.0000000", time_zone="Eastern Standard Time"
        )
        prefers = []

        def _get(request_configuration=None):
            preference = None
            if request_configuration is not None:
                preference = request_configuration.headers.get("Prefer")
            prefers.append(preference or None)
            return local_view if preference else utc_view

        builder, mock_client = _series_master_builder(None, current=utc_view)
        builder.get = AsyncMock(side_effect=_get)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-02T10:00:00",
            end="2026-11-02T10:30:00",
            config=_CFG_LA,
        )

        assert prefers[-1] == {'outlook.timezone="Eastern Standard Time"'}
        assert builder.patch.call_args[0][0].recurrence is None

    @pytest.mark.parametrize(
        "given",
        [
            "weekdays",
            {
                "pattern": {
                    "type": "weekly",
                    "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                },
                "range": {"type": "numbered", "numberOfOccurrences": 10},
            },
        ],
        ids=["shorthand", "the-same-days-written-out"],
    )
    async def test_the_refusals_remedy_sends_the_days_given(self, given):
        """What the refusal says to do, done: `recurrence` with the days wanted, beside `start`.

        The days equal the stored ones, so the echo path took them for the
        series' own and moved them three days from the Saturday: Thu–Mon,
        `updated`. They are the caller's now, sent as written.
        """
        from msgraph.generated.models.day_of_week import DayOfWeek

        builder, mock_client = _series_master_builder(
            None, current=_weekdays_begun_on_a_saturday("America/New_York")
        )
        with pytest.raises(ValueError, match="Pass `recurrence` with the days you want"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                config=_CFG_LA,
            )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T09:00:00",
            end="2026-11-03T09:30:00",
            recurrence=given,
            config=_CFG_LA,
        )

        sent = builder.patch.call_args[0][0].recurrence
        assert sent.pattern.days_of_week == [
            DayOfWeek.Monday,
            DayOfWeek.Tuesday,
            DayOfWeek.Wednesday,
            DayOfWeek.Thursday,
            DayOfWeek.Friday,
        ]
        assert sent.range.start_date == date(2026, 11, 3)

    async def test_a_zone_change_refuses_rather_than_measure_from_the_saturday(self):
        """The zone change's own re-send measured from `startDate` too; it is refused alike."""
        builder, mock_client = _series_master_builder(
            None, current=_weekdays_begun_on_a_saturday("America/New_York")
        )

        with pytest.raises(ValueError, match="Pass `recurrence` with the days you want"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-02T08:00:00",
                end="2026-11-02T08:30:00",
                timezone="America/Chicago",
                config=_CFG_LA,
            )
        builder.patch.assert_not_called()


def _a_series(pattern: dict, first: date, *, rng: dict | None = None, zone="America/New_York"):
    """A clean series master whose range starts ``first`` at 09:00 in ``zone``.

    The plain GET reports its start projected into UTC, as Graph does.
    """
    from zoneinfo import ZoneInfo

    local = datetime(first.year, first.month, first.day, 9, tzinfo=ZoneInfo(zone))
    recurrence = build_event_recurrence(
        {"pattern": pattern, "range": rng or {"type": "numbered", "numberOfOccurrences": 3}},
        start=f"{first.isoformat()}T09:00:00",
        zone=zone,
    )
    current = _current_event(
        date_time=f"{local.astimezone(timezone.utc):%Y-%m-%dT%H:%M:%S}.0000000",
        start_zone=zone,
        event_type="seriesMaster",
        recurrence=recurrence,
    )
    return _series_master_builder(recurrence, current=current)


_WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday"]


class TestWhatAMovedSeriesRepeatsOn:
    """#98 review: a series moved to another day keeps its days where it can.

    A day the pattern already repeats on keeps the pattern and moves only the
    range: weekdays started on a Wednesday stay weekdays. One day moves to the
    new one. Several weekdays moved to a day they skip are refused, since
    nothing says which days were meant; shifting them all put weekdays onto
    the weekend, `updated`. The reply's `note` says what was sent.
    """

    @pytest.mark.parametrize(
        ("days", "start"),
        [
            (_WEEKDAYS, "2026-11-04T09:00:00"),
            (_WEEKDAYS, "2026-11-18T10:00:00"),
            (["monday", "thursday"], "2026-11-05T09:00:00"),
        ],
        ids=["weekdays-to-wednesday", "weekdays-later-and-new-time", "mon-thu-to-thursday"],
    )
    async def test_a_day_it_already_repeats_on_keeps_its_days(self, days, start):
        builder, mock_client = _a_series({"type": "weekly", "daysOfWeek": days}, date(2026, 11, 2))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start=start,
            end=start[:11] + "11:00:00",
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == days
        assert sent["range"]["startDate"] == start[:10]
        named = ", ".join(d.capitalize() for d in days)
        assert f"every week on {named}, from {start[:10]}, 3 times" in result["note"]

    @pytest.mark.parametrize(
        ("days", "start"),
        [(["monday", "thursday"], "2026-11-03T09:00:00"), (_WEEKDAYS, "2026-11-07T09:00:00")],
        ids=["mon-thu-to-tuesday", "weekdays-to-saturday"],
    )
    async def test_several_days_moved_to_one_they_skip_are_refused(self, days, start):
        from outlook_mcp.errors import OutlookMCPError

        builder, mock_client = _a_series({"type": "weekly", "daysOfWeek": days}, date(2026, 11, 2))

        with pytest.raises(OutlookMCPError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start=start,
                end=start[:11] + "09:30:00",
                config=_CFG_LA,
            )

        assert excinfo.value.code == "series_days_ambiguous"
        assert "Pass `recurrence`" in str(excinfo.value)
        builder.patch.assert_not_called()

    @pytest.mark.parametrize(
        ("pattern", "first", "start", "expected"),
        [
            ({"type": "daily", "interval": 1}, "2026-11-02", "2026-11-05", {}),
            # N1: the same day of the next month was refused as "a different month".
            ({"type": "absoluteMonthly", "dayOfMonth": 15}, "2026-11-15", "2026-12-15", {}),
            (
                {"type": "absoluteMonthly", "dayOfMonth": 15},
                "2026-11-15",
                "2026-11-20",
                {"dayOfMonth": 20},
            ),
            (
                {"type": "relativeMonthly", "daysOfWeek": ["monday"], "index": "first"},
                "2026-11-02",
                "2026-12-07",
                {},
            ),
        ],
        ids=["daily", "monthly-next-month", "monthly-another-day", "first-monday-next-month"],
    )
    async def test_other_patterns_keep_or_move_their_day(self, pattern, first, start, expected):
        builder, mock_client = _a_series(pattern, date.fromisoformat(first))

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start=f"{start}T09:00:00",
            end=f"{start}T09:30:00",
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        stored = serialize_recurrence(builder.get.return_value.recurrence)
        assert sent["pattern"] == {**stored["pattern"], **expected}
        assert sent["range"]["startDate"] == start

    async def test_an_end_date_moves_as_far_as_the_start(self):
        """Three Mondays to 2026-11-16, moved to Tuesday: three Tuesdays to the 17th."""
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday"]},
            date(2026, 11, 2),
            rng={"type": "endDate", "endDate": "2026-11-16"},
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T09:00:00",
            end="2026-11-03T09:30:00",
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == ["tuesday"]
        assert (sent["range"]["startDate"], sent["range"]["endDate"]) == (
            "2026-11-03",
            "2026-11-17",
        )
        assert "every week on Tuesday, from 2026-11-03 until 2026-11-17" in result["note"]

    async def test_a_change_within_the_day_carries_no_note(self):
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": _WEEKDAYS}, date(2026, 11, 2)
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-02T10:00:00",
            end="2026-11-02T10:30:00",
            config=_CFG_LA,
        )

        assert builder.patch.call_args[0][0].recurrence is None
        assert "note" not in result

    async def test_a_zone_change_that_moves_the_days_says_so_and_one_that_does_not_is_quiet(self):
        """The zone change's own re-send reports only a rule that schedules other days."""
        builder, mock_client = _series_master_builder(_thursday_utc_series())
        moved = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-04T18:00:00",
            end="2026-11-04T18:30:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )
        builder, mock_client = _series_master_builder(_thursday_utc_series())
        kept = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-05T02:00:00",
            end="2026-11-05T02:30:00",
            timezone="Europe/London",
            config=_CFG_LA,
        )

        assert "every week on Wednesday, from 2026-11-04, 3 times" in moved["note"]
        assert "note" not in kept

    async def test_a_moved_master_gets_no_times_only_correction(self):
        """#262's read-back sends the times alone when Graph places one elsewhere.

        On a master that is #98's shape again. A patch that carried the
        recurrence is not checked, so Graph's reply here, an hour off, changes
        nothing: one PATCH, the recurrence on it, no `Prefer` for the check.
        """
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday"]}, date(2026, 11, 2)
        )
        builder.patch = AsyncMock(
            return_value=MagicMock(
                id="AAMkAG123=",
                start=MagicMock(date_time="2026-11-03T15:00:00.0000000", time_zone="UTC"),
                end=MagicMock(date_time="2026-11-03T15:30:00.0000000", time_zone="UTC"),
            )
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T14:00:00Z",
            end="2026-11-03T14:30:00Z",
            config=_CFG_LA,
        )

        assert builder.patch.await_count == 1
        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-11-03T09:00:00"
        assert patched.recurrence is not None
        assert not builder.patch.call_args.kwargs["request_configuration"].headers.get("Prefer")

    @pytest.mark.parametrize(
        ("first", "moves"),
        [("2027-03-28", True), ("2027-04-04", False)],
        ids=["to-another-day", "within-the-day"],
    )
    async def test_a_z_start_left_as_typed_cannot_move_a_series(self, first, moves):
        """N2: 15:30Z on 2027-04-03 is 02:30 on Sunday the 4th in Sydney, in the repeated hour.

        `_typed_in_zones` leaves it as typed, and Graph files a typed `Z` in UTC,
        where that is Saturday: a range moved to the Sunday would disagree with
        its start. So a move is refused; a change within the day goes out as
        before, with no recurrence.
        """
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["sunday"]},
            date.fromisoformat(first),
            zone="Australia/Sydney",
        )
        times = {"start": "2027-04-03T15:30:00Z", "end": "2027-04-03T16:30:00Z"}

        if moves:
            with pytest.raises(ValueError, match="could not be converted into local time"):
                await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **times)
            builder.patch.assert_not_called()
        else:
            await update_event(mock_client, event_id="AAMkAG123=", config=_CFG_LA, **times)
            patched = builder.patch.call_args[0][0]
            assert patched.start.date_time == "2027-04-03T15:30:00Z"
            assert patched.recurrence is None


def _utc_series(days: list[str], first: str):
    """A UTC-anchored weekly series at 02:00Z from ``first``, as pre-zone series are stored."""
    recurrence = build_event_recurrence(
        {
            "pattern": {
                "type": "weekly",
                "interval": 1,
                "daysOfWeek": days,
                "firstDayOfWeek": "sunday",
            },
            "range": {"type": "numbered", "numberOfOccurrences": 6},
        },
        start=f"{first}T02:00:00Z",
        zone="UTC",
    )
    current = _current_event(
        date_time=f"{first}T02:00:00.0000000",
        start_zone="UTC",
        event_type="seriesMaster",
        recurrence=recurrence,
    )
    return _series_master_builder(recurrence, current=current)


class TestARecurrencePassedWithAMove:
    """#98 re-review S1: `recurrence` passed back beside a start on another day.

    Several days are sent exactly as written, the range starting on the new
    start: they were shifted with it when they equalled the stored pattern.
    One day keeps the echo rule agents rely on to move a weekly meeting: it
    follows the new start's weekday.
    """

    @pytest.mark.parametrize(
        ("days", "given", "start"),
        [
            (_WEEKDAYS, "weekdays", "2026-11-04"),
            (["monday", "thursday"], None, "2026-11-03"),
            (["monday", "thursday"], None, "2026-11-05"),
        ],
        ids=["weekdays-to-wednesday", "mon-thu-to-tuesday", "mon-thu-to-thursday"],
    )
    async def test_several_days_are_sent_as_written(self, days, given, start):
        builder, mock_client = _a_series({"type": "weekly", "daysOfWeek": days}, date(2026, 11, 2))
        echoed = given or serialize_recurrence(builder.get.return_value.recurrence)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start=f"{start}T09:00:00",
            end=f"{start}T09:30:00",
            recurrence=echoed,
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == days
        assert sent["range"]["startDate"] == start

    async def test_one_day_follows_the_new_start(self):
        """Main's echo rule, kept: a Monday meeting handed back beside Tuesday becomes Tuesday's."""
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday"]}, date(2026, 11, 2)
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T09:00:00",
            end="2026-11-03T09:30:00",
            recurrence=serialize_recurrence(builder.get.return_value.recurrence),
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == ["tuesday"]
        assert sent["range"]["startDate"] == "2026-11-03"

    async def test_the_refusal_says_the_days_given_are_sent_exactly(self):
        from outlook_mcp.errors import OutlookMCPError

        _, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday", "thursday"]}, date(2026, 11, 2)
        )

        with pytest.raises(OutlookMCPError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                config=_CFG_LA,
            )
        assert "exactly those days are sent" in str(excinfo.value)


class TestAZoneChangeThatAlsoMovesTheDay:
    """#98 re-review S2: `timezone` beside a start the caller moved to another day.

    The zone's own share still shifts the days (the same instants read in the
    new zone), and the caller's move on top follows the same-zone rules: days
    kept where the new day is in the pattern, one day moved, several days on a
    day they skip refused, an end date moved as far.
    """

    async def test_weekdays_started_on_a_wednesday_stay_weekdays(self):
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": _WEEKDAYS}, date(2026, 11, 2)
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-04T09:00:00",
            end="2026-11-04T09:30:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        patched = builder.patch.call_args[0][0]
        sent = serialize_recurrence(patched.recurrence)
        assert patched.start.time_zone == "America/Los_Angeles"
        assert sent["pattern"]["daysOfWeek"] == _WEEKDAYS
        assert sent["range"]["startDate"] == "2026-11-04"

    async def test_several_days_moved_to_one_they_skip_are_refused(self):
        from outlook_mcp.errors import OutlookMCPError

        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday", "thursday"]}, date(2026, 11, 2)
        )

        with pytest.raises(OutlookMCPError) as excinfo:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                timezone="America/Los_Angeles",
                config=_CFG_LA,
            )
        assert excinfo.value.code == "series_days_ambiguous"
        builder.patch.assert_not_called()

    async def test_an_end_date_moves_as_far_as_the_start(self):
        builder, mock_client = _a_series(
            {"type": "weekly", "daysOfWeek": ["monday"]},
            date(2026, 11, 2),
            rng={"type": "endDate", "endDate": "2026-11-16"},
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-03T09:00:00",
            end="2026-11-03T09:30:00",
            timezone="America/Los_Angeles",
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == ["tuesday"]
        assert (sent["range"]["startDate"], sent["range"]["endDate"]) == (
            "2026-11-03",
            "2026-11-17",
        )

    @pytest.mark.parametrize("echo", [False, True], ids=["re-sent", "echoed"])
    async def test_a_pure_re_anchor_still_shifts_every_day(self, echo):
        """The control: Tuesday to Saturday at 02:00Z is Monday to Friday at 18:00 in Los Angeles.

        Same instants, so the zone alone moved the local days, and all of them
        move: neither refused as a multi-day move nor kept on the UTC days.
        """
        builder, mock_client = _utc_series(
            ["tuesday", "wednesday", "thursday", "friday", "saturday"], "2026-11-03"
        )
        given = serialize_recurrence(builder.get.return_value.recurrence) if echo else None

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-02T18:00:00",
            end="2026-11-02T18:30:00",
            timezone="America/Los_Angeles",
            recurrence=given,
            config=_CFG_LA,
        )

        sent = serialize_recurrence(builder.patch.call_args[0][0].recurrence)
        assert sent["pattern"]["daysOfWeek"] == _WEEKDAYS
        assert sent["range"]["startDate"] == "2026-11-02"


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
        mock_client = async_sdk()
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
        mock_client = async_sdk()
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


# ── Calendar writes that send email need `mail_send` ─────────────────────────
# An invitation carries the event's subject and body to every attendee, and an
# RSVP comment goes to the organizer. Under a policy that allows
# `calendar_write` and withholds `mail_send` — the README's "calendar-only"
# example — those were a way to send mail anyway.

_CFG_CAL_ONLY = Config(client_id="test", allow_categories=["calendar_write"])
_CFG_CAL_AND_SEND = Config(client_id="test", allow_categories=["calendar_write", "mail_send"])


def _meeting(*, is_organizer: bool = True, attendees: int = 1):
    """An existing event as a plain GET returns it, with other people on it."""
    event = _current_event()
    event.attendees = [MagicMock() for _ in range(attendees)]
    event.is_organizer = is_organizer
    return event


def _client_with(event):
    builder = _make_event_builder()
    builder.get = AsyncMock(return_value=event)
    builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
    client = MagicMock()
    client.me.events.by_event_id = MagicMock(return_value=builder)
    return builder, client


class TestCalendarWritesThatSendEmail:
    async def test_create_event_with_attendees_needs_mail_send(self):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Sync"))

        with pytest.raises(PermissionDeniedError) as exc:
            await create_event(
                mock_client,
                subject="Sync",
                start="2026-04-15T14:00:00Z",
                end="2026-04-15T15:00:00Z",
                body="mailbox contents",
                attendees=["stranger@example.com"],
                config=_CFG_CAL_ONLY,
            )

        assert "mail_send" in str(exc.value)
        mock_client.me.events.post.assert_not_called()

    async def test_create_event_without_attendees_stays_calendar_only(self):
        mock_client = async_sdk()
        mock_client.me.events.post = AsyncMock(return_value=_created("Focus time"))

        result = await create_event(
            mock_client,
            subject="Focus time",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            config=_CFG_CAL_ONLY,
        )

        assert result["status"] == "created"

    async def test_create_event_with_attendees_passes_once_mail_send_is_allowed(self):
        mock_client = async_sdk()
        mock_client.me.events.post = AsyncMock(return_value=_created("Sync"))

        result = await create_event(
            mock_client,
            subject="Sync",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            attendees=["alice@example.com"],
            config=_CFG_CAL_AND_SEND,
        )

        assert result["status"] == "created"
        assert len(_posted(mock_client).attendees) == 1

    async def test_update_event_adding_attendees_needs_mail_send(self):
        """Refused on the argument alone, before Graph is asked anything."""
        builder, client = _client_with(_current_event())

        with pytest.raises(PermissionDeniedError) as exc:
            await update_event(
                client,
                event_id="AAMkAG123=",
                attendees=["stranger@example.com"],
                config=_CFG_CAL_ONLY,
            )

        assert "mail_send" in str(exc.value)
        builder.get.assert_not_called()
        builder.patch.assert_not_called()

    @pytest.mark.parametrize(
        "change",
        [
            {"subject": "mailbox contents"},
            {"body": "mailbox contents"},
            {"location": "mailbox contents"},
            # Adding an online meeting writes join details into it. That the
            # update reaches attendees is inferred (#70): the live tier cannot
            # have attendees, so the worse case is taken.
            {"is_online": True},
        ],
        ids=lambda change: next(iter(change)),
    )
    async def test_rewording_a_meeting_with_attendees_needs_mail_send(self, change):
        """Graph mails the update to the people already on the meeting."""
        builder, client = _client_with(_meeting())

        with pytest.raises(PermissionDeniedError) as exc:
            await update_event(client, event_id="AAMkAG123=", config=_CFG_CAL_ONLY, **change)

        assert "mail_send" in str(exc.value)
        builder.patch.assert_not_called()

    async def test_an_online_meeting_on_an_event_with_nobody_on_it_stays_calendar_only(self):
        builder, client = _client_with(_meeting(attendees=0))

        result = await update_event(
            client, event_id="AAMkAG123=", is_online=True, config=_CFG_CAL_ONLY
        )

        assert result["status"] == "updated"
        builder.patch.assert_called_once()

    async def test_rewording_an_event_with_nobody_on_it_stays_calendar_only(self):
        builder, client = _client_with(_meeting(attendees=0))

        result = await update_event(
            client, event_id="AAMkAG123=", subject="Renamed", config=_CFG_CAL_ONLY
        )

        assert result["status"] == "updated"
        builder.patch.assert_called_once()

    @pytest.mark.parametrize("is_organizer", [True, False, None])
    async def test_whose_meeting_it_is_makes_no_difference(self, is_organizer):
        """An attendee's own copy is not exempt.

        The edit itself stays in their calendar, but the response Exchange
        sends when they accept or decline is built from that copy — so a
        reworded subject followed by a bare RSVP reaches the organizer, who is
        whoever sent the invite. Any event with other people on it is refused.
        """
        builder, client = _client_with(_meeting(is_organizer=is_organizer))

        with pytest.raises(PermissionDeniedError):
            await update_event(
                client, event_id="AAMkAG123=", subject="mailbox contents", config=_CFG_CAL_ONLY
            )

        builder.patch.assert_not_called()

    async def test_clearing_attendees_stays_calendar_only(self):
        """A cancellation carries no caller-written text."""
        builder, client = _client_with(_current_event())

        result = await update_event(
            client, event_id="AAMkAG123=", attendees=[], config=_CFG_CAL_ONLY
        )

        assert result["status"] == "updated"

    async def test_an_unrestricted_server_pays_no_extra_read_for_a_reword(self):
        builder, client = _client_with(_meeting())

        await update_event(client, event_id="AAMkAG123=", subject="Renamed", config=_CFG)

        builder.get.assert_not_called()
        builder.patch.assert_called_once()

    async def test_rsvp_with_a_message_needs_mail_send(self):
        builder, client = _client_with(_current_event())

        with pytest.raises(PermissionDeniedError) as exc:
            await rsvp(
                client,
                event_id="AAMkAG123=",
                response="accept",
                message="mailbox contents",
                config=_CFG_CAL_ONLY,
            )

        assert "mail_send" in str(exc.value)
        builder.accept.post.assert_not_called()

    async def test_rsvp_without_a_message_stays_calendar_only(self):
        builder, client = _client_with(_current_event())

        result = await rsvp(client, event_id="AAMkAG123=", response="accept", config=_CFG_CAL_ONLY)

        assert result["status"] == "accepted"
        builder.accept.post.assert_called_once()


# ── create_event: which calendar ──────────────────────────────────────


def _calendar(name: str, cal_id: str, can_edit: bool | None = True):
    """A calendar as `/me/calendars` returns it; `can_edit=None` is a calendar
    whose listing carries no `canEdit` at all."""
    cal = MagicMock()
    cal.name = name
    cal.id = cal_id
    cal.can_edit = can_edit
    return cal


_CALENDARS = [
    _calendar("Calendar", "DEFAULT1="),
    _calendar("Kids + School", "KIDS77="),
    _calendar("Birthdays", "BDAY42=", can_edit=False),
]


def _client_with_calendars(calendars=None):
    """A client whose `/me/calendars` lists ``calendars`` and whose default and
    per-calendar event collections both accept a POST."""
    client = MagicMock()
    client.me.calendars.get = AsyncMock(
        return_value=MagicMock(value=list(_CALENDARS if calendars is None else calendars))
    )
    client.me.calendars.get.return_value.odata_next_link = None
    client.me.events.post = AsyncMock(return_value=_created("Dentist", "AAMkAGdef="))
    secondary = MagicMock()
    secondary.events.post = AsyncMock(return_value=_created("Dentist", "AAMkAGsec="))
    client.me.calendars.by_calendar_id = MagicMock(return_value=secondary)
    return client, secondary


async def _create(client, **kwargs):
    return await create_event(
        client,
        subject="Dentist",
        start="2026-11-03T09:00:00",
        end="2026-11-03T09:30:00",
        config=_CFG,
        **kwargs,
    )


class TestCreateEventInAChosenCalendar:
    """`calendar` puts the event where `outlook_list_events` would read it back.

    The default calendar stays `POST /me/events` with no extra round trip, which
    is every call made before this argument existed.
    """

    async def test_omitted_calendar_posts_to_the_default_with_no_listing(self):
        client, secondary = _client_with_calendars()

        result = await _create(client)

        assert result["event_id"] == "AAMkAGdef="
        client.me.events.post.assert_called_once()
        client.me.calendars.get.assert_not_called()
        secondary.events.post.assert_not_called()

    @pytest.mark.parametrize("alias", ["primary", "PRIMARY", "", "   "])
    async def test_primary_and_blank_mean_the_default_calendar(self, alias):
        client, secondary = _client_with_calendars()

        await _create(client, calendar=alias)

        client.me.events.post.assert_called_once()
        client.me.calendars.get.assert_not_called()
        secondary.events.post.assert_not_called()

    async def test_a_display_name_posts_to_that_calendar(self):
        client, secondary = _client_with_calendars()

        result = await _create(client, calendar="kids + school")

        client.me.calendars.by_calendar_id.assert_called_once_with("KIDS77=")
        secondary.events.post.assert_called_once()
        client.me.events.post.assert_not_called()
        assert result["event_id"] == "AAMkAGsec="
        # The event itself is the one that would have gone to the default.
        posted = secondary.events.post.call_args[0][0]
        assert posted.subject == "Dentist"

    async def test_an_id_posts_to_that_calendar(self):
        client, secondary = _client_with_calendars()

        await _create(client, calendar="KIDS77=")

        client.me.calendars.by_calendar_id.assert_called_once_with("KIDS77=")
        secondary.events.post.assert_called_once()

    async def test_an_unknown_calendar_is_refused_naming_what_exists(self):
        client, secondary = _client_with_calendars()

        with pytest.raises(ValueError) as exc:
            await _create(client, calendar="Holidays")

        message = str(exc.value)
        assert "Calendar 'Holidays' not found" in message
        assert "Kids + School" in message
        client.me.events.post.assert_not_called()
        secondary.events.post.assert_not_called()

    async def test_an_ambiguous_name_is_refused_with_the_candidates(self):
        client, secondary = _client_with_calendars(
            [_calendar("Calendar", "DEFAULT1="), _calendar("Work", "W1="), _calendar("work", "W2=")]
        )

        with pytest.raises(ValueError) as exc:
            await _create(client, calendar="Work")

        assert "ambiguous" in str(exc.value)
        assert "W1=" in str(exc.value) and "W2=" in str(exc.value)
        secondary.events.post.assert_not_called()

    async def test_a_read_only_calendar_is_refused_before_the_post(self):
        """Graph's answer would be a 403 whose hint blames the account type."""
        client, secondary = _client_with_calendars()

        with pytest.raises(ValueError) as exc:
            await _create(client, calendar="Birthdays")

        message = str(exc.value)
        assert "Birthdays" in message
        assert "read-only" in message
        # Names where the event can go instead, and leaves out where it cannot.
        assert "Kids + School" in message
        assert "Calendar" in message
        assert message.count("Birthdays") == 1
        client.me.calendars.by_calendar_id.assert_not_called()
        secondary.events.post.assert_not_called()
        client.me.events.post.assert_not_called()

    async def test_a_read_only_calendar_given_by_id_is_refused_too(self):
        client, secondary = _client_with_calendars()

        with pytest.raises(ValueError, match="read-only"):
            await _create(client, calendar="BDAY42=")

        secondary.events.post.assert_not_called()

    async def test_a_calendar_that_reports_no_can_edit_is_left_to_graph(self):
        """Only an explicit `canEdit: false` refuses; Graph has the last word otherwise."""
        client, secondary = _client_with_calendars(
            [_calendar("Calendar", "DEFAULT1="), _calendar("Team", "TEAM9=", can_edit=None)]
        )

        await _create(client, calendar="Team")

        secondary.events.post.assert_called_once()

    async def test_a_bad_argument_is_refused_before_the_calendar_listing(self):
        """The listing is the only request this argument adds; a call that will
        be refused anyway must not pay for it."""
        client, _ = _client_with_calendars()

        with pytest.raises(ValueError):
            await create_event(
                client,
                subject="Dentist",
                start="not a date",
                end="2026-11-03T09:30:00",
                calendar="Kids + School",
                config=_CFG,
            )
        with pytest.raises(ValueError):
            await _create(client, calendar="Kids + School", show_as="maybe")
        with pytest.raises(ValueError):
            await _create(client, calendar="Kids + School", recurrence="fortnightly-ish")

        client.me.calendars.get.assert_not_called()

    async def test_a_read_only_config_is_refused_before_the_calendar_listing(self):
        client, _ = _client_with_calendars()

        with pytest.raises(ReadOnlyError):
            await create_event(
                client,
                subject="Dentist",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                calendar="Kids + School",
                config=_CFG_RO,
            )

        client.me.calendars.get.assert_not_called()

    async def test_attendees_still_need_mail_send_on_another_calendar(self):
        client, secondary = _client_with_calendars()

        with pytest.raises(PermissionDeniedError):
            await create_event(
                client,
                subject="Sync",
                start="2026-11-03T09:00:00",
                end="2026-11-03T09:30:00",
                attendees=["someone@example.com"],
                calendar="Kids + School",
                config=_CFG_CAL_ONLY,
            )

        client.me.calendars.get.assert_not_called()
        secondary.events.post.assert_not_called()


class TestCreateEventCalendarOnTheWire:
    """The URL the real SDK pipeline sends, not the builder a mock was handed.

    A mock asserts which request builder was called; it cannot see the path that
    builder produces, nor how kiota encodes an ID ending in `=`. So this builds
    the client the way `msgraph` does (as `test_me_rewrite_reaches_the_wire.py`
    does) over a mock transport, and asserts on what reaches it.
    """

    @staticmethod
    def _client(calendars: list[dict]):
        import httpx
        from azure.core.credentials import AccessToken
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as _gra
        from msgraph_core import GraphClientFactory

        class _Credential:
            def get_token(self, *args, **kwargs):
                return AccessToken("not-a-real-token", 9_999_999_999)

        sent: list[httpx.Request] = []

        def handle(request: httpx.Request) -> httpx.Response:
            sent.append(request)
            if request.method == "GET" and request.url.path.endswith("/me/calendars"):
                return httpx.Response(200, json={"value": calendars})
            if request.method == "POST" and request.url.path.endswith("/events"):
                return httpx.Response(201, json={"id": "AAMkAGnew=", "subject": "Dentist"})
            raise AssertionError(f"unexpected request: {request.method} {request.url}")

        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options,
            client=httpx.AsyncClient(transport=httpx.MockTransport(handle)),
        )
        auth = AzureIdentityAuthenticationProvider(
            _Credential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client)), sent

    _LISTED = [
        {"id": "AAMkDEF=", "name": "Calendar", "canEdit": True},
        {"id": "AAMkKIDS=", "name": "Kids + School", "canEdit": True},
    ]

    async def test_a_named_calendar_is_posted_to_its_own_events_collection(self):
        import json

        client, sent = self._client(self._LISTED)

        result = await _create(client, calendar="Kids + School")

        assert result["event_id"] == "AAMkAGnew="
        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/calendars"),
            ("POST", "/v1.0/me/calendars/AAMkKIDS=/events"),
        ]
        assert json.loads(sent[-1].content)["subject"] == "Dentist"

    async def test_the_default_calendar_is_one_post_to_me_events(self):
        client, sent = self._client(self._LISTED)

        await _create(client)

        assert [(r.method, r.url.path) for r in sent] == [("POST", "/v1.0/me/events")]


# ── event keys ────────────────────────────────────────────────────────

_KEY_PROP = "String {51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey"


class TestEventKeyWrites:
    async def test_create_stamps_the_key_on_the_event(self):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Flight"))

        await create_event(
            mock_client,
            subject="Flight",
            start="2026-11-03T09:00:00",
            end="2026-11-03T11:00:00",
            key="flight:AB12CD:LHR-JFK",
            config=_CFG,
        )

        (prop,) = _posted(mock_client).single_value_extended_properties
        assert prop.id == _KEY_PROP
        assert prop.value == "flight:AB12CD:LHR-JFK"

    async def test_create_without_a_key_sends_no_property(self):
        mock_client = AsyncMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Flight"))

        await create_event(
            mock_client,
            subject="Flight",
            start="2026-11-03T09:00:00",
            end="2026-11-03T11:00:00",
            config=_CFG,
        )

        assert _posted(mock_client).single_value_extended_properties is None

    async def test_create_refuses_an_invalid_key_before_sending(self):
        mock_client = AsyncMock()

        with pytest.raises(ValueError, match="Invalid event key"):
            await create_event(
                mock_client,
                subject="Flight",
                start="2026-11-03T09:00:00",
                end="2026-11-03T11:00:00",
                key="has space",
                config=_CFG,
            )

        mock_client.me.events.post.assert_not_called()

    async def test_update_sets_the_key_and_nothing_else(self):
        builder, client = _client_with(_current_event())

        await update_event(client, event_id="AAMkAG123=", key="booking:ABC123", config=_CFG)

        patched = builder.patch.call_args[0][0]
        (prop,) = patched.single_value_extended_properties
        assert (prop.id, prop.value) == (_KEY_PROP, "booking:ABC123")
        assert patched.subject is None and patched.start is None and patched.recurrence is None

    async def test_keying_a_series_master_reads_nothing_first(self):
        """A key is not a reshape: no anchor read, no occurrence check, no pin.

        Graph restores a series' changed occurrences when its times or pattern
        change; a key changes neither. The live tier checks that against Graph.
        """
        master = _current_event(event_type="seriesMaster", edited=[MagicMock()])
        builder, client = _client_with(master)

        await update_event(client, event_id="AAMkAG123=", key="booking:ABC123", config=_CFG)

        builder.get.assert_not_called()
        # No pin, asked of the header alone: the update may still carry a request
        # configuration for something else, such as switching kiota's retries off.
        sent = builder.patch.call_args.kwargs.get("request_configuration")
        assert sent is None or not sent.headers.get("If-Match")

    async def test_update_refuses_an_invalid_key_before_any_request(self):
        builder, client = _client_with(_current_event())

        with pytest.raises(ValueError, match="Invalid event key"):
            await update_event(client, event_id="AAMkAG123=", key="", config=_CFG)

        builder.get.assert_not_called()
        builder.patch.assert_not_called()

    async def test_a_key_on_an_event_with_attendees_needs_no_mail_send(self):
        """The reword gate covers what Outlook shows attendees (subject, body,
        location and an online meeting); a key is not one of them, so it pays for
        no read."""
        builder, client = _client_with(_meeting())

        await update_event(
            client, event_id="AAMkAG123=", key="booking:ABC123", config=_CFG_CAL_ONLY
        )

        builder.patch.assert_called_once()


# ── is_online: what Graph does with it differs by account (#70) ───────


def _reply(is_online_meeting):
    """A write reply whose `is_online_meeting` is set explicitly.

    Never left to MagicMock: an auto-created attribute is neither a bool nor
    None, so a test of "reported as Graph states it" would pass on a reply
    that states nothing.
    """
    return MagicMock(id="AAMkAG123=", subject="Sync", is_online_meeting=is_online_meeting)


class TestOnlineMeeting:
    @pytest.mark.parametrize("stated", [True, False])
    async def test_update_reports_is_online_as_graph_states_it(self, stated):
        """`False` is what an account that ignores the request answers — the
        caller has to see it on the call itself, not a bare `updated`."""
        builder, client = _client_with(_current_event())
        builder.patch = AsyncMock(return_value=_reply(stated))

        result = await update_event(client, event_id="AAMkAG123=", is_online=True, config=_CFG)

        assert builder.patch.call_args[0][0].is_online_meeting is True
        assert result["is_online"] is stated

    async def test_update_omits_is_online_when_the_reply_does_not_carry_it(self):
        """Absent is not "no meeting"; reporting `False` here would invent one."""
        builder, client = _client_with(_current_event())
        builder.patch = AsyncMock(return_value=_reply(None))

        result = await update_event(client, event_id="AAMkAG123=", is_online=True, config=_CFG)

        assert "is_online" not in result

    async def test_update_without_is_online_reports_nothing_about_it(self):
        builder, client = _client_with(_current_event())
        builder.patch = AsyncMock(return_value=_reply(True))

        result = await update_event(client, event_id="AAMkAG123=", subject="X", config=_CFG)

        assert result == {"status": "updated", "event_id": "AAMkAG123="}

    @pytest.mark.parametrize("config", [_CFG, _CFG_CAL_ONLY], ids=["open", "restricted"])
    async def test_update_refuses_is_online_false_before_any_request(self, config):
        """Graph accepts `isOnlineMeeting: false` and keeps the meeting — probed
        2026-11-09, the reply and a fresh GET both still reporting it 90 s later.
        `test_patching_online_off_keeps_the_meeting` is the live premise guard.

        Under a restricted policy as well, because that is where the reads before
        the PATCH happen: an unrestricted call reads nothing either way, so a
        refusal moved below the first network call would pass there."""
        builder, client = _client_with(_current_event())

        with pytest.raises(ValueError, match="cannot remove an online meeting") as refused:
            await update_event(client, event_id="AAMkAG123=", is_online=False, config=config)

        # Deleting comes first: an agent that only re-creates is left with two events,
        # and the attendees get no cancellation.
        assert "delete the event and create it again" in str(refused.value)
        builder.get.assert_not_called()
        builder.calendar.get.assert_not_called()
        builder.patch.assert_not_called()

    @pytest.mark.parametrize("stated", [True, False])
    async def test_create_reports_is_online_as_graph_states_it(self, stated):
        mock_client = async_sdk()
        mock_client.me.events.post = AsyncMock(return_value=_reply(stated))

        result = await create_event(
            mock_client,
            subject="Sync",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            is_online=True,
            config=_CFG,
        )

        assert _posted(mock_client).is_online_meeting is True
        assert result["is_online"] is stated

    async def test_create_omits_is_online_when_the_reply_does_not_carry_it(self):
        mock_client = async_sdk()
        mock_client.me.events.post = AsyncMock(return_value=_reply(None))

        result = await create_event(
            mock_client,
            subject="Sync",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            is_online=True,
            config=_CFG,
        )

        assert "is_online" not in result

    async def test_create_without_is_online_keeps_its_old_result(self):
        """The default `False` is not reported: it cannot be told apart from an
        argument nobody passed, and every create would grow a key."""
        mock_client = async_sdk()
        mock_client.me.events.post = AsyncMock(return_value=_reply(False))

        result = await create_event(
            mock_client,
            subject="Sync",
            start="2026-04-15T14:00:00Z",
            end="2026-04-15T15:00:00Z",
            config=_CFG,
        )

        assert result == {"status": "created", "event_id": "AAMkAG123=", "subject": "Sync"}


class TestRelativeEventTimes:
    """`start="+2h"` is resolved before it is sent, not forwarded as text (#74).

    Before, both tools validated the relative form and then sent the caller's
    own string, which Graph refused with a 400 naming a `'DateTime'` parameter
    this API does not have. The mock suite could not see it: a mock accepts
    `date_time="+2h"` as happily as a real time.

    The resolved moment goes out as wall-clock time in the event's anchor zone,
    with no `Z` or offset: handed either, Graph anchors a single event in UTC
    whatever `timeZone` says (#236). So these assert the local text and the
    zone beside it, and rebuild the instant from the two to check the moment.
    """

    # Thursday 01:00 UTC, which is Wednesday 17:00 in Los Angeles and 20:00 in
    # New York. The UTC and local dates differ, so a value converted in the
    # wrong zone, or a recurrence derived from the UTC date, cannot pass.
    NOW = datetime(2026, 11, 12, 1, 0, 0, tzinfo=timezone.utc)

    @pytest.fixture
    def clock(self, monkeypatch):
        """Each read of the clock is one second later than the last.

        So a tool that resolved `start` and `end` against two separate reads
        would put an hour and a second between `+2h` and `+3h`.
        """
        reads = []

        def ticking():
            reads.append(self.NOW + timedelta(seconds=len(reads)))
            return reads[-1]

        monkeypatch.setattr(calendar_write, "_now", ticking)
        return reads

    @staticmethod
    def _fixed(monkeypatch, moment):
        monkeypatch.setattr(calendar_write, "_now", lambda: moment)

    @staticmethod
    def _instant(local: str, zone: str) -> datetime:
        """The moment a naive `dateTime` names in `zone`, in UTC."""
        from zoneinfo import ZoneInfo

        return datetime.fromisoformat(local).replace(tzinfo=ZoneInfo(zone)).astimezone(timezone.utc)

    @staticmethod
    def _updatable(start_zone: str, end_zone: str | None = None):
        builder = _make_event_builder()
        builder.get = AsyncMock(
            return_value=_current_event(start_zone=start_zone, end_zone=end_zone)
        )
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)
        return builder, mock_client

    # ── create ────────────────────────────────────────────────────────

    async def test_create_writes_local_time_from_one_clock_read(self, clock):
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        await create_event(mock_client, subject="Standup", start="+2h", end="+3h", config=_CFG_LA)

        event = _posted(mock_client)
        assert event.start.date_time == "2026-11-11T19:00:00"  # 03:00Z, Los Angeles
        assert event.end.date_time == "2026-11-11T20:00:00"
        assert event.start.time_zone == event.end.time_zone == "America/Los_Angeles"
        assert self._instant(event.start.date_time, "America/Los_Angeles") == (
            self.NOW + timedelta(hours=2)
        )
        assert len(clock) == 1

    async def test_create_converts_across_a_daylight_saving_change(self, monkeypatch):
        """Saturday 12:00 PST plus 24 hours is Sunday 13:00 PDT, not 12:00."""
        self._fixed(monkeypatch, datetime(2027, 3, 13, 20, 0, tzinfo=timezone.utc))
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Brunch"))

        await create_event(mock_client, subject="Brunch", start="+24h", end="+25h", config=_CFG_LA)

        assert _posted(mock_client).start.date_time == "2027-03-14T13:00:00"

    async def test_a_relative_start_anchors_a_new_series_on_its_local_day(self, clock):
        """`+24h` is Friday 01:00 UTC: Thursday afternoon in Los Angeles."""
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Review"))

        await create_event(
            mock_client,
            subject="Review",
            start="+24h",
            end="+25h",
            recurrence="weekly",
            config=_CFG_LA,
        )

        recurrence = _posted(mock_client).recurrence
        assert [d.value for d in recurrence.pattern.days_of_week] == ["thursday"]
        assert recurrence.range.start_date == date(2026, 11, 12)

    # ── update ────────────────────────────────────────────────────────

    async def test_update_writes_local_time_in_the_events_own_zone(self, clock):
        """New York's wall clock, not the configured zone's (UTC here)."""
        builder, mock_client = self._updatable("America/New_York")

        await update_event(mock_client, event_id="AAMkAG123=", start="+2h", end="+3h", config=_CFG)

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-11-11T22:00:00"
        assert patched.end.date_time == "2026-11-11T23:00:00"
        assert patched.start.time_zone == "America/New_York"
        assert len(clock) == 1

    async def test_each_end_is_placed_in_its_own_zone(self, clock):
        """A flight leaves New York and lands in Los Angeles; each end keeps its zone."""
        builder, mock_client = self._updatable("America/New_York", "America/Los_Angeles")

        await update_event(mock_client, event_id="AAMkAG123=", start="+2h", end="+8h", config=_CFG)

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-11-11T22:00:00"  # 03:00Z in New York
        assert patched.end.date_time == "2026-11-12T01:00:00"  # 09:00Z in Los Angeles

    @pytest.mark.parametrize(
        ("start", "end", "sent_start", "sent_end"),
        [
            ("+2h", "2026-11-12T23:30:00", "2026-11-11T22:00:00", "2026-11-12T23:30:00"),
            # A typed offset loses only the offset: the same local time (#236).
            (
                "2026-11-11T21:00:00-05:00",
                "+3h",
                "2026-11-11T21:00:00",
                "2026-11-11T23:00:00",
            ),
        ],
    )
    async def test_a_typed_end_beside_a_relative_one_keeps_its_own_time(
        self, clock, start, end, sent_start, sent_end
    ):
        builder, mock_client = self._updatable("America/New_York")

        await update_event(mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG)

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == sent_start
        assert patched.end.date_time == sent_end

    async def test_the_same_series_change_in_an_iana_zone_is_sent_on_its_local_day(self, clock):
        builder, mock_client = self._updatable("America/Los_Angeles")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="+24h",
            end="+25h",
            recurrence="weekly",
            config=_CFG,
        )

        recurrence = builder.patch.call_args[0][0].recurrence
        assert [d.value for d in recurrence.pattern.days_of_week] == ["thursday"]

    # ── refusals ──────────────────────────────────────────────────────

    @pytest.mark.parametrize(("start", "end"), [("+2h", "+3h"), ("+2h", None), (None, "+3h")])
    async def test_a_relative_time_on_a_windows_zoned_event_is_refused(self, clock, start, end):
        """`zoneinfo` cannot convert into "Pacific Standard Time".

        The series this server created came back named that way when measured
        (2026-11-10, a consumer mailbox), so this is not rare. Sending `Z`
        instead would re-anchor the event in UTC (#236), so it is refused before
        anything is sent.
        """
        builder, mock_client = self._updatable("Pacific Standard Time")

        with pytest.raises(ValueError, match="Pacific Standard Time") as exc:
            await update_event(
                mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG
            )

        assert "Nothing was modified" in str(exc.value)
        assert "without a zone or offset" in str(exc.value)
        builder.patch.assert_not_called()

    async def test_a_zone_less_time_on_a_windows_zoned_event_is_sent(self, clock):
        """The remedy the refusal names really works: wall-clock time is already local."""
        builder, mock_client = self._updatable("Pacific Standard Time")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-12T17:00:00",
            end="2026-11-12T18:00:00",
            recurrence="weekly",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-11-12T17:00:00"
        assert patched.start.time_zone == "Pacific Standard Time"
        assert [d.value for d in patched.recurrence.pattern.days_of_week] == ["thursday"]

    async def test_a_custom_zone_event_keeps_its_own_refusal(self, clock):
        """Graph's custom-zone sentinel has no zone to place a time in; `_anchored` says so."""
        builder, mock_client = self._updatable("tzone://Microsoft/Custom")

        with pytest.raises(ValueError, match="Could not read the zone"):
            await update_event(
                mock_client, event_id="AAMkAG123=", start="+2h", end="+3h", config=_CFG
            )

        builder.patch.assert_not_called()

    # Los Angeles leaves daylight saving at 02:00 on 2027-11-07, so 01:00-02:00
    # happens twice: 08:00-09:00Z (PDT) and 09:00-10:00Z (PST). From 06:30Z,
    # +2h is 01:30 the first time and +3h is 01:30 the second.
    REPEAT = datetime(2027, 11, 7, 6, 30, tzinfo=timezone.utc)

    def test_the_repeated_interval_really_is_repeated(self):
        """Guard the premise: if tzdata drops this transition, the tests below prove nothing."""
        from zoneinfo import ZoneInfo

        first = (self.REPEAT + timedelta(hours=2)).astimezone(ZoneInfo("America/Los_Angeles"))
        second = (self.REPEAT + timedelta(hours=3)).astimezone(ZoneInfo("America/Los_Angeles"))
        assert first.strftime("%H:%M") == second.strftime("%H:%M") == "01:30"
        assert first.utcoffset() != second.utcoffset()

    @pytest.mark.parametrize(("start", "end"), [("+2h", "+5h"), ("+3h", "+5h"), ("+1h", "+2h")])
    async def test_create_refuses_a_time_in_the_repeated_interval(self, monkeypatch, start, end):
        self._fixed(monkeypatch, self.REPEAT)
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Late"))

        with pytest.raises(ValueError, match="repeats when daylight saving ends") as exc:
            await create_event(mock_client, subject="Late", start=start, end=end, config=_CFG_LA)

        # It names the local time that happens twice, and offers no offset as
        # the remedy: a typed one anchors a single event in UTC (#236).
        message = str(exc.value)
        assert "2027-11-07T01:30:00 happens twice" in message
        assert "Pass a time outside that interval." in message
        assert "-07:00" not in message and "-08:00" not in message
        assert "offset" not in message
        mock_client.me.events.post.assert_not_called()

    async def test_create_refuses_it_with_a_recurrence_too(self, monkeypatch):
        self._fixed(monkeypatch, self.REPEAT)
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Late"))

        with pytest.raises(ValueError, match="repeats when daylight saving ends"):
            await create_event(
                mock_client,
                subject="Late",
                start="+2h",
                end="+5h",
                recurrence="weekly",
                config=_CFG_LA,
            )

        mock_client.me.events.post.assert_not_called()

    async def test_create_sends_times_either_side_of_the_repeated_interval(self, monkeypatch):
        """00:30 PDT and 02:30 PST: a pair around the repeat, neither end inside it."""
        self._fixed(monkeypatch, self.REPEAT)
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Late"))

        await create_event(mock_client, subject="Late", start="+1h", end="+4h", config=_CFG_LA)

        event = _posted(mock_client)
        assert (event.start.date_time, event.end.date_time) == (
            "2027-11-07T00:30:00",
            "2027-11-07T02:30:00",
        )

    @pytest.mark.parametrize(("start", "end"), [("+2h", None), (None, "+3h")])
    async def test_update_refuses_a_time_in_the_repeated_interval(self, monkeypatch, start, end):
        self._fixed(monkeypatch, self.REPEAT)
        builder, mock_client = self._updatable("America/Los_Angeles")

        with pytest.raises(ValueError, match="repeats when daylight saving ends"):
            await update_event(
                mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG
            )

        builder.patch.assert_not_called()

    async def test_update_sends_times_either_side_of_the_repeated_interval(self, monkeypatch):
        self._fixed(monkeypatch, self.REPEAT)
        builder, mock_client = self._updatable("America/Los_Angeles")

        await update_event(mock_client, event_id="AAMkAG123=", start="+1h", end="+4h", config=_CFG)

        patched = builder.patch.call_args[0][0]
        assert (patched.start.date_time, patched.end.date_time) == (
            "2027-11-07T00:30:00",
            "2027-11-07T02:30:00",
        )

    # ── what the result reports (#74) ─────────────────────────────────

    async def test_create_reports_each_relative_end_as_written(self, clock):
        """So an agent that passed `start="7d"` can see it made a past event."""
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))

        result = await create_event(
            mock_client, subject="Standup", start="7d", end="+3h", config=_CFG_LA
        )

        assert result["start"] == "2026-11-04T17:00:00 (America/Los_Angeles)"
        assert result["end"] == "2026-11-11T20:00:00 (America/Los_Angeles)"

    async def test_update_reports_each_relative_end_in_its_own_zone(self, clock):
        builder, mock_client = self._updatable("America/New_York", "America/Los_Angeles")

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start="+2h", end="+8h", config=_CFG
        )

        assert result["start"] == "2026-11-11T22:00:00 (America/New_York)"
        assert result["end"] == "2026-11-12T01:00:00 (America/Los_Angeles)"

    async def test_a_typed_end_is_not_reported(self, clock):
        """Only what the server worked out is echoed; typed input the caller knows."""
        builder, mock_client = self._updatable("America/New_York")

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start="+2h", end="2026-11-12T23:30:00", config=_CFG
        )

        assert result["start"] == "2026-11-11T22:00:00 (America/New_York)"
        assert "end" not in result

    async def test_iso_times_keep_the_old_results(self, clock):
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Standup"))
        created = await create_event(
            mock_client,
            subject="Standup",
            start="2026-11-12T09:00:00",
            end="2026-11-12T10:00:00",
            config=_CFG_LA,
        )
        builder, update_client = self._updatable("America/New_York")
        updated = await update_event(
            update_client,
            event_id="AAMkAG123=",
            start="2026-11-12T09:00:00",
            end="2026-11-12T10:00:00",
            config=_CFG,
        )

        assert created == {"status": "created", "event_id": "AAMkAGnew=", "subject": "Standup"}
        assert updated == {"status": "updated", "event_id": "AAMkAG123="}

    # ── all-day events (#74) ──────────────────────────────────────────

    @pytest.mark.parametrize(("start", "end"), [("+24h", "2026-11-14T00:00:00"), ("now", "+1d")])
    async def test_create_refuses_a_relative_all_day_time_before_sending(self, clock, start, end):
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Off"))

        with pytest.raises(ValueError, match="cannot place an all-day event"):
            await create_event(
                mock_client, subject="Off", start=start, end=end, is_all_day=True, config=_CFG_LA
            )

        mock_client.me.events.post.assert_not_called()

    async def test_update_refuses_it_before_any_request_when_the_caller_says_all_day(self, clock):
        builder, mock_client = self._updatable("America/New_York")

        with pytest.raises(ValueError, match="cannot place an all-day event"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="+24h",
                end="+48h",
                is_all_day=True,
                config=_CFG,
            )

        builder.get.assert_not_called()
        builder.patch.assert_not_called()

    async def test_update_refuses_it_when_only_the_stored_event_is_all_day(self, clock):
        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=_current_event(is_all_day=True))
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        with pytest.raises(ValueError, match="cannot place an all-day event"):
            await update_event(mock_client, event_id="AAMkAG123=", start="+24h", config=_CFG)

        builder.patch.assert_not_called()

    async def test_a_typed_midnight_on_an_all_day_event_is_still_sent(self, clock):
        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=_current_event(is_all_day=True))
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-14T00:00:00",
            end="2026-11-15T00:00:00",
            config=_CFG,
        )

        builder.patch.assert_called_once()


class TestATypedInstantKeepsItsZone:
    """A `Z` or offset `start`/`end` goes out as wall-clock time in the anchor zone (#236).

    Handed a `Z` or an offset, Graph keeps the instant and anchors a *single*
    event in UTC whatever `timeZone` says; handed the same moment as wall-clock
    time, it keeps the zone. Measured live on a consumer mailbox, 2026-11-10. A
    mock stores either, so these assert the text sent and rebuild the instant
    from it, and the live tier asks Graph which zone it kept.
    """

    LA = "America/Los_Angeles"

    @staticmethod
    def _instant(local: str, zone: str) -> datetime:
        from zoneinfo import ZoneInfo

        return datetime.fromisoformat(local).replace(tzinfo=ZoneInfo(zone)).astimezone(timezone.utc)

    @staticmethod
    def _poster():
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(return_value=_created("Dinner"))
        reads_no_times(mock_client.me.events)
        return mock_client

    @staticmethod
    def _updatable(start_zone: str, end_zone: str | None = None, **kwargs):
        builder = _make_event_builder()
        builder.get = AsyncMock(
            return_value=_current_event(start_zone=start_zone, end_zone=end_zone, **kwargs)
        )
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)
        return builder, mock_client

    # ── create ────────────────────────────────────────────────────────

    @pytest.mark.parametrize(
        ("start", "end"),
        [
            ("2026-12-13T02:00:00Z", "2026-12-13T03:00:00Z"),
            ("2026-12-12T18:00:00-08:00", "2026-12-12T22:00:00-05:00"),
            ("2026-12-13T07:30:00+05:30", "2026-12-13T03:00:00+00:00"),
        ],
    )
    async def test_create_sends_the_same_moment_as_local_time(self, start, end):
        mock_client = self._poster()

        await create_event(
            mock_client, subject="Dinner", start=start, end=end, timezone=self.LA, config=_CFG
        )

        event = _posted(mock_client)
        assert event.start.date_time == "2026-12-12T18:00:00"
        assert event.end.date_time == "2026-12-12T19:00:00"
        assert event.start.time_zone == event.end.time_zone == self.LA
        assert self._instant(event.start.date_time, self.LA) == datetime.fromisoformat(
            start.replace("Z", "+00:00")
        )

    async def test_create_converts_into_the_configured_zone_by_default(self):
        mock_client = self._poster()

        await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG_LA,
        )

        event = _posted(mock_client)
        assert (event.start.date_time, event.start.time_zone) == ("2026-12-12T18:00:00", self.LA)

    async def test_create_keeps_a_fraction_of_a_second(self):
        """The moment is kept to the microsecond. (Graph's own seven digits are
        refused by input validation on Python 3.10, so six are used here.)"""
        mock_client = self._poster()

        await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-13T02:00:00.123456Z",
            end="2026-12-13T03:00:00.500Z",
            timezone=self.LA,
            config=_CFG,
        )

        event = _posted(mock_client)
        assert event.start.date_time == "2026-12-12T18:00:00.123456"
        assert event.end.date_time == "2026-12-12T19:00:00.500000"

    async def test_a_zone_less_time_is_sent_as_written(self):
        mock_client = self._poster()

        await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-12T18:00:00",
            end="2026-12-12T19:00:00",
            timezone=self.LA,
            config=_CFG,
        )

        assert _posted(mock_client).start.date_time == "2026-12-12T18:00:00"

    async def test_a_series_from_a_z_start_lands_on_its_local_day(self):
        """Thursday 02:00Z is Wednesday 18:00 in Los Angeles; the pattern says Wednesday."""
        mock_client = self._poster()

        await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-17T02:00:00Z",
            end="2026-12-17T03:00:00Z",
            recurrence="weekly",
            timezone=self.LA,
            config=_CFG,
        )

        event = _posted(mock_client)
        assert event.start.date_time == "2026-12-16T18:00:00"
        assert [d.value for d in event.recurrence.pattern.days_of_week] == ["wednesday"]
        assert event.recurrence.range.start_date == date(2026, 12, 16)

    async def test_an_all_day_event_is_sent_as_written(self):
        """Graph stores an all-day event in UTC whatever it is sent, so nothing is converted."""
        mock_client = self._poster()

        await create_event(
            mock_client,
            subject="Off",
            start="2026-12-13T00:00:00Z",
            end="2026-12-14T00:00:00Z",
            is_all_day=True,
            config=_CFG_LA,
        )

        event = _posted(mock_client)
        assert (event.start.date_time, event.start.time_zone) == ("2026-12-13T00:00:00Z", "UTC")

    @pytest.mark.parametrize(
        ("start", "end"),
        [
            # 01:30 happens twice in Los Angeles on 2027-11-07: the start is in it.
            ("2027-11-07T09:30:00Z", "2027-11-07T11:00:00Z"),
            # 00:30 local is fine; the end, 01:30 PDT, is the repeated one.
            ("2027-11-07T07:30:00Z", "2027-11-07T08:30:00Z"),
        ],
    )
    async def test_a_moment_in_the_repeated_hour_sends_both_ends_as_written(self, start, end):
        """Local time cannot say which 01:30 is meant, so that end keeps its `Z`.

        And so does the other one: converting only one would file the two ends
        in different zones, a state the old behaviour (both in UTC) never made.
        Refusing would break input that works today.
        """
        mock_client = self._poster()

        await create_event(
            mock_client, subject="Late", start=start, end=end, timezone=self.LA, config=_CFG
        )

        event = _posted(mock_client)
        assert (event.start.date_time, event.end.date_time) == (start, end)

    @pytest.mark.parametrize(
        ("start", "end", "zone"),
        [
            # East of UTC, the last second of year 9999 is past the last date.
            ("2026-12-13T02:00:00Z", "9999-12-31T23:59:59Z", "Asia/Tokyo"),
            # West of UTC, the first moment of year 1 is before the first date.
            ("0001-01-01T00:00:00Z", "2026-12-13T03:00:00Z", "America/Los_Angeles"),
        ],
    )
    async def test_a_moment_outside_the_years_python_can_hold_is_sent_as_written(
        self, start, end, zone
    ):
        """`astimezone` overflows there; that is no reason to fail a call that worked before."""
        mock_client = self._poster()

        await create_event(
            mock_client, subject="Forever", start=start, end=end, timezone=zone, config=_CFG
        )

        event = _posted(mock_client)
        assert (event.start.date_time, event.end.date_time) == (start, end)

    # ── update ────────────────────────────────────────────────────────

    async def test_update_converts_into_the_events_own_zone(self):
        builder, mock_client = self._updatable(self.LA)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-12-12T18:00:00"
        assert patched.end.date_time == "2026-12-12T19:00:00"
        assert patched.start.time_zone == self.LA

    async def test_update_converts_into_a_requested_zone(self):
        builder, mock_client = self._updatable(self.LA)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-12T18:00:00-08:00",
            end="2026-12-12T19:00:00-08:00",
            timezone="America/New_York",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-12-12T21:00:00"
        assert patched.start.time_zone == "America/New_York"

    async def test_each_end_is_converted_into_its_own_zone(self):
        """A flight leaves New York and lands in Los Angeles."""
        builder, mock_client = self._updatable("America/New_York", self.LA)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T13:00:00Z",
            end="2026-12-13T19:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert (patched.start.date_time, patched.start.time_zone) == (
            "2026-12-13T08:00:00",
            "America/New_York",
        )
        assert (patched.end.date_time, patched.end.time_zone) == ("2026-12-13T11:00:00", self.LA)

    async def test_update_converts_an_end_patched_alone(self):
        builder, mock_client = self._updatable(self.LA)

        await update_event(
            mock_client, event_id="AAMkAG123=", end="2026-12-13T03:00:00Z", config=_CFG
        )

        assert builder.patch.call_args[0][0].end.date_time == "2026-12-12T19:00:00"

    async def test_update_sends_both_ends_as_written_when_one_is_in_the_repeated_hour(self):
        builder, mock_client = self._updatable(self.LA)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2027-11-07T07:30:00Z",
            end="2027-11-07T08:30:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert (patched.start.date_time, patched.end.date_time) == (
            "2027-11-07T07:30:00Z",
            "2027-11-07T08:30:00Z",
        )

    async def test_an_end_in_an_unknown_zone_keeps_the_start_as_written_too(self):
        """Start in Los Angeles, end in a zone no table names: the end cannot convert."""
        builder, mock_client = self._updatable(self.LA, "Nowhere Standard Time")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert (patched.start.date_time, patched.end.date_time) == (
            "2026-12-13T02:00:00Z",
            "2026-12-13T03:00:00Z",
        )

    async def test_a_custom_zone_event_is_still_refused(self):
        builder, mock_client = self._updatable("tzone://Microsoft/Custom")

        with pytest.raises(ValueError, match="Could not read the zone"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-12-13T02:00:00Z",
                end="2026-12-13T03:00:00Z",
                config=_CFG,
            )

        builder.patch.assert_not_called()

    async def test_update_converts_across_the_date_line(self):
        """11:00 at UTC-12 on New Year's Eve is 13:00 on New Year's Day at UTC+14."""
        builder, mock_client = self._updatable("Pacific/Kiritimati")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-31T11:00:00-12:00",
            end="2026-12-31T12:00:00-12:00",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2027-01-01T13:00:00"
        assert patched.end.date_time == "2027-01-01T14:00:00"

    async def test_an_unknown_zone_name_is_sent_as_written(self):
        """Neither `zoneinfo` nor the Windows table knows it, so the old behaviour stays."""
        builder, mock_client = self._updatable("Nowhere Standard Time")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-12-13T02:00:00Z"
        assert patched.start.time_zone == "Nowhere Standard Time"

    @pytest.mark.parametrize(
        ("windows", "iana"),
        [
            ("Pacific Standard Time", "America/Los_Angeles"),
            ("W. Europe Standard Time", "Europe/Berlin"),
            ("Tokyo Standard Time", "Asia/Tokyo"),
        ],
    )
    async def test_a_windows_named_zone_converts_like_its_iana_equivalent(self, windows, iana):
        """Graph names an Outlook-made event's zone in Windows terms; it is sent back as is."""
        sent = {}
        for zone in (windows, iana):
            builder, mock_client = self._updatable(zone)
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-12-13T02:00:00Z",
                end="2026-12-13T03:00:00Z",
                config=_CFG,
            )
            patched = builder.patch.call_args[0][0]
            sent[zone] = (patched.start.date_time, patched.end.date_time)
            assert (patched.start.time_zone, patched.end.time_zone) == (zone, zone)

        assert sent[windows] == sent[iana]
        assert not sent[windows][0].endswith("Z")

    async def test_a_z_start_and_recurrence_on_a_windows_event_use_its_local_day(self):
        """#235: Thursday 01:00Z is Wednesday 17:00 Pacific, so the series is Wednesday's."""
        builder, mock_client = self._updatable("Pacific Standard Time")

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-11-12T01:00:00Z",
            recurrence="weekly",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert patched.start.date_time == "2026-11-11T17:00:00"
        assert [d.value for d in patched.recurrence.pattern.days_of_week] == ["wednesday"]
        assert patched.recurrence.range.start_date == date(2026, 11, 11)

    async def test_a_stored_all_day_event_is_sent_as_written(self):
        builder, mock_client = self._updatable(self.LA, is_all_day=True)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T00:00:00Z",
            end="2026-12-14T00:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args[0][0]
        assert (patched.start.date_time, patched.start.time_zone) == ("2026-12-13T00:00:00Z", "UTC")

    async def test_a_typed_instant_is_not_reported(self):
        """Only what the server worked out is echoed; the caller typed this moment."""
        builder, mock_client = self._updatable(self.LA)

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert result == {"status": "updated", "event_id": "AAMkAG123="}


def _read_back(start_utc: str | None, end_utc: str | None, *, zone: str | None = "UTC"):
    """An event's times as the check reads them: `GET ...?$select=start,end`, `Prefer` UTC.

    Seven fractional digits, Graph's own spelling. `None` leaves that end out.
    """
    event = MagicMock(id="AAMkAG123=")
    event.start = (
        None if start_utc is None else MagicMock(date_time=f"{start_utc}.0000000", time_zone=zone)
    )
    event.end = (
        None if end_utc is None else MagicMock(date_time=f"{end_utc}.0000000", time_zone=zone)
    )
    return event


#: The moments most tests below type, 02:00Z to 03:00Z, as Graph reads them back.
_AS_TYPED = ("2026-12-13T02:00:00", "2026-12-13T03:00:00")
#: The same two an hour late: what a disagreement over a zone's rules looks like.
_AN_HOUR_LATE = ("2026-12-13T03:00:00", "2026-12-13T04:00:00")


def _write_reply(
    start: str = "2026-12-12T18:00:00",
    end: str = "2026-12-12T19:00:00",
    zone: str = "America/Los_Angeles",
):
    """A write's reply as Graph sent it live on 2026-11-25: in the event's own zone.

    `Prefer: outlook.timezone="UTC"` on the POST did not change that. 18:00 Los
    Angeles is 02:00Z, so by default this is the typed moment in a zone the
    check must not read.
    """
    reply = MagicMock(id="AAMkAG123=", subject="Dinner")
    reply.start = MagicMock(date_time=f"{start}.0000000", time_zone=zone)
    reply.end = MagicMock(date_time=f"{end}.0000000", time_zone=zone)
    return reply


def _prefer(call) -> set:
    """The `Prefer` values a mocked SDK call was sent with."""
    config = call.kwargs.get("request_configuration")
    return set((config.headers.get("Prefer") if config is not None else None) or set())


def _is_read_back(config) -> bool:
    """Whether a mocked `.get()` was the check's read: `$select=start,end`."""
    query = getattr(config, "query_parameters", None)
    return getattr(query, "select", None) == ["start", "end"]


def _read_backs(builder) -> list:
    """The check's reads a mocked request builder was sent, in order."""
    return [
        c
        for c in builder.get.call_args_list
        if _is_read_back(c.kwargs.get("request_configuration"))
    ]


def _requests(builder) -> list[str]:
    """What a mocked request builder was sent, in order; `read` is the check's GET."""
    sent = []
    for name, _, kwargs in builder.method_calls:
        if name == "get":
            sent.append("read" if _is_read_back(kwargs.get("request_configuration")) else "get")
        elif name == "patch":
            sent.append("patch")
    return sent


def _answering(current, *read_backs) -> AsyncMock:
    """A `.get()` answering a plain read with `current` and the check's reads in turn.

    An exception among `read_backs` is raised for that read. Past the last one,
    the check gets an event with no times, which it cannot read.
    """
    queue = list(read_backs)

    async def get(request_configuration=None):
        if not _is_read_back(request_configuration):
            return current
        answer = queue.pop(0) if queue else _read_back(None, None)
        if isinstance(answer, BaseException):
            raise answer
        return answer

    return AsyncMock(side_effect=get)


def _graph_error(status: int):
    from kiota_abstractions.api_error import APIError

    return APIError(f"Graph answered {status}", response_status_code=status)


class TestAConvertedTimeIsCheckedAgainstGraph:
    """A converted `Z`/offset time is checked against where Graph put it (#262 review).

    The conversion uses this host's zone rules (IANA, through the Windows table
    for a Windows name); Graph reads the local time back with its own. Where the
    two disagree for a zone and date, the event would land at another moment and
    the tool would still say `created`/`updated`. So a write that converted reads
    the event's times back in UTC, and when one is elsewhere, the typed values
    are sent again as written: the moment is right, the zone is lost, and the
    result says so.

    Read with a GET, not off the write's reply. Graph ignores `Prefer:
    outlook.timezone` on a write's reply and answers in the event's own zone
    (measured live, 2026-11-25), so a check that read the reply never ran. The
    mocked write replies below have that shape, or one that would mislead.
    """

    UTC = 'outlook.timezone="UTC"'

    @staticmethod
    def _updatable(zone: str, *read_backs, reply=None, **current):
        builder = _make_event_builder()
        builder.get = _answering(_current_event(start_zone=zone, **current), *read_backs)
        builder.patch = AsyncMock(return_value=_write_reply() if reply is None else reply)
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)
        return builder, mock_client

    # ── update ────────────────────────────────────────────────────────

    @pytest.mark.parametrize("zone", ["America/Los_Angeles", "Pacific Standard Time"])
    async def test_a_read_back_at_the_typed_moment_needs_nothing_more(self, zone):
        builder, mock_client = self._updatable(zone, _read_back(*_AS_TYPED))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        # The plain read for the event's zone, the write, and one read back.
        assert _requests(builder) == ["get", "patch", "read"]
        assert result == {"status": "updated", "event_id": "AAMkAG123="}

    async def test_the_read_back_is_the_two_times_in_utc_and_the_patch_asks_nothing(self):
        """`GET /me/events/{id}?$select=start,end` with `Prefer: outlook.timezone="UTC"`.

        Graph honours that header on a GET: the read tools' UTC times rest on
        it. On the PATCH it did nothing, so the PATCH no longer carries it.
        """
        builder, mock_client = self._updatable("America/Los_Angeles", _read_back(*_AS_TYPED))

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert {c.args for c in mock_client.me.events.by_event_id.call_args_list} == {
            ("AAMkAG123=",)
        }
        (read,) = _read_backs(builder)
        config = read.kwargs["request_configuration"]
        assert config.query_parameters.select == ["start", "end"]
        assert _prefer(read) == {self.UTC}
        assert _prefer(builder.patch.call_args) == set()

    @pytest.mark.parametrize(
        "reply",
        [
            _write_reply(),
            _write_reply(*_AN_HOUR_LATE, zone="UTC"),
        ],
        ids=["in the event's own zone, as live", "an hour late in UTC"],
    )
    async def test_the_write_reply_decides_nothing(self, reply):
        """The read back is at the typed moment, so nothing moved, whatever the reply says."""
        builder, mock_client = self._updatable(
            "America/Los_Angeles", _read_back(*_AS_TYPED), reply=reply
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read"]
        assert "note" not in result

    @pytest.mark.parametrize(
        ("zone", "label"),
        [("America/Havana", "America/Havana"), ("Cuba Standard Time", "Cuba Standard Time")],
    )
    async def test_a_read_back_an_hour_off_is_put_back_as_typed(self, zone, label):
        """Graph's rules for the zone are an hour from this host's: the event landed late.

        The write's reply is at the typed moment, in UTC: the read back decides.
        """
        builder, mock_client = self._updatable(
            zone,
            _read_back(*_AN_HOUR_LATE),
            _read_back(*_AS_TYPED),
            reply=_write_reply(*_AS_TYPED, zone="UTC"),
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-12T22:00:00-05:00",
            config=_CFG,
        )

        # Read back, put back, and read back once more to confirm it held.
        assert _requests(builder) == ["get", "patch", "read", "patch", "read"]
        first, fix = (call.args[0] for call in builder.patch.call_args_list)
        assert not first.start.date_time.endswith("Z")
        # The corrective patch carries the times exactly as typed, beside the same
        # zone name, and nothing else.
        assert (fix.start.date_time, fix.start.time_zone) == ("2026-12-13T02:00:00Z", label)
        assert (fix.end.date_time, fix.end.time_zone) == ("2026-12-12T22:00:00-05:00", label)
        assert fix.subject is None and fix.recurrence is None
        assert all(_prefer(call) == set() for call in builder.patch.call_args_list)
        assert all(_prefer(read) == {self.UTC} for read in _read_backs(builder))
        assert result["status"] == "updated"
        assert "time zone" in result["note"] and "UTC" in result["note"]

    async def test_only_the_converted_end_is_put_back(self):
        builder, mock_client = self._updatable(
            "America/Los_Angeles",
            _read_back("2026-12-13T03:00:00", None),
            _read_back("2026-12-13T02:00:00", None),
        )

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start="2026-12-13T02:00:00Z", config=_CFG
        )

        fix = builder.patch.call_args_list[1].args[0]
        assert fix.start.date_time == "2026-12-13T02:00:00Z"
        assert fix.end is None
        assert "note" in result

    @pytest.mark.parametrize(
        "read_back",
        [
            _read_back(None, None),
            _read_back(*_AN_HOUR_LATE, zone="Pacific Standard Time"),
            MagicMock(id="AAMkAG123="),
            None,
        ],
        ids=["no start or end", "not in UTC", "no usable fields", "no event"],
    )
    async def test_a_read_back_that_cannot_say_is_not_guessed_at(self, read_back):
        """Fails open, as before the check: no correction and no `note`."""
        builder, mock_client = self._updatable("America/Los_Angeles", read_back)

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read"]
        assert result == {"status": "updated", "event_id": "AAMkAG123="}

    @pytest.mark.parametrize(
        "failure",
        [
            _graph_error(404),
            _graph_error(429),
            _graph_error(503),
            httpx.ConnectError("connection refused"),
            httpx.ReadTimeout("timed out"),
        ],
        ids=["404", "429", "503", "no connection", "timed out"],
    )
    async def test_a_read_back_that_fails_is_not_guessed_at(self, failure):
        """The write succeeded; Graph or the network failing the read is no reason to
        say otherwise, or to move it."""
        builder, mock_client = self._updatable("America/Havana", failure)

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read"]
        assert result == {"status": "updated", "event_id": "AAMkAG123="}

    async def test_a_bug_in_the_read_back_is_not_taken_for_graph_failing(self):
        """Only Graph's and the connection's failures fail open; anything else is ours."""
        builder, mock_client = self._updatable("America/Havana", TypeError("a bug"))

        with pytest.raises(TypeError, match="a bug"):
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-12-13T02:00:00Z",
                end="2026-12-13T03:00:00Z",
                config=_CFG,
            )

    async def test_the_read_back_is_sent_once_with_kiotas_retry_off(self):
        """The write is done: a throttled read retried after `Retry-After` would hold
        back its answer for minutes, and a client that gives up may create it again."""
        from kiota_http.middleware.options.retry_handler_option import RetryHandlerOption

        builder, mock_client = self._updatable(
            "America/Havana", _read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED)
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        reads = _read_backs(builder)
        assert len(reads) == 2  # the check's read and the one after the correction
        for read in reads:
            options = list(read.kwargs["request_configuration"].options)
            retry = [o for o in options if isinstance(o, RetryHandlerOption)]
            assert len(retry) == 1, options
            assert retry[0].should_retry is False

    # ── a read that still shows the event as it was says nothing ──────

    @staticmethod
    def _stored(start_utc: str, end_utc: str | None = None):
        """A single Havana event as a plain GET returns it, both ends in UTC."""
        current = _current_event(date_time=f"{start_utc}.0000000", start_zone="America/Havana")
        current.end = (
            None if end_utc is None else MagicMock(date_time=f"{end_utc}.0000000", time_zone="UTC")
        )
        return current

    @pytest.mark.parametrize(
        ("stored", "start", "end"),
        [
            # An hour later, start only: what a zone rule's gap also looks like.
            (("2026-12-13T01:00:00", "2026-12-13T03:00:00"), "2026-12-13T02:00:00Z", None),
            # Half an hour earlier.
            (("2026-12-13T02:30:00", "2026-12-13T03:00:00"), "2026-12-13T02:00:00Z", None),
            # Three hours later, both ends.
            (
                ("2026-12-12T23:00:00", "2026-12-13T00:00:00"),
                "2026-12-13T02:00:00Z",
                "2026-12-13T03:00:00Z",
            ),
            # Only the start moved; the end is sent again unchanged.
            (
                ("2026-12-13T01:00:00", "2026-12-13T03:00:00"),
                "2026-12-13T02:00:00Z",
                "2026-12-13T03:00:00Z",
            ),
        ],
        ids=["an hour later", "half an hour earlier", "both ends 3h", "start only of two"],
    )
    async def test_a_read_back_still_at_the_old_time_is_left_alone(self, stored, start, end):
        """A plain move whose read-back echoes the time before it: not "corrected" (S1).

        Taken at its word, the gap is the move itself, which passes for a zone
        rule's: the event would be patched again, filed in UTC, and its
        attendees mailed twice, under a note blaming the zone's rules.
        """
        builder, mock_client = self._updatable("America/Havana")
        builder.get = _answering(self._stored(*stored), _read_back(*stored))

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG
        )

        assert _requests(builder) == ["get", "patch", "read"]
        assert result == {"status": "updated", "event_id": "AAMkAG123="}

    async def test_a_read_back_elsewhere_than_the_old_time_is_still_corrected(self):
        """The guard is the old time exactly: an hour past the typed one is a zone's doing."""
        builder, mock_client = self._updatable("America/Havana")
        builder.get = _answering(
            self._stored("2026-12-13T01:00:00", "2026-12-13T03:00:00"),
            _read_back("2026-12-13T03:00:00", "2026-12-13T03:00:00"),
            _read_back("2026-12-13T02:00:00", "2026-12-13T03:00:00"),
        )

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start="2026-12-13T02:00:00Z", config=_CFG
        )

        assert _requests(builder) == ["get", "patch", "read", "patch", "read"]
        assert "so the time is right" in result["note"]

    @pytest.mark.parametrize(
        ("start", "end"),
        [
            ("2026-12-12T18:00:00", "2026-12-12T19:00:00"),  # zone-less: nothing converted
            ("2027-11-07T07:30:00Z", "2027-11-07T08:30:00Z"),  # repeated hour: sent as typed
        ],
    )
    async def test_a_write_that_converted_nothing_reads_nothing_back(self, start, end):
        builder, mock_client = self._updatable(
            "America/Los_Angeles", _read_back("2001-01-01T00:00:00", "2001-01-01T01:00:00")
        )

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start=start, end=end, config=_CFG
        )

        assert _requests(builder) == ["get", "patch"]
        assert _prefer(builder.patch.call_args) == set()
        assert "note" not in result

    async def test_a_failed_correction_says_where_the_event_is(self):
        from outlook_mcp.errors import GraphAPIError, OutlookMCPError

        builder, mock_client = self._updatable("Cuba Standard Time", _read_back(*_AN_HOUR_LATE))
        builder.patch.side_effect = [
            _write_reply(),
            GraphAPIError(400, "ErrorInvalidRequest", "nope", action=None),
        ]

        with pytest.raises(OutlookMCPError) as exc:
            await update_event(
                mock_client,
                event_id="AAMkAG123=",
                start="2026-12-13T02:00:00Z",
                end="2026-12-13T03:00:00Z",
                config=_CFG,
            )

        message = str(exc.value)
        assert "2026-12-13T03:00:00" in message  # where Graph put it
        assert "2026-12-13T02:00:00" in message  # where it was asked to go
        assert "nope" in message
        assert "outlook_get_event" in message
        # The way to put it right names the tool that moves an event, and the
        # input that needs no conversion at all.
        assert "outlook_update_event" in message
        assert "without a zone or offset" in message
        # Nothing is read back after a correction that was refused.
        assert _requests(builder) == ["get", "patch", "read", "patch"]

    # ── after a correction, the event is read back once more ──────────

    async def test_a_correction_that_held_is_reported_as_right(self):
        builder, mock_client = self._updatable(
            "America/Havana", _read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED)
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read", "patch", "read"]
        note = result["note"]
        assert "start at 2026-12-13T03:00:00 UTC instead of 2026-12-13T02:00:00 UTC" in note
        assert "so the time is right" in note and "filed in UTC" in note

    async def test_a_correction_graph_did_not_keep_is_reported_not_claimed(self):
        """Read back still elsewhere: the `note` says where, and claims nothing."""
        builder, mock_client = self._updatable(
            "America/Havana",
            _read_back(*_AN_HOUR_LATE),
            _read_back("2026-12-13T04:00:00", "2026-12-13T05:00:00"),
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        # One correction only: a second would be a guess.
        assert _requests(builder) == ["get", "patch", "read", "patch", "read"]
        note = result["note"]
        assert "start at 2026-12-13T03:00:00 UTC instead of 2026-12-13T02:00:00 UTC" in note
        assert "saved again as typed" in note
        assert "start at 2026-12-13T04:00:00 UTC instead of 2026-12-13T02:00:00 UTC" in note
        assert "outlook_get_event" in note
        assert "the time is right" not in note

    @pytest.mark.parametrize(
        "confirmation",
        [_graph_error(503), _read_back(None, None), _read_back(*_AN_HOUR_LATE)],
        ids=["fails", "cannot say", "still as the first read found it"],
    )
    async def test_a_confirmation_that_cannot_say_claims_nothing(self, confirmation):
        """Not "the time is right", and not "still elsewhere" either.

        A read after the correction that repeats the first read's times is most
        likely the event as it was a moment before, so it is not taken for a
        correction Graph ignored; like a failed read, it confirms nothing.
        """
        builder, mock_client = self._updatable(
            "America/Havana", _read_back(*_AN_HOUR_LATE), confirmation
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read", "patch", "read"]
        note = result["note"]
        assert "start at 2026-12-13T03:00:00 UTC instead of 2026-12-13T02:00:00 UTC" in note
        assert "saved again as typed" in note and "filed in UTC" in note
        assert "did not confirm the time" in note and "outlook_get_event" in note
        assert "the time is right" not in note
        assert "reading the event back puts" not in note

    async def test_an_unexplained_difference_is_not_read_again(self):
        """Nothing was written, so there is nothing to confirm."""
        builder, mock_client = self._updatable(
            "America/Havana", _read_back("2026-12-14T02:00:00", "2026-12-14T03:00:00")
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch", "read"]
        assert "not changed" in result["note"]

    # ── only a single event is checked ────────────────────────────────

    @pytest.mark.parametrize(
        "read_back",
        [
            # Monday's first occurrence, two days on: what a weekday series may echo.
            ("2026-12-14T17:00:00", "2026-12-14T17:30:00"),
            # An hour late: what a zone-rule disagreement looks like, and what a
            # single event would be corrected for.
            ("2026-12-12T18:00:00", "2026-12-12T18:30:00"),
        ],
        ids=["first occurrence", "an hour late"],
    )
    @pytest.mark.parametrize(
        ("event_type", "checked"),
        [
            ("singleInstance", True),
            ("seriesMaster", False),
            ("occurrence", False),
            ("exception", False),
        ],
    )
    async def test_only_an_update_to_a_single_event_is_checked(
        self, event_type, checked, read_back
    ):
        """A master's times can be its first occurrence, not the start sent (S1).

        Saturday 17:00Z sent to a weekday series can come back as Monday's. That
        is no zone rule's doing, and a corrective patch would file the series in
        UTC, or be refused because it carries no recurrence. So only an event
        that stands alone is checked; the others are sent as converted, as before
        the check existed, with no read back, no second write and no `note`,
        whatever Graph would say.
        """
        from tests.test_write_payloads_reach_the_wire import wire

        # A master always carries its recurrence: a weekday series begun on
        # Saturday 2026-12-12, which the GET reports, so this patch is a change
        # within its day and sends no rule of its own (#98).
        series = (
            {
                "date_time": "2026-12-12T17:00:00.0000000",
                "recurrence": build_event_recurrence(
                    "weekdays", start="2026-12-12T09:00:00", zone="America/Los_Angeles"
                ),
            }
            if event_type == "seriesMaster"
            else {}
        )
        builder, mock_client = self._updatable(
            "America/Los_Angeles",
            _read_back(*read_back),
            _read_back("2026-12-12T17:00:00", "2026-12-12T17:30:00"),
            event_type=event_type,
            **series,
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-12T17:00:00Z",
            end="2026-12-12T17:30:00Z",
            config=_CFG,
        )

        first = builder.patch.call_args_list[0]
        body = wire(first.args[0])
        # Converted either way (#236): the check is what differs.
        assert '"dateTime": "2026-12-12T09:00:00"' in body
        assert '"timeZone": "America/Los_Angeles"' in body
        assert _prefer(first) == set()
        if not checked:
            assert _read_backs(builder) == []
            assert builder.patch.call_count == 1
            assert "note" not in result
        else:
            # An hour is put back and read again; two days is no zone rule's
            # doing, so it is only reported.
            corrected = read_back[0].endswith("T18:00:00")
            assert len(_read_backs(builder)) == (2 if corrected else 1)
            assert builder.patch.call_count == (2 if corrected else 1)
            assert "note" in result

    async def test_a_single_event_typed_by_the_sdk_enum_is_checked(self):
        """Graph's reply is parsed into `EventType`, not a string; both must read as single."""
        from msgraph.generated.models.event_type import EventType

        builder, mock_client = self._updatable("America/Havana")
        current = _current_event(start_zone="America/Havana")
        current.type = EventType.SingleInstance
        builder.get = _answering(current, _read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert builder.patch.call_count == 2
        assert "note" in result

    async def test_an_event_whose_type_cannot_be_read_is_not_checked(self):
        builder, mock_client = self._updatable("America/Havana")
        current = _current_event(start_zone="America/Havana")
        current.type = None
        builder.get = _answering(current, _read_back(*_AN_HOUR_LATE))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert _requests(builder) == ["get", "patch"]
        assert _prefer(builder.patch.call_args) == set()
        assert "note" not in result

    async def test_an_update_that_removes_a_recurrence_is_not_checked(self):
        """`remove_recurrence` sends an explicit null recurrence; that request is not checked."""
        builder, mock_client = self._updatable("America/Havana", _read_back(*_AN_HOUR_LATE))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            remove_recurrence=True,
            config=_CFG,
        )

        patched = builder.patch.call_args
        assert patched.args[0].additional_data == {"recurrence": None}
        assert _requests(builder) == ["get", "patch"]
        assert _prefer(patched) == set()
        assert "note" not in result

    async def test_an_update_that_makes_a_series_is_not_checked(self):
        """A single event given a recurrence becomes a master in the same request."""
        builder, mock_client = self._updatable(
            "America/Los_Angeles", _read_back("2026-12-14T17:00:00", "2026-12-14T17:30:00")
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-12T17:00:00Z",
            end="2026-12-12T17:30:00Z",
            recurrence="weekdays",
            config=_CFG,
        )

        patched = builder.patch.call_args
        assert patched.args[0].recurrence is not None
        assert _read_backs(builder) == []
        assert _prefer(patched) == set()
        assert builder.patch.call_count == 1
        assert "note" not in result

    # ── only a difference a zone rule can make is corrected ───────────

    @pytest.mark.parametrize(
        ("start_utc", "end_utc"),
        [
            ("2026-12-14T02:00:00", "2026-12-14T03:00:00"),  # a whole day
            ("2026-12-13T05:00:01", "2026-12-13T06:00:01"),  # just past three hours
            ("2026-12-13T02:00:30", "2026-12-13T03:00:30"),  # not a whole minute
            ("2026-12-13T03:00:00", "2026-12-15T03:00:00"),  # one end fits, the other does not
        ],
        ids=["a day", "over three hours", "half a minute", "one end too far"],
    )
    async def test_a_difference_no_zone_rule_explains_is_reported_not_changed(
        self, start_utc, end_utc
    ):
        """Zone offsets are whole minutes, and two rule sets disagree by at most a few hours.

        Anything else did not come from the conversion, so a corrective write
        would be a guess: the event is left as Graph placed it, and the `note`
        says where that is, so the caller can check it with the user.
        """
        builder, mock_client = self._updatable("America/Havana", _read_back(start_utc, end_utc))

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert builder.patch.call_count == 1
        note = result["note"]
        assert f"start at {start_utc} UTC instead of 2026-12-13T02:00:00 UTC" in note
        assert "not changed" in note and "outlook_get_event" in note
        assert "filed in UTC" not in note

    @pytest.mark.parametrize(
        ("start_utc", "end_utc"),
        [
            ("2026-12-13T05:00:00", "2026-12-13T06:00:00"),  # three hours late
            ("2026-12-12T23:00:00", "2026-12-13T00:00:00"),  # three hours early
            ("2026-12-13T02:30:00", "2026-12-13T03:30:00"),  # Lord Howe's half hour
            ("2026-12-13T02:15:00", "2026-12-13T03:15:00"),  # a quarter-hour zone
        ],
    )
    async def test_a_difference_a_zone_rule_can_make_is_put_back(self, start_utc, end_utc):
        builder, mock_client = self._updatable(
            "America/Havana", _read_back(start_utc, end_utc), _read_back(*_AS_TYPED)
        )

        result = await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        assert builder.patch.call_count == 2
        assert "note" in result

    @pytest.mark.parametrize(
        ("typed", "read", "corrected"),
        [
            ("2026-12-13T02:00:00.900Z", "2026-12-13T02:00:00.0000000", False),
            ("2026-12-13T02:00:00.900Z", "2026-12-13T03:00:00.0000000", True),
            ("2026-12-13T02:00:00Z", "2026-12-13T03:00:00", True),
            ("2026-12-13T02:00:00Z", "2026-12-13T03:00:00Z", True),
        ],
        ids=["under a second", "an hour less a fraction", "no fraction", "trailing Z"],
    )
    async def test_a_fraction_of_a_second_and_graphs_spellings(self, typed, read, corrected):
        """Graph may drop a fraction it was sent; its times may or may not carry digits."""
        answer = MagicMock(id="AAMkAG123=")
        answer.start = MagicMock(date_time=read, time_zone="UTC")
        answer.end = None
        builder, mock_client = self._updatable("America/Havana", answer)

        result = await update_event(mock_client, event_id="AAMkAG123=", start=typed, config=_CFG)

        assert builder.patch.call_count == (2 if corrected else 1)
        assert ("note" in result) is corrected

    @pytest.mark.parametrize(
        ("typed", "read", "corrected"),
        [
            # Three hours early with the fraction dropped: 3 h 0.9 s from the
            # time typed, exactly three hours from its whole second.
            ("2026-12-13T02:00:00.900Z", "2026-12-12T23:00:00.0000000", True),
            ("2026-12-13T02:00:00.900Z", "2026-12-13T05:00:00.0000000", True),
            # Not a whole minute from either, though within a second of one.
            ("2026-12-13T02:00:00Z", "2026-12-13T02:00:59.6000000", False),
            # An hour and a fraction no dropped digit accounts for.
            ("2026-12-13T02:00:00.400Z", "2026-12-13T03:00:00.7000000", False),
        ],
        ids=["3h early", "3h late", "59.6 seconds", "an hour and a fraction"],
    )
    async def test_the_edges_are_the_same_either_side(self, typed, read, corrected):
        answer = MagicMock(id="AAMkAG123=")
        answer.start = MagicMock(date_time=read, time_zone="UTC")
        answer.end = None
        builder, mock_client = self._updatable("America/Havana", answer)

        result = await update_event(mock_client, event_id="AAMkAG123=", start=typed, config=_CFG)

        assert builder.patch.call_count == (2 if corrected else 1)
        assert ("not changed" in result["note"]) is not corrected

    async def test_the_note_names_the_end_that_moved_and_where(self):
        builder, mock_client = self._updatable(
            "America/Havana",
            _read_back("2026-12-13T03:00:00", None),
            _read_back("2026-12-13T02:00:00", None),
        )

        result = await update_event(
            mock_client, event_id="AAMkAG123=", start="2026-12-13T02:00:00Z", config=_CFG
        )

        note = result["note"]
        assert "start at 2026-12-13T03:00:00 UTC instead of 2026-12-13T02:00:00 UTC" in note
        # Only the start was sent again, so only the start is filed in UTC.
        assert "The start was saved again" in note
        assert "start and end" not in note

    async def test_the_corrective_patch_on_the_wire(self):
        """Only the times, as typed, beside each end's own zone; no `Prefer`, no pin."""
        import json

        from tests.test_write_payloads_reach_the_wire import wire

        builder, mock_client = self._updatable(
            "America/Havana", _read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED)
        )

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        fix = builder.patch.call_args_list[1]
        # `wire` marks the model clean, so it is serialized once and read twice.
        body = wire(fix.args[0])
        assert json.loads(body) == {
            "@odata.type": "#microsoft.graph.event",
            "start": {"dateTime": "2026-12-13T02:00:00Z", "timeZone": "America/Havana"},
            "end": {"dateTime": "2026-12-13T03:00:00Z", "timeZone": "America/Havana"},
        }
        headers = fix.kwargs["request_configuration"].headers
        assert not headers.get("Prefer") and not headers.get("If-Match")

    # ── create ────────────────────────────────────────────────────────

    @staticmethod
    def _poster(*read_backs, reply=None):
        mock_client = MagicMock()
        mock_client.me.events.post = AsyncMock(
            return_value=_write_reply() if reply is None else reply
        )
        fixer = MagicMock()
        fixer.get = _answering(None, *read_backs)
        fixer.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client.me.events.by_event_id = MagicMock(return_value=fixer)
        return mock_client, fixer

    @pytest.mark.parametrize(
        ("read_backs", "corrected"),
        [
            ((_read_back(*_AS_TYPED),), False),
            ((_read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED)), True),
        ],
        ids=["kept", "moved"],
    )
    async def test_create_reads_the_event_back_and_puts_a_moved_one_back(
        self, read_backs, corrected
    ):
        # The POST's reply is in the event's own zone, as live: it decides nothing.
        mock_client, fixer = self._poster(*read_backs)

        result = await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            timezone="America/Havana",
            config=_CFG,
        )

        assert _prefer(mock_client.me.events.post.call_args) == set()
        assert {c.args for c in mock_client.me.events.by_event_id.call_args_list} == {
            ("AAMkAG123=",)
        }
        assert all(_prefer(read) == {self.UTC} for read in _read_backs(fixer))
        assert _requests(fixer) == (["read", "patch", "read"] if corrected else ["read"])
        assert ("note" in result) is corrected
        if corrected:
            fix = fixer.patch.call_args.args[0]
            assert (fix.start.date_time, fix.end.date_time) == (
                "2026-12-13T02:00:00Z",
                "2026-12-13T03:00:00Z",
            )
            assert fix.start.time_zone == "America/Havana"
            assert _prefer(fixer.patch.call_args) == set()

    async def test_create_without_a_typed_instant_reads_nothing_back(self):
        mock_client, fixer = self._poster(_read_back("2001-01-01T00:00:00", "2001-01-01T01:00:00"))

        await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-12T18:00:00",
            end="2026-12-12T19:00:00",
            timezone="America/Los_Angeles",
            config=_CFG,
        )

        assert _prefer(mock_client.me.events.post.call_args) == set()
        mock_client.me.events.by_event_id.assert_not_called()
        assert _requests(fixer) == []

    @pytest.mark.parametrize(
        "failure",
        [_graph_error(404), _graph_error(503), httpx.ConnectError("connection refused")],
        ids=["404", "503", "no connection"],
    )
    async def test_a_create_whose_read_back_fails_is_still_created(self, failure):
        """The event exists: a failed read neither moves it nor reports an error."""
        mock_client, fixer = self._poster(failure)

        result = await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            timezone="America/Havana",
            config=_CFG,
        )

        assert result == {"status": "created", "event_id": "AAMkAG123=", "subject": "Dinner"}
        assert _requests(fixer) == ["read"]

    async def test_a_create_reply_without_an_id_reads_nothing_back(self):
        """The SDK builds no request for a missing id; the event exists either way."""
        reply = _write_reply()
        reply.id = None
        mock_client, fixer = self._poster(_read_back(*_AN_HOUR_LATE), reply=reply)

        result = await create_event(
            mock_client,
            subject="Dinner",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            timezone="America/Havana",
            config=_CFG,
        )

        mock_client.me.events.by_event_id.assert_not_called()
        assert result == {"status": "created", "event_id": None, "subject": "Dinner"}

    @pytest.mark.parametrize(
        "read_back",
        [
            ("2026-12-14T17:00:00", "2026-12-14T17:30:00"),
            ("2026-12-12T18:00:00", "2026-12-12T18:30:00"),
        ],
        ids=["first occurrence", "an hour late"],
    )
    @pytest.mark.parametrize(
        "recurrence",
        [
            "weekdays",
            {
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]},
                "range": {"type": "numbered", "numberOfOccurrences": 2},
            },
        ],
        ids=["weekdays", "thursdays"],
    )
    async def test_a_series_create_is_not_checked(self, recurrence, read_back):
        """Saturday 17:00Z for a series that never meets on a Saturday (S1).

        Graph accepts a start off the pattern and puts the occurrences on the
        pattern's days, so its times may be the first occurrence (Monday or
        Thursday), days from the start typed. Corrected, the series would be
        filed in UTC, or refused because the patch carries no recurrence. The
        series is converted as before (#236) and not read back, even when Graph
        would say an hour off, which on a single event would be corrected.
        """
        from tests.test_write_payloads_reach_the_wire import wire

        mock_client, fixer = self._poster(_read_back(*read_back))

        result = await create_event(
            mock_client,
            subject="Standup",
            start="2026-12-12T17:00:00Z",
            end="2026-12-12T17:30:00Z",
            timezone="America/Los_Angeles",
            recurrence=recurrence,
            config=_CFG,
        )

        post = mock_client.me.events.post.call_args
        body = wire(post.args[0])
        assert '"dateTime": "2026-12-12T09:00:00"' in body
        assert '"timeZone": "America/Los_Angeles"' in body
        assert '"recurrence"' in body
        assert _prefer(post) == set()
        mock_client.me.events.by_event_id.assert_not_called()
        assert _requests(fixer) == []
        assert "note" not in result

    async def test_a_failed_correction_on_create_names_the_new_event(self):
        """The event exists; creating it again would make a second one and mail it again."""
        from outlook_mcp.errors import GraphAPIError, OutlookMCPError

        mock_client, fixer = self._poster(_read_back(*_AN_HOUR_LATE))
        fixer.patch.side_effect = GraphAPIError(400, "ErrorInvalidRequest", "nope", action=None)

        with pytest.raises(OutlookMCPError) as exc:
            await create_event(
                mock_client,
                subject="Dinner",
                start="2026-12-13T02:00:00Z",
                end="2026-12-13T03:00:00Z",
                timezone="America/Havana",
                config=_CFG,
            )

        message = str(exc.value)
        assert exc.value.code == "event_time_not_restored"
        assert "created (id AAMkAG123=)" in message
        assert "outlook_update_event" in message
        assert "not by creating it again" in message

    async def test_create_in_a_named_calendar_reads_and_corrects_it_there(self):
        """The read back and the corrective patch go to the collection it was posted to."""
        client, secondary = _client_with_calendars()
        secondary.events.post = AsyncMock(return_value=_write_reply())
        fixer = MagicMock()
        fixer.get = _answering(None, _read_back(*_AN_HOUR_LATE), _read_back(*_AS_TYPED))
        fixer.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        secondary.events.by_event_id = MagicMock(return_value=fixer)

        result = await create_event(
            client,
            subject="Dentist",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            timezone="America/Havana",
            calendar="Kids + School",
            config=_CFG,
        )

        client.me.calendars.by_calendar_id.assert_called_with("KIDS77=")
        assert {c.args for c in secondary.events.by_event_id.call_args_list} == {("AAMkAG123=",)}
        client.me.events.by_event_id.assert_not_called()
        assert _requests(fixer) == ["read", "patch", "read"]
        assert "note" in result


class TestWindowsNameLookupIsExact:
    """The table is read with the name exactly as Graph gave it (#262 review).

    Graph returns its own spelling; a case-folded or trimmed match would convert
    into a zone for a name Graph never used. So anything else falls back and is
    sent as written, the behaviour before the table.
    """

    @pytest.mark.parametrize(
        "zone",
        ["pacific standard time", " Pacific Standard Time ", "Mid-Atlantic Standard Time"],
    )
    async def test_a_name_not_spelled_as_graph_spells_it_is_sent_as_written(self, zone):
        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=_current_event(start_zone=zone))
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client,
            event_id="AAMkAG123=",
            start="2026-12-13T02:00:00Z",
            end="2026-12-13T03:00:00Z",
            config=_CFG,
        )

        patched = builder.patch.call_args.args[0]
        assert (patched.start.date_time, patched.start.time_zone) == ("2026-12-13T02:00:00Z", zone)

    @pytest.mark.parametrize(
        ("zone", "local"),
        [
            ("UTC", "2026-12-13T02:00:00"),
            ("India Standard Time", "2026-12-13T07:30:00"),
            ("AUS Eastern Standard Time", "2026-12-13T13:00:00"),
        ],
    )
    async def test_a_pinned_local_time_for_zones_outside_the_us(self, zone, local):
        builder = _make_event_builder()
        builder.get = AsyncMock(return_value=_current_event(start_zone=zone))
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        await update_event(
            mock_client, event_id="AAMkAG123=", start="2026-12-13T02:00:00Z", config=_CFG
        )

        assert builder.patch.call_args.args[0].start.date_time == local
