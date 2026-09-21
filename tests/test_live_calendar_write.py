"""Write-tier live guards for recurring events (#41).

Why this tier exists: the mock suite asserts what we *build*. It cannot see
Graph reject what we built. A `PatternedRecurrence` that is perfectly
well-formed to the SDK still returns ErrorInvalidRecurrenceRange if the range
disagrees with the series master's start — exactly the failure mode this fix
had to get right — and 609 green mock tests would not notice.

Read `tests/conftest.py` before adding anything here: no attendees, no
unbounded ranges, calendar only, and everything created is deleted in a
`finally`.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
"""

from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta
from datetime import time as dt_time
from zoneinfo import ZoneInfo

import pytest

from outlook_mcp.tools.calendar_read import get_event, list_events
from outlook_mcp.tools.calendar_write import create_event, delete_event, update_event
from tests.conftest import LIVE_WRITE_SUBJECT

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]


def _anchor_monday() -> date:
    """A Monday about a month out — far enough not to clutter the working week."""
    d = date.today() + timedelta(days=30)
    return d + timedelta(days=(0 - d.weekday()) % 7)


@asynccontextmanager
async def _temporary_event(client, config, subject_suffix: str = "", **kwargs):
    """Create an event, hand back its id, and always delete it.

    ``subject_suffix`` narrows the marker for a test that has to find its own
    event again in a calendar listing; the shared prefix still makes anything a
    crash leaks greppable in the UI.
    """
    created = await create_event(
        client.sdk_client,
        subject=LIVE_WRITE_SUBJECT + subject_suffix,
        config=config,
        **kwargs,
    )
    event_id = created["event_id"]
    try:
        yield event_id
    finally:
        await delete_event(client.sdk_client, event_id, config=config)


class TestRecurringSeriesAreAccepted:
    async def test_dict_recurrence_creates_a_series_master(
        self, real_graph_client, live_write_config
    ):
        """#41's payload shape, bounded: Graph stores a real series, not an occurrence."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T09:00:00Z",
            end=f"{monday.isoformat()}T09:30:00Z",
            recurrence={
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {"type": "numbered", "numberOfOccurrences": 2},
            },
        ) as event_id:
            detail = await get_event(real_graph_client.sdk_client, event_id)
            # The bug: this came back "singleInstance" with recurrence None.
            assert detail["type"] == "seriesMaster"
            assert detail["recurrence"] is not None
            assert detail["recurrence"]["pattern"]["type"] == "weekly"
            assert detail["recurrence"]["pattern"]["daysOfWeek"] == ["monday"]
            assert detail["recurrence"]["range"]["type"] == "numbered"
            assert detail["recurrence"]["range"]["numberOfOccurrences"] == 2

    async def test_start_date_is_defaulted_to_something_graph_accepts(
        self, real_graph_client, live_write_config
    ):
        """We omit range.startDate and fill it in ourselves; prove Graph agrees."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T11:00:00Z",
            end=f"{monday.isoformat()}T11:30:00Z",
            recurrence={
                "pattern": {"type": "daily", "interval": 1},
                "range": {"type": "numbered", "numberOfOccurrences": 2},
            },
        ) as event_id:
            detail = await get_event(real_graph_client.sdk_client, event_id)
            assert detail["type"] == "seriesMaster"
            assert detail["recurrence"]["range"]["startDate"] == monday.isoformat()

    async def test_shorthand_creates_a_series(self, real_graph_client, live_write_config):
        """The documented "weekly" shorthand — silently a no-op before this fix.

        Shorthands expand to an open-ended range, which this tier bans, so the
        pattern is checked here and the range is deliberately overridden to a
        bounded one by passing the expanded object instead. This asserts the
        expansion Graph accepts, not the noEnd range.
        """
        monday = _anchor_monday()
        from outlook_mcp.tools._recurrence import build_event_recurrence, serialize_recurrence

        expanded = serialize_recurrence(
            build_event_recurrence("weekly", start=f"{monday.isoformat()}T13:00:00Z")
        )
        expanded["range"] = {"type": "numbered", "numberOfOccurrences": 2}

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T13:00:00Z",
            end=f"{monday.isoformat()}T13:30:00Z",
            recurrence=expanded,
        ) as event_id:
            detail = await get_event(real_graph_client.sdk_client, event_id)
            assert detail["type"] == "seriesMaster"
            assert detail["recurrence"]["pattern"]["daysOfWeek"] == ["monday"]


class TestNonRecurringIsUnchanged:
    async def test_plain_event_is_a_single_instance(self, real_graph_client, live_write_config):
        """Control: `type` discriminates, so the assertions above mean something."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T15:00:00Z",
            end=f"{monday.isoformat()}T15:30:00Z",
        ) as event_id:
            detail = await get_event(real_graph_client.sdk_client, event_id)
            assert detail["type"] == "singleInstance"
            assert detail["recurrence"] is None


class TestGraphRejectsAMismatchedRange:
    async def test_conflicting_start_date_never_reaches_graph(
        self, real_graph_client, live_write_config
    ):
        """We reject this locally; the point is that the local rule matches Graph's."""
        monday = _anchor_monday()

        with pytest.raises(ValueError, match="startDate"):
            await create_event(
                real_graph_client.sdk_client,
                subject=LIVE_WRITE_SUBJECT,
                start=f"{monday.isoformat()}T17:00:00Z",
                end=f"{monday.isoformat()}T17:30:00Z",
                recurrence={
                    "pattern": {"type": "daily", "interval": 1},
                    "range": {
                        "type": "numbered",
                        "numberOfOccurrences": 2,
                        "startDate": (monday + timedelta(days=3)).isoformat(),
                    },
                },
                config=live_write_config,
            )


class TestUpdatingIntoASeries:
    async def test_update_converts_a_single_event_into_a_series(
        self, real_graph_client, live_write_config
    ):
        """The #41 follow-on: there was no path to add recurrence to an existing event."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T19:00:00Z",
            end=f"{monday.isoformat()}T19:30:00Z",
        ) as event_id:
            before = await get_event(real_graph_client.sdk_client, event_id)
            assert before["type"] == "singleInstance"

            await update_event(
                real_graph_client.sdk_client,
                event_id=event_id,
                recurrence={
                    "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                    "range": {"type": "numbered", "numberOfOccurrences": 2},
                },
                config=live_write_config,
            )

            after = await get_event(real_graph_client.sdk_client, event_id)
            assert after["type"] == "seriesMaster"
            assert after["recurrence"]["pattern"]["daysOfWeek"] == ["monday"]
            # The anchor was read off the event itself — no `start` was passed.
            assert after["recurrence"]["range"]["startDate"] == monday.isoformat()


class TestPatchingEventFlags:
    """`attendees` is deliberately NOT exercised here.

    Patching it makes Outlook email invitations to everyone on the list and
    cancellations to anyone dropped. The no-attendees rule for this tier
    (see tests/conftest.py) exists precisely for that, and it outranks the
    coverage. The attendee path is unit-tested against the built payload; its
    Graph behavior is documented in the tool docstring, not asserted here.

    `is_online` is absent because Graph ignores isOnlineMeeting on consumer
    mailboxes — see test_online_meeting_is_not_supported_on_personal_accounts.
    """

    async def test_update_can_make_a_midnight_event_all_day(
        self, real_graph_client, live_write_config
    ):
        """Graph needs the bounds resent with isAllDay; this proves the rule we enforce."""
        monday = _anchor_monday()
        tuesday = monday + timedelta(days=1)

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T00:00:00Z",
            end=f"{tuesday.isoformat()}T00:00:00Z",
        ) as event_id:
            await update_event(
                real_graph_client.sdk_client,
                event_id=event_id,
                start=f"{monday.isoformat()}T00:00:00Z",
                end=f"{tuesday.isoformat()}T00:00:00Z",
                is_all_day=True,
                config=live_write_config,
            )

            assert (await get_event(real_graph_client.sdk_client, event_id))["is_all_day"] is True

    async def test_subject_edit_leaves_other_fields_alone(
        self, real_graph_client, live_write_config
    ):
        """A partial patch must not blank what it didn't mention."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T22:00:00Z",
            end=f"{monday.isoformat()}T22:30:00Z",
            location="Room 101",
        ) as event_id:
            await update_event(
                real_graph_client.sdk_client,
                event_id=event_id,
                subject=f"{LIVE_WRITE_SUBJECT} (renamed)",
                config=live_write_config,
            )

            after = await get_event(real_graph_client.sdk_client, event_id)
            assert after["location"] == "Room 101"
            assert after["is_all_day"] is False

    async def test_online_meeting_is_not_supported_on_personal_accounts(
        self, real_graph_client, live_write_config, consumer_mailbox_only
    ):
        """Pins the reason `is_online` is absent from update_event.

        Graph accepts isOnlineMeeting on a consumer mailbox and silently drops
        it. If Microsoft ever starts honouring it, this test fails and tells us
        the parameter is worth adding.

        ``consumer_mailbox_only`` is load-bearing, not decoration. The assertion
        is a claim about personal accounts and the name says so, but nothing
        checked it until 2026-09-17 — so on a work or school account, where
        Graph really does honour isOnlineMeeting, this failed correctly and
        meant nothing. It did that on two contributor PRs before anyone ran it
        against a consumer mailbox and found it green.
        """
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T23:00:00Z",
            end=f"{monday.isoformat()}T23:30:00Z",
            is_online=True,
        ) as event_id:
            assert (await get_event(real_graph_client.sdk_client, event_id))["is_online"] is False


class TestRemovingRecurrence:
    async def test_remove_recurrence_turns_a_series_back_into_one_event(
        self, real_graph_client, live_write_config
    ):
        """Ending a series without deleting it — previously impossible through this server.

        This is the assertion mocks cannot make: the SDK omits a field set to
        None, so the payload that actually reaches Graph is the whole question.
        """
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            start=f"{monday.isoformat()}T16:00:00Z",
            end=f"{monday.isoformat()}T16:30:00Z",
            recurrence={
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {"type": "numbered", "numberOfOccurrences": 3},
            },
        ) as event_id:
            assert (await get_event(real_graph_client.sdk_client, event_id))[
                "type"
            ] == "seriesMaster"

            await update_event(
                real_graph_client.sdk_client,
                event_id=event_id,
                remove_recurrence=True,
                config=live_write_config,
            )

            after = await get_event(real_graph_client.sdk_client, event_id)
            assert after["type"] == "singleInstance"
            assert after["recurrence"] is None
            # The first occurrence's time survives the conversion.
            assert "16:00:00" in after["start"]


# ── Time zone anchoring ──────────────────────────────────────────────────
#
# A zone name is server-side behaviour end to end: Graph decides what a series
# means, and the mock suite can only assert the string we put in `timeZone`.
# It was the literal "UTC" for every event this tool ever created, which is
# self-consistent, well-formed, accepted, and wrong — the series drifts an hour
# the week daylight saving ends and reports `status: created` throughout.

_DST_ZONE = "America/Los_Angeles"


def _next_dst_transition(zone_name: str) -> date | None:
    """The next date ``zone_name``'s UTC offset changes, searched from a week out.

    Computed rather than hardcoded: a fixed 2026-11-01 would quietly stop
    testing anything the moment it fell into the past, which is the shape of
    live test that reads as covered and proves nothing.
    """
    zone = ZoneInfo(zone_name)

    def offset_on(day: date):
        return datetime.combine(day, dt_time(12, 0), tzinfo=zone).utcoffset()

    cursor = date.today() + timedelta(days=7)
    previous = offset_on(cursor)
    for _ in range(400):
        cursor += timedelta(days=1)
        current = offset_on(cursor)
        if current != previous:
            return cursor
        previous = current
    return None


def _utc_instant(summary_start: str) -> str:
    """The ``HH:MM:SS`` of a listing's start, which Graph returns in UTC.

    The UTC instant is the only thing that pins this. Comparing local wall
    clocks cannot: PEP 495 has intra-zone comparison ignore ``fold``, so two
    moments an hour apart in UTC can compare equal in their own zone.
    """
    date_time, _, zone = summary_start.partition(" (")
    assert zone.rstrip(")") == "UTC", f"listing stopped returning UTC: {summary_start}"
    return date_time[11:19]


class TestTimeZoneAnchoring:
    async def test_the_transition_dates_really_do_straddle_a_dst_change(self):
        """Guard the premise, so the real assertion cannot quietly stop proving it.

        If tzdata ever drops daylight saving for this zone — a live proposal in
        more than one jurisdiction — the occurrences below stay an hour apart
        in UTC for a reason that has nothing to do with our anchoring, and the
        test passes while testing nothing.
        """
        transition = _next_dst_transition(_DST_ZONE)
        if transition is None:
            pytest.skip(
                f"{_DST_ZONE} has no UTC offset change in the next 400 days — this "
                f"zone no longer observes daylight saving, so nothing here can "
                f"distinguish a zone anchor from a UTC one"
            )

        zone = ZoneInfo(_DST_ZONE)
        before = datetime.combine(transition - timedelta(days=3), dt_time(9, 0), tzinfo=zone)
        after = datetime.combine(transition + timedelta(days=4), dt_time(9, 0), tzinfo=zone)
        assert before.utcoffset() != after.utcoffset()

    async def test_a_series_holds_its_local_hour_across_a_dst_change(
        self, real_graph_client, live_write_config
    ):
        """The bug, stated as an assertion: 09:00 stays 09:00 in the named zone.

        Anchored in UTC — what this tool sent unconditionally before the fix —
        both occurrences land on the *same* UTC instant, which means the second
        one moved an hour in the only terms the user cares about. Anchored in
        the zone, the UTC instants differ by exactly the offset change and the
        local hour holds.

        Note what is asserted: UTC instants, not local wall clocks. Comparing
        the latter would pass either way.
        """
        transition = _next_dst_transition(_DST_ZONE)
        if transition is None:
            pytest.skip(f"{_DST_ZONE} no longer changes its UTC offset; nothing to straddle")

        first = transition - timedelta(days=3)
        second = first + timedelta(days=7)
        assert second > transition, "the second occurrence must land after the change"

        suffix = " tz-anchor"
        async with _temporary_event(
            real_graph_client,
            live_write_config,
            subject_suffix=suffix,
            # Deliberately zone-less: the zone argument is what gives it meaning.
            start=f"{first.isoformat()}T09:00:00",
            end=f"{first.isoformat()}T09:30:00",
            timezone=_DST_ZONE,
            recurrence={
                "pattern": {"type": "daily", "interval": 7},
                "range": {"type": "numbered", "numberOfOccurrences": 2},
            },
        ):
            listing = await list_events(
                real_graph_client.sdk_client,
                after=f"{(first - timedelta(days=1)).isoformat()}T00:00:00Z",
                before=f"{(second + timedelta(days=1)).isoformat()}T00:00:00Z",
                count=100,
                timezone="UTC",
            )

        ours = [e for e in listing["events"] if e["subject"] == LIVE_WRITE_SUBJECT + suffix]
        assert len(ours) == 2, (
            f"expected 2 expanded occurrences, found {len(ours)} in a window of "
            f"{listing['count']} events — cannot tell a held anchor from a drifted "
            f"one without both"
        )

        zone = ZoneInfo(_DST_ZONE)
        shift = datetime.combine(second, dt_time(9, 0), tzinfo=zone).utcoffset() - datetime.combine(
            first, dt_time(9, 0), tzinfo=zone
        ).utcoffset()

        starts = [_utc_instant(e["start"]) for e in ours]
        assert starts[0] != starts[1], (
            "both occurrences are at the same UTC instant, so the series is "
            "anchored in UTC and the second one has moved an hour locally — "
            "this is the bug"
        )
        instants = [
            datetime.strptime(f"{day.isoformat()} {clock}", "%Y-%m-%d %H:%M:%S")
            for day, clock in zip((first, second), starts)
        ]
        assert (instants[1] - instants[0]) - timedelta(days=7) == -shift

    async def test_the_anchor_zone_survives_the_round_trip(
        self, real_graph_client, live_write_config
    ):
        """What the user sees: create it in a zone, read it back, it says so."""
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            subject_suffix=" tz-roundtrip",
            start=f"{monday.isoformat()}T09:00:00",
            end=f"{monday.isoformat()}T09:30:00",
            timezone=_DST_ZONE,
        ) as event_id:
            detail = await get_event(real_graph_client.sdk_client, event_id)

            assert detail["original_start_time_zone"] == _DST_ZONE
            assert detail["original_end_time_zone"] == _DST_ZONE
            # 09:00 in that zone is never 09:00 UTC, so this also proves Graph
            # read the zone-less datetime as wall-clock time in the named zone
            # rather than as UTC.
            assert "T09:00:00" not in detail["start"]

    async def test_a_utc_series_can_be_re_anchored_into_a_zone(
        self, real_graph_client, live_write_config
    ):
        """Every series this tool created before the fix is stored in UTC.

        Graph refuses a start/end patch that would change a series master's
        zone — `400 ErrorPropertyValidationFailure`, which names no zone and no
        property — unless the recurrence is re-sent with it. `update_event`
        does that silently, so the repair path works on existing data; without
        it this call is the 400.
        """
        monday = _anchor_monday()

        async with _temporary_event(
            real_graph_client,
            live_write_config,
            subject_suffix=" tz-reanchor",
            start=f"{monday.isoformat()}T16:00:00Z",
            end=f"{monday.isoformat()}T16:30:00Z",
            timezone="UTC",
            recurrence={
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {"type": "numbered", "numberOfOccurrences": 2},
            },
        ) as event_id:
            before = await get_event(real_graph_client.sdk_client, event_id)
            assert before["original_start_time_zone"] == "UTC"

            await update_event(
                real_graph_client.sdk_client,
                event_id=event_id,
                start=f"{monday.isoformat()}T09:00:00",
                end=f"{monday.isoformat()}T09:30:00",
                timezone=_DST_ZONE,
                config=live_write_config,
            )

            after = await get_event(real_graph_client.sdk_client, event_id)
            assert after["type"] == "seriesMaster", "the re-sent recurrence kept it a series"
            assert after["original_start_time_zone"] != "UTC"
            assert after["recurrence"]["range"]["numberOfOccurrences"] == 2
