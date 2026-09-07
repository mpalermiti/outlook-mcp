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
from datetime import date, timedelta

import pytest

from outlook_mcp.tools.calendar_read import get_event
from outlook_mcp.tools.calendar_write import create_event, delete_event, update_event
from tests.conftest import LIVE_WRITE_SUBJECT

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]


def _anchor_monday() -> date:
    """A Monday about a month out — far enough not to clutter the working week."""
    d = date.today() + timedelta(days=30)
    return d + timedelta(days=(0 - d.weekday()) % 7)


@asynccontextmanager
async def _temporary_event(client, config, **kwargs):
    """Create an event, hand back its id, and always delete it."""
    created = await create_event(
        client.sdk_client,
        subject=LIVE_WRITE_SUBJECT,
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
