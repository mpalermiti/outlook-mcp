"""Live query-shape guards for the intent read tools (agenda, open_loops, catch_up,
brief). Read-only.

Run with: scripts/live-tests.sh live (RELEASING.md 1b), against the project's
test accounts. Deselected by default, so the offline suite and CI never run them.

The offline tests (``test_intent_agenda.py`` and the others) assert what each tool
*sends*. These assert what Graph *does* with it, for every new ``$filter``,
``$orderby``, ``$search``, ``$select`` and delta shape the read tools build: a loud
400 fails the call, and the quiet failure, a 200 with the wrong rows, fails an
assertion on the rows. Each query is built from the same constant or helper the
tool uses, so the two cannot drift apart. Where a guard settles a probe from the
contracts (§14), its id is named.

They read whatever mailbox is signed in and skip when it lacks the data an
assertion needs. Every request is a read: the live guard refuses anything else
before it leaves (``tests/live_guard.py``).
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.intents import agenda
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.reading import read_view
from outlook_mcp.intents.resolver import open_lookup

pytestmark = [pytest.mark.live, pytest.mark.asyncio]


@pytest.fixture
async def lookup(real_graph_client):
    """A counted lookup over the signed-in mailbox, with room for harvesting."""
    async with open_lookup(
        real_graph_client,
        RequestBudget(200),
        now=datetime.now(timezone.utc),
        zone="UTC",
        signed_in=None,
    ) as found:
        yield found


# ── agenda (§12.4) ────────────────────────────────────────────────────────


async def test_agenda_calendar_view_takes_its_select(lookup):
    """A2: the calendar view with agenda's `$select`, `iCalUId` and `location` included.
    An all-day event must come back as a floating midnight-to-midnight date, which is
    how agenda reads it (``reading.event_span``)."""
    calendars = await fetch_all_calendars(lookup.sdk)
    if not calendars:
        pytest.skip("no calendars")
    default = next((c for c in calendars if c.is_default_calendar), calendars[0])
    now = datetime.now(timezone.utc)
    rows, _ = await read_view(
        lookup,
        f"/me/calendars/{default.id}/calendarView",
        now - timedelta(days=30),
        now + timedelta(days=30),
        agenda.EVENT_SELECT,
        pages=1,
    )
    if not rows:
        pytest.skip("no events within 30 days of now")
    for row in rows:
        assert "iCalUId" in row and "showAs" in row and "start" in row, sorted(row)
    for row in [r for r in rows if r.get("isAllDay") is True]:
        assert row["start"]["dateTime"][11:19] == "00:00:00", row["start"]
        assert row["end"]["dateTime"][11:19] == "00:00:00", row["end"]


async def test_agenda_end_to_end_stays_within_its_ceiling(real_graph_client):
    env = await agenda.read_agenda(
        real_graph_client, window="this_week", config_zone="UTC", loaded_tools=frozenset()
    )
    assert env["tool"] == "outlook_agenda"
    assert env["outcome"] == "answered", env["failed_sections"]
    assert env["graph_requests"] <= 25
    assert len(env["sections"][0]["items"]) == 7
