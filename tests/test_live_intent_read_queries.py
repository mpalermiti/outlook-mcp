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
from urllib.parse import parse_qsl

import pytest

from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import agenda, flags, open_loops
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.envelope import utc
from outlook_mcp.intents.people import parse_time
from outlook_mcp.intents.reading import read_view
from outlook_mcp.intents.resolver import SENT_PATH, _conversation_filter, open_lookup

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


# ── open_loops (§12.3) and the follow-up flags (§8.2) ─────────────────────


async def test_l1_sent_items_in_the_window_newest_first(lookup):
    """L1: Sent Items filtered and ordered on `sentDateTime`, with the eventMessage cast
    if Graph takes it (P-SI1), and without it if not, as open_loops falls back."""
    since = utc(datetime.now(timezone.utc) - timedelta(days=90))
    query = {
        "$filter": f"sentDateTime ge {since}",
        "$orderby": "sentDateTime desc",
        "$select": open_loops.SENT_SELECT,
        "$top": 100,
    }
    try:
        body = await lookup.graph.get(SENT_PATH, query)
    except GraphAPIError as error:
        if error.status_code != 400:
            raise
        body = await lookup.graph.get(SENT_PATH, {**query, "$select": open_loops.SENT_SELECT_PLAIN})
    rows = body.get("value") or []
    if not rows:
        pytest.skip("no letters sent in 90 days")
    sent = [r["sentDateTime"] for r in rows]
    assert sent == sorted(sent, reverse=True)
    assert all(s >= since for s in sent), "the filter must bound the rows"
    assert all("flag" in r and "conversationId" in r for r in rows)


@pytest.mark.parametrize("folder", ["inbox", "archive"])
async def test_l2_l3_a_folder_in_the_window_newest_first(lookup, folder):
    since = utc(datetime.now(timezone.utc) - timedelta(days=90))
    try:
        body = await lookup.graph.get(
            f"/me/mailFolders/{folder}/messages",
            {
                "$filter": f"receivedDateTime ge {since}",
                "$orderby": "receivedDateTime desc",
                "$select": open_loops.RECEIVED_SELECT,
                "$top": 100,
            },
        )
    except GraphAPIError as error:
        if folder == "archive" and error.status_code == 404:
            pytest.skip("this account has no Archive folder; open_loops says so")
        raise
    rows = body.get("value") or []
    if not rows:
        pytest.skip(f"nothing in {folder} within 90 days")
    received = [r["receivedDateTime"] for r in rows]
    assert received == sorted(received, reverse=True)
    assert all(r >= since for r in received)
    assert all("inferenceClassification" in r for r in rows)


async def test_p_d3a_f1_flagged_mail_newest_first(lookup):
    """P-D3a: the flag filter behind the received-time floor, ordered newest first. A
    400 here means the fallback scan is what runs; the guard says so and skips."""
    try:
        body = await lookup.graph.get(
            "/me/messages",
            {
                "$filter": flags.F1_FILTER,
                "$orderby": "receivedDateTime desc",
                "$select": flags.F1_SELECT,
                "$top": flags.F1_TOP,
            },
        )
    except GraphAPIError as error:
        if error.status_code == 400:
            pytest.skip("P-D3a: Graph refuses the flag filter here; the fallback scan runs")
        raise
    rows = body.get("value") or []
    if not rows:
        pytest.skip("no flagged mail")
    assert all(flags.is_flagged(r) for r in rows), "only flagged mail may come back"
    received = [r["receivedDateTime"] for r in rows]
    assert received == sorted(received, reverse=True)


async def test_f2_later_mail_in_a_conversation_oldest_first(lookup):
    """F2: `receivedDateTime gt` then the conversation, ascending: only later mail."""
    body = await lookup.graph.get(
        "/me/messages", {"$select": "id,conversationId,receivedDateTime", "$top": 50}
    )
    rows = body.get("value") or []
    if not rows:
        pytest.skip("no mail")
    first = rows[-1]
    (url,) = flags.reply_urls([(first, parse_time(first["receivedDateTime"]))])
    path, _, query = url.partition("?")
    later = await lookup.graph.get(path, dict(parse_qsl(query)))
    found = later.get("value") or []
    times = [r["receivedDateTime"] for r in found]
    assert times == sorted(times)
    assert all(t > first["receivedDateTime"] for t in times)
    assert first["id"] not in {r["id"] for r in found}


async def test_l8_l9_a_search_and_a_conversations_latest(lookup):
    body = await lookup.graph.get(
        "/me/messages", {"$search": '"the"', "$select": open_loops.SEARCH_SELECT, "$top": 25}
    )
    rows = [r for r in body.get("value") or [] if r.get("conversationId")]
    if not rows:
        pytest.skip("nothing matched a common word")
    latest = await lookup.graph.get(
        "/me/messages",
        {
            "$filter": _conversation_filter(rows[0]["conversationId"]),
            "$orderby": "receivedDateTime desc",
            "$top": 1,
            "$select": open_loops.LATEST_SELECT,
        },
    )
    (row,) = latest.get("value") or []
    assert row["conversationId"] == rows[0]["conversationId"]


async def test_open_loops_end_to_end_stays_within_its_ceiling(real_graph_client):
    env = await open_loops.read_open_loops(
        real_graph_client, window="30d", zone="UTC", loaded_tools=frozenset()
    )
    assert env["tool"] == "outlook_open_loops"
    assert env["outcome"] == "answered", env["failed_sections"]
    assert env["graph_requests"] <= 38
