"""Graph's real cursors must survive the URL guard.

`require_graph_url` is deliberately stricter than "is this a URL": it demands
https, rejects control characters, and compares the whole `netloc`, so a
userinfo prefix or an explicit port is refused. That strictness is the point —
it closes the parser differentials that let a poisoned cursor through.

It is also the risk. Every delta call routes its cursor through that guard, so
a guard stricter than what Graph actually emits breaks delta sync completely —
and the mocked suite cannot see it, because the mocks return cursors we wrote
ourselves. The existing live tier covers mail query shapes and preflight only
checks that the delta endpoints answer; neither drives `fetch_delta_pages`.

These close that gap: real call, real cursor, through the real guard.
"""

import pytest

from outlook_mcp.tools._delta import require_graph_url
from outlook_mcp.tools.calendar_delta import list_events_delta
from outlook_mcp.tools.contacts_delta import list_contacts_delta
from outlook_mcp.tools.mail_delta import list_inbox_delta

pytestmark = [pytest.mark.live, pytest.mark.asyncio]


async def test_mail_delta_cursor_round_trips(real_graph_client):
    """A full sync round, then the returned cursor replayed through the guard."""
    result = await list_inbox_delta(real_graph_client, folder="inbox", page_size=5)

    token = result["delta_token"]
    assert token, "Graph returned no cursor — cannot verify the guard against it"
    # The assertion that matters: whatever Graph emitted is something we will
    # accept back. A raise here means the guard is stricter than Graph.
    assert require_graph_url(token, source="live-check") == token

    # And it is actually usable: feeding it back must not be refused.
    second = await list_inbox_delta(
        real_graph_client, folder="inbox", page_size=5, delta_token=token
    )
    assert "messages" in second


async def test_calendar_delta_cursor_round_trips(real_graph_client):
    result = await list_events_delta(
        real_graph_client,
        start="2026-05-21T00:00:00Z",
        end="2026-05-28T00:00:00Z",
        page_size=5,
    )
    token = result["delta_token"]
    assert token, "Graph returned no cursor"
    assert require_graph_url(token, source="live-check") == token


async def test_contacts_delta_cursor_round_trips(real_graph_client):
    result = await list_contacts_delta(real_graph_client, page_size=5)
    token = result["delta_token"]
    assert token, "Graph returned no cursor"
    assert require_graph_url(token, source="live-check") == token


async def test_a_real_cursor_has_no_port_or_userinfo(real_graph_client):
    """Pin the two assumptions the guard's netloc equality actually rests on.

    If Microsoft ever starts emitting a port or a userinfo prefix in a
    deltaLink, this fails loudly here rather than silently breaking every
    delta-sync caller in production.
    """
    from urllib.parse import urlsplit

    result = await list_inbox_delta(real_graph_client, folder="inbox", page_size=5)
    parsed = urlsplit(result["delta_token"])

    assert parsed.netloc == "graph.microsoft.com", parsed.netloc
    assert parsed.port is None
    assert parsed.username is None and parsed.password is None
