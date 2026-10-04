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

The same goes, more sharply, for the per-tool endpoint check: it matches the
cursor's *path* against a pattern written from what Graph was seen to emit —
`/me/mailFolders('<id>')/messages/delta` and so on. If Graph spells a link
another way, every delta call fails on its own cursor. So each tool's
deltaLink is put through `resource=` here, and so is each tool's mid-sync
nextLink where the mailbox is large enough to produce one.
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
    assert require_graph_url(token, source="live-check", resource="mail") == token

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
    assert require_graph_url(token, source="live-check", resource="calendar") == token


async def test_contacts_delta_cursor_round_trips(real_graph_client):
    result = await list_contacts_delta(real_graph_client, page_size=5)
    token = result["delta_token"]
    assert token, "Graph returned no cursor"
    assert require_graph_url(token, source="live-check", resource="contacts") == token


# The other cursor shape: a nextLink handed back when the per-call cap stops a
# walk. `page_size=1` caps a call at four items, so a resource with more than
# that returns a nextLink (`$skiptoken`) rather than a deltaLink. One test per
# tool, and a skip that names the shape when the mailbox is too small to show
# it — a run that never saw a mail nextLink must not read as having checked one.


async def test_a_mid_sync_mail_nextlink_passes_the_mail_check(real_graph_client):
    result = await list_inbox_delta(real_graph_client, folder="inbox", page_size=1)
    if not result["has_more"]:
        pytest.skip(
            "The inbox has four messages or fewer, so Graph returned no mail nextLink — "
            "that cursor shape is unverified by this run"
        )
    token = result["delta_token"]
    assert require_graph_url(token, source="live-check", resource="mail") == token
    resumed = await list_inbox_delta(
        real_graph_client, folder="inbox", page_size=1, delta_token=token
    )
    assert "messages" in resumed


async def test_a_mid_sync_calendar_nextlink_passes_the_calendar_check(real_graph_client):
    result = await list_events_delta(
        real_graph_client,
        start="2026-01-01T00:00:00Z",
        end="2026-12-31T00:00:00Z",
        page_size=1,
    )
    if not result["has_more"]:
        pytest.skip(
            "Four events or fewer in 2026, so Graph returned no calendar nextLink — "
            "that cursor shape is unverified by this run"
        )
    token = result["delta_token"]
    assert require_graph_url(token, source="live-check", resource="calendar") == token
    resumed = await list_events_delta(
        real_graph_client, start=None, end=None, page_size=1, delta_token=token
    )
    assert "events" in resumed


async def test_a_mid_sync_contacts_nextlink_passes_the_contacts_check(real_graph_client):
    result = await list_contacts_delta(real_graph_client, page_size=1)
    if not result["has_more"]:
        pytest.skip(
            "Four contacts or fewer, so Graph returned no contacts nextLink — "
            "that cursor shape is unverified by this run"
        )
    token = result["delta_token"]
    assert require_graph_url(token, source="live-check", resource="contacts") == token
    resumed = await list_contacts_delta(real_graph_client, page_size=1, delta_token=token)
    assert "contacts" in resumed


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
