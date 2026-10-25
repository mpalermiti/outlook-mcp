"""Tests for the composed ``changes_since`` digest tool.

All tests mock the underlying v1.9.0 delta tool functions
(``list_inbox_delta`` / ``list_events_delta`` / ``list_contacts_delta``)
so we exercise the digest's composition / classification / window-filter
/ resync-recovery logic in isolation.
"""

from __future__ import annotations

import asyncio
from datetime import datetime, timedelta, timezone
from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest

from outlook_mcp.tools import digest
from outlook_mcp.tools.digest import changes_since

# ── Fixture builders ──────────────────────────────────────────────────


def _mock_graph_client():
    client = MagicMock()
    client.credential = MagicMock()
    return client


def _mail_item(
    *,
    id="m1",
    from_email="sender@test.com",
    from_name="Sender",
    subject="Test",
    received=None,  # default: now
    importance="normal",
    flag="notFlagged",
    is_deleted=False,
):
    if received is None:
        received = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if is_deleted:
        return {"id": id, "is_deleted": True}
    return {
        "id": id,
        "subject": subject,
        "from_email": from_email,
        "from_name": from_name,
        "received": received,
        "is_read": False,
        "importance": importance,
        "preview": "",
        "has_attachments": False,
        "categories": [],
        "flag": flag,
        "conversation_id": "c1",
        "classification": "focused",
        "is_deleted": False,
    }


def _event_item(
    *,
    id="e1",
    subject="Meeting",
    start="2026-05-22T14:00:00.0000000 (UTC)",
    end="2026-05-22T15:00:00.0000000 (UTC)",
    is_deleted=False,
):
    if is_deleted:
        return {"id": id, "is_deleted": True}
    return {
        "id": id,
        "subject": subject,
        "start": start,
        "end": end,
        "location": "",
        "is_all_day": False,
        "organizer": "Alice",
        "response_status": "accepted",
        "is_online": False,
        "is_deleted": False,
    }


def _contact_item(*, id="con1", is_deleted=False):
    if is_deleted:
        return {"id": id, "is_deleted": True}
    return {
        "id": id,
        "display_name": "Test Contact",
        "email": "c@test.com",
        "phone": "",
        "company": "",
        "is_deleted": False,
    }


def _delta_page(items, *, token="tok-next", has_more=False, key="messages"):
    return {key: items, "delta_token": token, "has_more": has_more}


def _mail_response(items, *, token="mail-tok-1", has_more=False):
    return _delta_page(items, token=token, has_more=has_more, key="messages")


def _events_response(items, *, token="events-tok-1", has_more=False):
    return _delta_page(items, token=token, has_more=has_more, key="events")


def _contacts_response(items, *, token="contacts-tok-1", has_more=False):
    return _delta_page(items, token=token, has_more=has_more, key="contacts")


def _empty_patches():
    """Return patch contexts that mock all three deltas to empty single-page responses."""
    return (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response([], token="mt")),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([], token="et")),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([], token="ct")),
        ),
    )


# ── 1. First call, empty mailbox ──────────────────────────────────────


@pytest.mark.asyncio
async def test_first_call_empty_returns_zero_counts_and_tokens():
    pm, pe, pc = _empty_patches()
    with pm, pe, pc:
        result = await changes_since(_mock_graph_client())

    assert result["mail"]["new_count"] == 0
    assert result["mail"]["modified_count"] == 0
    assert result["mail"]["removed_count"] == 0
    assert result["mail"]["urgent_flagged"] == []
    assert result["mail"]["by_sender"] == {}
    assert result["events"]["new"] == []
    assert result["events"]["modified"] == []
    assert result["events"]["cancelled"] == []
    assert result["contacts"]["new_count"] == 0
    assert result["contacts"]["modified_count"] == 0
    assert result["contacts"]["removed_count"] == 0
    assert result["delta_tokens"] == {"mail": "mt", "events": "et", "contacts": "ct"}
    assert "from" in result["window"] and "to" in result["window"]
    assert "_meta" not in result


# ── 2. First call, mail mix (high, flagged, normal) ───────────────────


@pytest.mark.asyncio
async def test_first_call_mail_high_flagged_normal_urgent_grouping():
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    items = [
        _mail_item(id="m1", from_email="a@x.com", importance="high", received=now_iso),
        _mail_item(id="m2", from_email="b@x.com", flag="flagged", received=now_iso),
        _mail_item(id="m3", from_email="a@x.com", received=now_iso),
    ]
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        result = await changes_since(_mock_graph_client())

    assert result["mail"]["new_count"] == 3
    assert len(result["mail"]["urgent_flagged"]) == 2
    ids = {u["id"] for u in result["mail"]["urgent_flagged"]}
    assert ids == {"m1", "m2"}
    # by_sender groups by lowercased from_email
    assert result["mail"]["by_sender"] == {"a@x.com": 2, "b@x.com": 1}


# ── 3. First call, fallback_window_hours filters out old items ────────


@pytest.mark.asyncio
async def test_first_call_window_filter_excludes_old_items():
    now = datetime.now(timezone.utc)
    recent = (now - timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
    too_old = (now - timedelta(hours=12)).strftime("%Y-%m-%dT%H:%M:%SZ")
    items = [
        _mail_item(id="recent", received=recent, importance="high"),
        _mail_item(id="old", received=too_old, importance="high"),
    ]
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        result = await changes_since(_mock_graph_client(), fallback_window_hours=4)

    assert result["mail"]["new_count"] == 1
    assert [u["id"] for u in result["mail"]["urgent_flagged"]] == ["recent"]


# ── 4. Subsequent mail call: 2 new + 1 deleted ────────────────────────


@pytest.mark.asyncio
async def test_subsequent_mail_call_counts_new_and_removed():
    items = [
        _mail_item(id="n1"),
        _mail_item(id="n2"),
        _mail_item(id="d1", is_deleted=True),
    ]
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items, token="new-mail-tok")),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "prior-mail", "events": None, "contacts": None},
        )

    assert result["mail"]["new_count"] == 2
    assert result["mail"]["removed_count"] == 1
    assert result["delta_tokens"]["mail"] == "new-mail-tok"


# ── 5. Subsequent events call: 1 cancelled (tombstone) ────────────────


@pytest.mark.asyncio
async def test_subsequent_events_call_cancelled_only():
    items = [_event_item(id="ev-gone", is_deleted=True)]
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response([])),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response(items, token="ev-tok-2")),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"events": "prior-events"},
        )

    assert len(result["events"]["cancelled"]) == 1
    assert result["events"]["cancelled"][0]["id"] == "ev-gone"
    assert result["events"]["modified"] == []
    assert result["events"]["new"] == []


# ── 6. Calendar bootstrap passes correct start/end ────────────────────


@pytest.mark.asyncio
async def test_first_call_calendar_passes_window_args():
    mock_events = AsyncMock(return_value=_events_response([]))
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response([])),
        ),
        patch.object(digest, "list_events_delta", new=mock_events),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        await changes_since(_mock_graph_client(), fallback_window_hours=6)

    # First call should be on the events tool. Inspect kwargs.
    _, kwargs = mock_events.call_args
    start_arg = kwargs["start"]
    end_arg = kwargs["end"]
    # Both should be ISO8601-Z strings.
    start_dt = datetime.strptime(start_arg, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
    end_dt = datetime.strptime(end_arg, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
    # End must be ~7 days after now; start must be ~6 hours before now.
    now = datetime.now(timezone.utc)
    assert (now - start_dt) < timedelta(hours=6, minutes=1)
    assert (now - start_dt) > timedelta(hours=5, minutes=59)
    assert (end_dt - now) > timedelta(days=6, hours=23)
    assert (end_dt - now) < timedelta(days=7, hours=1)
    assert kwargs["delta_token"] is None


# ── 7. Mixed tokens — mail present, events/contacts missing ───────────


@pytest.mark.asyncio
async def test_mixed_tokens_each_resource_independent():
    mock_mail = AsyncMock(return_value=_mail_response([], token="mt"))
    mock_events = AsyncMock(return_value=_events_response([], token="et"))
    mock_contacts = AsyncMock(return_value=_contacts_response([], token="ct"))
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(digest, "list_events_delta", new=mock_events),
        patch.object(digest, "list_contacts_delta", new=mock_contacts),
    ):
        await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "prior-mail-tok"},
        )

    # Mail used the token; events/contacts were bootstrap calls.
    assert mock_mail.call_args.kwargs["delta_token"] == "prior-mail-tok"
    assert mock_events.call_args.kwargs["delta_token"] is None
    # Events got bootstrap window args; contacts has no window args at all.
    assert mock_events.call_args.kwargs["start"] is not None
    assert mock_events.call_args.kwargs["end"] is not None
    assert mock_contacts.call_args.kwargs["delta_token"] is None


# ── 8. syncStateNotFound (410) on mail token → resync ────────────────


@pytest.mark.asyncio
async def test_mail_sync_state_not_found_triggers_resync():
    # First call (with token) raises 410; second call (no token) returns
    # an empty snapshot.
    fake_response = MagicMock()
    fake_response.status_code = 410
    err = httpx.HTTPStatusError("sync state lost", request=MagicMock(), response=fake_response)

    call_args_seen: list[dict] = []

    async def flaky_mail(client, *, page_size, delta_token, received_since):
        call_args_seen.append({"delta_token": delta_token})
        if delta_token is not None:
            raise err
        return _mail_response([], token="fresh-mail-tok")

    with (
        patch.object(digest, "list_inbox_delta", new=AsyncMock(side_effect=flaky_mail)),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([], token="et")),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([], token="ct")),
        ),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "stale-mail-tok"},
        )

    assert result.get("_meta", {}).get("resync") == ["mail"]
    assert result["delta_tokens"]["mail"] == "fresh-mail-tok"
    # The other resources should not have resynced.
    assert result["delta_tokens"]["events"] == "et"
    assert result["delta_tokens"]["contacts"] == "ct"
    # Two mail attempts: stale-token first, then no-token resync.
    assert [c["delta_token"] for c in call_args_seen] == ["stale-mail-tok", None]


# ── 9. Internal pagination drains across has_more pages ──────────────


@pytest.mark.asyncio
async def test_internal_pagination_drains_multiple_pages():
    page1 = _mail_response(
        [_mail_item(id="p1-1"), _mail_item(id="p1-2")],
        token="page2-tok",
        has_more=True,
    )
    page2 = _mail_response(
        [_mail_item(id="p2-1")],
        token="final-mail-tok",
        has_more=False,
    )
    mock_mail = AsyncMock(side_effect=[page1, page2])
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        # Provide a token so we're on the token-driven path; window
        # filtering is off so all 3 items count.
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "start-tok"},
        )

    assert result["mail"]["new_count"] == 3
    assert mock_mail.await_count == 2
    assert result["delta_tokens"]["mail"] == "final-mail-tok"


# ── 10. by_sender capped at 5 ─────────────────────────────────────────


@pytest.mark.asyncio
async def test_by_sender_capped_at_top_five():
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    # 8 distinct senders with varying counts.
    items: list[dict] = []
    # Sender weights: a=8, b=7, c=6, d=5, e=4, f=3, g=2, h=1 (top 5: a-e)
    for letter, count in [
        ("a", 8),
        ("b", 7),
        ("c", 6),
        ("d", 5),
        ("e", 4),
        ("f", 3),
        ("g", 2),
        ("h", 1),
    ]:
        for n in range(count):
            items.append(
                _mail_item(id=f"{letter}{n}", from_email=f"{letter}@x.com", received=now_iso)
            )

    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([])),
        ),
    ):
        result = await changes_since(_mock_graph_client())

    by_sender = result["mail"]["by_sender"]
    assert len(by_sender) == 5
    assert set(by_sender.keys()) == {"a@x.com", "b@x.com", "c@x.com", "d@x.com", "e@x.com"}
    assert by_sender["a@x.com"] == 8
    assert by_sender["e@x.com"] == 4


# ── 11. No permission category needed — read-only ────────────────────


def test_digest_does_not_call_check_permission():
    """Sanity: digest.py composes read-only deltas; no permissions module import."""
    import outlook_mcp.tools.digest as digest_mod

    src = open(digest_mod.__file__).read()
    assert "check_permission" not in src
    assert "from outlook_mcp.permissions" not in src


# ── 12. Resources are fetched concurrently, not sequentially ──────────


@pytest.mark.asyncio
async def test_changes_since_fetches_resources_concurrently():
    """The three delta resources are fetched concurrently (asyncio.gather).

    Rendezvous barrier: each mocked delta call blocks until all three have
    started. Under sequential awaits, the first call blocks on the barrier
    forever (the other two never start), so its inner wait_for times out and
    changes_since raises. Under concurrent gather, all three start, the barrier
    releases, and the digest completes normally. This is what makes the
    most-used recurring tool ~2-3x lower latency.
    """
    started = 0
    all_started = asyncio.Event()

    async def _rendezvous():
        nonlocal started
        started += 1
        if started == 3:
            all_started.set()
        # If the three don't run concurrently, this never releases.
        await asyncio.wait_for(all_started.wait(), timeout=1.0)

    async def fake_inbox(client, *, page_size, delta_token, received_since):
        await _rendezvous()
        return _mail_response([], token="mt")

    async def fake_events(client, *, start, end, page_size, delta_token):
        await _rendezvous()
        return _events_response([], token="et")

    async def fake_contacts(client, *, page_size, delta_token, modified_since):
        await _rendezvous()
        return _contacts_response([], token="ct")

    with (
        patch.object(digest, "list_inbox_delta", new=fake_inbox),
        patch.object(digest, "list_events_delta", new=fake_events),
        patch.object(digest, "list_contacts_delta", new=fake_contacts),
    ):
        result = await changes_since(_mock_graph_client())

    assert result["delta_tokens"] == {"mail": "mt", "events": "et", "contacts": "ct"}


# ── 13. A first run is a window, not the whole mailbox ───────────────
#
# The first call used to take Graph's full snapshot of the inbox, keep the
# window client-side, and stop after five delta calls (~1,000 items) without
# saying so. The token it handed back was the snapshot's nextLink, so the next
# call took the token path, which applies no window: the rest of the mailbox's
# history came back as new mail, and every old high-importance or flagged
# message as urgent.


def _window_bound(fallback_window_hours: int) -> datetime:
    return datetime.now(timezone.utc) - timedelta(hours=fallback_window_hours)


@pytest.mark.asyncio
async def test_first_call_asks_graph_for_mail_in_the_window_only():
    mock_mail = AsyncMock(return_value=_mail_response([], token="mt"))
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(_mock_graph_client(), fallback_window_hours=6)

    since = mock_mail.call_args.kwargs["received_since"]
    assert abs(since - _window_bound(6)) < timedelta(minutes=1)
    # The bound Graph filters on is the window the response reports.
    assert result["window"]["from"] == since.strftime("%Y-%m-%dT%H:%M:%SZ")


@pytest.mark.asyncio
async def test_first_call_counts_only_contacts_changed_in_the_window():
    mock_contacts = AsyncMock(return_value=_contacts_response([], token="ct"))
    with (
        patch.object(digest, "list_inbox_delta", new=AsyncMock(return_value=_mail_response([]))),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(digest, "list_contacts_delta", new=mock_contacts),
    ):
        await changes_since(_mock_graph_client(), fallback_window_hours=6)

    since = mock_contacts.call_args.kwargs["modified_since"]
    assert abs(since - _window_bound(6)) < timedelta(minutes=1)


@pytest.mark.asyncio
async def test_a_token_call_sends_no_window():
    """A token already carries its round's filter; a changed contact is a change."""
    mock_mail = AsyncMock(return_value=_mail_response([], token="mt"))
    mock_contacts = AsyncMock(return_value=_contacts_response([], token="ct"))
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(digest, "list_contacts_delta", new=mock_contacts),
    ):
        await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "prior-mail", "events": "prior-ev", "contacts": "prior-c"},
        )

    assert mock_mail.call_args.kwargs["received_since"] is None
    assert mock_contacts.call_args.kwargs["modified_since"] is None


@pytest.mark.asyncio
async def test_a_resync_asks_for_the_window_again():
    fake_response = MagicMock()
    fake_response.status_code = 410
    err = httpx.HTTPStatusError("sync state lost", request=MagicMock(), response=fake_response)

    async def flaky_mail(client, *, page_size, delta_token, received_since):
        if delta_token is not None:
            raise err
        return _mail_response([], token="fresh-mail")

    async def flaky_contacts(client, *, page_size, delta_token, modified_since):
        if delta_token is not None:
            raise err
        return _contacts_response([], token="fresh-contacts")

    mock_mail = AsyncMock(side_effect=flaky_mail)
    mock_contacts = AsyncMock(side_effect=flaky_contacts)
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(digest, "list_contacts_delta", new=mock_contacts),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "stale", "contacts": "stale"},
            fallback_window_hours=6,
        )

    assert result["_meta"]["resync"] == ["mail", "contacts"]
    bound = _window_bound(6)
    assert abs(mock_mail.call_args.kwargs["received_since"] - bound) < timedelta(minutes=1)
    assert abs(mock_contacts.call_args.kwargs["modified_since"] - bound) < timedelta(minutes=1)


@pytest.mark.asyncio
async def test_has_more_says_which_resource_stopped_early():
    """Five delta calls is the digest's ceiling; the caller has to be told."""
    mock_mail = AsyncMock(
        side_effect=[
            _mail_response([_mail_item(id=f"m{i}")], token=f"next-{i}", has_more=True)
            for i in range(digest._MAX_INTERNAL_PAGES)
        ]
    )
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(_mock_graph_client())

    assert mock_mail.await_count == digest._MAX_INTERNAL_PAGES
    assert result["has_more"] == {"mail": True, "events": False, "contacts": False}
    assert result["delta_tokens"]["mail"] == f"next-{digest._MAX_INTERNAL_PAGES - 1}"


@pytest.mark.parametrize(
    "resource,patched,response",
    [
        ("mail", "list_inbox_delta", _mail_response),
        ("events", "list_events_delta", _events_response),
        ("contacts", "list_contacts_delta", _contacts_response),
    ],
)
@pytest.mark.asyncio
async def test_a_reply_with_no_new_link_keeps_the_callers_token(resource, patched, response):
    """No link means "your token is still good" — never a reason to store None.

    `fetch_delta_pages` returns no token when Graph answers with neither a
    nextLink nor a deltaLink, so the caller keeps using the one it has. The
    digest returned that None in `delta_tokens`, and a caller storing the dict
    as told lost its place: the next call started a fresh sync.
    """
    quiet = {
        "list_inbox_delta": _mail_response([], token="mt"),
        "list_events_delta": _events_response([], token="et"),
        "list_contacts_delta": _contacts_response([], token="ct"),
    }
    quiet[patched] = response([], token=None)
    with (
        patch.object(
            digest, "list_inbox_delta", new=AsyncMock(return_value=quiet["list_inbox_delta"])
        ),
        patch.object(
            digest, "list_events_delta", new=AsyncMock(return_value=quiet["list_events_delta"])
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value=quiet["list_contacts_delta"]),
        ),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "keep-m", "events": "keep-e", "contacts": "keep-c"},
        )

    assert result["delta_tokens"][resource] == f"keep-{resource[0]}"


@pytest.mark.asyncio
async def test_the_first_call_window_reaches_the_mail_request():
    """The bound is on the URL Graph receives, not only on a keyword argument."""
    seen: list[str] = []

    async def fake_get(url, headers=None):
        seen.append(url)
        response = MagicMock()
        response.status_code = 200
        response.raise_for_status = MagicMock()
        response.json = MagicMock(
            return_value={
                "value": [],
                "@odata.deltaLink": (
                    "https://graph.microsoft.com/v1.0/me/mailFolders('INBOX_ID')"
                    "/messages/delta?$deltatoken=abc"
                ),
            }
        )
        return response

    http = MagicMock()
    http.get = AsyncMock(side_effect=fake_get)
    http.__aenter__ = AsyncMock(return_value=http)
    http.__aexit__ = AsyncMock(return_value=False)
    with (
        patch(
            "outlook_mcp.tools.mail_delta.resolve_folder_id",
            new=AsyncMock(return_value="INBOX_ID"),
        ),
        patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=http),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(_mock_graph_client(), fallback_window_hours=6)

    assert seen == [
        "https://graph.microsoft.com/v1.0/me/mailFolders/INBOX_ID/messages/delta?$top=50"
        f"&$filter=receivedDateTime%20ge%20{result['window']['from']}"
    ]


@pytest.mark.asyncio
@pytest.mark.parametrize("ignored", [False, True])
async def test_mail_bootstrap_reports_a_filter_ignored_by_graph(ignored):
    recent = _mail_item(id="recent", importance="high")
    old = _mail_item(id="old", received="2000-01-01T00:00:00Z", importance="high")
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(
                return_value=_mail_response([recent, old] if ignored else [recent], token="done")
            ),
        ),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(_mock_graph_client())
    assert result["mail"]["new_count"] == 1
    assert [m["id"] for m in result["mail"]["urgent_flagged"]] == ["recent"]
    assert ("mail_window_fallback" in result.get("_meta", {})) is ignored
    assert result["has_more"]["mail"] is False
    assert "mail_bootstrap_from" not in result["delta_tokens"]


@pytest.mark.asyncio
async def test_ignored_mail_filter_keeps_its_bound_across_the_cap_and_then_releases_it():
    old = _mail_item(id="old", received="2000-01-01T00:00:00Z", importance="high")
    recent = _mail_item(id="recent", importance="high")
    mail = AsyncMock(
        side_effect=[
            _mail_response([old], token="unfinished", has_more=True),
            _mail_response([old, recent], token="finished"),
            _mail_response([old], token="later-change"),
        ]
    )
    with (
        patch.object(digest, "_MAX_INTERNAL_PAGES", 1),
        patch.object(digest, "list_inbox_delta", new=mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        first = await changes_since(_mock_graph_client(), fallback_window_hours=6)
        assert first["has_more"]["mail"] is True
        assert first["mail"]["new_count"] == 0
        assert first["delta_tokens"]["mail_bootstrap_from"] == first["window"]["from"]
        second = await changes_since(
            _mock_graph_client(), delta_tokens=first["delta_tokens"], fallback_window_hours=0
        )
        assert second["mail"]["new_count"] == 1
        assert second["window"]["from"] == first["window"]["from"]
        assert [m["id"] for m in second["mail"]["urgent_flagged"]] == ["recent"]
        assert second["has_more"]["mail"] is False
        assert "mail_bootstrap_from" not in second["delta_tokens"]
        assert "mail_window_fallback" in second["_meta"]
        third = await changes_since(_mock_graph_client(), delta_tokens=second["delta_tokens"])
    # After bootstrap, a genuine change to historical mail is still a change.
    assert third["mail"]["new_count"] == 1
    assert mail.call_args_list[1].kwargs["delta_token"] == "unfinished"
    assert (
        mail.call_args_list[1].kwargs["received_since"].strftime("%Y-%m-%dT%H:%M:%SZ")
        == first["window"]["from"]
    )
    assert mail.call_args_list[2].kwargs["received_since"] is None


@pytest.mark.asyncio
async def test_mail_resync_reports_both_the_resync_and_the_ignored_filter():
    response = httpx.Response(410, request=httpx.Request("GET", "https://graph.microsoft.com"))
    error = httpx.HTTPStatusError("stale", request=response.request, response=response)
    with (
        patch.object(
            digest,
            "list_inbox_delta",
            new=AsyncMock(
                side_effect=[
                    error,
                    _mail_response([_mail_item(received="2000-01-01T00:00:00Z")], token="fresh"),
                ]
            ),
        ),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(_mock_graph_client(), delta_tokens={"mail": "stale"})
    assert result["mail"]["new_count"] == 0
    assert result["_meta"]["resync"] == ["mail"]
    assert "mail_window_fallback" in result["_meta"]


@pytest.mark.asyncio
@pytest.mark.parametrize("bound", ["bad", 123])
async def test_invalid_mail_bootstrap_context_is_refused_before_graph(bound):
    mail = AsyncMock()
    with patch.object(digest, "list_inbox_delta", new=mail):
        with pytest.raises(ValueError, match="mail_bootstrap_from"):
            await changes_since(
                _mock_graph_client(), delta_tokens={"mail": "next", "mail_bootstrap_from": bound}
            )
    mail.assert_not_called()


@pytest.mark.asyncio
async def test_a_null_mail_bootstrap_bound_is_absent_on_an_ordinary_round():
    historical = _mail_item(received="2000-01-01T00:00:00Z", importance="high")
    mail = AsyncMock(return_value=_mail_response([historical], token="next"))
    with (
        patch.object(digest, "list_inbox_delta", new=mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(
            _mock_graph_client(), delta_tokens={"mail": "current", "mail_bootstrap_from": None}
        )
    assert result["mail"]["new_count"] == 1
    assert len(result["mail"]["urgent_flagged"]) == 1
    assert mail.call_args.kwargs["received_since"] is None
    assert "mail_window_fallback" not in result.get("_meta", {})
    assert "mail_bootstrap_from" not in result["delta_tokens"]


@pytest.mark.asyncio
async def test_a_resync_window_reports_its_new_bound_instead_of_the_stale_bootstrap_bound():
    response = httpx.Response(410, request=httpx.Request("GET", "https://graph.microsoft.com"))
    error = httpx.HTTPStatusError("stale", request=response.request, response=response)
    mail = AsyncMock(side_effect=[error, _mail_response([], token="fresh")])
    with (
        patch.object(digest, "list_inbox_delta", new=mail),
        patch.object(digest, "list_events_delta", new=AsyncMock(return_value=_events_response([]))),
        patch.object(
            digest, "list_contacts_delta", new=AsyncMock(return_value=_contacts_response([]))
        ),
    ):
        result = await changes_since(
            _mock_graph_client(),
            delta_tokens={"mail": "stale", "mail_bootstrap_from": "2000-01-01T00:00:00Z"},
        )
    assert result["window"]["from"] == mail.call_args.kwargs["received_since"].strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    assert result["_meta"]["resync"] == ["mail"]
