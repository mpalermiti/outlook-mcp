"""Tests for the composed ``changes_since`` digest tool.

All tests mock the underlying v1.9.0 delta tool functions
(``list_inbox_delta`` / ``list_events_delta`` / ``list_contacts_delta``)
so we exercise the digest's composition / classification / window-filter
/ resync-recovery logic in isolation.
"""

from __future__ import annotations

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


def _event_item(*, id="e1", subject="Meeting", start="2026-05-22T14:00:00.0000000 (UTC)",
                end="2026-05-22T15:00:00.0000000 (UTC)", is_deleted=False):
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
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items, token="new-mail-tok")),
        ),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response([])),
        ),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response(items, token="ev-tok-2")),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response([])),
        ),
        patch.object(digest, "list_events_delta", new=mock_events),
        patch.object(
            digest, "list_contacts_delta",
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
    err = httpx.HTTPStatusError(
        "sync state lost", request=MagicMock(), response=fake_response
    )

    call_args_seen: list[dict] = []

    async def flaky_mail(client, *, page_size, delta_token):
        call_args_seen.append({"delta_token": delta_token})
        if delta_token is not None:
            raise err
        return _mail_response([], token="fresh-mail-tok")

    with (
        patch.object(digest, "list_inbox_delta", new=AsyncMock(side_effect=flaky_mail)),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([], token="et")),
        ),
        patch.object(
            digest, "list_contacts_delta",
            new=AsyncMock(return_value=_contacts_response([], token="ct"))),
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
        token="page2-tok", has_more=True,
    )
    page2 = _mail_response(
        [_mail_item(id="p2-1")],
        token="final-mail-tok", has_more=False,
    )
    mock_mail = AsyncMock(side_effect=[page1, page2])
    with (
        patch.object(digest, "list_inbox_delta", new=mock_mail),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
    for letter, count in [("a", 8), ("b", 7), ("c", 6), ("d", 5),
                          ("e", 4), ("f", 3), ("g", 2), ("h", 1)]:
        for n in range(count):
            items.append(
                _mail_item(id=f"{letter}{n}", from_email=f"{letter}@x.com", received=now_iso)
            )

    with (
        patch.object(
            digest, "list_inbox_delta",
            new=AsyncMock(return_value=_mail_response(items)),
        ),
        patch.object(
            digest, "list_events_delta",
            new=AsyncMock(return_value=_events_response([])),
        ),
        patch.object(
            digest, "list_contacts_delta",
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
