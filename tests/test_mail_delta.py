"""Tests for mail delta tool."""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.tools.mail_delta import _format_message_delta, list_inbox_delta

# ── Helpers ──────────────────────────────────────────────────────────


def _raw_message(**overrides) -> dict:
    """Build a raw Graph JSON message dict (the wire shape)."""
    base = {
        "id": "AAMkAG123=",
        "subject": "Test Subject",
        "from": {"emailAddress": {"address": "sender@test.com", "name": "Sender"}},
        "receivedDateTime": "2026-05-21T10:00:00Z",
        "isRead": False,
        "importance": "normal",
        "bodyPreview": "Preview text",
        "hasAttachments": False,
        "categories": [],
        "flag": {"flagStatus": "notFlagged"},
        "conversationId": "conv123",
        "inferenceClassification": "focused",
    }
    base.update(overrides)
    return base


def _mock_graph_client():
    """Build a minimal graph_client that delta tools accept.

    Delta tools need ``.sdk_client`` (used by ``resolve_folder_id`` on
    the first call only) and ``.credential`` (used to mint raw bearer
    tokens for the httpx call).
    """
    client = MagicMock()
    client.credential = MagicMock()
    return client


def _patch_resolve(folder_id: str = "INBOX_ID"):
    return patch(
        "outlook_mcp.tools.mail_delta.resolve_folder_id",
        new=AsyncMock(return_value=folder_id),
    )


def _http_response(body: dict, status: int = 200):
    """Mock an httpx.Response with .json() and .raise_for_status()."""
    r = MagicMock()
    r.status_code = status
    r.json = MagicMock(return_value=body)
    r.raise_for_status = MagicMock()
    return r


def _async_client_with(responses):
    """Patch httpx.AsyncClient so its .get() walks a queue of responses."""
    responses = list(responses)

    async def fake_get(url, headers=None):
        return responses.pop(0)

    fake_client = MagicMock()
    fake_client.get = AsyncMock(side_effect=fake_get)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    return patch(
        "outlook_mcp.tools._delta.httpx.AsyncClient",
        return_value=fake_client,
    ), fake_client


# ── Formatter ────────────────────────────────────────────────────────


class TestFormatMessageDelta:
    def test_maps_basic_fields(self):
        out = _format_message_delta(_raw_message())
        assert out["id"] == "AAMkAG123="
        assert out["subject"] == "Test Subject"
        assert out["from_email"] == "sender@test.com"
        assert out["from_name"] == "Sender"
        assert out["is_read"] is False
        assert out["importance"] == "normal"
        assert out["flag"] == "notFlagged"
        assert out["conversation_id"] == "conv123"
        assert out["classification"] == "focused"

    def test_missing_from_node(self):
        out = _format_message_delta(_raw_message(**{"from": None}))
        assert out["from_email"] == ""
        assert out["from_name"] == ""

    def test_missing_subject(self):
        out = _format_message_delta(_raw_message(subject=None))
        assert out["subject"] == "(no subject)"


# ── First call ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_first_call_returns_messages_and_delta_token():
    body = {
        "value": [_raw_message(id="m1"), _raw_message(id="m2")],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/me/mailFolders('INBOX')/messages/delta?$deltatoken=abc",
    }
    patch_client, _ = _async_client_with([_http_response(body)])
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(_mock_graph_client(), folder="inbox", page_size=50)

    assert len(result["messages"]) == 2
    assert result["messages"][0]["id"] == "m1"
    assert all(m["is_deleted"] is False for m in result["messages"])
    assert result["delta_token"] == body["@odata.deltaLink"]
    assert result["has_more"] is False


# ── Subsequent call with delta_token ─────────────────────────────────


@pytest.mark.asyncio
async def test_subsequent_call_uses_delta_token_as_url():
    prior_token = "https://graph.microsoft.com/v1.0/me/mailFolders('INBOX')/messages/delta?$deltatoken=abc"
    new_delta = "https://graph.microsoft.com/v1.0/me/mailFolders('INBOX')/messages/delta?$deltatoken=def"
    body = {
        "value": [_raw_message(id="changed1", subject="Updated")],
        "@odata.deltaLink": new_delta,
    }
    patch_client, fake_client = _async_client_with([_http_response(body)])
    with _patch_resolve(), patch_client:
        # resolve_folder_id should NOT be called on a delta-token call
        result = await list_inbox_delta(
            _mock_graph_client(), folder="inbox", delta_token=prior_token
        )

    # The delta_token URL was used verbatim
    called_url = fake_client.get.call_args.args[0]
    assert called_url == prior_token

    assert len(result["messages"]) == 1
    assert result["messages"][0]["subject"] == "Updated"
    assert result["delta_token"] == new_delta
    assert result["has_more"] is False


# ── Safety cap ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_cap_reached_returns_nextlink_and_has_more():
    """200 messages with page_size=50 → cap=200 hit exactly; nextLink returned."""
    page = lambda i, link: {  # noqa: E731
        "value": [_raw_message(id=f"m{i*50 + n}") for n in range(50)],
        "@odata.nextLink": link,
    }
    # Cap is page_size * 4 = 200; deliver 4 nextLink pages of 50 each, the
    # last one also a nextLink so the cap triggers.
    responses = [
        _http_response(page(0, "https://graph/page2")),
        _http_response(page(1, "https://graph/page3")),
        _http_response(page(2, "https://graph/page4")),
        _http_response(page(3, "https://graph/page5")),
    ]
    patch_client, fake_client = _async_client_with(responses)
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(_mock_graph_client(), folder="inbox", page_size=50)

    assert len(result["messages"]) == 200
    assert result["has_more"] is True
    assert result["delta_token"] == "https://graph/page5"
    # 4 HTTP calls total — the 4th came back with a nextLink that we hand
    # to the caller instead of following.
    assert fake_client.get.await_count == 4


# ── Final page reached after several pages ───────────────────────────


@pytest.mark.asyncio
async def test_follows_nextlink_until_deltalink():
    """Two nextLink pages then a deltaLink — should follow both, return deltaLink."""
    responses = [
        _http_response({
            "value": [_raw_message(id="m1")],
            "@odata.nextLink": "https://graph/p2",
        }),
        _http_response({
            "value": [_raw_message(id="m2")],
            "@odata.deltaLink": "https://graph/delta-final",
        }),
    ]
    patch_client, fake_client = _async_client_with(responses)
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(_mock_graph_client(), folder="inbox", page_size=50)

    assert len(result["messages"]) == 2
    assert result["delta_token"] == "https://graph/delta-final"
    assert result["has_more"] is False
    assert fake_client.get.await_count == 2


# ── Tombstones (@removed) ────────────────────────────────────────────


@pytest.mark.asyncio
async def test_removed_item_collapses_to_id_only():
    body = {
        "value": [
            _raw_message(id="m1"),
            {"id": "m2-deleted", "@removed": {"reason": "deleted"}},
        ],
        "@odata.deltaLink": "https://graph/delta",
    }
    patch_client, _ = _async_client_with([_http_response(body)])
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(_mock_graph_client())

    assert result["messages"][0] == {**result["messages"][0], "is_deleted": False}
    tomb = result["messages"][1]
    assert tomb == {"id": "m2-deleted", "is_deleted": True}
    # No other fields on a tombstone — agents shouldn't see stale data.
    assert set(tomb.keys()) == {"id", "is_deleted"}


# ── No changes (empty value, only deltaLink) ─────────────────────────


@pytest.mark.asyncio
async def test_no_changes_returns_empty_list_and_delta_token():
    body = {
        "value": [],
        "@odata.deltaLink": "https://graph/delta-same",
    }
    patch_client, _ = _async_client_with([_http_response(body)])
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(
            _mock_graph_client(), delta_token="https://graph/prior-delta"
        )

    assert result["messages"] == []
    assert result["delta_token"] == "https://graph/delta-same"
    assert result["has_more"] is False


@pytest.mark.asyncio
async def test_no_links_at_all_returns_none_token():
    """If Graph returns neither nextLink nor deltaLink, hand back None."""
    body = {"value": []}
    patch_client, _ = _async_client_with([_http_response(body)])
    with _patch_resolve(), patch_client:
        result = await list_inbox_delta(_mock_graph_client())

    assert result["messages"] == []
    assert result["delta_token"] is None
    assert result["has_more"] is False
