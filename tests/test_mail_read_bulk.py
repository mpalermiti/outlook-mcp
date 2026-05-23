"""Tests for bulk-read (``read_messages``) — Graph ``$batch`` GET path.

These tests mock the raw httpx POST to ``$batch`` instead of the SDK fluent
chain. ``read_messages`` deliberately bypasses the SDK (raw httpx) the same
way the delta tools do, so unit tests target the wire layer directly.
"""

from __future__ import annotations

import json
from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest

from outlook_mcp.tools.mail_read import (
    BATCH_URL,
    MAX_BATCH_SIZE,
    read_messages,
)


def _fake_graph_client(token: str = "fake-token") -> MagicMock:
    """Build a fake ``GraphClient`` exposing the .credential attribute."""
    client = MagicMock()
    client.credential.get_token = MagicMock(return_value=MagicMock(token=token))
    return client


def _raw_message(
    msg_id: str,
    *,
    subject: str = "Hello",
    body_html: str = "<p>hi</p>",
    from_email: str = "alice@example.com",
    from_name: str = "Alice",
    is_read: bool = False,
    importance: str = "normal",
    has_attachments: bool = False,
    extended_props: list[dict] | None = None,
) -> dict:
    """Build a Graph-shape JSON message dict."""
    msg: dict = {
        "id": msg_id,
        "subject": subject,
        "from": {"emailAddress": {"address": from_email, "name": from_name}},
        "toRecipients": [
            {"emailAddress": {"address": "bob@example.com", "name": "Bob"}},
        ],
        "ccRecipients": [],
        "receivedDateTime": "2026-05-22T10:00:00Z",
        "body": {"contentType": "html", "content": body_html},
        "isRead": is_read,
        "importance": importance,
        "hasAttachments": has_attachments,
        "attachments": [],
        "categories": [],
        "flag": {"flagStatus": "notFlagged"},
        "conversationId": f"conv-{msg_id}",
    }
    if extended_props is not None:
        msg["singleValueExtendedProperties"] = extended_props
    return msg


def _batch_response(subresponses: list[dict]) -> MagicMock:
    """Build a fake httpx Response whose .json() yields the $batch payload."""
    resp = MagicMock()
    resp.status_code = 200
    resp.raise_for_status = MagicMock()
    resp.json = MagicMock(return_value={"responses": subresponses})
    return resp


def _ok_sub(req_id: str, body: dict) -> dict:
    return {"id": req_id, "status": 200, "headers": {}, "body": body}


def _err_sub(req_id: str, status: int, code: str, message: str) -> dict:
    return {
        "id": req_id,
        "status": status,
        "headers": {},
        "body": {"error": {"code": code, "message": message}},
    }


@pytest.fixture
def patched_httpx():
    """Patch ``httpx.AsyncClient`` so ``read_messages`` doesn't hit the network.

    Yields ``(post_mock, set_response)``. ``set_response(resp)`` configures
    the next ``client.post(...)`` to return ``resp``.
    """
    with patch("outlook_mcp.tools.mail_read.httpx.AsyncClient") as cm:
        client_instance = MagicMock()
        post_mock = AsyncMock()
        client_instance.post = post_mock
        cm.return_value.__aenter__ = AsyncMock(return_value=client_instance)
        cm.return_value.__aexit__ = AsyncMock(return_value=False)

        def set_response(resp):
            post_mock.return_value = resp

        yield post_mock, set_response


# ── Input validation ────────────────────────────────────────────────────


class TestInputValidation:
    @pytest.mark.asyncio
    async def test_empty_list_raises(self):
        client = _fake_graph_client()
        with pytest.raises(ValueError, match="must not be empty"):
            await read_messages(client, [])

    @pytest.mark.asyncio
    async def test_over_cap_raises(self):
        client = _fake_graph_client()
        ids = [f"AAA{i}=" for i in range(MAX_BATCH_SIZE + 1)]
        with pytest.raises(ValueError, match="Maximum"):
            await read_messages(client, ids)

    @pytest.mark.asyncio
    async def test_malformed_id_raises_before_http(self, patched_httpx):
        post_mock, _ = patched_httpx
        client = _fake_graph_client()
        with pytest.raises(ValueError, match="invalid characters"):
            await read_messages(client, ["bad id with spaces!"])
        # No HTTP call must have happened.
        post_mock.assert_not_called()


# ── Happy paths ──────────────────────────────────────────────────────────


class TestBatchReadHappyPath:
    @pytest.mark.asyncio
    async def test_single_id(self, patched_httpx):
        post_mock, set_response = patched_httpx
        set_response(_batch_response([_ok_sub("0", _raw_message("AAA=", subject="One"))]))

        result = await read_messages(_fake_graph_client(), ["AAA="])

        assert result == {
            "messages": [result["messages"][0]],
            "failures": [],
            "requested": 1,
            "succeeded": 1,
            "failed": 0,
        }
        assert result["messages"][0]["subject"] == "One"
        assert result["messages"][0]["id"] == "AAA="
        # The POST was to $batch with one sub-request.
        post_mock.assert_called_once()
        call = post_mock.call_args
        assert call.args[0] == BATCH_URL
        body = json.loads(call.kwargs["content"])
        assert body == {
            "requests": [
                {"id": "0", "method": "GET", "url": "/me/messages/AAA%3D"},
            ]
        }

    @pytest.mark.asyncio
    async def test_multiple_all_succeed(self, patched_httpx):
        _, set_response = patched_httpx
        ids = ["AAA=", "BBB=", "CCC="]
        set_response(_batch_response([
            _ok_sub("0", _raw_message("AAA=", subject="A")),
            _ok_sub("1", _raw_message("BBB=", subject="B")),
            _ok_sub("2", _raw_message("CCC=", subject="C")),
        ]))

        result = await read_messages(_fake_graph_client(), ids)
        assert result["requested"] == 3
        assert result["succeeded"] == 3
        assert result["failed"] == 0
        # Ordering matches input.
        assert [m["id"] for m in result["messages"]] == ids
        assert [m["subject"] for m in result["messages"]] == ["A", "B", "C"]

    @pytest.mark.asyncio
    async def test_ordering_preserved_even_when_graph_reorders(self, patched_httpx):
        """Graph spec allows responses in any order — input order is what we return."""
        _, set_response = patched_httpx
        ids = ["ID_C=", "ID_A=", "ID_B="]
        # Graph returns them shuffled relative to input.
        set_response(_batch_response([
            _ok_sub("1", _raw_message("ID_A=", subject="msgA")),
            _ok_sub("2", _raw_message("ID_B=", subject="msgB")),
            _ok_sub("0", _raw_message("ID_C=", subject="msgC")),
        ]))

        result = await read_messages(_fake_graph_client(), ids)
        # Input order: [C, A, B]
        assert [m["subject"] for m in result["messages"]] == ["msgC", "msgA", "msgB"]
        assert [m["id"] for m in result["messages"]] == ["ID_C=", "ID_A=", "ID_B="]


# ── Failure paths ────────────────────────────────────────────────────────


class TestBatchReadFailures:
    @pytest.mark.asyncio
    async def test_partial_failure_404_mixed(self, patched_httpx):
        _, set_response = patched_httpx
        ids = ["AAA=", "BBB=", "CCC="]
        set_response(_batch_response([
            _ok_sub("0", _raw_message("AAA=", subject="A")),
            _err_sub("1", 404, "ErrorItemNotFound", "The specified object was not found."),
            _ok_sub("2", _raw_message("CCC=", subject="C")),
        ]))

        result = await read_messages(_fake_graph_client(), ids)
        assert result["requested"] == 3
        assert result["succeeded"] == 2
        assert result["failed"] == 1
        # messages[] excludes the failed one, ordering preserved for survivors.
        assert [m["subject"] for m in result["messages"]] == ["A", "C"]
        # The failure references the original (input) id.
        assert result["failures"] == [
            {
                "id": "BBB=",
                "status": 404,
                "code": "ErrorItemNotFound",
                "message": "The specified object was not found.",
            }
        ]
        # The invariant.
        assert result["requested"] == result["succeeded"] + result["failed"]

    @pytest.mark.asyncio
    async def test_all_fail(self, patched_httpx):
        _, set_response = patched_httpx
        ids = ["AAA=", "BBB=", "CCC="]
        set_response(_batch_response([
            _err_sub("0", 404, "ErrorItemNotFound", "gone"),
            _err_sub("1", 404, "ErrorItemNotFound", "gone"),
            _err_sub("2", 403, "ErrorAccessDenied", "no"),
        ]))

        result = await read_messages(_fake_graph_client(), ids)
        assert result["messages"] == []
        assert result["succeeded"] == 0
        assert result["failed"] == 3
        # Ordering of failures matches input order.
        assert [f["id"] for f in result["failures"]] == ids
        assert [f["status"] for f in result["failures"]] == [404, 404, 403]

    @pytest.mark.asyncio
    async def test_transport_level_500_raises(self, patched_httpx):
        """A 500 on the whole $batch is NOT swallowed into failures[]."""
        post_mock, _ = patched_httpx
        resp = MagicMock()
        resp.status_code = 500
        resp.raise_for_status = MagicMock(side_effect=httpx.HTTPStatusError(
            "boom", request=MagicMock(), response=MagicMock(status_code=500),
        ))
        post_mock.return_value = resp

        with pytest.raises(httpx.HTTPStatusError):
            await read_messages(_fake_graph_client(), ["AAA="])


# ── Format / concise / deferred-send ─────────────────────────────────────


def _captured_subrequests(post_mock) -> list[dict]:
    body = json.loads(post_mock.call_args.kwargs["content"])
    return body["requests"]


class TestFormatVariants:
    @pytest.mark.asyncio
    async def test_format_text_default(self, patched_httpx):
        _, set_response = patched_httpx
        set_response(_batch_response([_ok_sub("0", _raw_message("AAA=", body_html="<p>Hi</p>"))]))
        result = await read_messages(_fake_graph_client(), ["AAA="], format="text")
        msg = result["messages"][0]
        # text only: body present (sanitized), body_html is None
        assert msg["body"] != ""
        assert "Hi" in msg["body"]
        assert msg["body_html"] is None

    @pytest.mark.asyncio
    async def test_format_html(self, patched_httpx):
        post_mock, set_response = patched_httpx
        set_response(_batch_response([
            _ok_sub("0", _raw_message("AAA=", body_html="<p>Hello</p>")),
        ]))
        result = await read_messages(_fake_graph_client(), ["AAA="], format="html")
        msg = result["messages"][0]
        assert msg["body"] == ""
        assert msg["body_html"] == "<p>Hello</p>"
        # The sub-request URL stays at /me/messages/<id> — body format
        # is projected post-fetch, matching outlook_read_message.
        subs = _captured_subrequests(post_mock)
        assert subs[0]["url"] == "/me/messages/AAA%3D"

    @pytest.mark.asyncio
    async def test_format_full(self, patched_httpx):
        _, set_response = patched_httpx
        set_response(_batch_response([
            _ok_sub("0", _raw_message("AAA=", body_html="<p>Hello</p>")),
        ]))
        result = await read_messages(_fake_graph_client(), ["AAA="], format="full")
        msg = result["messages"][0]
        # full: both populated
        assert msg["body"] != ""
        assert "Hello" in msg["body"]
        assert msg["body_html"] == "<p>Hello</p>"

    @pytest.mark.asyncio
    async def test_concise_drops_body(self, patched_httpx):
        _, set_response = patched_httpx
        long_body = "<p>" + ("xy" * 500) + "</p>"
        set_response(_batch_response([_ok_sub("0", _raw_message("AAA=", body_html=long_body))]))

        result = await read_messages(_fake_graph_client(), ["AAA="], concise=True)
        msg = result["messages"][0]
        assert "body" not in msg
        assert "body_html" not in msg
        assert "body_preview" in msg
        assert len(msg["body_preview"]) <= 200
        # Headers preserved.
        assert msg["subject"] == "Hello"
        assert msg["from_email"] == "alice@example.com"
        assert "to" in msg
        assert "cc" in msg


class TestIncludeDeferredSend:
    @pytest.mark.asyncio
    async def test_extended_property_filter_in_url(self, patched_httpx):
        """The sub-request URL adds $expand=singleValueExtendedProperties($filter=...)."""
        post_mock, set_response = patched_httpx
        set_response(_batch_response([
            _ok_sub("0", _raw_message(
                "AAA=",
                extended_props=[
                    {"id": "SystemTime 0x3fef", "value": "2026-06-01T15:00:00Z"},
                ],
            )),
        ]))

        result = await read_messages(
            _fake_graph_client(), ["AAA="], include_deferred_send=True
        )
        msg = result["messages"][0]
        # The deferred value is surfaced.
        assert msg["deferred_send_datetime"] == "2026-06-01T15:00:00Z"
        # The sub-request URL contains the expand+filter.
        subs = _captured_subrequests(post_mock)
        url = subs[0]["url"]
        assert "$expand=singleValueExtendedProperties" in url
        # The $filter clause is included verbatim (matches outlook_read_message's
        # raw-URL builder: `=`, space, and single-quotes are kept literal).
        assert "$filter=id eq 'SystemTime 0x3FEF'" in url

    @pytest.mark.asyncio
    async def test_deferred_send_none_when_property_absent(self, patched_httpx):
        _, set_response = patched_httpx
        set_response(_batch_response([
            _ok_sub("0", _raw_message("AAA=", extended_props=[])),
        ]))

        result = await read_messages(
            _fake_graph_client(), ["AAA="], include_deferred_send=True
        )
        assert result["messages"][0]["deferred_send_datetime"] is None
