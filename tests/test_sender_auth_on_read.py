"""`include_sender_auth` on the read tools: asked for only when asked, and delivered (#130).

Graph returns ``internetMessageHeaders`` only when ``$select`` names it, and once
``$select`` names anything Graph returns nothing else. So the opt-in read names
every field the formatter reads, and the default read must stay exactly what it
was: no ``$select`` at all. Both are asserted on the URL the real msgraph
pipeline sends (built as ``test_me_rewrite_reaches_the_wire.py`` builds it), on
the ``$batch`` sub-request URLs, and on what an MCP client finally receives.
"""

from __future__ import annotations

import asyncio
import json
from unittest.mock import AsyncMock, MagicMock, patch
from urllib.parse import parse_qs, unquote, urlsplit

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from kiota_serialization_json.json_parse_node import JsonParseNode
from mcp.client import Client
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph.generated.models.message import Message
from msgraph_core import GraphClientFactory

from outlook_mcp.tools import mail_read

MESSAGE_ID = "AAMkAGI2THVSAAA="

MS_PASS = (
    "spf=pass (sender IP is 192.0.2.10) smtp.mailfrom=contoso.com; "
    "dkim=pass (signature was verified) header.d=contoso.com;"
    "dmarc=pass action=none header.from=contoso.com;compauth=pass reason=100"
)
MS_FAIL = (
    "spf=fail (sender IP is 203.0.113.9) smtp.mailfrom=contoso.com; "
    "dkim=none (message not signed) header.d=none;"
    "dmarc=fail action=none header.from=contoso.com;compauth=fail reason=001"
)


def _graph_message(auth_results: str = MS_PASS) -> dict:
    """What Graph answers for the opt-in read: the selected fields only."""
    return {
        "id": MESSAGE_ID,
        "subject": "Invoice 4471",
        "from": {"emailAddress": {"address": "billing@contoso.com", "name": "Contoso Billing"}},
        "toRecipients": [{"emailAddress": {"address": "user@example.com", "name": "User"}}],
        "ccRecipients": [],
        "receivedDateTime": "2026-10-20T09:00:00Z",
        "body": {"contentType": "html", "content": "<p>Please pay.</p>"},
        "isRead": False,
        "importance": "normal",
        "hasAttachments": False,
        "categories": [],
        "flag": {"flagStatus": "notFlagged"},
        "conversationId": "conv-1",
        "internetMessageHeaders": [
            {"name": "Received", "value": "from mail.contoso.com"},
            {"name": "Authentication-Results", "value": auth_results},
            # A sender can write this into the message; it is below Microsoft's.
            {"name": "Authentication-Results", "value": "dkim=pass header.d=contoso.com"},
        ],
    }


# ── The URL the real msgraph pipeline sends ─────────────────────────────


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def _read_through_the_real_sdk(message_id: str = MESSAGE_ID, **kwargs) -> tuple[str, dict]:
    """``read_message`` against the real msgraph client; (URL sent, result)."""
    sent: list[str] = []

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(str(request.url))
        if "$filter" in request.url.params:
            return httpx.Response(200, json={"value": [{"id": MESSAGE_ID}]})
        return httpx.Response(200, json=_graph_message())

    http_client = GraphClientFactory.create_with_default_middleware(
        options=getattr(_gra, "options", None),
        client=httpx.AsyncClient(transport=httpx.MockTransport(transport)),
    )
    auth = AzureIdentityAuthenticationProvider(
        _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
    )
    client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))
    result = asyncio.run(mail_read.read_message(client, message_id, **kwargs))
    assert len(sent) == (2 if "@" in message_id else 1)
    return sent[-1], result


def _query(url: str) -> dict[str, str]:
    return {k: v[0] for k, v in parse_qs(urlsplit(url).query).items()}


def _selected(url: str) -> set[str]:
    return set(unquote(_query(url)["$select"]).split(","))


def test_the_default_read_sends_no_select():
    url, result = _read_through_the_real_sdk()
    assert urlsplit(url).query == ""
    assert "sender_authentication" not in result


def test_the_opt_in_read_selects_the_headers_and_every_field_it_formats():
    url, result = _read_through_the_real_sdk(include_sender_auth=True)
    selected = _selected(url)
    assert "internetMessageHeaders" in selected
    # Everything the default read showed is still there.
    assert {"subject", "from", "body", "toRecipients", "receivedDateTime"} <= selected
    assert result["subject"] == "Invoice 4471"
    assert result["body"] == "<p>Please pay.</p>"
    assert result["sender_authentication"]["authenticated"] is True


def test_the_opt_in_read_keeps_the_deferred_send_expansion():
    url, result = _read_through_the_real_sdk(include_sender_auth=True, include_deferred_send=True)
    query = _query(url)
    assert unquote(query["$expand"]).startswith("singleValueExtendedProperties($filter=")
    assert "internetMessageHeaders" in _selected(url)
    assert result["sender_authentication"]["authenticated"] is True
    assert "deferred_send_datetime" in result


def test_the_deferred_send_read_alone_is_unchanged():
    url, _ = _read_through_the_real_sdk(include_deferred_send=True)
    assert "$select" not in _query(url)


def test_the_verdict_comes_from_the_deserialized_headers():
    """The SDK path reads ``internet_message_headers`` the SDK filled from JSON.

    A formatter reading a misspelt attribute off a MagicMock would pass a mock
    test; this one goes through kiota's deserializer.
    """
    _, result = _read_through_the_real_sdk(include_sender_auth=True)
    assert result["sender_authentication"] == {
        "authenticated": True,
        "from_domain": "contoso.com",
        "dkim": "pass",
        "dmarc": "pass",
        "compauth": "pass",
        "reason": "dmarc=pass",
    }


def test_the_sdk_and_batch_paths_add_the_same_field_in_the_same_place():
    """outlook_read_messages promises outlook_read_message's shape, byte for byte.

    Compared on the field this adds. The two paths already differ on `received`
    for a kiota-deserialized message (`str(datetime)` against Graph's string),
    which is outside this change.
    """
    _, single = _read_through_the_real_sdk(include_sender_auth=True)
    batched = mail_read._format_read_message_from_raw(
        _graph_message(), "text", False, False, include_sender_auth=True
    )
    assert list(batched)[-1] == list(single)[-1] == "sender_authentication"
    assert json.dumps(batched["sender_authentication"]) == json.dumps(
        single["sender_authentication"]
    )


# ── The $batch sub-requests ─────────────────────────────────────────────


def _fake_graph_client() -> MagicMock:
    client = MagicMock()
    client.credential.get_token = MagicMock(return_value=MagicMock(token="fake-token"))
    return client


def _batch_response(body: dict) -> MagicMock:
    resp = MagicMock()
    resp.raise_for_status = MagicMock()
    resp.json = MagicMock(
        return_value={"responses": [{"id": "0", "status": 200, "headers": {}, "body": body}]}
    )
    return resp


@pytest.fixture
def batch_post():
    with patch("outlook_mcp.tools.mail_read.httpx.AsyncClient") as cm:
        client_instance = MagicMock()
        client_instance.post = AsyncMock(return_value=_batch_response(_graph_message(MS_FAIL)))
        cm.return_value.__aenter__ = AsyncMock(return_value=client_instance)
        cm.return_value.__aexit__ = AsyncMock(return_value=False)
        yield client_instance.post


def _subrequest_url(post: AsyncMock) -> str:
    return json.loads(post.await_args.kwargs["content"])["requests"][0]["url"]


@pytest.mark.asyncio
async def test_the_default_batch_read_is_unchanged(batch_post):
    result = await mail_read.read_messages(_fake_graph_client(), [MESSAGE_ID])
    assert _subrequest_url(batch_post) == "/me/messages/AAMkAGI2THVSAAA%3D"
    assert "sender_authentication" not in result["messages"][0]


@pytest.mark.asyncio
async def test_the_opt_in_batch_read_selects_the_headers(batch_post):
    result = await mail_read.read_messages(
        _fake_graph_client(), [MESSAGE_ID], include_sender_auth=True
    )
    url = _subrequest_url(batch_post)
    assert urlsplit(url).path == "/me/messages/AAMkAGI2THVSAAA%3D"
    assert _selected(url) == set(mail_read._SENDER_AUTH_SELECT.split(","))
    verdict = result["messages"][0]["sender_authentication"]
    assert verdict["authenticated"] is False
    assert verdict["reason"] == "dmarc=fail"


@pytest.mark.asyncio
async def test_the_opt_in_batch_read_keeps_the_deferred_send_expansion(batch_post):
    await mail_read.read_messages(
        _fake_graph_client(), [MESSAGE_ID], include_sender_auth=True, include_deferred_send=True
    )
    url = _subrequest_url(batch_post)
    assert unquote(_query(url)["$expand"]).startswith("singleValueExtendedProperties($filter=")
    assert "internetMessageHeaders" in _selected(url)


# ── What an MCP client receives ─────────────────────────────────────────


async def _call(tool: str, arguments: dict) -> dict:
    from outlook_mcp import server as server_mod

    async with Client(server_mod.mcp) as client:
        result = await client.call_tool(tool, arguments)
    assert not result.is_error, result.content
    return json.loads(result.content[0].text)


@pytest.mark.asyncio
async def test_outlook_read_message_delivers_the_verdict_to_the_client():
    from outlook_mcp import server as server_mod

    # The SDK object the mock returns is built the way kiota builds it.
    message = JsonParseNode(_graph_message(MS_FAIL)).get_object_value(Message)
    graph = MagicMock()
    graph.sdk_client.me.messages.by_message_id.return_value.get = AsyncMock(return_value=message)

    with patch.object(server_mod, "_get_graph_client", return_value=graph):
        opted_in = await _call(
            "outlook_read_message", {"message_id": MESSAGE_ID, "include_sender_auth": True}
        )
        default = await _call("outlook_read_message", {"message_id": MESSAGE_ID})

    assert opted_in["sender_authentication"] == {
        "authenticated": False,
        "from_domain": "contoso.com",
        "dkim": None,
        "dmarc": "fail",
        "compauth": "fail",
        "reason": "dmarc=fail",
    }
    assert "sender_authentication" not in default
    assert {k: v for k, v in opted_in.items() if k != "sender_authentication"} == default


@pytest.mark.asyncio
async def test_outlook_read_messages_delivers_the_verdict_to_the_client(batch_post):
    from outlook_mcp import server as server_mod

    with patch.object(server_mod, "_get_graph_client", return_value=_fake_graph_client()):
        result = await _call(
            "outlook_read_messages", {"message_ids": [MESSAGE_ID], "include_sender_auth": True}
        )

    assert result["messages"][0]["sender_authentication"]["reason"] == "dmarc=fail"
    assert "internetMessageHeaders" in _selected(_subrequest_url(batch_post))


def test_the_message_id_read_selects_sender_authentication_headers():
    url, result = _read_through_the_real_sdk("<invoice@example.com>", include_sender_auth=True)
    assert "internetMessageHeaders" in _selected(url)
    assert result["sender_authentication"]["authenticated"] is True
