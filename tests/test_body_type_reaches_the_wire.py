"""Does ``format="text"`` ask Graph for a text body? (#124)

Graph returns a message body in the type it was stored in, which for most mail
is HTML, unless the request carries ``Prefer: outlook.body-content-type="text"``.
Without that header ``format="text"`` returned the same HTML as ``format="html"``:
layout code the model paid for on every read, and text the sender hid
(``display:none``, comments) presented as readable. A mock of the SDK chain
cannot see request headers, so this builds the client the way ``msgraph``
does, hands it a mock transport, and reads the headers the real pipeline
sends, as ``test_message_id_reaches_the_wire.py`` does for the URL.

``read_message`` reaches Graph three ways (plain GET, ``$select`` for sender
auth, a raw URL for the deferred-send expansion); each is checked.
"""

from __future__ import annotations

import inspect

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory

from outlook_mcp.tools.mail_read import read_message

TEXT_BODY = 'outlook.body-content-type="text"'


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def _graph(body: dict) -> tuple[GraphServiceClient, list[httpx.Request]]:
    """A real SDK client whose transport answers every request with ``body``."""
    sent: list[httpx.Request] = []

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(request)
        return httpx.Response(200, json=body)

    factory = GraphClientFactory.create_with_default_middleware
    kwargs = {}
    params = inspect.signature(factory).parameters
    if "options" in params:
        kwargs["options"] = getattr(_gra, "options", None)
    if "client" in params:
        kwargs["client"] = httpx.AsyncClient(transport=httpx.MockTransport(transport))
    auth = AzureIdentityAuthenticationProvider(
        _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
    )
    client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, factory(**kwargs)))
    return client, sent


_MESSAGE = {"id": "AAMk=", "subject": "s", "body": {"contentType": "text", "content": "hi"}}

_PATHS = [
    pytest.param({}, id="plain"),
    pytest.param({"include_sender_auth": True}, id="sender-auth-select"),
    pytest.param({"include_deferred_send": True}, id="deferred-send-raw-url"),
]


@pytest.mark.parametrize("extra", _PATHS)
@pytest.mark.asyncio
async def test_text_asks_graph_for_a_text_body(extra):
    client, sent = _graph(_MESSAGE)

    await read_message(client, "AAMk=", format="text", **extra)

    assert len(sent) == 1
    assert sent[0].headers.get("Prefer") == TEXT_BODY


@pytest.mark.parametrize("fmt", ["html", "full"])
@pytest.mark.parametrize("extra", _PATHS)
@pytest.mark.asyncio
async def test_html_and_full_keep_the_stored_body(fmt, extra):
    client, sent = _graph(_MESSAGE)

    await read_message(client, "AAMk=", format=fmt, **extra)

    assert "Prefer" not in sent[0].headers


@pytest.mark.parametrize("fmt", ["html", "full"])
@pytest.mark.asyncio
async def test_concise_previews_the_text_body_whatever_the_format(fmt):
    """``concise`` returns no body, only ``body_preview``: preview the text one."""
    client, sent = _graph(_MESSAGE)

    await read_message(client, "AAMk=", format=fmt, concise=True)

    assert sent[0].headers.get("Prefer") == TEXT_BODY
