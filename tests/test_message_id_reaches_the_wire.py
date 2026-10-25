"""Does a Message-ID reach Graph's ``$filter`` exactly as the caller gave it?

``outlook_read_message`` takes a Message-ID header value in place of a Graph id
and finds the message with ``$filter=internetMessageId eq '...'``. Message-IDs
carry ``+``, ``/`` and ``=`` (55 of 200 real ones had a ``/``, 7 a ``+``), and in
a query string a literal ``+`` is a space: an encoder that left it alone would
send a different Message-ID and find nothing, with no error. The mock suite
only sees the filter string handed to the SDK, so this builds the client the
way ``msgraph`` does, hands it a mock transport, and decodes the URL the real
pipeline sends, as ``test_me_rewrite_reaches_the_wire.py`` does for ``/me``.
"""

from __future__ import annotations

import inspect
from urllib.parse import parse_qs, urlsplit

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

ME = "https://graph.microsoft.com/v1.0/me"


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def _graph(responses: list[dict]) -> tuple[GraphServiceClient, list[httpx.URL]]:
    """A real SDK client whose transport answers with ``responses`` in order."""
    sent: list[httpx.URL] = []
    queue = list(responses)

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(request.url)
        return httpx.Response(200, json=queue.pop(0))

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


@pytest.mark.parametrize(
    "message_id",
    [
        "<CAF1a+b/c=d@mail.example.com>",
        "<o'brien&co#1%20x@example.com>",
        "1234.5678@example.com",
    ],
)
@pytest.mark.asyncio
async def test_the_message_id_reaches_the_filter_intact(message_id):
    client, sent = _graph([{"value": [{"id": "AAMkFound="}]}, {"id": "AAMkFound=", "subject": "s"}])

    result = await read_message(client, message_id)

    lookup = sent[0]
    assert str(lookup).startswith(f"{ME}/messages?"), lookup
    literal = message_id.replace("'", "''")
    assert parse_qs(urlsplit(str(lookup)).query) == {
        "$filter": [f"internetMessageId eq '{literal}'"],
        "$select": ["id"],
        "$top": ["10"],
    }
    assert str(sent[1]) == f"{ME}/messages/AAMkFound%3D"
    assert result["id"] == "AAMkFound="
