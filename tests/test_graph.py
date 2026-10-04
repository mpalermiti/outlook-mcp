"""Tests for Graph client factory."""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.errors import AuthRequiredError
from outlook_mcp.graph import GraphClient


def test_graph_client_requires_credential():
    """GraphClient raises without credential."""
    with pytest.raises(AuthRequiredError):
        GraphClient(credential=None)


def test_graph_client_init():
    """GraphClient initializes with a credential and creates sdk_client."""
    mock_credential = MagicMock()
    client = GraphClient(credential=mock_credential)
    assert client.sdk_client is not None


# ── The SDK path only authenticates requests to Graph ───────────────────────
# Kiota's auth provider attaches a token to whatever URL a request names, and
# with no allow-list "every host is valid". Six call sites follow a Graph
# `@odata.nextLink` through `with_url`. Those links come from Graph today, but
# the guarantee belongs in the client, not in where the links come from — the
# raw delta path already has it (`require_graph_url`).

import time  # noqa: E402

from azure.core.credentials import AccessToken  # noqa: E402
from kiota_abstractions.request_information import RequestInformation  # noqa: E402


class _RecordingCredential:
    """Hands out a token for whatever scope it is asked, and remembers the asks."""

    def __init__(self):
        self.scopes: list[tuple[str, ...]] = []

    def get_token(self, *scopes, **_kwargs):
        self.scopes.append(scopes)
        return AccessToken("SECRET-GRAPH-TOKEN", int(time.time()) + 3600)


def _auth_provider(monkeypatch):
    """The provider GraphClient really builds, captured on its way into the adapter."""
    from outlook_mcp import graph

    captured = {}
    real_adapter = graph.GraphRequestAdapter

    def recording_adapter(auth_provider, *args, **kwargs):
        captured["provider"] = auth_provider
        return real_adapter(auth_provider, *args, **kwargs)

    monkeypatch.setattr(graph, "GraphRequestAdapter", recording_adapter)
    credential = _RecordingCredential()
    GraphClient(credential=credential)
    return captured["provider"], credential


async def _authenticate(provider, url: str) -> RequestInformation:
    request = RequestInformation()
    request.url = url
    await provider.authenticate_request(request)
    return request


@pytest.mark.asyncio
async def test_a_graph_request_gets_the_token(monkeypatch):
    provider, credential = _auth_provider(monkeypatch)

    request = await _authenticate(provider, "https://graph.microsoft.com/v1.0/me/messages")

    assert request.headers.contains("Authorization")
    assert credential.scopes == [("https://graph.microsoft.com/.default",)]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "url",
    [
        "https://evil.example/v1.0/me/messages?$skiptoken=x",
        "https://graph.microsoft.com.evil.example/v1.0/me/messages",
        # Another Microsoft resource: a token minted for it is still a token.
        "https://outlook.office.com/api/v2.0/me/messages",
    ],
)
async def test_any_other_host_gets_no_token_and_none_is_minted(monkeypatch, url):
    provider, credential = _auth_provider(monkeypatch)

    request = await _authenticate(provider, url)

    assert not request.headers.contains("Authorization")
    assert credential.scopes == []
