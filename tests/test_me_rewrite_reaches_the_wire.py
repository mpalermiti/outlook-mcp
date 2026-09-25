"""Tier 4 of the silent-no-op audit: does ``/me`` actually reach the wire as ``/me``?

Every ``/me`` request this server makes is built by the SDK as
``/users/me-token-to-replace`` and rewritten to ``/me`` by a middleware that
msgraph-core installs. kiota 1.13.0 moved per-request options from a
monkey-patched ``request.options`` attribute to ``request.extensions``; the
msgraph-core release available at the time still read the old attribute, so
the rewrite silently stopped firing and Graph received the literal placeholder:
``400 me-token-to-replace is invalid`` on every call (#80).

Nothing in the offline suite could see it — the mocks never reach the
middleware — and the fresh-install canary imported the package, counted the
tools and never sent a request. The lock file pinned a working kiota, so every
developer and CI run was green while every new user's install was broken.

So this test builds the client exactly the way ``msgraph`` does, hands the
factory an httpx client whose transport is a mock, and asserts on the URL the
real middleware pipeline hands to that transport. It runs against whatever
kiota/msgraph-core the environment resolved, which is the point: under the
lock it pins the rewrite; in the fresh-install and published-install jobs it
fails the moment a dependency resolves outside the range that works.
"""

from __future__ import annotations

import inspect

import httpx
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def sent_url_for_me() -> str:
    """The absolute URL the real middleware pipeline sends for ``client.me.get()``.

    Importable by the CI canaries, which run it against a fresh resolution.
    """
    sent: list[str] = []

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(str(request.url))
        return httpx.Response(200, json={"id": "probe"})

    factory = GraphClientFactory.create_with_default_middleware
    kwargs = {}
    params = inspect.signature(factory).parameters
    # msgraph registers the /me rewrite as a default request option on its
    # adapter module; hand the same options to the factory so the middleware
    # under test is the one users actually get.
    if "options" in params:
        kwargs["options"] = getattr(_gra, "options", None)
    if "client" in params:
        kwargs["client"] = httpx.AsyncClient(transport=httpx.MockTransport(transport))
    http_client = factory(**kwargs)
    auth = AzureIdentityAuthenticationProvider(
        _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
    )
    client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))

    import asyncio

    async def call():
        await client.me.get()

    asyncio.run(call())
    assert sent, "the mock transport was never reached — the probe is wired to the wrong client"
    return sent[-1]


def test_me_is_sent_as_me_not_as_the_placeholder():
    url = sent_url_for_me()
    assert "me-token-to-replace" not in url, (
        f"/me reached the wire as the SDK placeholder: {url}\n"
        "The msgraph-core rewrite did not fire. This is what a fresh install on "
        "kiota >= 1.13 with msgraph-core <= 1.5.1 does (#80). Check the resolved "
        "versions against the caps in pyproject.toml."
    )
    assert url.endswith("/v1.0/me"), url


if __name__ == "__main__":  # so CI can run it as a script against a fresh venv
    print(sent_url_for_me())
    test_me_is_sent_as_me_not_as_the_placeholder()
    print("OK: /me reaches the wire as /me")
