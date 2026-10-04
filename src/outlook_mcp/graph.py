"""Microsoft Graph client factory."""

from __future__ import annotations

from typing import Any

from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphRequestAdapter, GraphServiceClient

from outlook_mcp.errors import AuthRequiredError

# The only host the SDK client may authenticate a request to. The raw delta
# path pins the same host for the same reason (`tools/_delta.py`).
GRAPH_HOST = "graph.microsoft.com"


class GraphClient:
    """Wrapper around the Microsoft Graph SDK client."""

    def __init__(self, credential: Any) -> None:
        if credential is None:
            raise AuthRequiredError()
        # Disable CAE (Continuous Access Evaluation) — the default enables it,
        # which forces a fresh interactive auth flow instead of using the
        # cached token from `outlook-mcp auth`. It would also persist through
        # a SECOND cache: every cache name gets a signal-file suffix in
        # ~/.IdentityService (.nocae for this non-CAE cache, .cae for a CAE
        # one), and the two suffixes are two signal files writing the same
        # host-wide Keychain item — each believing it owns the item, they
        # clobber each other's writes. One non-CAE cache, one signal file:
        # ~/.IdentityService/outlook-mcp.nocae.
        #
        # `allowed_hosts` because kiota's default is "every host is valid": the
        # provider asks the credential for a token scoped to whatever host a
        # request names and attaches it. Every SDK request here goes to Graph,
        # including the `@odata.nextLink` pages followed with `with_url`, so a
        # request anywhere else is one this server never meant to authenticate
        # — it goes out with no token, and none is minted for it.
        auth_provider = AzureIdentityAuthenticationProvider(
            credential, is_cae_enabled=False, allowed_hosts=[GRAPH_HOST]
        )
        request_adapter = GraphRequestAdapter(auth_provider)
        self.sdk_client = GraphServiceClient(request_adapter=request_adapter)
        # Retained so delta-query tools can mint raw bearer tokens for direct
        # httpx calls to Graph's *delta endpoints — the SDK's typed delta
        # builders rebuild URL templates from query-parameter dataclasses and
        # silently drop the ``@removed`` annotation we need to surface to
        # callers, so the delta module bypasses the SDK.
        self.credential = credential
