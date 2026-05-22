"""Microsoft Graph client factory."""

from __future__ import annotations

from typing import Any

from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphRequestAdapter, GraphServiceClient

from outlook_mcp.errors import AuthRequiredError


class GraphClient:
    """Wrapper around the Microsoft Graph SDK client."""

    def __init__(self, credential: Any) -> None:
        if credential is None:
            raise AuthRequiredError()
        # Disable CAE (Continuous Access Evaluation) — the default enables it,
        # which forces a fresh interactive auth flow instead of using the
        # cached token from `outlook-mcp auth`.
        auth_provider = AzureIdentityAuthenticationProvider(
            credential, is_cae_enabled=False
        )
        request_adapter = GraphRequestAdapter(auth_provider)
        self.sdk_client = GraphServiceClient(request_adapter=request_adapter)
        # Retained so delta-query tools can mint raw bearer tokens for direct
        # httpx calls to Graph's *delta endpoints — the SDK's typed delta
        # builders rebuild URL templates from query-parameter dataclasses and
        # silently drop the ``@removed`` annotation we need to surface to
        # callers, so the delta module bypasses the SDK.
        self.credential = credential
