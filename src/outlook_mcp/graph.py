"""Microsoft Graph client factory."""

from __future__ import annotations

from typing import Any

from msgraph import GraphServiceClient

from outlook_mcp.errors import AuthRequiredError


class GraphClient:
    """Wrapper around the Microsoft Graph SDK client."""

    def __init__(self, credential: Any) -> None:
        if credential is None:
            raise AuthRequiredError()
        self.sdk_client = GraphServiceClient(credentials=credential)
