"""Owner metadata from real SDK deserialization, without changing consumer requests."""

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory

from outlook_mcp.tools.user import list_calendars


class Credential:
    def get_token(self, *args, **kwargs):
        return AccessToken("synthetic-token", 9_999_999_999)


@pytest.mark.asyncio
@pytest.mark.parametrize("private_access", [True, False, None])
async def test_owner_metadata_survives_deserialization_and_pagination(private_access):
    requests = []
    next_url = "https://graph.microsoft.com/v1.0/me/calendars?$skiptoken=synthetic"

    def handle(request):
        requests.append(request)
        if len(requests) == 1:
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "first",
                            "name": "Calendar",
                            "owner": {"name": "Alex", "address": "alex@example.com"},
                        }
                    ],
                    "@odata.nextLink": next_url,
                },
            )
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "second",
                        "name": "Calendar",
                        "canEdit": False,
                        "owner": {
                            "name": "Morgan\r\n\x1b[31m",
                            "address": "morgan@example.com\x00",
                        },
                        "canViewPrivateItems": private_access,
                    }
                ]
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handle)) as raw:
        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options, client=raw
        )
        auth = AzureIdentityAuthenticationProvider(
            Credential(), scopes=["https://graph.microsoft.com/.default"]
        )
        client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))
        result = await list_calendars(client)

    assert result["count"] == 2
    first, second = result["calendars"]
    assert first["owner"] == {"name": "Alex", "address": "alex@example.com"}
    assert first["can_view_private_items"] is None
    assert second["owner"] == {"name": "Morgan  ", "address": "morgan@example.com"}
    assert second["can_edit"] is False
    assert second["can_view_private_items"] is private_access
    # Exactly the previous consumer request and its nextLink: no $select or extra lookup.
    assert len(requests) == 2
    assert requests[0].method == "GET" and requests[0].url.path == "/v1.0/me/calendars"
    assert dict(requests[0].url.params) == {"$top": "100"}
    assert str(requests[1].url) == next_url


@pytest.mark.asyncio
@pytest.mark.parametrize("owner,expected", [(None, None), ({}, {"name": None, "address": None})])
async def test_missing_owner_metadata_is_not_invented(owner, expected):
    from types import SimpleNamespace
    from unittest.mock import AsyncMock

    from msgraph.generated.models.calendar import Calendar
    from msgraph.generated.models.email_address import EmailAddress

    calendar = Calendar(
        id="synthetic", name="Calendar", owner=EmailAddress(**owner) if owner is not None else None
    )
    graph = SimpleNamespace(
        me=SimpleNamespace(
            calendars=SimpleNamespace(
                get=AsyncMock(return_value=SimpleNamespace(value=[calendar], odata_next_link=None)),
            )
        )
    )
    result = await list_calendars(graph)
    assert result["calendars"][0]["owner"] == expected
    assert result["calendars"][0]["can_view_private_items"] is None
