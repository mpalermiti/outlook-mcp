"""Invitation resolution and meeting metadata through the real SDK, with synthetic HTTP."""

from contextlib import asynccontextmanager

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as adapter
from msgraph_core import GraphClientFactory

from outlook_mcp.tools.calendar_read import get_event


class Credential:
    def get_token(self, *args, **kwargs):
        return AccessToken("synthetic-token", 9_999_999_999)


@asynccontextmanager
async def graph(handler):
    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as raw:
        http = GraphClientFactory.create_with_default_middleware(
            options=adapter.options, client=raw
        )
        auth = AzureIdentityAuthenticationProvider(
            Credential(), scopes=["https://graph.microsoft.com/.default"]
        )
        yield GraphServiceClient(request_adapter=adapter.GraphRequestAdapter(auth, http))


def event(**extra):
    return {"id": "recipient-event", "subject": "Planning", "type": "singleInstance", **extra}


@pytest.mark.asyncio
@pytest.mark.parametrize("value", [True, False, None])
async def test_existing_event_read_has_metadata_without_changing_request(value):
    requests = []

    def handle(request):
        requests.append(request)
        return httpx.Response(
            200, json=event(isOrganizer=value, allowNewTimeProposals=value, isCancelled=value)
        )

    async with graph(handle) as client:
        result = await get_event(client, "recipient-event")
    assert result["is_organizer"] is value
    assert result["allow_new_time_proposals"] is value
    assert result["is_cancelled"] is value
    assert "invitation" not in result
    assert len(requests) == 1 and requests[0].method == "GET"
    assert requests[0].url.path == "/v1.0/me/events/recipient-event"
    assert dict(requests[0].url.params) == {
        "$expand": (
            "singleValueExtendedProperties($filter=id eq 'String "
            "{51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey')"
        )
    }


@pytest.mark.asyncio
async def test_attendee_proposal_is_preserved_with_its_zone():
    proposal = {
        "start": {"dateTime": "2027-01-12T10:00:00", "timeZone": "W. Europe Standard Time"},
        "end": {"dateTime": "2027-01-12T10:30:00", "timeZone": "W. Europe Standard Time"},
    }
    async with graph(
        lambda _: httpx.Response(
            200,
            json=event(
                attendees=[
                    {
                        "emailAddress": {"name": "Morgan", "address": "morgan@example.com"},
                        "proposedNewTime": proposal,
                    },
                    {"emailAddress": {"address": "alex@example.com"}},
                ]
            ),
        )
    ) as client:
        result = await get_event(client, "recipient-event")
    assert result["attendees"][0]["proposed_new_time"] == {
        "start": {"date_time": "2027-01-12T10:00:00", "time_zone": "W. Europe Standard Time"},
        "end": {"date_time": "2027-01-12T10:30:00", "time_zone": "W. Europe Standard Time"},
    }
    assert result["attendees"][1]["proposed_new_time"] is None


@pytest.mark.asyncio
@pytest.mark.parametrize("outdated", [True, False, None])
async def test_invitation_resolves_recipient_event_then_reads_current_details(outdated):
    requests = []

    def handle(request):
        requests.append(request)
        if request.url.path == "/v1.0/me/messages/invitation-id":
            assert dict(request.url.params) == {"$expand": "microsoft.graph.eventMessage/event"}
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.eventMessageRequest",
                    "id": "invitation-id",
                    "isOutOfDate": outdated,
                    "event": event(subject="Old invitation"),
                },
            )
        assert request.url.path == "/v1.0/me/events/recipient-event"
        return httpx.Response(200, json=event(subject="Current meeting", isOrganizer=False))

    async with graph(handle) as client:
        result = await get_event(client, message_id="invitation-id")
    assert result["id"] == "recipient-event" and result["subject"] == "Current meeting"
    assert result["invitation"] == {"message_id": "invitation-id", "is_out_of_date": outdated}
    assert len(requests) == 2 and all(r.method == "GET" for r in requests)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "body",
    [
        {"@odata.type": "#microsoft.graph.message", "id": "regular-mail"},
        {
            "@odata.type": "#microsoft.graph.eventMessageRequest",
            "id": "invitation-id",
            "event": None,
        },
        {"@odata.type": "#microsoft.graph.eventMessageRequest", "event": {}},
    ],
)
async def test_missing_association_is_refused_without_subject_search(body):
    requests = []

    def handle(request):
        requests.append(request)
        return httpx.Response(200, json=body)

    async with graph(handle) as client:
        with pytest.raises(
            ValueError,
            match=(
                r"no associated calendar event \(the meeting may have been cancelled or removed\)"
            ),
        ):
            await get_event(client, message_id="invitation-id")
    assert len(requests) == 1


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "arguments",
    [{}, {"event_id": "one", "message_id": "two"}, {"event_id": ""}, {"message_id": "../bad"}],
)
async def test_choose_exactly_one_valid_identifier_before_graph(arguments):
    def handle(request):
        pytest.fail("invalid input reached Graph")

    async with graph(handle) as client:
        with pytest.raises(ValueError):
            await get_event(client, **arguments)


@pytest.mark.asyncio
async def test_message_id_header_uses_existing_exact_resolver():
    requests = []

    def handle(request):
        requests.append(request)
        if request.url.path == "/v1.0/me/messages":
            assert request.url.params["$filter"] == "internetMessageId eq '<invite@example.com>'"
            return httpx.Response(200, json={"value": [{"id": "invitation-id"}]})
        if request.url.path == "/v1.0/me/messages/invitation-id":
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.eventMessageRequest",
                    "id": "invitation-id",
                    "event": event(),
                },
            )
        return httpx.Response(200, json=event())

    async with graph(handle) as client:
        result = await get_event(client, message_id="<invite@example.com>")
    assert result["invitation"]["message_id"] == "invitation-id"
    assert len(requests) == 3


@pytest.mark.asyncio
async def test_server_forwards_invitation_to_the_real_reader(monkeypatch):
    from types import SimpleNamespace

    from outlook_mcp import server

    def handle(request):
        if request.url.path.startswith("/v1.0/me/messages/"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.eventMessageRequest",
                    "id": "invitation-id",
                    "event": event(),
                },
            )
        return httpx.Response(200, json=event(isOrganizer=False))

    async with graph(handle) as client:
        monkeypatch.setattr(
            server, "_get_graph_client", lambda ctx: SimpleNamespace(sdk_client=client)
        )
        result = await server.outlook_get_event(None, message_id="invitation-id")
    assert result["id"] == "recipient-event" and result["is_organizer"] is False
