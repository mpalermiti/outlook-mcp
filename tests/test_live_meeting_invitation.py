"""Read-only checks of Graph's invitation-to-calendar relationship.

Both guards list ID-only messages filtered on their message class
(PidTagMessageClass), as ``find_events_by_key`` filters events on an extended
property. Plain mail is checked on the first 50 ``IPM.Note`` messages.
Invitations are rarer, so their guard follows ``@odata.nextLink`` through
``IPM.Schedule.Meeting.Request`` messages, at most 20 pages, until it finds one
with a surviving event, expanding at most 50 invitations. A listed invitation
the SDK does not build as an ``EventMessageRequest`` fails the run. Each guard
skips when the mailbox lacks what it needs, and the invitation skip says how
much it read. No fixture sends or changes anything.
"""

from contextlib import aclosing

import pytest
from msgraph.generated.models.event_message_request import EventMessageRequest
from msgraph.generated.users.item.messages.item.message_item_request_builder import (
    MessageItemRequestBuilder,
)
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

from outlook_mcp.pagination import build_request_config
from outlook_mcp.tools.calendar_read import get_event

pytestmark = [pytest.mark.live, pytest.mark.asyncio]

PAGE_SIZE = 50
MAX_PAGES = 20
MAX_INVITATIONS = 50
MESSAGE_CLASS_PROPERTY_ID = "String 0x001A"  # PidTagMessageClass
PLAIN_MAIL = "IPM.Note"
INVITATION = "IPM.Schedule.Meeting.Request"


def message_class_filter(message_class: str) -> str:
    return (
        f"singleValueExtendedProperties/Any(ep: ep/id eq '{MESSAGE_CLASS_PROPERTY_ID}' "
        f"and ep/value eq '{message_class}')"
    )


async def message_pages(sdk, message_class: str, max_pages: int = 1):
    """Yield ID-only pages of one message class, following @odata.nextLink up to max_pages."""
    page = await sdk.me.messages.get(
        request_configuration=build_request_config(
            MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters,
            {"$filter": message_class_filter(message_class), "$select": "id", "$top": PAGE_SIZE},
        )
    )
    pages = 0
    while page is not None:
        pages += 1
        yield page.value or []
        if pages >= max_pages or not page.odata_next_link:
            return
        page = await sdk.me.messages.with_url(page.odata_next_link).get()


async def invitation_pages(sdk):
    """Yield pages of invitations, failing on any the SDK did not build as one.

    The filter selects invitations by message class, so a listed message that
    is not an ``EventMessageRequest`` means Graph no longer marks invitations'
    ``@odata.type`` in an ID-only listing, the premise these guards rely on.
    """
    async with aclosing(message_pages(sdk, INVITATION, MAX_PAGES)) as pages:
        async for messages in pages:
            for message in messages:
                if not isinstance(message, EventMessageRequest):
                    pytest.fail(
                        f"Graph listed {INVITATION} message {message.id} without "
                        "@odata.type #microsoft.graph.eventMessageRequest, so the SDK "
                        f"built a {type(message).__name__}, not an EventMessageRequest"
                    )
            yield messages


@pytest.fixture
async def plain_mail_sample(real_graph_client):
    async with aclosing(message_pages(real_graph_client.sdk_client, PLAIN_MAIL)) as pages:
        return await anext(pages, [])


async def test_plain_mail_has_no_event_rather_than_a_graph_error(
    real_graph_client, plain_mail_sample
):
    """The subtype expand must be accepted even when the message is ordinary mail."""
    message = next((m for m in plain_mail_sample if m.id), None)
    if message is None:
        pytest.skip(f"No {PLAIN_MAIL} message among the first {PAGE_SIZE} listed")

    with pytest.raises(ValueError, match="no associated calendar event"):
        await get_event(real_graph_client.sdk_client, message_id=message.id)


async def test_invitation_resolves_to_the_mailbox_event(real_graph_client):
    """An expanded invitation must identify an event readable through this user's /me."""
    sdk = real_graph_client.sdk_client
    pages = invitations = 0
    async with aclosing(invitation_pages(sdk)) as listing:
        async for messages in listing:
            pages += 1
            for message in messages:
                if not message.id:
                    continue
                invitations += 1
                if await _resolves_to_current_event(sdk, message):
                    return
                if invitations >= MAX_INVITATIONS:
                    pytest.skip(
                        f"No invitation with a calendar event among {invitations} "
                        f"{INVITATION} messages (expansion cap reached; {pages} pages read)"
                    )

    pytest.skip(
        f"No invitation with a calendar event among {invitations} {INVITATION} messages "
        f"({pages} of at most {MAX_PAGES} pages read)"
    )


async def _resolves_to_current_event(sdk, message) -> bool:
    """Assert one invitation's result; False when its meeting has no calendar event."""
    expanded = await sdk.me.messages.by_message_id(message.id).get(
        request_configuration=build_request_config(
            MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters,
            {"$expand": "microsoft.graph.eventMessage/event"},
        )
    )
    associated = getattr(expanded, "event", None)
    if associated is None or not associated.id:
        return False  # Cancelled or removed meetings need not have a calendar event.

    current = await sdk.me.events.by_event_id(associated.id).get()
    result = await get_event(sdk, message_id=message.id)

    assert result["id"] == current.id == associated.id
    assert result["is_organizer"] is current.is_organizer
    assert result["allow_new_time_proposals"] is current.allow_new_time_proposals
    assert result["is_cancelled"] is current.is_cancelled
    assert result["invitation"] == {
        "message_id": message.id,
        "is_out_of_date": expanded.is_out_of_date,
    }
    return True
