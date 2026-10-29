"""The draft and calendar guards decide on fields Graph has to actually return.

`require_draft` lets a call through only when `isDraft` comes back `True`, and
the calendar `mail_send` gate reads `attendees` off a plain GET of the event.
The offline suite mocks both reads, so it shows what the code does with a value
and never that Graph puts the value there. A guard reading a field Graph leaves
out fails safe for drafts (every call refused) and open for the calendar
(nothing gated), and neither would show offline.

Read-only, like the rest of the live tier: GETs only. Each check skips by name
when the mailbox has nothing to check it against.
"""

import pytest
from kiota_abstractions.base_request_configuration import RequestConfiguration

from outlook_mcp.tools.calendar_write import _has_attendees
from outlook_mcp.tools.mail_drafts import require_draft

pytestmark = [pytest.mark.live, pytest.mark.asyncio]


async def _newest_message_id(sdk, folder: str) -> str | None:
    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (  # noqa: E501
        MessagesRequestBuilder,
    )

    query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(top=1, select=["id"])
    page = await sdk.me.mail_folders.by_mail_folder_id(folder).messages.get(
        request_configuration=RequestConfiguration(query_parameters=query)
    )
    return page.value[0].id if page and page.value else None


async def _newest_draft_id(sdk) -> str | None:
    """The newest item in Drafts that Graph calls a draft.

    The Drafts folder can hold messages that are not drafts, such as an old sent
    message moved there. When one was the newest item, this check failed on the
    mailbox, not the guard. Look a little further, and skip when there is no
    draft at all.
    """
    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (  # noqa: E501
        MessagesRequestBuilder,
    )

    query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=25, select=["id", "isDraft"]
    )
    page = await sdk.me.mail_folders.by_mail_folder_id("drafts").messages.get(
        request_configuration=RequestConfiguration(query_parameters=query)
    )
    messages = page.value if page and page.value else []
    return next((m.id for m in messages if m.is_draft is True), None)


async def test_require_draft_accepts_a_real_draft(real_graph_client):
    sdk = real_graph_client.sdk_client
    draft_id = await _newest_draft_id(sdk)
    if draft_id is None:
        pytest.skip("No draft among the 25 newest items in Drafts, so there is none to check")

    await require_draft(sdk, draft_id, "live-check")  # must not raise


async def test_require_draft_refuses_a_received_message(real_graph_client):
    sdk = real_graph_client.sdk_client
    message_id = await _newest_message_id(sdk, "inbox")
    if message_id is None:
        pytest.skip("The inbox is empty, so there is no received message to check")

    with pytest.raises(ValueError, match="not a draft"):
        await require_draft(sdk, message_id, "live-check")


async def _events_with_and_without_attendees(sdk) -> tuple[str | None, str | None]:
    from msgraph.generated.users.item.events.events_request_builder import (
        EventsRequestBuilder,
    )

    query = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
        top=50, select=["id", "attendees"]
    )
    page = await sdk.me.events.get(
        request_configuration=RequestConfiguration(query_parameters=query)
    )
    with_guests = without_guests = None
    for event in page.value if page and page.value else []:
        if event.attendees and with_guests is None:
            with_guests = event.id
        if not event.attendees and without_guests is None:
            without_guests = event.id
    return with_guests, without_guests


async def test_the_calendar_gate_sees_attendees_on_a_plain_get(real_graph_client):
    """The gate reads the event the way `update_event` does: a GET with no `$select`."""
    sdk = real_graph_client.sdk_client
    with_guests, without_guests = await _events_with_and_without_attendees(sdk)
    if with_guests is None:
        pytest.skip(
            "None of the 50 most recent events has attendees, so the gate's "
            "'meeting with guests' reading is unverified by this run"
        )

    event = await sdk.me.events.by_event_id(with_guests).get()
    assert _has_attendees(event)

    if without_guests is not None:
        solo = await sdk.me.events.by_event_id(without_guests).get()
        assert not _has_attendees(solo)
