"""Calendar write tools: create_event, update_event, delete_event, rsvp."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.validation import validate_datetime, validate_email, validate_graph_id


def _check_read_only(read_only: bool, tool_name: str) -> None:
    """Raise ReadOnlyError if server is in read-only mode."""
    if read_only:
        raise ReadOnlyError(tool_name)


async def create_event(
    graph_client: Any,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    is_all_day: bool = False,
    is_online: bool = False,
    recurrence: str | None = None,
    read_only: bool = False,
) -> dict:
    """Create a calendar event.

    Validates inputs, builds a Graph Event object, and posts via
    graph_client.me.events.post().
    """
    _check_read_only(read_only, "outlook_create_event")

    # Validate datetime inputs
    validate_datetime(start)
    validate_datetime(end)

    # Validate attendee emails if provided
    validated_attendees = []
    if attendees:
        validated_attendees = [validate_email(e) for e in attendees]

    from msgraph.generated.models.attendee import Attendee
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.event import Event
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.location import Location

    event = Event()
    event.subject = subject

    event.start = DateTimeTimeZone()
    event.start.date_time = start
    event.start.time_zone = "UTC"

    event.end = DateTimeTimeZone()
    event.end.date_time = end
    event.end.time_zone = "UTC"

    event.is_all_day = is_all_day
    event.is_online_meeting = is_online

    if location:
        event.location = Location()
        event.location.display_name = location

    if body:
        event.body = ItemBody()
        event.body.content = body
        event.body.content_type = BodyType.Text

    if validated_attendees:
        event.attendees = []
        for email in validated_attendees:
            att = Attendee()
            att.email_address = EmailAddress()
            att.email_address.address = email
            event.attendees.append(att)

    response = await graph_client.me.events.post(event)

    return {
        "status": "created",
        "event_id": response.id,
        "subject": response.subject,
    }


async def update_event(
    graph_client: Any,
    event_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    read_only: bool = False,
) -> dict:
    """Update an existing calendar event.

    Only patches changed fields.
    """
    _check_read_only(read_only, "outlook_update_event")
    event_id = validate_graph_id(event_id)

    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.event import Event
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.location import Location

    event = Event()

    if subject is not None:
        event.subject = subject

    if start is not None:
        validate_datetime(start)
        event.start = DateTimeTimeZone()
        event.start.date_time = start
        event.start.time_zone = "UTC"

    if end is not None:
        validate_datetime(end)
        event.end = DateTimeTimeZone()
        event.end.date_time = end
        event.end.time_zone = "UTC"

    if location is not None:
        event.location = Location()
        event.location.display_name = location

    if body is not None:
        event.body = ItemBody()
        event.body.content = body
        event.body.content_type = BodyType.Text

    response = await graph_client.me.events.by_event_id(event_id).patch(event)

    return {
        "status": "updated",
        "event_id": response.id,
    }


async def delete_event(
    graph_client: Any,
    event_id: str,
    read_only: bool = False,
) -> dict:
    """Delete a calendar event."""
    _check_read_only(read_only, "outlook_delete_event")
    event_id = validate_graph_id(event_id)

    await graph_client.me.events.by_event_id(event_id).delete()

    return {"status": "deleted", "event_id": event_id}


async def rsvp(
    graph_client: Any,
    event_id: str,
    response: str,
    message: str | None = None,
    read_only: bool = False,
) -> dict:
    """RSVP to a calendar event.

    response must be one of: accept, decline, tentative.
    """
    _check_read_only(read_only, "outlook_rsvp")
    event_id = validate_graph_id(event_id)

    event_builder = graph_client.me.events.by_event_id(event_id)

    if response == "accept":
        from msgraph.generated.users.item.events.item.accept.accept_post_request_body import (
            AcceptPostRequestBody,
        )

        request_body = AcceptPostRequestBody()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await event_builder.accept.post(request_body)
        return {"status": "accepted", "event_id": event_id}

    elif response == "decline":
        from msgraph.generated.users.item.events.item.decline.decline_post_request_body import (
            DeclinePostRequestBody,
        )

        request_body = DeclinePostRequestBody()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await event_builder.decline.post(request_body)
        return {"status": "declined", "event_id": event_id}

    elif response == "tentative":
        from msgraph.generated.users.item.events.item.tentatively_accept import (  # noqa: E501
            tentatively_accept_post_request_body,
        )

        cls = tentatively_accept_post_request_body.TentativelyAcceptPostRequestBody
        request_body = cls()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await event_builder.tentatively_accept.post(request_body)
        return {"status": "tentativelyAccepted", "event_id": event_id}

    else:
        raise ValueError(f"response must be accept/decline/tentative; got {response}")
