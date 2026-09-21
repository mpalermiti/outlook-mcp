"""Calendar write tools: create_event, update_event, delete_event, rsvp."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_CALENDAR_WRITE, check_permission
from outlook_mcp.tools._recurrence import build_event_recurrence, serialize_recurrence
from outlook_mcp.validation import (
    validate_datetime,
    validate_email,
    validate_event_timezone,
    validate_graph_id,
)


def _is_series_master(event: Any) -> bool:
    """Whether Graph considers this event the master of a recurring series."""
    event_type = getattr(event, "type", None)
    if event_type is None:
        return False
    return (getattr(event_type, "value", None) or str(event_type)) == "seriesMaster"


def _resend_recurrence(event: Any, *, anchor: str | None) -> Any:
    """The event's own recurrence, rebuilt as a payload safe to send back.

    Not ``event.recurrence`` handed straight over: that model came off the wire
    with its unset fields explicitly ``None``, and the request adapter
    serializes through the backing store, which writes those nulls out — onto
    the *parent*, under their Python names. Round-tripping through
    ``serialize_recurrence`` (which omits what is unset) and the strict builder
    produces a fresh model carrying only the fields Graph actually sent.

    Two fields are dropped rather than echoed:

    ``range.startDate``, so the builder re-derives it from ``anchor`` — a patch
    that moves the series to a different day has to move the range with it, and
    Graph refuses a range beginning on a different day than the first
    occurrence.

    ``range.recurrenceTimeZone``, because it is the *old* zone. Sending it back
    beside a new ``start.timeZone`` is a self-contradictory patch, and Graph
    answers the contradiction with the same opaque
    ``400 ErrorPropertyValidationFailure`` this function exists to avoid.
    Omitted, Graph derives the range's zone from the event's own — verified
    live 2026-09-21, a series created with no ``recurrenceTimeZone`` at all
    came back carrying the event's.
    """
    payload = serialize_recurrence(getattr(event, "recurrence", None))
    if not payload:
        return None

    start = anchor or getattr(getattr(event, "start", None), "date_time", None)
    if not start:
        raise ValueError(
            "Could not read the event's current start to re-anchor its recurrence "
            "for a timezone change; pass `start` and `end` alongside `timezone`."
        )

    stale = {"startDate", "recurrenceTimeZone"}
    payload["range"] = {k: v for k, v in (payload.get("range") or {}).items() if k not in stale}
    return build_event_recurrence(payload, start=start)


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
    recurrence: dict | str | None = None,
    timezone: str | None = None,
    *,
    config: Config,
) -> dict:
    """Create a calendar event.

    Validates inputs, builds a Graph Event object, and posts via
    graph_client.me.events.post().

    ``recurrence`` accepts a Graph recurrence object, a JSON string of one, or
    a shorthand ("daily", "weekdays", "weekly", "monthly", "yearly"). Setting
    it makes the created event a series master rather than a single occurrence.

    ``timezone`` is the IANA zone the event is *anchored* in, defaulting to
    ``config.timezone``. It is the zone that governs the series, not a
    conversion applied to ``start``: an offset or ``Z`` in the datetime still
    pins the instant (Graph honours it), while the zone decides what the second
    occurrence of a weekly series does when daylight saving ends. Anchored in
    ``UTC`` — which is what this tool sent unconditionally before 1.23.0 — a
    09:00 series silently becomes 08:00 the week the clocks go back.

    A zone-less ``start``/``end`` ("2026-10-28T09:00:00") is read by Graph as
    wall-clock time *in that zone*, which is what ``config.timezone`` has
    always meant for input interpretation everywhere else in this server.

    The datetime string is passed through as the caller wrote it rather than
    normalized to UTC first. That is load-bearing twice over: normalizing a
    naive value needs a zone to resolve it against, and resolving it against
    the host's clock is the bug 1.15.0 removed; and Graph requires an all-day
    event to start on a midnight boundary, which a UTC conversion of local
    midnight is not.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")

    # Validate datetime inputs
    validate_datetime(start)
    validate_datetime(end)
    zone = validate_event_timezone(timezone or config.timezone)

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
    event.start.time_zone = zone

    event.end = DateTimeTimeZone()
    event.end.date_time = end
    event.end.time_zone = zone

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

    if recurrence:
        event.recurrence = build_event_recurrence(recurrence, start=start)

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
    recurrence: dict | str | None = None,
    remove_recurrence: bool = False,
    attendees: list[str] | None = None,
    is_all_day: bool | None = None,
    timezone: str | None = None,
    *,
    config: Config,
) -> dict:
    """Update an existing calendar event.

    Only patches changed fields. ``None`` means "leave alone" for every
    argument, so ``False`` and ``[]`` are real instructions, not absences.

    ``attendees`` **replaces the whole collection** — Graph has no add-one
    operation, so pass the full intended list, and expect Outlook to email
    invitations to everyone on it and cancellations to anyone dropped. ``[]``
    removes them all. This is the one argument here with an outward-facing
    side effect.

    ``is_all_day`` requires ``start`` and ``end`` in the *same* call — Graph
    returns ``ErrorInvalidRequest: Missing parameters: Event.Start`` for a
    lone isAllDay patch — and both must fall on midnight boundaries. We reject
    the incomplete call locally rather than pass through that error.

    ``timezone`` re-anchors the event, and requires ``start`` and ``end`` in
    the same call for the same reason: Graph rejects a ``start`` patch that
    carries no ``timeZone`` at all (``TimeZoneNotSupportedException`` on the
    empty string), so the zone is never an independent edit. Omitted, the
    event keeps the zone it is already stored in — patching a colleague's
    09:00 New York meeting must not quietly move it to this server's zone,
    which is why this defaults differently from ``create_event``. Reading the
    stored zone costs one GET, issued only when ``start``/``end`` is in play.

    Changing the zone of a *series master* needs the recurrence re-sent in the
    same patch; without it Graph answers ``400 ErrorPropertyValidationFailure``
    with nothing naming the zone. When the caller has not supplied a new
    recurrence, the event's existing one is read and sent back unchanged.

    There is deliberately no ``is_online`` here. Graph accepts ``isOnlineMeeting``
    on a personal (consumer) account and silently ignores it — a created or
    patched event comes back ``isOnlineMeeting: False, onlineMeetingProvider:
    'unknown'``. Teams meetings need a work/school account, which this server
    does not target. Adding the parameter would only ship a convincing no-op.

    Setting ``recurrence`` turns a single event into a series, or replaces the
    pattern of an existing one. It takes the same shapes as ``create_event``.
    Graph anchors the range on the series master's start, so when ``start``
    isn't part of the same patch the event's current start is read first.
    Passing ``None`` leaves any existing recurrence untouched — this is a
    partial patch.

    ``remove_recurrence=True`` turns a series master back into a single event,
    keeping the first occurrence's time. It is a separate flag rather than a
    sentinel value on ``recurrence`` because ``None`` there already means
    "leave alone", and the two are mutually exclusive.

    It has to go through ``additional_data``: Graph clears a series with an
    explicit ``"recurrence": null``, and ``event.recurrence = None`` does not
    produce one. The adapter serializes through the backing store, which writes
    a top-level null *beside* the object body rather than inside it, and then
    refuses the mixed document with ``ValueError("Invalid Json output")`` — so
    the plain assignment is not a silent no-op, it is an unsendable request.
    ``test_setting_event_recurrence_none_is_unsendable`` pins that; if kiota
    ever starts emitting a usable top-level null, it fails and this can be
    simplified.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_update_event")
    event_id = validate_graph_id(event_id)

    if timezone is not None and (start is None or end is None):
        raise ValueError(
            "timezone requires start and end in the same call. It is the zone those "
            "datetimes are read in and the series is anchored to, not a conversion "
            "applied to the times already stored — Graph rejects a start patch that "
            "carries no timeZone, so the two travel together. To re-anchor an event "
            "without moving it, pass its current start and end with the new timezone."
        )

    validated_attendees = None
    if attendees is not None:
        validated_attendees = [validate_email(e) for e in attendees]

    # The event as Graph currently holds it, fetched at most once and only when
    # something actually needs it: the zone a start/end patch must carry, the
    # anchor date for a recurrence sent without a start, or the existing
    # recurrence a zone change has to bring along.
    _current: Any = None

    async def current_event() -> Any:
        nonlocal _current
        if _current is None:
            _current = await graph_client.me.events.by_event_id(event_id).get()
        return _current

    from msgraph.generated.models.attendee import Attendee
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.event import Event
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.location import Location

    event = Event()

    if subject is not None:
        event.subject = subject

    zone: str | None = None
    if start is not None or end is not None:
        if timezone is not None:
            zone = validate_event_timezone(timezone)
        else:
            # Echoed back verbatim, not validated: Graph stores Windows zone
            # names ("Pacific Standard Time") as readily as IANA ones, and a
            # name it gave us is by definition one it accepts.
            stored = getattr(getattr(await current_event(), "start", None), "time_zone", None)
            zone = stored or validate_event_timezone(config.timezone)

    if start is not None:
        validate_datetime(start)
        event.start = DateTimeTimeZone()
        event.start.date_time = start
        event.start.time_zone = zone

    if end is not None:
        validate_datetime(end)
        event.end = DateTimeTimeZone()
        event.end.date_time = end
        event.end.time_zone = zone

    if location is not None:
        event.location = Location()
        event.location.display_name = location

    if body is not None:
        event.body = ItemBody()
        event.body.content = body
        event.body.content_type = BodyType.Text

    if validated_attendees is not None:
        event.attendees = []
        for email in validated_attendees:
            att = Attendee()
            att.email_address = EmailAddress()
            att.email_address.address = email
            event.attendees.append(att)

    if is_all_day is not None:
        if start is None or end is None:
            raise ValueError(
                "is_all_day requires start and end in the same call; Graph rejects a lone "
                "isAllDay patch with 'Missing parameters: Event.Start'. Both must be "
                "midnight boundaries, e.g. start=2026-10-22T00:00:00Z, end=2026-10-23T00:00:00Z"
            )
        event.is_all_day = is_all_day

    if remove_recurrence:
        if recurrence is not None:
            raise ValueError(
                "Pass either recurrence or remove_recurrence, not both — they ask for "
                "opposite things"
            )
        # `event.recurrence = None` does not serialize (see the docstring);
        # additional_data survives as the explicit JSON null Graph needs.
        event.additional_data = {**(event.additional_data or {}), "recurrence": None}

    if recurrence is not None:
        anchor = start
        if anchor is None:
            anchor = getattr(getattr(await current_event(), "start", None), "date_time", None)
            if not anchor:
                raise ValueError(
                    "Could not read the event's current start to anchor the recurrence "
                    "range; pass `start` alongside `recurrence`."
                )
        event.recurrence = build_event_recurrence(recurrence, start=anchor)
    elif zone is not None and not remove_recurrence:
        # A start/end patch that changes a series master's zone is refused with
        # `400 ErrorPropertyValidationFailure` unless the recurrence travels
        # with it; the same patch succeeds when it does. Verified live against a
        # consumer mailbox, 2026-09-21 — both directions, and with the zone
        # unchanged as the control.
        existing = await current_event()
        stored_zone = getattr(getattr(existing, "start", None), "time_zone", None)
        if stored_zone and stored_zone != zone and _is_series_master(existing):
            event.recurrence = _resend_recurrence(existing, anchor=start)

    response = await graph_client.me.events.by_event_id(event_id).patch(event)

    return {
        "status": "updated",
        "event_id": response.id,
    }


async def delete_event(
    graph_client: Any,
    event_id: str,
    *,
    config: Config,
) -> dict:
    """Delete a calendar event."""
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_delete_event")
    event_id = validate_graph_id(event_id)

    await graph_client.me.events.by_event_id(event_id).delete()

    return {"status": "deleted", "event_id": event_id}


async def rsvp(
    graph_client: Any,
    event_id: str,
    response: str,
    message: str | None = None,
    *,
    config: Config,
) -> dict:
    """RSVP to a calendar event.

    response must be one of: accept, decline, tentative.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_rsvp")
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
