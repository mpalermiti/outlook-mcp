"""Calendar read tools: list_events, get_event."""

from __future__ import annotations

from datetime import datetime, timedelta
from datetime import timezone as dt_timezone
from typing import Any

from outlook_mcp.calendar_resolver import resolve_calendar_id
from outlook_mcp.pagination import (
    apply_pagination,
    build_request_config,
    decode_cursor_payload,
    encode_cursor_payload,
    wrap_nextlink,
)
from outlook_mcp.tools._recurrence import serialize_recurrence
from outlook_mcp.validation import (
    CONFIG_REMEDY_TEXT,
    resolve_timezone,
    sanitize_output,
    validate_datetime,
    validate_event_key,
    validate_graph_id,
)

_UTC_FORMAT = "%Y-%m-%dT%H:%M:%SZ"

# A `$select` and the formatter that reads the response are one contract written
# in two places and two spellings — camelCase on the wire, snake_case on the SDK
# model. Nothing in the language connects them, so they drift silently: Graph
# honours the `$select`, the SDK leaves the unasked-for attribute `None`, and the
# formatter reports it as `""`. That is issue #69, where `type` was read by
# `_format_event_summary` and never selected, so every listed event came back
# `type: ""` from 1.16.0 on.
#
# Spelling each `$select` once, at module scope, is what lets
# `tests/test_select_covers_the_formatter.py` hold the two halves together: both
# pairs have a `_PAIRS` row, checked in both directions. Add a field to the
# formatter and to its `$select`, or to neither.
_SUMMARY_SELECT = (
    "id,subject,start,end,location,isAllDay,organizer,responseStatus,isOnlineMeeting,showAs,type"
)

# An event's attachments as `get_event` lists them: what each one is, never its
# bytes. Without the `$select`, `contentBytes` of every file comes back, base64.
_ATTACHMENT_SELECT = "id,name,contentType,size,isInline"

# Concise mode needs `attendees` and `isOrganizer` to compute its two derived
# fields, and nothing else — keep it tight so a day-at-a-glance scan never pulls
# event bodies.
_CONCISE_SELECT = "id,subject,start,end,location,isAllDay,isOrganizer,isOnlineMeeting,attendees"


# Where an event's `key` lives: a named extended property this server owns.
# Outlook does not display it, and a `$filter` can match on it, which is what
# `find_events_by_key` does. The GUID was generated once for this purpose.
# NEVER CHANGE IT: every event already stamped becomes unfindable by key the
# moment it moves. Written by `calendar_write`, read here.
EVENT_KEY_PROPERTY_ID = "String {51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey"

# Expands that one property and no other. An event can carry extended
# properties other software wrote; without the inner filter they would all come
# back on every read.
_EVENT_KEY_EXPAND = f"singleValueExtendedProperties($filter=id eq '{EVENT_KEY_PROPERTY_ID}')"

# More than this many events sharing one key is not a lookup any more; the
# answer says `has_more` rather than paging through them.
_FIND_LIMIT = 25


def event_key(event: Any) -> str | None:
    """The event's key, or ``None``.

    Compared case-insensitively: Graph does not promise to echo a property id
    in the case it was written in, and it lowercases the hex of the
    `SystemTime 0x3FEF` property `mail_read` reads.

    Called beside the formatters rather than from them: the key comes from an
    ``$expand``, not from the ``$select`` that ``test_select_covers_the_formatter``
    holds each formatter to.
    """
    wanted = EVENT_KEY_PROPERTY_ID.casefold()
    for prop in getattr(event, "single_value_extended_properties", None) or []:
        if (getattr(prop, "id", None) or "").casefold() == wanted:
            return getattr(prop, "value", None)
    return None


def _with_key(summary: dict, event: Any) -> dict:
    """``summary`` plus ``key`` when the event has one.

    Absent rather than ``None`` on a listing: most events carry no key, and a
    null on every row of a page is a cost every listing would pay.
    """
    key = event_key(event)
    if key is not None:
        summary["key"] = key
    return summary


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _compute_calendar_range(
    days: int,
    after: str | None,
    before: str | None,
    timezone: str,
) -> tuple[str, str]:
    """Compute UTC start/end for calendarView.

    Uses explicit after/before if provided, otherwise computes
    relative to "now" in the configured timezone.
    """
    tz = resolve_timezone(timezone, remedy=CONFIG_REMEDY_TEXT)

    def _now_plus(days_ahead: int) -> str:
        # The guard is load-bearing, not a micro-optimisation. PEP 495 makes
        # arithmetic on an aware datetime ignore ``fold`` and return fold=0, so
        # ``+ timedelta(days=0)`` is NOT the identity: during the repeated hour
        # after a DST fall-back it silently picks the first pass, moving the
        # window an hour early. ``start`` must keep the fold ``datetime.now``
        # gave it. (``end`` resets fold too, as it did before this refactor.)
        moment = datetime.now(tz)
        if days_ahead:
            moment += timedelta(days=days_ahead)
        return moment.astimezone(dt_timezone.utc).strftime(_UTC_FORMAT)

    start_utc = validate_datetime(after, timezone) if after else _now_plus(0)
    end_utc = validate_datetime(before, timezone) if before else _now_plus(days)

    return start_utc, end_utc


def _format_event_summary(event: Any) -> dict:
    """Convert Graph SDK event to summary dict."""
    organizer_name = ""
    if event.organizer and event.organizer.email_address:
        organizer_name = event.organizer.email_address.name or ""

    response = ""
    if event.response_status and event.response_status.response:
        response = (
            event.response_status.response.value
            if hasattr(event.response_status.response, "value")
            else str(event.response_status.response)
        )

    start_str = ""
    if event.start:
        start_str = f"{event.start.date_time} ({event.start.time_zone})"

    end_str = ""
    if event.end:
        end_str = f"{event.end.date_time} ({event.end.time_zone})"

    # `type` tells a recurring event from a one-off without fetching each one.
    # Which values can actually appear depends on the caller: this formatter
    # feeds both `list_events`, which reads `calendarView` and therefore sees
    # only `occurrence` / `exception` / `singleInstance`, and
    # `_format_event_detail`, which reads one event by id and can also see
    # `seriesMaster`. Documenting the master as a listing value would send an
    # agent filtering for recurring meetings to a value the listing never
    # returns — #69's silent wrong answer, one value over.
    event_type = ""
    if event.type is not None:
        event_type = event.type.value if hasattr(event.type, "value") else str(event.type)

    # Graph's `showAs` — what Outlook labels "Show as". Every event carries one
    # (Graph defaults it to `busy`), so an empty string here means the field was
    # not fetched rather than that the event has no status.
    show_as = ""
    if event.show_as is not None:
        show_as = event.show_as.value if hasattr(event.show_as, "value") else str(event.show_as)

    return {
        "id": event.id,
        "subject": sanitize_output(event.subject or "(no subject)"),
        "start": start_str,
        "end": end_str,
        "location": sanitize_output(
            event.location.display_name if event.location and event.location.display_name else ""
        ),
        "is_all_day": bool(event.is_all_day),
        "organizer": sanitize_output(organizer_name),
        "response_status": response,
        "is_online": bool(event.is_online_meeting),
        "type": event_type,
        "show_as": show_as,
    }


def _format_event_detail(event: Any) -> dict:
    """Convert Graph SDK event to full detail dict.

    ``start``/``end`` are UTC because that is what Graph returns without a
    ``Prefer: outlook.timezone`` header, and what every other datetime in this
    server's responses is. ``original_start_time_zone`` is therefore the only
    way a caller can see the zone the event is actually anchored in — which is
    not cosmetic: a weekly series anchored in UTC shifts an hour in local terms
    the moment daylight saving ends, and one anchored in a named zone does not.
    """
    summary = _format_event_summary(event)

    organizer = {}
    if event.organizer and event.organizer.email_address:
        organizer = {
            "name": sanitize_output(event.organizer.email_address.name or ""),
            "email": event.organizer.email_address.address or "",
        }

    attendees = []
    for att in event.attendees or []:
        entry = {}
        if att.email_address:
            entry["name"] = sanitize_output(att.email_address.name or "")
            entry["email"] = att.email_address.address or ""
        if att.status and att.status.response:
            entry["response"] = (
                att.status.response.value
                if hasattr(att.status.response, "value")
                else str(att.status.response)
            )
        attendees.append(entry)

    body = ""
    if event.body and event.body.content:
        body = sanitize_output(event.body.content, multiline=True)

    online_meeting_url = None
    if event.online_meeting and event.online_meeting.join_url:
        online_meeting_url = event.online_meeting.join_url

    return {
        **summary,
        "organizer": organizer,
        "body": body,
        "attendees": attendees,
        "online_meeting_url": online_meeting_url,
        "original_start_time_zone": event.original_start_time_zone,
        "original_end_time_zone": event.original_end_time_zone,
        "recurrence": serialize_recurrence(event.recurrence),
        "categories": list(event.categories or []),
    }


def _format_event_concise(event: Any) -> dict:
    """Return a token-efficient event dict for concise-mode listings.

    Keeps: id, subject, start, end, location, is_all_day, is_organizer,
    is_online_meeting, attendees_count. Drops, relative to the normal listing:
    organizer (object), response_status, `type` and `show_as` — concise mode is
    for day-at-a-glance scans where the series/one-off distinction isn't worth
    the tokens; use the normal listing when it is. (Body, the full attendees
    list and categories are detail-only and on no listing shape.)

    `show_as` is dropped for the same reason and on the same terms as
    `response_status`: both are status fields this shape has always traded
    away, and adding one to the highest-volume listing is a token cost every
    caller pays on every scan. The normal listing carries it.
    """
    start_str = ""
    if event.start:
        start_str = f"{event.start.date_time} ({event.start.time_zone})"

    end_str = ""
    if event.end:
        end_str = f"{event.end.date_time} ({event.end.time_zone})"

    return {
        "id": event.id,
        "subject": sanitize_output(event.subject or "(no subject)"),
        "start": start_str,
        "end": end_str,
        "location": sanitize_output(
            event.location.display_name if event.location and event.location.display_name else ""
        ),
        "is_all_day": bool(event.is_all_day),
        "is_organizer": bool(getattr(event, "is_organizer", False)),
        "is_online_meeting": bool(event.is_online_meeting),
        "attendees_count": len(event.attendees or []),
    }


async def list_events(
    graph_client: Any,
    days: int = 7,
    after: str | None = None,
    before: str | None = None,
    count: int = 50,
    timezone: str = "UTC",
    cursor: str | None = None,
    concise: bool = False,
    calendar: str | None = None,
) -> dict:
    """List calendar events in a date range (expands recurring instances).

    The calendarView endpoint requires startDateTime and endDateTime.
    If after/before are not provided, they are computed from `days`
    relative to "now" in the configured timezone.

    concise: when True, return a compact event shape — drops ``organizer``,
    ``response_status``, ``type``, ``show_as``; adds ``is_organizer`` and
    ``attendees_count``. Default False preserves the existing shape.

    calendar: which calendar to read. None, blank or "primary" keeps the
    default calendar's ``/me/calendarView``; otherwise a display name or ID
    resolved via ``calendar_resolver.resolve_calendar_id`` — the only way to
    reach secondary calendars. A cursor continues the listing it came from:
    the calendar is stored in it, so a later page neither re-resolves the
    name nor drifts to a different calendar when the argument is omitted.
    """
    count = _clamp(count, 1, 100)
    start_utc, end_utc = _compute_calendar_range(days, after, before, timezone)
    # Decode before any network call, so a bad cursor is refused for free.
    payload = decode_cursor_payload(cursor) if cursor else {}

    if "calendar" in payload:
        # Caller-held state that goes into a URL path: validate, never trust.
        calendar_id = payload["calendar"]
        if calendar_id is not None:
            if not isinstance(calendar_id, str):
                raise ValueError("Invalid pagination cursor")
            calendar_id = validate_graph_id(calendar_id)
    else:
        calendar_id = await resolve_calendar_id(graph_client, calendar)

    query_params = apply_pagination({}, count, cursor)
    query_params["start_date_time"] = start_utc
    query_params["end_date_time"] = end_utc
    query_params["$orderby"] = "start/dateTime"
    query_params["$select"] = _CONCISE_SELECT if concise else _SUMMARY_SELECT
    if not concise:
        # Concise scans do not carry the key, so they do not pay to expand it.
        query_params["$expand"] = _EVENT_KEY_EXPAND

    from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
        CalendarViewRequestBuilder,
    )

    # The per-calendar view has its own request-builder class in the SDK, but
    # its query-parameters dataclass is field-for-field the same and kiota
    # reads the parameters off the object, not the class — one import serves
    # both paths.
    if calendar_id is None:
        view = graph_client.me.calendar_view
    else:
        view = graph_client.me.calendars.by_calendar_id(calendar_id).calendar_view

    req_config = build_request_config(
        CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters, query_params
    )
    response = await view.get(request_configuration=req_config)

    if concise:
        events = [_format_event_concise(e) for e in (response.value or [])]
    else:
        events = [_with_key(_format_event_summary(e), e) for e in (response.value or [])]

    next_cursor = wrap_nextlink(response.odata_next_link)
    if next_cursor is not None:
        # Pin the calendar into the cursor (None = the default calendar), so
        # the next page is unambiguously a continuation of this listing.
        next_cursor = encode_cursor_payload(
            {**decode_cursor_payload(next_cursor), "calendar": calendar_id}
        )

    return {
        "events": events,
        "count": len(events),
        "has_more": response.odata_next_link is not None,
        "cursor": next_cursor,
    }


def _format_event_attachment(attachment: Any) -> dict:
    """One attachment as ``get_event`` lists it; the name is the sender's text."""
    content_type = attachment.content_type
    return {
        "id": attachment.id,
        "name": sanitize_output(attachment.name or ""),
        "size": attachment.size,
        "content_type": sanitize_output(content_type) if content_type else None,
        "is_inline": bool(attachment.is_inline),
    }


async def _event_attachments(builder: Any) -> list[dict]:
    """Every attachment on the event ``builder`` addresses, following ``@odata.nextLink``."""
    from msgraph.generated.users.item.events.item.attachments.attachments_request_builder import (  # noqa: E501
        AttachmentsRequestBuilder,
    )

    config = build_request_config(
        AttachmentsRequestBuilder.AttachmentsRequestBuilderGetQueryParameters,
        {"$select": _ATTACHMENT_SELECT},
    )
    response = await builder.attachments.get(request_configuration=config)
    found = list(response.value or []) if response else []
    while isinstance(getattr(response, "odata_next_link", None), str) and response.odata_next_link:
        response = await builder.attachments.with_url(response.odata_next_link).get()
        found.extend((response.value or []) if response else [])
    return [_format_event_attachment(a) for a in found]


async def get_event(
    graph_client: Any,
    event_id: str,
) -> dict:
    """Get full details for a single event, with its ``key`` (``None`` when unset)
    and its attachments listed.

    The attachments are a second request, made only when Graph reports
    ``hasAttachments``, so an event without files costs what it always did.
    """
    event_id = validate_graph_id(event_id)

    from msgraph.generated.users.item.events.item.event_item_request_builder import (
        EventItemRequestBuilder,
    )

    config = build_request_config(
        EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters,
        {"$expand": _EVENT_KEY_EXPAND},
    )
    builder = graph_client.me.events.by_event_id(event_id)
    event = await builder.get(request_configuration=config)
    attachments = await _event_attachments(builder) if event.has_attachments is True else []

    return {**_format_event_detail(event), "key": event_key(event), "attachments": attachments}


async def find_events_by_key(
    graph_client: Any,
    key: str,
    calendar: str | None = None,
) -> dict:
    """The events in one calendar whose key is ``key``.

    ``GET /me/events`` (or ``/me/calendars/{id}/events``) filtered on the key
    property, so the search has no date window: a booking months out is found
    as readily as tomorrow's. It returns single events and series masters,
    never occurrences — the master is the one to update, and patching an
    occurrence would make it an exception instead.

    Without ``calendar`` the search covers the default calendar, which is what
    ``/me/events`` lists (Graph's user resource: "Default is to show Events
    under the Default Calendar"); ``calendar`` is resolved as for
    ``list_events``. Nothing matching is an empty answer, not an error.
    """
    key = validate_event_key(key)
    calendar_id = await resolve_calendar_id(graph_client, calendar)

    from msgraph.generated.users.item.events.events_request_builder import EventsRequestBuilder

    # The key is validated to characters that need no escaping. Its quotes are
    # doubled all the same, as mail_read does for its literals, so the literal
    # stays closed if that alphabet ever grows.
    literal = key.replace("'", "''")
    query = {
        "$filter": (
            f"singleValueExtendedProperties/Any(ep: ep/id eq '{EVENT_KEY_PROPERTY_ID}' "
            f"and ep/value eq '{literal}')"
        ),
        "$expand": _EVENT_KEY_EXPAND,
        "$select": _SUMMARY_SELECT,
        "$top": _FIND_LIMIT,
    }
    # As in list_events: the per-calendar builder's query-parameters class is
    # field-for-field the same, and kiota reads them off the object.
    events = (
        graph_client.me.events
        if calendar_id is None
        else graph_client.me.calendars.by_calendar_id(calendar_id).events
    )
    response = await events.get(
        request_configuration=build_request_config(
            EventsRequestBuilder.EventsRequestBuilderGetQueryParameters, query
        )
    )
    found = [_with_key(_format_event_summary(e), e) for e in (response.value or [])]
    return {
        "key": key,
        "events": found,
        "count": len(found),
        "has_more": response.odata_next_link is not None,
    }
