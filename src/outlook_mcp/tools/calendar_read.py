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

# Concise mode needs `attendees` and `isOrganizer` to compute its two derived
# fields, and nothing else — keep it tight so a day-at-a-glance scan never pulls
# event bodies.
_CONCISE_SELECT = "id,subject,start,end,location,isAllDay,isOrganizer,isOnlineMeeting,attendees"


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

    # `type` distinguishes a seriesMaster from a singleInstance / occurrence /
    # exception. Without it a listing cannot tell a recurring event from a
    # one-off without fetching each one.
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
        events = [_format_event_summary(e) for e in (response.value or [])]

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


async def get_event(
    graph_client: Any,
    event_id: str,
) -> dict:
    """Get full details for a single event."""
    event_id = validate_graph_id(event_id)

    event = await graph_client.me.events.by_event_id(event_id).get()

    return _format_event_detail(event)
