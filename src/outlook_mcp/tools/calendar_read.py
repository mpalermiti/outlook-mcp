"""Calendar read tools: list_events, get_event."""

from __future__ import annotations

from datetime import datetime, timedelta
from datetime import timezone as dt_timezone
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from outlook_mcp.folder_resolver import _looks_like_graph_id
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.tools._recurrence import serialize_recurrence
from outlook_mcp.validation import sanitize_output, validate_datetime, validate_graph_id

_UTC_FORMAT = "%Y-%m-%dT%H:%M:%SZ"


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _has_time_zone_database() -> bool:
    """Whether *any* IANA database is reachable, asked by resolving a key that
    every database has.

    Importability of ``tzdata`` is a proxy for this, not the thing itself: a
    POSIX host with ``/usr/share/zoneinfo`` and no ``tzdata`` installed has a
    database, and would have been told its install was broken.
    """
    try:
        ZoneInfo("UTC")
    except (ZoneInfoNotFoundError, ValueError):
        return False
    return True


def _resolve_timezone(name: str) -> ZoneInfo:
    """Return the ``ZoneInfo`` for ``name``, or say why it would not load.

    ``name`` is server configuration rather than a tool argument, but an
    unhandled ``ZoneInfoNotFoundError`` reaches the model as a message-free
    ``Error executing tool outlook_list_events`` — ``_wrap_tool_errors`` keeps
    the text of an *unexpected* exception on the server, which is right for a
    crash and wrong for a misconfiguration. Raising ``ValueError`` routes it
    through ``ToolInputError`` instead, so the message survives the trip. This
    is what ``validate_datetime`` already does with the same config value.

    The two ways the lookup fails need different fixes, so they get different
    messages: a name no database contains is a typo in ``config.json``, while
    no database at all is a broken install — ``tzdata`` is a dependency exactly
    so that Windows and slim Linux images have one.

    ``ValueError`` is caught alongside ``ZoneInfoNotFoundError`` because zoneinfo
    raises it, not the subclass, for a path-shaped key: ``/etc/localtime`` is a
    plausible thing to put in a config file and would otherwise escape both the
    truncation and the hint. ``validate_datetime`` already catches both.
    """
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        if not _has_time_zone_database():
            raise ValueError(
                f"Invalid timezone: {name[:50]} — this host has no IANA time zone "
                "database, so no zone name would resolve. Reinstall "
                "outlook-graph-mcp to pick up its `tzdata` dependency."
            ) from exc
        raise ValueError(
            f"Invalid timezone: {name[:50]} — not a zone name the IANA database "
            "contains. Set `timezone` in ~/.outlook-mcp/config.json to a name "
            "like America/Los_Angeles or UTC."
        ) from exc


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
    tz = _resolve_timezone(timezone)

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


async def _resolve_calendar_id(graph_client: Any, calendar_ref: str) -> str | None:
    """Resolve a user-supplied calendar reference to a Graph calendar ID.

    Accepts (in priority order):
    - "primary" (case-insensitive): the default calendar, returned as ``None``
      so the caller keeps the ``/me/calendarView`` path.
    - Graph calendar IDs (passed through after syntactic validation).
    - Calendar display names (case-insensitive, matched against ``/me/calendars``).

    Raises ValueError when the name is not found or is ambiguous. Both messages
    name the calendars that do exist, because that list is one
    ``outlook_list_calendars`` call away from fixing the reference — the same
    stance `folder_resolver` takes for mail folders.
    """
    trimmed = calendar_ref.strip()
    if not trimmed:
        raise ValueError("Calendar reference must not be empty")
    if trimmed.lower() == "primary":
        return None
    if _looks_like_graph_id(trimmed):
        return validate_graph_id(trimmed)

    response = await graph_client.me.calendars.get()
    calendars = list(response.value or [])
    matches = [cal for cal in calendars if cal.name and cal.name.lower() == trimmed.lower()]
    available = ", ".join(cal.name or "(unnamed)" for cal in calendars)

    if not matches:
        raise ValueError(
            f"Calendar '{trimmed[:50]}' not found. Available calendars: {available}. "
            "Pass a display name or a Graph calendar ID from outlook_list_calendars."
        )
    if len(matches) > 1:
        raise ValueError(
            f"Calendar name '{trimmed[:50]}' is ambiguous "
            f"({len(matches)} matches). Pass a Graph calendar ID instead."
        )
    return matches[0].id


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
    }


def _format_event_detail(event: Any) -> dict:
    """Convert Graph SDK event to full detail dict."""
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
        "recurrence": serialize_recurrence(event.recurrence),
        "categories": list(event.categories or []),
    }


def _format_event_concise(event: Any) -> dict:
    """Return a token-efficient event dict for concise-mode listings.

    Keeps: id, subject, start, end, location, is_all_day, is_organizer,
    is_online_meeting, attendees_count. Drops: body, organizer (object),
    response_status, categories, full attendees list, and `type` — concise
    mode is for day-at-a-glance scans where the series/one-off distinction
    isn't worth the tokens; use the normal listing when it is.
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
    ``response_status``, ``categories``; adds ``is_organizer`` and
    ``attendees_count``. Default False preserves the existing shape.

    calendar: which calendar to read. None (or "primary") keeps the default
    calendar's ``/me/calendarView``; otherwise a display name or Graph calendar
    ID resolved via ``_resolve_calendar_id`` — the only way to reach secondary
    calendars.
    """
    count = _clamp(count, 1, 100)
    start_utc, end_utc = _compute_calendar_range(days, after, before, timezone)
    calendar_id = await _resolve_calendar_id(graph_client, calendar) if calendar else None

    query_params = apply_pagination({}, count, cursor)
    query_params["start_date_time"] = start_utc
    query_params["end_date_time"] = end_utc
    query_params["$orderby"] = "start/dateTime"
    if concise:
        # We need attendees + isOrganizer to compute the concise fields.
        # Keep the select tight to avoid pulling full event bodies.
        query_params["$select"] = (
            "id,subject,start,end,location,isAllDay,isOrganizer,isOnlineMeeting,attendees"
        )
    else:
        query_params["$select"] = (
            "id,subject,start,end,location,isAllDay,"
            "organizer,responseStatus,isOnlineMeeting,categories"
        )

    if calendar_id is None:
        from msgraph.generated.users.item.calendar_view import (
            calendar_view_request_builder as cv,
        )

        view = graph_client.me.calendar_view
    else:
        from msgraph.generated.users.item.calendars.item.calendar_view import (
            calendar_view_request_builder as cv,
        )

        view = graph_client.me.calendars.by_calendar_id(calendar_id).calendar_view

    req_config = build_request_config(
        cv.CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters,
        query_params,
    )
    response = await view.get(request_configuration=req_config)

    if concise:
        events = [_format_event_concise(e) for e in (response.value or [])]
    else:
        events = [_format_event_summary(e) for e in (response.value or [])]

    return {
        "events": events,
        "count": len(events),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }


async def get_event(
    graph_client: Any,
    event_id: str,
) -> dict:
    """Get full details for a single event."""
    event_id = validate_graph_id(event_id)

    event = await graph_client.me.events.by_event_id(event_id).get()

    return _format_event_detail(event)
