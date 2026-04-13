"""Calendar read tools: list_events, get_event."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from outlook_mcp.pagination import apply_pagination, wrap_nextlink
from outlook_mcp.validation import sanitize_output, validate_datetime, validate_graph_id


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
    tz = ZoneInfo(timezone)

    if after:
        start_utc = validate_datetime(after)
    else:
        now_local = datetime.now(tz)
        start_utc = now_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

    if before:
        end_utc = validate_datetime(before)
    else:
        now_local = datetime.now(tz)
        end_local = now_local + timedelta(days=days)
        end_utc = end_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

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

    recurrence = None
    if event.recurrence:
        recurrence = str(event.recurrence)

    return {
        **summary,
        "organizer": organizer,
        "body": body,
        "attendees": attendees,
        "online_meeting_url": online_meeting_url,
        "recurrence": recurrence,
        "categories": list(event.categories or []),
    }


async def list_events(
    graph_client: Any,
    days: int = 7,
    after: str | None = None,
    before: str | None = None,
    count: int = 50,
    timezone: str = "UTC",
    cursor: str | None = None,
) -> dict:
    """List calendar events using calendarView.

    The calendarView endpoint requires startDateTime and endDateTime.
    If after/before are not provided, they are computed from `days`
    relative to "now" in the configured timezone.
    """
    count = _clamp(count, 1, 100)
    start_utc, end_utc = _compute_calendar_range(days, after, before, timezone)

    query_params = apply_pagination({}, count, cursor)
    query_params["start_date_time"] = start_utc
    query_params["end_date_time"] = end_utc
    query_params["$orderby"] = "start/dateTime"
    query_params["$select"] = (
        "id,subject,start,end,location,isAllDay,"
        "organizer,responseStatus,isOnlineMeeting,categories"
    )

    response = await graph_client.me.calendar_view.get(query_params=query_params)

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
