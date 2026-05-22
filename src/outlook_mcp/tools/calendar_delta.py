"""Calendar delta tool: ``list_events_delta``.

Wraps ``GET /me/calendarView/delta``.

Graph constraints worth knowing before reading the code:

- ``startDateTime`` + ``endDateTime`` are *required* on the first call —
  unlike ``/me/messages/delta`` there is no whole-mailbox sync, only
  windowed deltas. We validate both at the boundary so a missing arg
  fails fast as ``ValueError`` (passed through by the tool wrapper).
- ``$top`` is rejected with a 400. Page sizing is via the
  ``Prefer: odata.maxpagesize=N`` header instead.
- Subsequent calls use the returned deltaLink / nextLink URL verbatim —
  Graph encodes the window in the cursor.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import quote, urljoin

from outlook_mcp.tools._delta import GRAPH_BASE, fetch_delta_pages, format_delta_item
from outlook_mcp.validation import sanitize_output, validate_datetime


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _format_event_delta(raw: dict) -> dict:
    """Build the wire-shape event summary from a raw Graph JSON event.

    Mirrors ``calendar_read._format_event_summary`` field-for-field so
    callers don't have to branch on whether a given dict came from a
    delta call or a regular list call.
    """
    start_node = raw.get("start") or {}
    end_node = raw.get("end") or {}
    organizer_node = (raw.get("organizer") or {}).get("emailAddress") or {}
    response_node = raw.get("responseStatus") or {}
    location_node = raw.get("location") or {}

    start_str = ""
    if start_node:
        start_str = f"{start_node.get('dateTime', '')} ({start_node.get('timeZone', '')})"
    end_str = ""
    if end_node:
        end_str = f"{end_node.get('dateTime', '')} ({end_node.get('timeZone', '')})"

    return {
        "id": raw.get("id"),
        "subject": sanitize_output(raw.get("subject") or "(no subject)"),
        "start": start_str,
        "end": end_str,
        "location": sanitize_output(location_node.get("displayName") or ""),
        "is_all_day": bool(raw.get("isAllDay")),
        "organizer": sanitize_output(organizer_node.get("name") or ""),
        "response_status": response_node.get("response") or "",
        "is_online": bool(raw.get("isOnlineMeeting")),
    }


async def list_events_delta(
    graph_client: Any,
    start: str | None,
    end: str | None,
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List calendar-view changes within ``[start, end]`` since the last token.

    Args:
        graph_client: A ``GraphClient`` (delta tools need the
            ``credential`` attribute to mint raw bearer tokens).
        start: ISO 8601 window start. **Required** on the first call;
            ignored when ``delta_token`` is set (the cursor already
            encodes the window).
        end: ISO 8601 window end. Same as ``start``.
        page_size: Items per Graph page (1-100). Mapped to
            ``Prefer: odata.maxpagesize=N`` — Graph rejects ``$top`` on
            this endpoint.
        delta_token: Opaque cursor from a previous call.

    Returns:
        ``{events, delta_token, has_more}`` — see ``mail_delta`` for the
        full semantics of these fields. Tombstones come back as
        ``{id, is_deleted: True}``.
    """
    page_size = _clamp(page_size, 1, 100)
    credential = graph_client.credential

    initial_url = ""
    if not delta_token:
        if not start or not end:
            raise ValueError(
                "outlook_list_events_delta requires both 'start' and 'end' "
                "(ISO 8601 datetimes) on the first call — Graph's "
                "/me/calendarView/delta endpoint has no whole-calendar sync."
            )
        safe_start = validate_datetime(start)
        safe_end = validate_datetime(end)
        initial_url = urljoin(GRAPH_BASE, "me/calendarView/delta")
        initial_url = (
            f"{initial_url}?startDateTime={quote(safe_start, safe='')}"
            f"&endDateTime={quote(safe_end, safe='')}"
        )

    raw_items, next_token, has_more = await fetch_delta_pages(
        credential,
        initial_url=initial_url,
        delta_token=delta_token,
        page_size=page_size,
        headers={"Prefer": f"odata.maxpagesize={page_size}"},
    )

    events = [format_delta_item(item, _format_event_delta) for item in raw_items]

    return {
        "events": events,
        "delta_token": next_token,
        "has_more": has_more,
    }
