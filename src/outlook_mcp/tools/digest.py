"""Composed digest tool: ``changes_since``.

Wraps the three v1.9.0 delta tools (mail / events / contacts) into a
single structured "since last call" payload. Designed for recurring
agent loops — a morning brief or hourly inbox sweep that wants ONE
structured digest call instead of orchestrating three delta calls and
reasoning over the raw item shapes itself.

Tokens are caller-managed (matches the v1.9.0 delta tools). The caller
stores the returned ``delta_tokens`` dict and passes it back on the
next call. Each resource is independent — a missing or stale token for
one resource doesn't block the other two.

Behavior summary:

- **First call** (no tokens, or a key is missing): call the underlying
  delta tool with no token to bootstrap. Graph's first response is a
  *full snapshot*, not "what changed", so we filter to the last
  ``fallback_window_hours`` to avoid flooding the digest with historical
  items. For calendar, we pass ``[now - window, now + 7d]`` so the
  bootstrap snapshot captures recent + upcoming changes.
- **Subsequent call** (with tokens): use the token verbatim; drain
  pagination up to a safety cap; classify each returned item.
- **`syncStateNotFound` recovery**: if a stored token is too old, Graph
  returns 410. We drop the bad token, re-do that resource as a "first
  call", and surface ``_meta.resync`` in the response so the caller
  knows their old watermark was discarded.
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx

from outlook_mcp.tools.calendar_delta import list_events_delta
from outlook_mcp.tools.contacts_delta import list_contacts_delta
from outlook_mcp.tools.mail_delta import list_inbox_delta

# Safety cap on internal pagination drain. Each underlying delta call is
# already bounded to ``page_size * 4 = ~200`` items; we let the digest
# chain up to 5 of those for a hard ceiling of ~1,000 items per resource
# per call. Keeps the digest bounded on a very large bootstrap.
_MAX_INTERNAL_PAGES = 5

# Default per-page size for underlying delta calls.
_PAGE_SIZE = 50

# Look-ahead window for calendar bootstrap.
_CALENDAR_LOOKAHEAD_DAYS = 7

# Cap on the by_sender top-N.
_TOP_SENDERS = 5


def _parse_iso(value: str) -> datetime | None:
    """Parse a Graph ISO 8601 string to an aware UTC datetime. ``None`` on failure.

    Graph hands us strings like ``"2026-05-21T10:00:00Z"`` (mail
    received times) and ``"2026-05-22T14:00:00.0000000"`` (event
    dateTime — note the trailing fractional zeros and no zone). We
    normalize both to aware UTC.
    """
    if not value:
        return None
    s = value.strip()
    # Strip trailing Z → +00:00 for fromisoformat compatibility.
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def _is_sync_state_not_found(exc: Exception) -> bool:
    """Detect Graph's ``syncStateNotFound`` (HTTP 410) on a delta call.

    The underlying delta tools use raw httpx and call ``raise_for_status``,
    so a stale token surfaces as an ``httpx.HTTPStatusError`` with a 410
    status. We treat any 410 as a sync-state failure (Graph's documented
    code for it is ``syncStateNotFound`` / ``syncStateInvalid``).
    """
    if isinstance(exc, httpx.HTTPStatusError):
        return exc.response.status_code == 410
    return False


# ── Per-resource collectors ────────────────────────────────────────────


async def _drain_mail(
    graph_client: Any,
    *,
    delta_token: str | None,
) -> tuple[list[dict], str | None]:
    """Drain mail delta pages up to the internal cap. Returns (items, token)."""
    items: list[dict] = []
    token = delta_token
    last_token: str | None = None
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_inbox_delta(
            graph_client,
            page_size=_PAGE_SIZE,
            delta_token=token,
        )
        items.extend(page["messages"])
        last_token = page["delta_token"] if page["delta_token"] is not None else last_token
        if not page["has_more"]:
            break
        token = page["delta_token"]
    return items, last_token


async def _drain_events(
    graph_client: Any,
    *,
    start: str | None,
    end: str | None,
    delta_token: str | None,
) -> tuple[list[dict], str | None]:
    """Drain calendar delta pages up to the internal cap. Returns (items, token)."""
    items: list[dict] = []
    token = delta_token
    last_token: str | None = None
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_events_delta(
            graph_client,
            start=start,
            end=end,
            page_size=_PAGE_SIZE,
            delta_token=token,
        )
        items.extend(page["events"])
        last_token = page["delta_token"] if page["delta_token"] is not None else last_token
        if not page["has_more"]:
            break
        token = page["delta_token"]
    return items, last_token


async def _drain_contacts(
    graph_client: Any,
    *,
    delta_token: str | None,
) -> tuple[list[dict], str | None]:
    """Drain contacts delta pages up to the internal cap. Returns (items, token)."""
    items: list[dict] = []
    token = delta_token
    last_token: str | None = None
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_contacts_delta(
            graph_client,
            page_size=_PAGE_SIZE,
            delta_token=token,
        )
        items.extend(page["contacts"])
        last_token = page["delta_token"] if page["delta_token"] is not None else last_token
        if not page["has_more"]:
            break
        token = page["delta_token"]
    return items, last_token


# ── Classifiers ────────────────────────────────────────────────────────


def _classify_mail(
    items: list[dict],
    *,
    window_start: datetime | None,
) -> dict:
    """Build the mail section from a list of formatted message-delta items.

    ``window_start`` is non-None on the first-call/resync path — items
    whose ``received`` predates it are excluded from new_count /
    urgent_flagged / by_sender. (Delta's first call returns a full
    snapshot; we don't want the digest surfacing 1,000 historical
    messages.)

    On a token-driven call ``window_start`` is None — every returned
    item is by definition a change since the last token, so no filtering
    needed.
    """
    new_count = 0
    modified_count = 0
    removed_count = 0
    urgent: list[dict] = []
    sender_counter: Counter[str] = Counter()

    for item in items:
        if item.get("is_deleted"):
            removed_count += 1
            continue

        # Window filter on bootstrap.
        if window_start is not None:
            received_dt = _parse_iso(item.get("received") or "")
            if received_dt is None or received_dt < window_start:
                continue

        # On a token call we can't distinguish new vs modified from a
        # delta response alone (Graph returns both in the same shape).
        # The spec treats anything not deleted as "new" on the bootstrap
        # path; on the token path we also count it under new_count so
        # the caller has a single "how much landed" number.
        new_count += 1

        is_high = (item.get("importance") or "").lower() == "high"
        is_flagged = (item.get("flag") or "") == "flagged"
        if is_high or is_flagged:
            urgent.append({
                "id": item.get("id"),
                "subject": item.get("subject") or "",
                "from_email": item.get("from_email") or "",
                "from_name": item.get("from_name") or "",
                "received_at": item.get("received") or "",
                "is_high_importance": is_high,
                "is_flagged": is_flagged,
            })

        from_email = (item.get("from_email") or "").lower()
        if from_email:
            sender_counter[from_email] += 1

    by_sender = dict(sender_counter.most_common(_TOP_SENDERS))

    return {
        "new_count": new_count,
        "modified_count": modified_count,
        "removed_count": removed_count,
        "urgent_flagged": urgent,
        "by_sender": by_sender,
    }


def _classify_events(items: list[dict]) -> dict:
    """Build the events section.

    We can't tell new-vs-modified from a delta response alone (Graph
    sends both with the same shape — modified events don't carry a
    ``@changed`` annotation). Per spec, treat anything not cancelled /
    deleted as ``new``. ``modified`` is reserved for cases where we can
    affirmatively detect a change; for now it stays empty until we add
    server-side diff tracking.
    """
    new_list: list[dict] = []
    modified_list: list[dict] = []
    cancelled_list: list[dict] = []

    for item in items:
        if item.get("is_deleted"):
            # Hard-delete tombstones — same bucket as cancellations from
            # the caller's perspective (the event is gone).
            cancelled_list.append({
                "id": item.get("id"),
                "subject": "",
                "start": "",
            })
            continue

        # ``is_cancelled`` isn't in the formatter's output (the calendar
        # delta formatter doesn't surface it). Live items therefore all
        # land in "new"; cancellations only surface via @removed
        # tombstones above. Documented in the tool docstring.
        new_list.append({
            "id": item.get("id"),
            "subject": item.get("subject") or "",
            "start": item.get("start") or "",
            "end": item.get("end") or "",
            "organizer_email": "",  # formatter exposes organizer name, not email
        })

    return {
        "new": new_list,
        "modified": modified_list,
        "cancelled": cancelled_list,
    }


def _classify_contacts(items: list[dict]) -> dict:
    """Build the contacts section (counts only)."""
    new_count = 0
    modified_count = 0
    removed_count = 0
    for item in items:
        if item.get("is_deleted"):
            removed_count += 1
        else:
            new_count += 1
    return {
        "new_count": new_count,
        "modified_count": modified_count,
        "removed_count": removed_count,
    }


# ── Resource runners ───────────────────────────────────────────────────


async def _run_mail(
    graph_client: Any,
    *,
    token: str | None,
    window_start: datetime,
    resync_log: list[str],
) -> tuple[dict, str | None]:
    """Run the mail resource end-to-end with resync recovery."""
    bootstrap_window: datetime | None = window_start if token is None else None

    try:
        items, new_token = await _drain_mail(graph_client, delta_token=token)
    except Exception as exc:
        if token is not None and _is_sync_state_not_found(exc):
            resync_log.append("mail")
            items, new_token = await _drain_mail(graph_client, delta_token=None)
            bootstrap_window = window_start
        else:
            raise

    return _classify_mail(items, window_start=bootstrap_window), new_token


async def _run_events(
    graph_client: Any,
    *,
    token: str | None,
    bootstrap_start: str,
    bootstrap_end: str,
    resync_log: list[str],
) -> tuple[dict, str | None]:
    """Run the events resource end-to-end with resync recovery."""
    if token is None:
        items, new_token = await _drain_events(
            graph_client,
            start=bootstrap_start,
            end=bootstrap_end,
            delta_token=None,
        )
    else:
        try:
            items, new_token = await _drain_events(
                graph_client,
                start=None,
                end=None,
                delta_token=token,
            )
        except Exception as exc:
            if _is_sync_state_not_found(exc):
                resync_log.append("events")
                items, new_token = await _drain_events(
                    graph_client,
                    start=bootstrap_start,
                    end=bootstrap_end,
                    delta_token=None,
                )
            else:
                raise

    return _classify_events(items), new_token


async def _run_contacts(
    graph_client: Any,
    *,
    token: str | None,
    resync_log: list[str],
) -> tuple[dict, str | None]:
    """Run the contacts resource end-to-end with resync recovery."""
    try:
        items, new_token = await _drain_contacts(graph_client, delta_token=token)
    except Exception as exc:
        if token is not None and _is_sync_state_not_found(exc):
            resync_log.append("contacts")
            items, new_token = await _drain_contacts(graph_client, delta_token=None)
        else:
            raise

    return _classify_contacts(items), new_token


# ── Public entry point ────────────────────────────────────────────────


async def changes_since(
    graph_client: Any,
    delta_tokens: dict | None = None,
    fallback_window_hours: int = 24,
) -> dict:
    """Composed "since last call" digest across mail / events / contacts.

    See module docstring for full behavior. The function calls the
    underlying v1.9.0 delta *tool functions* directly (not the FastMCP
    wrappers) so it composes cleanly without going through the server.
    """
    now = datetime.now(timezone.utc)
    window_start = now - timedelta(hours=max(0, fallback_window_hours))
    bootstrap_end_dt = now + timedelta(days=_CALENDAR_LOOKAHEAD_DAYS)

    # Format with millisecond precision; Graph accepts ISO 8601 UTC.
    bootstrap_start_iso = window_start.strftime("%Y-%m-%dT%H:%M:%SZ")
    bootstrap_end_iso = bootstrap_end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    tokens = delta_tokens or {}
    mail_token = tokens.get("mail")
    events_token = tokens.get("events")
    contacts_token = tokens.get("contacts")

    resync_log: list[str] = []

    mail_section, mail_new_token = await _run_mail(
        graph_client,
        token=mail_token,
        window_start=window_start,
        resync_log=resync_log,
    )
    events_section, events_new_token = await _run_events(
        graph_client,
        token=events_token,
        bootstrap_start=bootstrap_start_iso,
        bootstrap_end=bootstrap_end_iso,
        resync_log=resync_log,
    )
    contacts_section, contacts_new_token = await _run_contacts(
        graph_client,
        token=contacts_token,
        resync_log=resync_log,
    )

    result: dict = {
        "mail": mail_section,
        "events": events_section,
        "contacts": contacts_section,
        "delta_tokens": {
            "mail": mail_new_token,
            "events": events_new_token,
            "contacts": contacts_new_token,
        },
        "window": {
            "from": bootstrap_start_iso,
            "to": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        },
    }

    if resync_log:
        result["_meta"] = {"resync": resync_log}

    return result
