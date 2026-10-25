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
  *full snapshot*, not "what changed", so each resource is held to the
  last ``fallback_window_hours``. Mail asks Graph for it with
  ``receivedDateTime ge``, the only kind of ``$filter`` message delta
  takes. Some consumer mailboxes accept but ignore it, so the client also
  enforces the bound and reports that fallback. If the snapshot exceeds the
  cap, `mail_bootstrap_from` travels inside `delta_tokens` until it finishes;
  continuing the snapshot cannot report historical mail as new. Calendar passes
  ``[now - window, now + 7d]`` as its view, so the bootstrap snapshot
  captures recent + upcoming changes. Contacts delta takes no filter, so
  contacts are kept only when their ``lastModifiedDateTime`` falls in the
  window.
- **Subsequent call** (with tokens): use the token verbatim; drain
  pagination up to a safety cap; classify each returned item.
- **``has_more``** (per resource, beside ``delta_tokens``): the safety cap
  stopped that resource before Graph finished the round. Its token resumes
  the round; call again with it.
- A reply with no new cursor means the old one is still good, so the
  caller's token is handed back rather than ``None``.
- **`syncStateNotFound` recovery**: if a stored token is too old, Graph
  returns 410. We drop the bad token, re-do that resource as a "first
  call", and surface ``_meta.resync`` in the response so the caller
  knows their old watermark was discarded.
"""

from __future__ import annotations

import asyncio
from collections import Counter
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx

from outlook_mcp.tools._delta import parse_graph_datetime
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
    received_since: datetime | None,
) -> tuple[list[dict], str | None, bool]:
    """Drain mail delta pages up to the internal cap. Returns (items, token, has_more).

    A page with no new cursor means the one in hand is still good, so the
    token returned is never ``None`` when the caller passed one.
    """
    items: list[dict] = []
    token = delta_token
    last_token: str | None = delta_token
    has_more = False
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_inbox_delta(
            graph_client,
            page_size=_PAGE_SIZE,
            delta_token=token,
            received_since=received_since,
        )
        items.extend(page["messages"])
        last_token = page["delta_token"] if page["delta_token"] is not None else last_token
        has_more = page["has_more"]
        if not has_more:
            break
        token = page["delta_token"]
    return items, last_token, has_more


async def _drain_events(
    graph_client: Any,
    *,
    start: str | None,
    end: str | None,
    delta_token: str | None,
) -> tuple[list[dict], str | None, bool]:
    """Drain calendar delta pages up to the internal cap. See ``_drain_mail``."""
    items: list[dict] = []
    token = delta_token
    last_token: str | None = delta_token
    has_more = False
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
        has_more = page["has_more"]
        if not has_more:
            break
        token = page["delta_token"]
    return items, last_token, has_more


async def _drain_contacts(
    graph_client: Any,
    *,
    delta_token: str | None,
    modified_since: datetime | None,
) -> tuple[list[dict], str | None, bool]:
    """Drain contacts delta pages up to the internal cap. See ``_drain_mail``."""
    items: list[dict] = []
    token = delta_token
    last_token: str | None = delta_token
    has_more = False
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_contacts_delta(
            graph_client,
            page_size=_PAGE_SIZE,
            delta_token=token,
            modified_since=modified_since,
        )
        items.extend(page["contacts"])
        last_token = page["delta_token"] if page["delta_token"] is not None else last_token
        has_more = page["has_more"]
        if not has_more:
            break
        token = page["delta_token"]
    return items, last_token, has_more


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
            received_dt = parse_graph_datetime(item.get("received") or "")
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
            urgent.append(
                {
                    "id": item.get("id"),
                    "subject": item.get("subject") or "",
                    "from_email": item.get("from_email") or "",
                    "from_name": item.get("from_name") or "",
                    "received_at": item.get("received") or "",
                    "is_high_importance": is_high,
                    "is_flagged": is_flagged,
                }
            )

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
            cancelled_list.append(
                {
                    "id": item.get("id"),
                    "subject": "",
                    "start": "",
                }
            )
            continue

        # ``is_cancelled`` isn't in the formatter's output (the calendar
        # delta formatter doesn't surface it). Live items therefore all
        # land in "new"; cancellations only surface via @removed
        # tombstones above. Documented in the tool docstring.
        new_list.append(
            {
                "id": item.get("id"),
                "subject": item.get("subject") or "",
                "start": item.get("start") or "",
                "end": item.get("end") or "",
                "organizer_email": "",  # formatter exposes organizer name, not email
            }
        )

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
    bootstrap_from: datetime | None = None,
) -> tuple[dict, str | None, bool, bool, datetime | None, bool]:
    """Run the mail resource end-to-end with resync recovery.

    Returns section, cursor, has_more, resynced, pending bootstrap bound and
    whether Graph ignored the filter. ``resynced`` is
    True when a stale token forced a full re-bootstrap. Each runner reports
    its own resync so the composed digest can order them deterministically
    (the runners execute concurrently and must not share mutable state).
    """
    bootstrap_window = window_start if token is None else bootstrap_from
    resynced = False

    try:
        items, new_token, has_more = await _drain_mail(
            graph_client, delta_token=token, received_since=bootstrap_window
        )
    except Exception as exc:
        if token is not None and _is_sync_state_not_found(exc):
            resynced = True
            bootstrap_window = window_start
            items, new_token, has_more = await _drain_mail(
                graph_client, delta_token=None, received_since=bootstrap_window
            )
        else:
            raise

    section = _classify_mail(items, window_start=bootstrap_window)
    filter_ignored = bootstrap_window is not None and any(
        not item.get("is_deleted")
        and (received := parse_graph_datetime(item.get("received") or "")) is not None
        and received < bootstrap_window
        for item in items
    )
    return (
        section,
        new_token,
        has_more,
        resynced,
        bootstrap_window if has_more else None,
        filter_ignored,
    )


async def _run_events(
    graph_client: Any,
    *,
    token: str | None,
    bootstrap_start: str,
    bootstrap_end: str,
) -> tuple[dict, str | None, bool, bool]:
    """Run the events resource end-to-end with resync recovery.

    Returns ``(section, new_token, has_more, resynced)`` — see ``_run_mail``.
    """
    resynced = False
    if token is None:
        items, new_token, has_more = await _drain_events(
            graph_client,
            start=bootstrap_start,
            end=bootstrap_end,
            delta_token=None,
        )
    else:
        try:
            items, new_token, has_more = await _drain_events(
                graph_client,
                start=None,
                end=None,
                delta_token=token,
            )
        except Exception as exc:
            if _is_sync_state_not_found(exc):
                resynced = True
                items, new_token, has_more = await _drain_events(
                    graph_client,
                    start=bootstrap_start,
                    end=bootstrap_end,
                    delta_token=None,
                )
            else:
                raise

    return _classify_events(items), new_token, has_more, resynced


async def _run_contacts(
    graph_client: Any,
    *,
    token: str | None,
    window_start: datetime,
) -> tuple[dict, str | None, bool, bool]:
    """Run the contacts resource end-to-end with resync recovery.

    Returns ``(section, new_token, has_more, resynced)`` — see ``_run_mail``.
    A bootstrap keeps only contacts changed inside the window: the snapshot
    is the whole address book, not what changed in it.
    """
    resynced = False
    try:
        items, new_token, has_more = await _drain_contacts(
            graph_client,
            delta_token=token,
            modified_since=window_start if token is None else None,
        )
    except Exception as exc:
        if token is not None and _is_sync_state_not_found(exc):
            resynced = True
            items, new_token, has_more = await _drain_contacts(
                graph_client, delta_token=None, modified_since=window_start
            )
        else:
            raise

    return _classify_contacts(items), new_token, has_more, resynced


# ── Public entry point ────────────────────────────────────────────────


async def changes_since(
    graph_client: Any,
    delta_tokens: dict | None = None,
    fallback_window_hours: int = 24,
) -> dict:
    """Composed "since last call" digest across mail / events / contacts.

    See module docstring for full behavior. The function calls the
    underlying v1.9.0 delta *tool functions* directly (not the MCPServer
    wrappers) so it composes cleanly without going through the server.
    """
    now = datetime.now(timezone.utc).replace(microsecond=0)
    window_start = now - timedelta(hours=max(0, fallback_window_hours))
    bootstrap_end_dt = now + timedelta(days=_CALENDAR_LOOKAHEAD_DAYS)

    # Format with millisecond precision; Graph accepts ISO 8601 UTC.
    bootstrap_start_iso = window_start.strftime("%Y-%m-%dT%H:%M:%SZ")
    bootstrap_end_iso = bootstrap_end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    tokens = delta_tokens or {}
    mail_token = tokens.get("mail")
    events_token = tokens.get("events")
    contacts_token = tokens.get("contacts")
    mail_bootstrap_from = None
    value = tokens.get("mail_bootstrap_from")
    if mail_token is not None and value is not None:
        mail_bootstrap_from = parse_graph_datetime(value) if isinstance(value, str) else None
        if mail_bootstrap_from is None:
            raise ValueError("mail_bootstrap_from must be the timestamp returned in delta_tokens.")

    # The three resources are independent — fetch them concurrently. Each
    # runner handles its own stale-token resync internally, so a 410 on one
    # resource never cancels the others; only a genuine error propagates and
    # cancels siblings (matching the previous sequential behavior).
    (
        (
            mail_section,
            mail_new_token,
            mail_has_more,
            mail_resynced,
            mail_bootstrap_pending,
            mail_filter_ignored,
        ),
        (events_section, events_new_token, events_has_more, events_resynced),
        (contacts_section, contacts_new_token, contacts_has_more, contacts_resynced),
    ) = await asyncio.gather(
        _run_mail(
            graph_client,
            token=mail_token,
            window_start=window_start,
            bootstrap_from=mail_bootstrap_from,
        ),
        _run_events(
            graph_client,
            token=events_token,
            bootstrap_start=bootstrap_start_iso,
            bootstrap_end=bootstrap_end_iso,
        ),
        _run_contacts(
            graph_client,
            token=contacts_token,
            window_start=window_start,
        ),
    )

    # Build the resync log in fixed resource order (gather completion order is
    # nondeterministic; the wire output must not be).
    resync_log = [
        name
        for name, resynced in (
            ("mail", mail_resynced),
            ("events", events_resynced),
            ("contacts", contacts_resynced),
        )
        if resynced
    ]

    result: dict = {
        "mail": mail_section,
        "events": events_section,
        "contacts": contacts_section,
        "delta_tokens": {
            "mail": mail_new_token,
            "events": events_new_token,
            "contacts": contacts_new_token,
        },
        # Beside the tokens because it qualifies them: True means the safety
        # cap stopped that resource mid-round, and its token resumes it.
        "has_more": {
            "mail": mail_has_more,
            "events": events_has_more,
            "contacts": contacts_has_more,
        },
        "window": {
            "from": (
                mail_bootstrap_from
                if mail_bootstrap_from is not None and not mail_resynced
                else window_start
            ).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "to": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        },
    }

    if mail_bootstrap_pending is not None:
        result["delta_tokens"]["mail_bootstrap_from"] = mail_bootstrap_pending.strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
    if resync_log:
        result["_meta"] = {"resync": resync_log}
    if mail_filter_ignored:
        result.setdefault("_meta", {})["mail_window_fallback"] = (
            "Graph ignored the receivedDateTime filter; mail was bounded client-side. "
            "Replace the stored delta_tokens dict with this one; never merge them. "
            "If has_more.mail is true, pass it back to finish the snapshot with the same bound."
        )

    return result
