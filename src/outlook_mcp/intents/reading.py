"""What the read tools share (contracts §12.2-§12.5): times, events, tasks and outcomes.

``catch_up``, ``open_loops``, ``agenda`` and ``brief`` each read a few Graph lists
and sort what came back into sections. The pieces they have in common live here:

- **Times.** An event's start and end as ``{utc, local, zone}`` (§1.4), and local
  days in the zone that reads the input and shows the times (§3.2). An all-day
  event is a floating date, read in that zone.
- **Events.** The event item of §1.4 and the busy rule of §12.4.
- **Tasks.** The owner's open To Do tasks, the flagged-email list skipped (§8.3).
- **Outcomes.** ``answered`` when every section finished, ``incomplete`` when one
  was skipped or failed: an answer that ran out of requests or time never reads
  as "nothing here".
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone, tzinfo
from typing import Any
from urllib.parse import quote
from zoneinfo import ZoneInfo

from outlook_mcp import sender_auth
from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents.envelope import others_wrote, sender_auth_result, utc, warn, why
from outlook_mcp.intents.people import (
    SentTable,
    is_valid_address,
    parse_time,
    recipients,
    sender_of,
)
from outlook_mcp.intents.resolver import (
    _RESPONSES,
    _SHOW_AS,
    Lookup,
    _address,
    _due,
    _enum,
    _flag,
    _graph_time,
    _headers_of,
    _list_name,
    _sender_warnings,
    _verdict_warnings,
)
from outlook_mcp.tools.todo import fetch_all_task_lists
from outlook_mcp.validation import CONFIG_REMEDY_TEXT, _first_instant_of_local_day, resolve_timezone

# §12.4: what makes an event busy. An all-day event blocks its days only when busy or away.
BUSY_SHOW_AS = frozenset({"busy", "tentative", "oof", "workingElsewhere"})
ALL_DAY_BLOCKS = frozenset({"busy", "oof"})
TASK_TOP = 100
# Calendar views are read in this order, so a capped read is the window's head.
VIEW_ORDER = "start/dateTime"
_IMPORTANCE = frozenset({"low", "normal", "high"})
_TASK_STATUS = frozenset({"notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"})

# The answer when every section finished, and when one did not (§1.2, plus the
# `incomplete` that find already answers: a section skipped for budget or time, or
# failed, is never an empty answer).
ANSWERED = "answered"
INCOMPLETE = "incomplete"


def zone_info(name: str, remedy: str | None = CONFIG_REMEDY_TEXT) -> tzinfo:
    """The zone ``name`` (an IANA name), checked as the window parser checks it."""
    resolve_timezone(name, remedy=remedy)
    return timezone.utc if name == "UTC" else ZoneInfo(name)


def local_iso(moment: datetime, tz: tzinfo) -> str:
    return moment.astimezone(tz).isoformat(timespec="seconds")


def when(moment: datetime | None, tz: tzinfo, zone_name: str) -> dict | None:
    """An instant as the envelope shows it (§1.4): UTC, local, and the zone's name."""
    if moment is None:
        return None
    return {"utc": utc(moment), "local": local_iso(moment, tz), "zone": zone_name}


def window_echo(value: str | None, start: datetime, end: datetime, zone_name: str) -> dict:
    """``understood.window`` (§3.2): the value as given, and both ends in UTC and local."""
    tz = zone_info(zone_name, remedy=None)
    return {
        "value": value,
        "start": utc(start),
        "end": utc(end),
        "zone": zone_name,
        "start_local": local_iso(start, tz),
        "end_local": local_iso(end, tz),
    }


def local_midnight(day: date, tz: tzinfo) -> datetime:
    """The first instant of ``day`` in ``tz``, as UTC (a midnight lost to a daylight-saving
    gap starts the day at the first instant after it, as ``resolve_window`` does)."""
    if tz is timezone.utc:
        return datetime.combine(day, datetime.min.time(), tzinfo=timezone.utc)
    return _first_instant_of_local_day(day, tz)  # type: ignore[arg-type]


def local_dates(start: datetime, end: datetime, tz: tzinfo) -> list[date]:
    """Every local date that ``[start, end)`` touches, in order."""
    first = start.astimezone(tz).date()
    last = (end - timedelta(microseconds=1)).astimezone(tz).date()
    return [first + timedelta(days=n) for n in range((last - first).days + 1)]


def outcome_of(sections: Iterable[dict]) -> str:
    """``answered`` unless a section was skipped or failed."""
    return ANSWERED if all(s["status"] == "ok" for s in sections) else INCOMPLETE


class Once:
    """One shared read for several sections of a call: made once, its value or error
    kept for every caller (catch_up's sent table, which mail and calendar both need)."""

    def __init__(self, make):
        self._make = make
        self._lock = asyncio.Lock()
        self._done = False
        self._value: Any = None
        self._error: BaseException | None = None

    async def get(self) -> Any:
        async with self._lock:
            if not self._done:
                try:
                    self._value = await self._make()
                except Exception as error:  # noqa: BLE001 - kept for each caller, re-raised
                    self._error = error
                self._done = True
        if self._error is not None:
            raise self._error
        return self._value


# ── Events ────────────────────────────────────────────────────────────────


def event_span(row: dict, tz: tzinfo) -> tuple[datetime | None, datetime | None]:
    """An event's start and end as aware UTC instants.

    An all-day event is a floating date: Graph writes it as midnight to midnight in
    whatever zone the reply is in (UTC here, since no ``outlook.timezone`` is asked
    for), so its dates are read as local days in ``tz``, not converted.
    """
    if row.get("isAllDay") is True:
        days = []
        for key in ("start", "end"):
            raw = (row.get(key) or {}).get("dateTime") if isinstance(row.get(key), dict) else None
            try:
                days.append(date.fromisoformat(raw[:10]) if isinstance(raw, str) else None)
            except ValueError:
                days.append(None)
        first, after = days
        if first is None:
            return None, None
        if after is None or after <= first:
            after = first + timedelta(days=1)
        return local_midnight(first, tz), local_midnight(after, tz)
    return _graph_time(row.get("start")), _graph_time(row.get("end"))


def response_of(row: dict) -> str:
    return _enum((row.get("responseStatus") or {}).get("response"), _RESPONSES, "none")


def show_as_of(row: dict) -> str:
    return _enum(row.get("showAs"), _SHOW_AS, "unknown")


def is_busy(row: dict) -> bool:
    """§12.4: not cancelled, not declined, and shown as taking the time."""
    if row.get("isCancelled") is True or response_of(row) == "declined":
        return False
    show_as = show_as_of(row)
    if row.get("isAllDay") is True:
        return show_as in ALL_DAY_BLOCKS
    return show_as in BUSY_SHOW_AS


def is_open_invite(row: dict) -> bool:
    """An invitation still waiting for the owner's answer: not theirs, not answered, not
    cancelled."""
    return (
        row.get("isOrganizer") is not True
        and response_of(row) in ("none", "notResponded")
        and row.get("isCancelled") is not True
    )


def _location(row: dict) -> str | None:
    place = row.get("location")
    name = place.get("displayName") if isinstance(place, dict) else None
    return name if isinstance(name, str) and name.strip() else None


def _event_marks(
    row: dict, *, today: date | None, start: datetime | None, tz: tzinfo, known: bool | None
) -> tuple[list[dict], list[dict]]:
    """The reasons and warnings every event item carries."""
    is_organizer = row.get("isOrganizer") is True
    response = response_of(row)
    reasons: list[dict] = []
    if is_organizer:
        reasons.append(why("organizer_you"))
    elif known:
        reasons.append(why("organizer_written_to"))
    if not is_organizer and response in ("none", "notResponded"):
        reasons.append(why("not_responded"))
    if today is not None and start is not None and start.astimezone(tz).date() == today:
        reasons.append(why("today"))
    if row.get("isCancelled") is True:
        reasons.append(why("cancelled"))
    warnings: list[dict] = []
    if show_as_of(row) == "free":
        warnings.append(warn("shows_free"))
    if response == "declined":
        warnings.append(warn("declined"))
    if row.get("isAllDay") is True:
        warnings.append(warn("all_day"))
    return reasons, warnings


def event_item(
    row: dict,
    *,
    calendar: str | None,
    tz: tzinfo,
    zone_name: str,
    now: datetime,
    table: SentTable | None = None,
) -> dict:
    """The event item of §1.4.

    With the sent table, the organizer's ``known`` is the known test (the owner
    organizes it, or has written to the organizer; Graph's ``isOrganizer`` only), and
    an open invitation from someone never written to carries ``stranger_invite``.
    Without it, ``known`` is null unless the owner organizes the event: a tool that
    reads no sent mail does not guess.
    """
    organizer, organizer_name = sender_of({"from": row.get("organizer")})
    organizer = _address(organizer)
    is_organizer = row.get("isOrganizer") is True
    if table is None:
        known: bool | None = True if is_organizer else None
    else:
        known = is_organizer or bool(organizer and table.get(organizer) is not None)
    start, end = event_span(row, tz)
    reasons, warnings = _event_marks(
        row, today=now.astimezone(tz).date(), start=start, tz=tz, known=known
    )
    if table is not None and not is_organizer:
        # Live round 1, P-SI1 (b): the owner's own event reads isOrganizer true; a
        # received meeting that names the owner as organizer reads false with response
        # `organizer`. So isOrganizer stays the test, never the organizer's address, and
        # that response on a meeting the owner doesn't organize is a stranger's
        # invitation too (as in find).
        if not known and response_of(row) in ("none", "notResponded", "organizer"):
            warnings.append(warn("stranger_invite"))
        warnings += _sender_warnings(organizer, organizer_name, table)
    attendees = row.get("attendees")
    return {
        "kind": "event",
        "id": row.get("id"),
        "series_master_id": row.get("seriesMasterId"),
        "start": when(start, tz, zone_name),
        "end": when(end, tz, zone_name),
        "all_day": row.get("isAllDay") is True,
        "calendar": calendar,
        "is_organizer": is_organizer,
        "organizer": {"address": organizer, "known": known},
        "response": response_of(row),
        "show_as": show_as_of(row),
        "is_cancelled": row.get("isCancelled") is True,
        "attendee_count": len(attendees) if isinstance(attendees, list) else None,
        "why": reasons,
        "warnings": warnings,
        "others_wrote": others_wrote(
            subject=row.get("subject"), location=_location(row), organizer_name=organizer_name
        ),
    }


async def read_view(
    lookup: Lookup,
    path: str,
    start: datetime,
    end: datetime,
    select: str,
    *,
    pages: int,
    top: int = 100,
) -> tuple[list[dict], bool]:
    """A calendar view over ``[start, end)``, in start order: at most ``pages`` pages of
    ``top``. Also whether more were left unread. In start order, a cap cuts the tail,
    so what was read is complete up to the last event read (review A1)."""
    body = await lookup.graph.get(
        path,
        {
            "startDateTime": utc(start),
            "endDateTime": utc(end),
            "$select": select,
            "$orderby": VIEW_ORDER,
            "$top": top,
        },
    )
    rows = list(body.get("value") or [])
    for _ in range(pages - 1):
        link = body.get("@odata.nextLink")
        if not isinstance(link, str) or not link:
            return rows, False
        body = await lookup.graph.follow(link)
        rows += body.get("value") or []
    link = body.get("@odata.nextLink")
    return rows, isinstance(link, str) and bool(link)


# ── Tasks ─────────────────────────────────────────────────────────────────


async def open_tasks(
    lookup: Lookup, *, lists_max: int, params: dict | None = None
) -> tuple[list[tuple[dict, str]], list[str]]:
    """The owner's open To Do tasks: ``(task, list id)`` pairs, and what was not read.

    The flagged-email list is skipped: it mirrors flagged mail (§8.3). At most
    ``lists_max`` lists, ``TASK_TOP`` tasks each; more of either goes into the second
    value (``task_lists_beyond_<n>``, ``tasks_beyond_100``). To Do ids are not
    immutable-id resources, so these replies don't speak for the id family.
    """
    lists = [t for t in await fetch_all_task_lists(lookup.sdk) if _list_name(t) != "flaggedEmails"]
    not_covered: list[str] = []
    if len(lists) > lists_max:
        not_covered.append(f"task_lists_beyond_{lists_max}")
    pairs: list[tuple[dict, str]] = []
    filtered = True
    for task_list in lists[:lists_max]:
        path = f"/me/todo/lists/{quote(task_list.id, safe='')}/tasks"
        extra = dict(params or {})
        body = None
        if filtered:
            try:
                # P-H1a: whether To Do takes this filter on a personal account. On a 400
                # the lists are read unfiltered, and open tasks picked out locally. The
                # refusal is remembered for the rest of the call: one extra request, not
                # one per list (review A2).
                body = await lookup.graph.get(
                    path,
                    {"$filter": "status ne 'completed'", "$top": TASK_TOP, **extra},
                    ids=False,
                )
            except GraphAPIError as error:
                if error.status_code != 400:
                    raise
                filtered = False
        if body is None:
            body = await lookup.graph.get(path, {"$top": TASK_TOP, **extra}, ids=False)
        rows = body.get("value") or []
        if (body.get("@odata.nextLink") or len(rows) >= TASK_TOP) and (
            "tasks_beyond_100" not in not_covered
        ):
            not_covered.append("tasks_beyond_100")
        pairs += [(row, task_list.id) for row in rows if row.get("status") != "completed"]
    return pairs, not_covered


def due_date(row: dict) -> date | None:
    """A task's due day (To Do due dates are days)."""
    return _due(row)


def task_item(row: dict, list_id: str, *, today: date, in_window: bool = False) -> dict:
    """The task item of §1.4, with its due reasons."""
    due = due_date(row)
    reasons: list[dict] = []
    if due is not None:
        if due == today:
            reasons.append(why("due_today"))
        elif due < today:
            reasons.append(why("overdue", due=due.isoformat()))
        elif in_window:
            reasons.append(why("due_in_window"))
    return {
        "kind": "task",
        "id": row.get("id"),
        "list_id": list_id,
        "due": due.isoformat() if due else None,
        "status": _enum(row.get("status"), _TASK_STATUS, "notStarted"),
        "importance": _enum(row.get("importance"), _IMPORTANCE, "normal"),
        "linked_message_id": None,
        "why": reasons,
        "warnings": [],
        "others_wrote": others_wrote(title=row.get("title")),
    }


def task_order(item: dict) -> tuple:
    """Due ascending (none last), then high importance first, then id (§12.2, §12.4)."""
    return (item["due"] or "9999-12-31", 0 if item["importance"] == "high" else 1, str(item["id"]))


def failed_like(name: str, failure: dict) -> dict:
    """``failure`` reported for another section that depended on the same read."""
    return {**failure, "section": name}


def as_str(value: Any) -> str | None:
    return value if isinstance(value, str) else None


# ── Mail: sender checks, bulk mail, and threads ───────────────────────────

HEADER_SELECT = "id,from,internetMessageHeaders"


@dataclass
class HeaderRead:
    """One message's headers, read: Microsoft's verdict on its From domain, and whether
    it is mailing-list or machine mail (§4.3)."""

    verdict: dict
    bulk: bool
    auto: bool


def _bulk_and_auto(headers: list[tuple[str, str]]) -> tuple[bool, bool]:
    """``bulk_sender``: ``List-Unsubscribe``, ``List-Id`` or ``Precedence: bulk|list``.
    ``auto_generated``: ``Auto-Submitted`` other than ``no``, ``Precedence:
    auto_reply|junk``, an ``X-Autoreply``/``X-Autorespond`` header (out-of-office
    replies), or a ``multipart/report`` body (a bounce). Names in any case."""
    bulk = auto = False
    for name, value in headers:
        key, said = name.strip().lower(), value.strip().lower()
        if key in ("list-unsubscribe", "list-id"):
            bulk = True
        elif key == "precedence" and said in ("bulk", "list"):
            bulk = True
        elif key == "precedence" and said in ("auto_reply", "junk"):
            auto = True
        elif key == "auto-submitted" and said != "no":
            auto = True
        elif key in ("x-autoreply", "x-autorespond"):
            auto = True
        elif key == "content-type" and said.startswith("multipart/report"):
            auto = True
    return bulk, auto


async def header_reads(lookup: Lookup, message_ids: list[str]) -> dict[str, HeaderRead]:
    """Each message's headers, by id: one batch of single-message reads, which select
    headers as documented (contracts F5). A read that failed is left out, so it reads
    as not checked."""
    wanted = [m for m in message_ids if isinstance(m, str)]
    if not wanted:
        return {}
    replies = await lookup.graph.batch(
        [f"/me/messages/{quote(m, safe='')}?$select={HEADER_SELECT}" for m in wanted]
    )
    found: dict[str, HeaderRead] = {}
    for message_id, reply in zip(wanted, replies, strict=True):
        body = reply.get("body") if isinstance(reply.get("body"), dict) else {}
        if 200 <= (reply.get("status") or 0) < 300 and "internetMessageHeaders" in body:
            headers = _headers_of(body)
            address, _ = sender_of(body)
            bulk, auto = _bulk_and_auto(headers)
            verdict = sender_auth.sender_authentication(headers, address or "")
            found[message_id] = HeaderRead(verdict=verdict, bulk=bulk, auto=auto)
    return found


def recipients_of(row: dict, keys: tuple[str, ...], me: frozenset[str]) -> list[str]:
    """Valid addresses on ``keys`` (To first, then Cc), the owner's own left out."""
    seen: list[str] = []
    for address, _ in recipients(row, keys):
        if address not in me and is_valid_address(address) and address not in seen:
            seen.append(address)
    return seen


def to_me(row: dict, me: frozenset[str]) -> str:
    if any(a in me for a, _ in recipients(row, ("toRecipients",))):
        return "to"
    if any(a in me for a, _ in recipients(row, ("ccRecipients",))):
        return "cc"
    return "none"


def written_reason(address: str | None, table: SentTable, basis: str) -> dict | None:
    """``sender_written_to {count, last, basis}`` for a sender the owner writes to."""
    entry = table.get(address) if address and address not in table.me else None
    if entry is None:
        return None
    params: dict = {"count": entry.count, "basis": basis}
    if entry.last is not None:
        params["last"] = utc(entry.last)
    return why("sender_written_to", **params)


def thread_item(
    latest: dict,
    *,
    waits_on: str,
    at: datetime | None,
    table: SentTable,
    read: HeaderRead | None,
    reasons: list[dict],
    waiting_for: list[str] | None = None,
) -> dict:
    """The thread item of §1.4, for the conversation whose latest message is ``latest``.

    ``waits_on`` is ``"them"`` when ``latest`` is the owner's own letter (in Sent
    Items), and ``"me"`` when it is someone else's; then ``read`` is its header read,
    or ``None`` when it was not checked.
    """
    address, name = sender_of(latest)
    valid = _address(address)
    mine = waits_on == "them"
    warnings: list[dict] = []
    if mine:
        known: bool = True
        auth = sender_auth_result(None)
    else:
        known = bool(valid and valid not in table.me and table.get(valid) is not None)
        auth = sender_auth_result(read.verdict if read else None)
        warnings += _verdict_warnings(auth["result"])
        warnings += _sender_warnings(valid, name, table)
        if address and not valid:
            warnings.append(warn("invalid_address"))
    conversation = latest.get("conversationId") or latest.get("id")
    item = {
        "kind": "thread",
        "id": conversation,
        "conversation_id": conversation,
        "last_message_id": latest.get("id"),
        "last_from": {"address": valid, "known": known},
        "last_at": utc(at) if at else None,
        "message_count": None,
        "waits_on": waits_on,
        "waiting_since": utc(at) if at else None,
        "sender_auth": auth,
        "why": reasons,
        "warnings": warnings,
        "others_wrote": others_wrote(
            subject=latest.get("subject"),
            preview=latest.get("bodyPreview"),
            name=name,
            address_text=address if address and not valid else None,
        ),
    }
    if waiting_for is not None:
        item["waiting_for"] = waiting_for
    return item


def message_item(
    row: dict,
    *,
    table: SentTable,
    read: HeaderRead | None,
    folder: str | None,
    owner: bool = False,
) -> dict:
    """The message item of §1.4. ``owner``: the message is in Sent Items, so its sender
    is the owner and there is nothing to check; otherwise ``read`` is its header read,
    or ``None`` when it was not checked."""
    address, name = sender_of(row)
    valid = _address(address)
    warnings: list[dict] = []
    if owner:
        known = True
        auth = sender_auth_result(None)
    else:
        known = bool(valid and valid not in table.me and table.get(valid) is not None)
        auth = sender_auth_result(read.verdict if read else None)
        warnings += _verdict_warnings(auth["result"])
        warnings += _sender_warnings(valid, name, table)
        if address and not valid:
            warnings.append(warn("invalid_address"))
    received = parse_time(row.get("receivedDateTime"))
    return {
        "kind": "message",
        "id": row.get("id"),
        "conversation_id": row.get("conversationId"),
        "received": utc(received) if received else None,
        "folder": folder,
        "from": {"address": valid, "known": known, "sender_auth": auth},
        "to_me": to_me(row, table.me),
        "has_attachments": row.get("hasAttachments") is True,
        "is_read": row.get("isRead") is True,
        "flag": _flag(row),
        "why": [],
        "warnings": warnings,
        "others_wrote": others_wrote(
            subject=row.get("subject"),
            preview=row.get("bodyPreview"),
            name=name,
            address_text=address if address and not valid else None,
        ),
    }
