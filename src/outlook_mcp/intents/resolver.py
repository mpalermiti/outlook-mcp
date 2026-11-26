"""The resolver (contracts §2): names, words and exact references to items, with reasons.

``outlook_find`` is this module exposed as a tool, and every intent tool that
takes a name or search words calls the same ``resolve``. It answers with an
outcome and ranked candidates, and the outcome is what makes acting safe:

- **exact**: the words were an id, an address or a Message-ID, and it was found;
- **single**: exactly one candidate, with no blocking warning, that passes the
  known test for its kind (for a person: the owner has sent them a letter);
- **candidates**: anything else that matched. A tool that acts writes nothing and
  returns these; a read tool may go on with a single unknown one;
- **nothing_found**.

Ranking uses only Sent Items and saved contacts (``people``); received mail only
discovers people, at no score. Every Graph request is counted (``transport``).

Where the code depends on how Graph behaves on a personal mailbox, the probe that
settles it is named (``P-xx``, contracts §14) and the conservative choice is the
one built.
"""

from __future__ import annotations

import asyncio
import contextlib
import math
import re
from collections.abc import AsyncIterator
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from typing import Any
from urllib.parse import quote
from zoneinfo import ZoneInfo

from outlook_mcp import sender_auth
from outlook_mcp._windows_zones import WINDOWS_TO_IANA
from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.errors import GraphAPIError, NoMatchError
from outlook_mcp.folder_resolver import (
    fetch_all_top_level_folders,
    resolve_folder_id,
    walk_folder_tree,
)
from outlook_mcp.intents.budget import NotSupported, RequestBudget
from outlook_mcp.intents.envelope import (
    has_blocking_warning,
    others_wrote,
    sender_auth_result,
    utc,
    warn,
    why,
)
from outlook_mcp.intents.people import (
    SentTable,
    _loose,
    build_sent_table,
    is_meeting_message,
    is_valid_address,
    name_match,
    norm_name,
    parse_time,
    person_score,
    person_warnings,
    recipients,
    sender_of,
)
from outlook_mcp.intents.selection import (
    MailSelection,
    attachment_kql,
    kql_words,
    matches_locally,
    search_kql,
)
from outlook_mcp.intents.transport import IntentGraph, build_url, counted
from outlook_mcp.tools.mail_read import RECEIVED_FLOOR, SUMMARY_SELECT
from outlook_mcp.tools.todo import fetch_all_task_lists
from outlook_mcp.validation import (
    has_control_characters,
    sanitize_output,
    validate_graph_id,
    validate_internet_message_id,
)

KINDS = ("person", "message", "thread", "file", "event", "task", "draft", "rule", "folder")
PURPOSES = ("read", "act")
WORDS_MAX = 200
LIMIT_MAX = 10

# §2.6, Q1: the newest letters, 100 a page, at most 3 pages.
SENT_PATH = "/me/mailFolders/sentitems/messages"
SENT_SELECT = "from,toRecipients,ccRecipients,bccRecipients,sentDateTime,conversationId"
SENT_PAGES = 3
RECIPIENTS_SELECT = "toRecipients,ccRecipients,bccRecipients,sentDateTime"
CONTACT_SEARCH_SELECT = "id,displayName,givenName,surname,emailAddresses"
CONTACT_EXACT_SELECT = "id,displayName,emailAddresses"
SEARCH_TOP = 25
# The candidates that get the per-person reads (Q5-Q7), in one batch.
ENRICHED = 5
FIRST_SEEN_TIMEOUT_S = 5.0
HEADER_SELECT = "id,from,internetMessageHeaders"
HEADER_READS = 5
# `isDraft` beyond the contract's list: a draft is not a message anyone sent, so it
# is left out of message, thread and file results (``find(kind="draft")`` finds it).
MESSAGE_SELECT = SUMMARY_SELECT + ",toRecipients,ccRecipients,parentFolderId,isDraft"
FILE_MESSAGES = 10
ATTACHMENT_SELECT = "id,name,size,contentType,isInline"
EVENT_SELECT = (
    "id,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,isCancelled,"
    "showAs,seriesMasterId,type,attendees"
)
EVENT_CALENDARS = 5
EVENT_PAGES = 3
EVENT_TOP = 100
EVENT_DEFAULT_BACK = timedelta(days=30)
EVENT_DEFAULT_AHEAD = timedelta(days=180)
TASK_LISTS = 8
TASK_TOP = 100
ID_MIN = 32  # words this long that pass the id check are tried as an id first (§2.2)

_RESPONSES = frozenset(
    {"none", "notResponded", "accepted", "tentativelyAccepted", "declined", "organizer"}
)
_SHOW_AS = frozenset({"free", "tentative", "busy", "oof", "workingElsewhere", "unknown"})
_TASK_STATUS = frozenset({"notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"})
_IMPORTANCE = frozenset({"low", "normal", "high"})
_MIME_TYPE = re.compile(r"^[a-z0-9!#$&^_.+-]+/[a-z0-9!#$&^_.+-]+$")
_NAME_WHY = {
    "address": "exact_address",
    "full": "exact_name",
    "first_or_last": "first_name",
    "prefix": "name_prefix",
    "contains": "name_contains",
    "local_part": "name_contains",
}


# ── Inputs and results ────────────────────────────────────────────────────


@dataclass(frozen=True)
class Filters:
    """What narrows a resolution (§2.1), already validated by the calling tool."""

    sender: str | None = None  # a person reference: words or an address
    after: datetime | None = None
    before: datetime | None = None
    window: tuple[datetime, datetime] | None = None
    has_file: bool | None = None
    folder: str | None = None
    task_list: str | None = None


@dataclass
class Candidate:
    """One ranked match: the item the envelope carries, and whether it is known."""

    item: dict
    known: bool
    sort_key: tuple = ()

    @property
    def blocking(self) -> bool:
        return has_blocking_warning(self.item)


@dataclass
class Resolution:
    kind: str
    words: str
    outcome: str
    target: Candidate | None
    candidates: list[Candidate]
    total: int
    understood: dict
    sources: list[str]

    def proceed(self, purpose: str) -> Candidate | None:
        """What a tool of ``purpose`` goes on with (§2.2 step 4), or ``None`` to return
        the candidates and write nothing. Nothing found raises ``NoMatchError``;
        ``find`` reports its outcome instead of calling this."""
        if self.outcome in ("exact", "single"):
            return self.target
        if self.outcome == "nothing_found":
            raise NoMatchError(self.kind, self.words, self.sources)
        only = self.candidates[0] if len(self.candidates) == 1 else None
        if purpose == "read" and only is not None and not only.blocking:
            return only  # fails only the known test: a read goes on, warnings and all
        return None


@dataclass
class Lookup:
    """Everything one resolution reads through, for one tool call."""

    graph: IntentGraph
    sdk: Any  # the SDK client, counted against the same budget
    now: datetime
    zone: str
    signed_in: str | None
    _sent: SentTable | None = None

    async def sent_table(self) -> SentTable:
        """Q1, once per call and never kept after it (house rule; decision D6 is open)."""
        if self._sent is None:
            body = await self.graph.get(
                SENT_PATH,
                {"$select": SENT_SELECT, "$orderby": "sentDateTime desc", "$top": 100},
            )
            rows = list(body.get("value") or [])
            for _ in range(SENT_PAGES - 1):
                link = body.get("@odata.nextLink")
                if not isinstance(link, str) or not link:
                    break
                body = await self.graph.follow(link)
                rows += body.get("value") or []
            self._sent = build_sent_table(rows, signed_in=self.signed_in)
        return self._sent


@contextlib.asynccontextmanager
async def open_lookup(
    graph_client: Any,
    budget: RequestBudget,
    *,
    now: datetime,
    zone: str,
    signed_in: str | None,
) -> AsyncIterator[Lookup]:
    """A ``Lookup`` over ``graph_client`` (its credential and SDK client) for one call."""
    async with IntentGraph(graph_client.credential, budget) as graph:
        yield Lookup(
            graph=graph,
            sdk=counted(graph_client.sdk_client, budget, graph.limiter),
            now=now,
            zone=zone,
            signed_in=signed_in,
        )


def check_words(words: object, what: str = "words") -> str:
    """1-200 characters, none of them control characters (§2.1)."""
    if not isinstance(words, str) or not words.strip():
        raise ValueError(f"{what} must be 1-{WORDS_MAX} characters of names or search words.")
    if len(words) > WORDS_MAX:
        raise ValueError(f"{what} is {len(words)} characters; at most {WORDS_MAX} are read.")
    if has_control_characters(words) or any(c in words for c in "\r\n\t"):
        raise ValueError(f"{what} must be one line with no control characters.")
    return words.strip()


async def resolve(
    lookup: Lookup,
    kind: str,
    words: str,
    filters: Filters,
    purpose: str,
    limit: int = 5,
) -> Resolution:
    """Resolve ``words`` to ranked ``kind`` candidates (§2.1), with the outcome of §2.2."""
    if kind not in KINDS:
        raise ValueError(f"kind must be one of: {', '.join(KINDS)}.")
    if purpose not in PURPOSES:
        raise ValueError(f"purpose must be one of: {', '.join(PURPOSES)}.")
    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= LIMIT_MAX:
        raise ValueError(f"limit must be a whole number from 1 to {LIMIT_MAX}.")
    words = check_words(words)
    if filters.sender is not None:
        check_words(filters.sender, "sender")

    if kind == "person":
        found = await _people(lookup, words, limit, light=False)
    elif kind in ("message", "thread", "file", "draft"):
        found = await _mail(lookup, kind, words, filters, purpose, limit)
    elif kind == "event":
        found = await _events(lookup, words, filters)
    elif kind == "task":
        found = await _tasks(lookup, words, filters)
    elif kind == "rule":
        found = await _rules(lookup, words)
    else:
        found = await _folders(lookup, words)
    return _finish(lookup, kind, words, filters, limit, found)


@dataclass
class _Found:
    """What one kind's search found, before the act-on-match rule is applied."""

    ranked: list[Candidate] = field(default_factory=list)
    exact: bool = False
    sources: list[str] = field(default_factory=list)
    # Set when the result is about another kind: an ambiguous sender's people.
    kind: str | None = None
    notes: dict = field(default_factory=dict)


def _finish(
    lookup: Lookup, kind: str, words: str, filters: Filters, limit: int, found: _Found
) -> Resolution:
    ranked = sorted(found.ranked, key=lambda c: c.sort_key)
    if found.exact and ranked:
        outcome, target, ranked = "exact", ranked[0], ranked[:1]
    elif len(ranked) == 1 and ranked[0].known and not ranked[0].blocking and found.kind is None:
        outcome, target = "single", ranked[0]
    elif ranked:
        outcome, target = "candidates", None
    else:
        outcome, target = "nothing_found", None
    understood = {
        "kind": kind,
        "words": sanitize_output(words),
        "filters": _filters_echo(filters),
        "queries": list(lookup.graph.queries),
        **found.notes,
    }
    if found.kind is not None:
        understood["matched_kind"] = found.kind
    return Resolution(
        kind=kind,
        words=words,
        outcome=outcome,
        target=target,
        candidates=ranked[:limit],
        total=len(ranked),
        understood=understood,
        sources=found.sources,
    )


def _filters_echo(filters: Filters) -> dict:
    return {
        "sender": sanitize_output(filters.sender) if filters.sender else None,
        "after": utc(filters.after) if filters.after else None,
        "before": utc(filters.before) if filters.before else None,
        "has_file": filters.has_file,
        "folder": sanitize_output(filters.folder) if filters.folder else None,
        "task_list": sanitize_output(filters.task_list) if filters.task_list else None,
    }


# ── Shared helpers ────────────────────────────────────────────────────────


def _id_shaped(words: str) -> bool:
    """Words that could be a Graph id: tried as one first, falling through on 400 or 404."""
    if len(words) < ID_MIN:
        return False
    try:
        validate_graph_id(words)
    except ValueError:
        return False
    return True


def _missing(error: GraphAPIError) -> bool:
    return error.status_code in (400, 404)


def label_score(words: str, label: str | None) -> int:
    """How well ``words`` name a label (§2.3): exact phrase 40, every word 30, at least
    half the words 15, any word 5. Both sides compare after ``norm_name``."""
    query = norm_name(words) or _loose(words)
    text = norm_name(label or "") or _loose(label or "")
    if not query or not text:
        return 0
    if f" {query} " in f" {text} ":
        return 40
    wanted = query.split()
    tokens = set(text.split())
    hits = sum(1 for w in wanted if w in tokens)
    if hits == len(wanted):
        return 30
    if hits * 2 >= len(wanted):
        return 15
    return 5 if hits else 0


def _graph_time(value: object) -> datetime | None:
    """A Graph ``dateTimeTimeZone`` as an aware UTC datetime, or ``None``.

    Calendar views come back in UTC (no ``outlook.timezone`` preference is sent);
    flags may carry an IANA or a Windows zone name.
    """
    if not isinstance(value, dict):
        return None
    raw = value.get("dateTime")
    if not isinstance(raw, str) or len(raw) < 19:
        return None
    try:
        naive = datetime.fromisoformat(raw[:19])
    except ValueError:
        return None
    name = value.get("timeZone") or "UTC"
    try:
        zone: Any = (
            timezone.utc
            if name in ("UTC", "tzone://Microsoft/Utc")
            else ZoneInfo(WINDOWS_TO_IANA.get(name, name))
        )
    except (ValueError, KeyError, OSError):
        return None
    return naive.replace(tzinfo=zone).astimezone(timezone.utc)


def _address(value: str | None) -> str | None:
    return value if value and is_valid_address(value) else None


def _enum(value: object, allowed: frozenset[str], default: str) -> str:
    return value if isinstance(value, str) and value in allowed else default


def _time_sort(moment: datetime | None) -> float:
    return -moment.timestamp() if moment else math.inf


def _sender_warnings(address: str | None, name: str | None, table: SentTable) -> list[dict]:
    """The people warnings on a message's sender or an event's organizer, without
    ``first_time``, which belongs to people being written to."""
    if address is None:
        return []
    if address in table.me:
        return []
    found = person_warnings(
        address, name, table, written=table.get(address) is not None, contacts=[]
    )
    return [w for w in found if w["code"] != "first_time"]


def _verdict_warnings(result: str) -> list[dict]:
    if result == "failed":
        return [warn("sender_failed")]
    if result == "unverified":
        return [warn("sender_unverified")]
    return []


def _headers_of(row: dict) -> list[tuple[str, str]]:
    found = []
    for header in row.get("internetMessageHeaders") or []:
        if isinstance(header, dict):
            name, value = header.get("name"), header.get("value")
            if isinstance(name, str) and isinstance(value, str):
                found.append((name, value))
    return found


async def _header_verdicts(lookup: Lookup, rows: list[dict]) -> dict[str, dict]:
    """Microsoft's verdict on the From domain of each message, by id: one batch of
    single-message reads, which select headers as documented (contracts F5)."""
    wanted = [r["id"] for r in rows if isinstance(r.get("id"), str)]
    if not wanted:
        return {}
    replies = await lookup.graph.batch(
        [f"/me/messages/{quote(m, safe='')}?$select={HEADER_SELECT}" for m in wanted]
    )
    verdicts: dict[str, dict] = {}
    for message_id, reply in zip(wanted, replies, strict=True):
        body = reply.get("body") if isinstance(reply.get("body"), dict) else {}
        if 200 <= (reply.get("status") or 0) < 300 and "internetMessageHeaders" in body:
            address, _ = sender_of(body)
            verdicts[message_id] = sender_auth.sender_authentication(
                _headers_of(body), address or ""
            )
    return verdicts


# ── People (§2.6 Q1-Q8) ───────────────────────────────────────────────────


@dataclass
class _Person:
    address: str
    name_points: int = 0
    name_kind: str | None = None
    contact_id: str | None = None
    contact_name: str | None = None
    sent_name: str | None = None
    received_name: str | None = None
    written: tuple[int, datetime | None, str] | None = None  # count, last, basis
    first_sent: datetime | None = None
    received_only: bool = False
    first_seen: datetime | None = None
    first_seen_basis: str = "none"
    verdict: dict | None = None

    def take_name(self, points: int, kind: str | None) -> None:
        if points > self.name_points:
            self.name_points, self.name_kind = points, kind

    def score(self, now: datetime) -> int:
        count, last = (self.written[0], self.written[1]) if self.written else (0, None)
        return person_score(
            name=self.name_points,
            count=count,
            last=last,
            saved_contact=self.contact_id is not None,
            now=now,
        )

    def sort_key(self, now: datetime) -> tuple:
        last = self.written[1] if self.written else None
        return (-self.score(now), 0 if self.written else 1, _time_sort(last), self.address)


def _q(words: str, prefix: str = "") -> str | None:
    """A KQL search value from the agent's words, or ``None`` when nothing is left."""
    cleaned = kql_words(words)
    return f'"{prefix}{cleaned}"' if cleaned else None


def _conversation_filter(conversation_id: str) -> str:
    """One conversation's mail, floor first as in ``mail_thread.list_thread``; sorted
    newest first here, so the latest message leads."""
    return (
        f"receivedDateTime ge {RECEIVED_FLOOR} and conversationId eq '{_literal(conversation_id)}'"
    )


def _literal(value: str) -> str:
    """``value`` inside an OData string literal: each quote doubled."""
    return value.replace("'", "''")


def _from_filter(address: str) -> str:
    """Mail from one address, with the received-time floor first (``RECEIVED_FLOOR``:
    Graph refuses an ``$orderby`` on received time unless its clause leads)."""
    return (
        f"receivedDateTime ge {RECEIVED_FLOOR} and "
        f"from/emailAddress/address eq '{_literal(address)}'"
    )


async def _people(lookup: Lookup, words: str, limit: int, *, light: bool) -> _Found:
    """Person candidates (§2.6). ``light`` is the narrowing path: Q1, Q2 and Q8 only, at
    most five requests, with no per-person reads."""
    table = await lookup.sent_table()
    people: dict[str, _Person] = {}
    contacts: list[tuple[str, str]] = []
    sources = ["sent mail"]

    def person(address: str) -> _Person:
        found = people.get(address)
        if found is None:
            found = people[address] = _Person(address)
            entry = table.get(address)
            if entry is not None:
                found.written = (entry.count, entry.last, "scan")
                found.sent_name, found.first_sent = entry.name, entry.first
        return found

    if is_valid_address(words.strip()):
        address = words.strip().lower()
        target = person(address)
        target.take_name(100, "address")
        if not light:
            sources.append("contacts")
            body = await lookup.graph.get(
                "/me/contacts",
                {
                    "$filter": f"emailAddresses/any(a:a/address eq '{_literal(address)}')",
                    "$select": CONTACT_EXACT_SELECT,
                    "$top": 5,
                },
            )
            for row in body.get("value") or []:
                target.contact_id = row.get("id") if isinstance(row.get("id"), str) else None
                target.contact_name = row.get("displayName")
                break
            await _enrich(lookup, [target])
        return _Found(
            ranked=[_person_candidate(target, table, contacts, lookup.now)],
            exact=True,
            sources=sources,
        )

    for entry in table.entries.values():
        points, kind = name_match(words, entry.name, entry.address)
        if points:
            person(entry.address).take_name(points, kind)
    from_sent = bool(people)

    search = _q(words)
    from_contacts = False
    if search:
        sources.append("contacts")
        body = await lookup.graph.get(
            "/me/contacts",
            {"$search": search, "$select": CONTACT_SEARCH_SELECT, "$top": SEARCH_TOP},
        )
        for row in body.get("value") or []:
            display = row.get("displayName") if isinstance(row.get("displayName"), str) else None
            given_surname = " ".join(
                str(row.get(k) or "") for k in ("givenName", "surname")
            ).strip()
            for email in row.get("emailAddresses") or []:
                address = str((email or {}).get("address") or "").strip().lower()
                if not is_valid_address(address) or address in table.me:
                    continue
                contacts.append((display or "", address))
                best = max(
                    name_match(words, display, address), name_match(words, given_surname, address)
                )
                if best[0]:
                    found = person(address)
                    found.take_name(*best)
                    found.contact_id = row.get("id") if isinstance(row.get("id"), str) else None
                    found.contact_name = display
                    from_contacts = True

    if search and not light and len(people) < limit:
        # Q4: people who only wrote to the owner. They are discovered, never ranked up.
        sources.append("received mail")
        body = await lookup.graph.get(
            "/me/messages",
            {"$search": _q(words, "from:"), "$select": "from,receivedDateTime", "$top": SEARCH_TOP},
        )
        for row in body.get("value") or []:
            address, name = sender_of(row)
            if not address or address in table.me or not is_valid_address(address):
                continue
            points, kind = name_match(words, name, address)
            if not points:
                continue
            known_already = address in people
            found = person(address)
            found.take_name(points, kind)
            found.received_name = found.received_name or name
            if not known_already and found.written is None:
                found.received_only = True

    if search and not from_sent and not from_contacts:
        # Q8: people written to more than the scanned letters ago.
        sources.append("older sent mail")
        body = await lookup.graph.get(
            SENT_PATH,
            {"$search": _q(words, "to:"), "$select": RECIPIENTS_SELECT, "$top": SEARCH_TOP},
        )
        hits: dict[str, list[datetime | None]] = {}
        for row in body.get("value") or []:
            if is_meeting_message(row):
                continue
            for address, name in recipients(row):
                if address in table.me or not is_valid_address(address):
                    continue
                points, kind = name_match(words, name, address)
                if points:
                    found = person(address)
                    found.take_name(points, kind)
                    found.sent_name = found.sent_name or name
                    hits.setdefault(address, []).append(parse_time(row.get("sentDateTime")))
        for address, times in hits.items():
            found = people[address]
            if found.written is None:
                stamps = [t for t in times if t is not None]
                found.written = (len(times), max(stamps) if stamps else None, "search")
                found.received_only = False

    ordered = sorted(people.values(), key=lambda p: p.sort_key(lookup.now))
    if not light:
        await _enrich(lookup, ordered[:ENRICHED])
    return _Found(
        ranked=[_person_candidate(p, table, contacts, lookup.now) for p in people.values()],
        sources=sources,
    )


async def _enrich(lookup: Lookup, chosen: list[_Person]) -> None:
    """The per-person reads for the top candidates (§2.6): the latest received message
    and its sender check (Q6), letters beyond the scan (Q7), then first contact (Q5)
    under its own timeout."""
    if not chosen:
        return
    urls: list[str] = []
    slots: list[tuple[str, _Person]] = []
    for p in chosen:
        # Q6: the latest message from them, with its headers for the sender check.
        urls.append(
            build_url(
                "/me/messages",
                {
                    "$filter": _from_filter(p.address),
                    "$orderby": "receivedDateTime desc",
                    "$top": 1,
                    "$select": HEADER_SELECT,
                },
            )
        )
        slots.append(("latest", p))
        if p.written is None:
            # Q7: letters to them older than the scan.
            urls.append(
                build_url(
                    SENT_PATH,
                    {
                        "$search": f'"recipients:{p.address}"',
                        "$select": RECIPIENTS_SELECT,
                        "$top": SEARCH_TOP,
                    },
                )
            )
            slots.append(("letters", p))
    replies = await lookup.graph.batch(urls)

    needs_read: list[tuple[str, _Person]] = []
    for (slot, p), reply in zip(slots, replies, strict=True):
        if not 200 <= (reply.get("status") or 0) < 300:
            continue  # this person's check is left as not done; nothing is assumed
        rows = (reply.get("body") or {}).get("value") or []
        if slot == "latest" and rows:
            row = rows[0]
            _, name = sender_of(row)
            p.received_name = p.received_name or name
            if "internetMessageHeaders" in row:
                p.verdict = sender_auth.sender_authentication(_headers_of(row), p.address)
            elif isinstance(row.get("id"), str):
                # P-F5: headers may not come back on a list. The fallback is the
                # documented single-message read (contracts §2.6 Q6).
                needs_read.append((row["id"], p))
        elif slot == "letters":
            # P-SE1: KQL `recipients:` may match more than the exact address. Only
            # letters that really list it count, and meeting messages never do.
            stamps = [
                parse_time(row.get("sentDateTime"))
                for row in rows
                if not is_meeting_message(row) and any(a == p.address for a, _ in recipients(row))
            ]
            if stamps:
                dated = [s for s in stamps if s is not None]
                p.written = (len(stamps), max(dated) if dated else None, "search")
                p.received_only = False
    if needs_read:
        verdicts = await _header_verdicts(lookup, [{"id": m} for m, _ in needs_read])
        for message_id, p in needs_read:
            p.verdict = verdicts.get(message_id)

    # Q5: the first message ever received from them.
    first_urls = [
        build_url(
            "/me/messages",
            {
                "$filter": _from_filter(p.address),
                "$orderby": "receivedDateTime asc",
                "$top": 1,
                "$select": "id,receivedDateTime",
            },
        )
        for p in chosen
    ]
    try:
        # P-LAT1: on a large mailbox this walk can be long, so it has its own limit.
        firsts = await asyncio.wait_for(lookup.graph.batch(first_urls), FIRST_SEEN_TIMEOUT_S)
    except asyncio.TimeoutError:
        for p in chosen:
            p.first_seen, p.first_seen_basis = None, "none"
        return
    for p, reply in zip(chosen, firsts, strict=True):
        rows = []
        if 200 <= (reply.get("status") or 0) < 300:
            rows = (reply.get("body") or {}).get("value") or []
        seen = parse_time(rows[0].get("receivedDateTime")) if rows else None
        if seen is not None:
            p.first_seen, p.first_seen_basis = seen, "received"
        elif p.first_sent is not None:
            p.first_seen, p.first_seen_basis = p.first_sent, "sent"


def _person_candidate(
    p: _Person, table: SentTable, contacts: list[tuple[str, str]], now: datetime
) -> Candidate:
    """A person candidate (§2.4)."""
    names = [n for n in (p.contact_name, p.sent_name, p.received_name) if n]
    warnings: dict[str, dict] = {}
    for name in names or [None]:
        for w in person_warnings(
            p.address, name, table, written=p.written is not None, contacts=contacts
        ):
            warnings.setdefault(w["code"], w)
    auth = sender_auth_result(p.verdict)
    if p.received_only:
        warnings.setdefault("received_only", warn("received_only"))
    if auth["result"] == "failed":
        warnings.setdefault("sender_failed", warn("sender_failed"))

    reasons = []
    if p.name_kind:
        reasons.append(why(_NAME_WHY[p.name_kind]))
    written_to = None
    if p.written is not None:
        count, last, basis = p.written
        written_to = {"count": count, "last": utc(last) if last else None, "basis": basis}
        reasons.append(why("written_to", **{k: v for k, v in written_to.items() if v is not None}))
    if p.contact_id is not None:
        reasons.append(why("saved_contact"))

    score = p.score(now)
    item = {
        "kind": "person",
        "id": f"person:{p.address}",
        "address": p.address,
        "contact_id": p.contact_id,
        "first_seen": utc(p.first_seen) if p.first_seen else None,
        "first_seen_basis": p.first_seen_basis,
        "written_to": written_to,
        "sender_auth": auth,
        "signals": {
            "score": score,
            "name_match": p.name_kind,
            "saved_contact": p.contact_id is not None,
        },
        "why": reasons,
        "warnings": list(warnings.values()),
        "others_wrote": others_wrote(name=names[0] if names else None),
    }
    return Candidate(item=item, known=p.written is not None, sort_key=p.sort_key(now))


# ── Mail: message, thread, file, draft ────────────────────────────────────


async def _mail(
    lookup: Lookup, kind: str, words: str, filters: Filters, purpose: str, limit: int
) -> _Found:
    select = MESSAGE_SELECT

    # §2.2 step 1: exact references.
    if kind in ("message", "draft") and _id_shaped(words):
        try:
            row = await lookup.graph.get(
                f"/me/messages/{quote(words, safe='')}", {"$select": select}
            )
        except GraphAPIError as error:
            if not _missing(error):
                raise
        else:
            if kind == "draft" and row.get("isDraft") is True:
                return _Found(
                    ranked=[_draft_candidate(row, words, search=False)],
                    exact=True,
                    sources=["drafts"],
                )
            if kind == "message":
                table = await lookup.sent_table()
                return _Found(
                    ranked=[_mail_candidate(row, table, {}, words, None, search=False)],
                    exact=True,
                    sources=["mail"],
                )
    if kind == "thread" and _id_shaped(words):
        try:
            body = await lookup.graph.get(
                "/me/messages",
                {
                    "$filter": _conversation_filter(words),
                    "$orderby": "receivedDateTime desc",
                    "$top": 10,
                    "$select": MESSAGE_SELECT,
                },
            )
        except GraphAPIError as error:
            if not _missing(error):
                raise
        else:
            rows = body.get("value") or []
            if rows:
                table = await lookup.sent_table()
                count = None if body.get("@odata.nextLink") else len(rows)
                candidate = _thread_candidate(rows, table, {}, words, search=False, count=count)
                return _Found(ranked=[candidate], exact=True, sources=["mail"])
    if kind == "message" and words.startswith("<") and words.endswith(">") and "@" in words:
        # The shape of `mail_read.resolve_message_id` (reuse of its query), read here so
        # the ids come back in this process's id family (contracts §6.3).
        literal = validate_internet_message_id(words).replace("'", "''")
        body = await lookup.graph.get(
            "/me/messages",
            {"$filter": f"internetMessageId eq '{literal}'", "$select": select, "$top": 10},
        )
        table = await lookup.sent_table()
        rows = body.get("value") or []
        ranked = [_mail_candidate(r, table, {}, words, None, search=False) for r in rows]
        return _Found(ranked=ranked, exact=len(rows) == 1, sources=["mail"])

    # The sender, through the light path unless it's an address.
    senders: list[str] = []
    if filters.sender:
        if is_valid_address(filters.sender.strip()):
            senders = [filters.sender.strip().lower()]
        else:
            # A name narrows through the light path (§2.6). An ambiguous sender is
            # answered with the people, so the next call can pass one address.
            light = await _people(lookup, filters.sender, limit, light=True)
            people = _finish(lookup, "person", filters.sender, Filters(), limit, light)
            if people.outcome == "nothing_found":
                return _Found(sources=people.sources, notes={"sender_resolved": None})
            chosen = people.proceed(purpose)
            if chosen is None:
                return _Found(ranked=people.candidates, sources=people.sources, kind="person")
            senders = [chosen.item["address"]]

    path = "/me/messages"
    if kind == "draft":
        path = "/me/mailFolders/drafts/messages"
    elif filters.folder:
        folder_id = await resolve_folder_id(lookup.sdk, filters.folder)
        path = f"/me/mailFolders/{quote(folder_id, safe='')}/messages"

    selection = MailSelection(
        words=words,
        folder=filters.folder,
        from_=senders,
        newer_than=filters.after,
        older_than=filters.before,
        has_attachments=True if kind == "file" else filters.has_file,
    )
    body = await lookup.graph.get(
        path, {"$search": search_kql(selection), "$select": select, "$top": SEARCH_TOP}
    )
    rows = [r for r in body.get("value") or [] if matches_locally(r, selection)]
    if kind == "file":
        rows += await _named_in_attachments(lookup, path, select, selection, rows)
    notes = {"selection": selection.echo(), "sender_resolved": senders[0] if senders else None}
    if kind == "draft":
        # The owner's own: no sent table is needed to know them.
        ranked = [_draft_candidate(r, words) for r in rows if r.get("isDraft") is not False]
        return _Found(ranked=ranked, sources=["drafts"], notes=notes)
    rows = [r for r in rows if r.get("isDraft") is not True]
    table = await lookup.sent_table()
    sources = ["mail"]

    # Provisional order, to choose whose headers are read.
    rows.sort(key=lambda r: _mail_sort(r, words, table))
    if kind == "thread":
        threads: dict[str, list[dict]] = {}
        for row in rows:
            threads.setdefault(str(row.get("conversationId") or row.get("id")), []).append(row)
        latest = [
            max(group, key=lambda r: r.get("receivedDateTime") or "") for group in threads.values()
        ]
        latest.sort(key=lambda r: _mail_sort(r, words, table))
        verdicts = await _header_verdicts(lookup, _others(latest, table)[:HEADER_READS])
        ranked = [
            _thread_candidate(
                threads[str(r.get("conversationId") or r.get("id"))],
                table,
                verdicts,
                words,
                search=True,
            )
            for r in latest
        ]
        return _Found(ranked=ranked, sources=sources, notes=notes)

    if kind == "file":
        with_files = rows[:FILE_MESSAGES]
        replies = await lookup.graph.batch(
            [
                f"/me/messages/{quote(r['id'], safe='')}/attachments?$select={ATTACHMENT_SELECT}"
                for r in with_files
            ]
        )
        found_files: list[tuple[dict, dict]] = []
        for row, reply in zip(with_files, replies, strict=True):
            if not 200 <= (reply.get("status") or 0) < 300:
                continue
            for attachment in (reply.get("body") or {}).get("value") or []:
                if attachment.get("isInline") is True:
                    continue
                if label_score(words, attachment.get("name")) > 0:
                    found_files.append((row, attachment))
        file_rows = list({r["id"]: r for r, _ in found_files}.values())
        verdicts = await _header_verdicts(lookup, _others(file_rows, table)[:HEADER_READS])
        ranked = [_file_candidate(r, a, table, verdicts, words) for r, a in found_files]
        return _Found(ranked=ranked, sources=sources, notes=notes)

    verdicts = await _header_verdicts(lookup, _others(rows, table)[:HEADER_READS])
    ranked = [_mail_candidate(r, table, verdicts, words, filters.folder) for r in rows]
    return _Found(ranked=ranked, sources=sources, notes=notes)


async def _named_in_attachments(
    lookup: Lookup, path: str, select: str, selection: MailSelection, found: list[dict]
) -> list[dict]:
    """Messages whose attachment names hold the words, beyond those ``found`` already.

    P-AT1: whether a personal mailbox takes KQL ``attachment:``. On a 400 the file
    search is what it was without it, the words search alone; nothing is raised.
    """
    query = attachment_kql(selection)
    if query is None:
        return []
    try:
        body = await lookup.graph.get(
            path, {"$search": query, "$select": select, "$top": SEARCH_TOP}
        )
    except GraphAPIError as error:
        if error.status_code != 400:
            raise
        return []
    seen = {r.get("id") for r in found}
    return [
        r
        for r in body.get("value") or []
        if r.get("id") not in seen and matches_locally(r, selection)
    ]


def _others(rows: list[dict], table: SentTable) -> list[dict]:
    """Rows not sent by the owner: only those get a sender check."""
    return [r for r in rows if (sender_of(r)[0] or "") not in table.me]


def _sender_known(address: str | None, table: SentTable) -> bool:
    return bool(address) and (address in table.me or table.get(address) is not None)


def _mail_sort(row: dict, words: str, table: SentTable) -> tuple:
    address, _ = sender_of(row)
    score = label_score(words, row.get("subject")) + 10
    score += 15 if _sender_known(address, table) else 0
    return (-score, _time_sort(parse_time(row.get("receivedDateTime"))), str(row.get("id")))


def _from(row: dict, table: SentTable, verdicts: dict[str, dict]) -> tuple[dict, dict, list[dict]]:
    """``from`` for an item, the text that goes under ``others_wrote``, and warnings."""
    address, name = sender_of(row)
    valid = _address(address)
    verdict = verdicts.get(str(row.get("id")))
    auth = sender_auth_result(verdict) if valid not in table.me else sender_auth_result(None)
    shown = {"address": valid, "known": _sender_known(valid, table), "sender_auth": auth}
    wrote: dict = {"name": name}
    warnings = _verdict_warnings(auth["result"]) + _sender_warnings(valid, name, table)
    if address and not valid:
        wrote["address_text"] = address
        warnings.append(warn("invalid_address"))
    return shown, wrote, warnings


def _search_warnings(search: bool) -> list[dict]:
    return [warn("search_ranked"), warn("index_lag_possible")] if search else []


def _to_me(row: dict, table: SentTable) -> str:
    if any(a in table.me for a, _ in recipients(row, ("toRecipients",))):
        return "to"
    if any(a in table.me for a, _ in recipients(row, ("ccRecipients",))):
        return "cc"
    return "none"


def _flag(row: dict) -> dict | None:
    flag = row.get("flag") if isinstance(row.get("flag"), dict) else {}
    status = flag.get("flagStatus")
    if status not in ("flagged", "complete"):
        return None
    start, due = _graph_time(flag.get("startDateTime")), _graph_time(flag.get("dueDateTime"))
    return {
        "status": status,
        "start": utc(start) if start else None,
        "due": utc(due) if due else None,
    }


def _mail_candidate(
    row: dict,
    table: SentTable,
    verdicts: dict[str, dict],
    words: str,
    folder: str | None,
    *,
    search: bool = True,
) -> Candidate:
    shown, wrote, warnings = _from(row, table, verdicts)
    received_at = parse_time(row.get("receivedDateTime"))
    score = label_score(words, row.get("subject")) + (10 if search else 0)
    score += 15 if shown["known"] else 0
    reasons = [why("search_match")] if search else []
    if shown["known"] and shown["address"] not in table.me:
        reasons.append(why("sender_written_to"))
    to_me = _to_me(row, table)
    if to_me == "to":
        reasons.append(why("you_are_on_to"))
    elif to_me == "cc":
        reasons.append(why("you_are_on_cc"))
    if row.get("hasAttachments") is True:
        reasons.append(why("has_attachments"))
    if row.get("importance") == "high":
        reasons.append(why("high_importance"))
    item = {
        "kind": "message",
        "id": row.get("id"),
        "conversation_id": row.get("conversationId"),
        "received": utc(received_at) if received_at else None,
        "folder": sanitize_output(folder) if folder else None,
        "folder_id": row.get("parentFolderId"),
        "from": shown,
        "to_me": to_me,
        "has_attachments": row.get("hasAttachments") is True,
        "is_read": row.get("isRead") is True,
        "flag": _flag(row),
        "first_seen": None,
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": reasons,
        "warnings": _search_warnings(search) + warnings,
        "others_wrote": others_wrote(
            subject=row.get("subject"),
            preview=row.get("bodyPreview"),
            name=wrote.get("name"),
            address_text=wrote.get("address_text"),
        ),
    }
    known = shown["known"]
    return Candidate(
        item=item, known=known, sort_key=(-score, _time_sort(received_at), str(row.get("id")))
    )


def _thread_candidate(
    rows: list[dict],
    table: SentTable,
    verdicts: dict[str, dict],
    words: str,
    *,
    search: bool,
    count: int | None = None,
) -> Candidate:
    last = max(rows, key=lambda r: r.get("receivedDateTime") or "")
    shown, wrote, warnings = _from(last, table, verdicts)
    last_at = parse_time(last.get("receivedDateTime"))
    mine = [r for r in rows if (sender_of(r)[0] or "") in table.me]
    score = label_score(words, last.get("subject")) + (10 if search else 0)
    score += 15 if shown["known"] or mine else 0
    reasons = [why("search_match")] if search else []
    if mine:
        stamps = [t for t in (parse_time(r.get("receivedDateTime")) for r in mine) if t]
        at = {"at": utc(max(stamps))} if stamps else {}
        reasons.append(why("you_wrote_in_thread", **at))
    elif shown["known"]:
        reasons.append(why("sender_written_to"))
    from_me = shown["address"] in table.me
    conversation = last.get("conversationId")
    item = {
        "kind": "thread",
        "id": conversation,
        "conversation_id": conversation,
        "last_message_id": last.get("id"),
        "last_from": {"address": shown["address"], "known": shown["known"]},
        "last_at": utc(last_at) if last_at else None,
        "message_count": count,
        "waits_on": "them" if from_me else "me",
        "waiting_since": utc(last_at) if last_at else None,
        "first_seen": None,
        "sender_auth": shown["sender_auth"],
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": reasons,
        "warnings": _search_warnings(search) + warnings,
        "others_wrote": others_wrote(
            subject=last.get("subject"),
            preview=last.get("bodyPreview"),
            name=wrote.get("name"),
            address_text=wrote.get("address_text"),
        ),
    }
    return Candidate(
        item=item,
        known=bool(mine) or shown["known"],
        sort_key=(-score, _time_sort(last_at), str(conversation)),
    )


def _file_candidate(
    row: dict, attachment: dict, table: SentTable, verdicts: dict[str, dict], words: str
) -> Candidate:
    shown, wrote, warnings = _from(row, table, verdicts)
    received_at = parse_time(row.get("receivedDateTime"))
    score = label_score(words, attachment.get("name")) + 10 + (15 if shown["known"] else 0)
    content_type = str(attachment.get("contentType") or "").strip().lower()
    size = attachment.get("size")
    reasons = [why("search_match")]
    if shown["known"] and shown["address"] not in table.me:
        reasons.append(why("sender_written_to"))
    item = {
        "kind": "file",
        "id": attachment.get("id"),
        "message_id": row.get("id"),
        "size": size if isinstance(size, int) and not isinstance(size, bool) else None,
        "content_type": content_type if _MIME_TYPE.fullmatch(content_type) else None,
        "from": shown,
        "received": utc(received_at) if received_at else None,
        "first_seen": None,
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": reasons,
        "warnings": _search_warnings(True) + warnings,
        "others_wrote": others_wrote(
            file_name=attachment.get("name"),
            subject=row.get("subject"),
            name=wrote.get("name"),
            address_text=wrote.get("address_text"),
        ),
    }
    return Candidate(
        item=item,
        known=shown["known"],
        sort_key=(-score, _time_sort(received_at), str(attachment.get("id"))),
    )


def _draft_candidate(row: dict, words: str, *, search: bool = True) -> Candidate:
    """A draft: the owner's own, so it always passes the known test."""
    score = label_score(words, row.get("subject")) + (10 if search else 0)
    to = [a for a, _ in recipients(row, ("toRecipients", "ccRecipients")) if is_valid_address(a)]
    item = {
        "kind": "draft",
        "id": row.get("id"),
        "conversation_id": row.get("conversationId"),
        "to": to,
        "has_attachments": row.get("hasAttachments") is True,
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": [why("search_match")] if search else [],
        "warnings": _search_warnings(search),
        "others_wrote": others_wrote(subject=row.get("subject"), preview=row.get("bodyPreview")),
    }
    received_at = parse_time(row.get("receivedDateTime"))
    return Candidate(
        item=item, known=True, sort_key=(-score, _time_sort(received_at), str(row.get("id")))
    )


# ── Events ────────────────────────────────────────────────────────────────


async def _events(lookup: Lookup, words: str, filters: Filters) -> _Found:
    """Events on the owner's calendars (§2.6): calendar views, matched locally on the
    subject and on attendee names and addresses, since Graph has no search on events."""
    if _id_shaped(words):
        try:
            row = await lookup.graph.get(
                f"/me/events/{quote(words, safe='')}", {"$select": EVENT_SELECT}
            )
        except GraphAPIError as error:
            if not _missing(error):
                raise
        else:
            table = await lookup.sent_table() if row.get("isOrganizer") is not True else None
            return _Found(
                ranked=[_event_candidate(row, None, table, words, lookup)],
                exact=True,
                sources=["your calendars"],
            )

    start, end = filters.window or (
        lookup.now - EVENT_DEFAULT_BACK,
        lookup.now + EVENT_DEFAULT_AHEAD,
    )
    calendars = await fetch_all_calendars(lookup.sdk)
    calendars.sort(key=lambda c: 0 if getattr(c, "is_default_calendar", False) else 1)
    matched: dict[str, tuple[dict, str]] = {}
    for calendar in calendars[:EVENT_CALENDARS]:
        body = await lookup.graph.get(
            f"/me/calendars/{quote(calendar.id, safe='')}/calendarView",
            {
                "startDateTime": utc(start),
                "endDateTime": utc(end),
                "$select": EVENT_SELECT,
                "$top": EVENT_TOP,
            },
        )
        for page in range(EVENT_PAGES):
            for row in body.get("value") or []:
                if isinstance(row.get("id"), str) and _event_match(words, row):
                    matched.setdefault(row["id"], (row, calendar.id))
            link = body.get("@odata.nextLink")
            if page == EVENT_PAGES - 1 or not isinstance(link, str) or not link:
                break
            body = await lookup.graph.follow(link)

    needs_table = any(row.get("isOrganizer") is not True for row, _ in matched.values())
    table = await lookup.sent_table() if needs_table else None
    ranked = [_event_candidate(row, cal, table, words, lookup) for row, cal in matched.values()]
    return _Found(
        ranked=ranked,
        sources=["your calendars"],
        notes={"window": {"start": utc(start), "end": utc(end)}},
    )


def _attendees(row: dict) -> list[tuple[str, str | None]]:
    """``(address, name)`` of each attendee, addresses lowercased."""
    return list(recipients(row, ("attendees",)))


def _event_match(words: str, row: dict) -> bool:
    if label_score(words, row.get("subject")):
        return True
    people = _attendees(row) + [p for p in [sender_of({"from": row.get("organizer")})] if p[0]]
    return any(name_match(words, name, address)[0] for address, name in people if address)


def _event_candidate(
    row: dict, calendar: str | None, table: SentTable | None, words: str, lookup: Lookup
) -> Candidate:
    organizer, organizer_name = sender_of({"from": row.get("organizer")})
    organizer = _address(organizer)
    is_organizer = row.get("isOrganizer") is True or bool(table and organizer in table.me)
    written = bool(table and organizer and table.get(organizer) is not None)
    response = _enum((row.get("responseStatus") or {}).get("response"), _RESPONSES, "none")
    show_as = _enum(row.get("showAs"), _SHOW_AS, "unknown")
    start, end = _graph_time(row.get("start")), _graph_time(row.get("end"))
    with_them = any(
        name_match(words, name, address)[0] for address, name in _attendees(row) if address
    )
    subject_points = label_score(words, row.get("subject"))
    score = max(subject_points, 30 if with_them else 0) + (15 if is_organizer or written else 0)

    reasons: list[dict] = []
    if is_organizer:
        reasons.append(why("organizer_you"))
    elif written:
        reasons.append(why("organizer_written_to"))
    if not is_organizer and response in ("none", "notResponded"):
        reasons.append(why("not_responded"))
    zone = ZoneInfo(lookup.zone) if lookup.zone != "UTC" else timezone.utc
    if start and start.astimezone(zone).date() == lookup.now.astimezone(zone).date():
        reasons.append(why("today"))
    if row.get("isCancelled") is True:
        reasons.append(why("cancelled"))
    if with_them:
        reasons.append(why("meeting_with_them"))

    warnings: list[dict] = []
    if not is_organizer and not written and response in ("none", "notResponded"):
        warnings.append(warn("stranger_invite"))
    if show_as == "free":
        warnings.append(warn("shows_free"))
    if response == "declined":
        warnings.append(warn("declined"))
    if row.get("isAllDay") is True:
        warnings.append(warn("all_day"))
    if table is not None and not is_organizer:
        warnings += _sender_warnings(organizer, organizer_name, table)

    def when(moment: datetime | None) -> dict | None:
        if moment is None:
            return None
        return {
            "utc": utc(moment),
            "local": moment.astimezone(zone).isoformat(),
            "zone": lookup.zone,
        }

    item = {
        "kind": "event",
        "id": row.get("id"),
        "series_master_id": row.get("seriesMasterId"),
        "start": when(start),
        "end": when(end),
        "all_day": row.get("isAllDay") is True,
        "calendar": calendar,
        "is_organizer": is_organizer,
        "organizer": {"address": organizer, "known": is_organizer or written},
        "response": response,
        "show_as": show_as,
        "is_cancelled": row.get("isCancelled") is True,
        "attendee_count": len(row.get("attendees") or []),
        "first_seen": None,
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": reasons,
        "warnings": warnings,
        "others_wrote": others_wrote(subject=row.get("subject"), organizer_name=organizer_name),
    }
    future = start is not None and start >= lookup.now
    tie = (0, start.timestamp()) if future and start else (1, _time_sort(start))
    return Candidate(
        item=item, known=is_organizer or written, sort_key=(-score, tie, str(row.get("id")))
    )


# ── Tasks ─────────────────────────────────────────────────────────────────


def _list_name(task_list: Any) -> str:
    known = getattr(task_list, "wellknown_list_name", None)
    return str(getattr(known, "value", known) or "")


async def _tasks(lookup: Lookup, words: str, filters: Filters) -> _Found:
    """Open tasks in the owner's lists (§2.6), matched on title. The flagged-email list
    is skipped: it mirrors flagged mail.

    To Do ids are not immutable-id resources, so these replies don't speak for the
    id family (``ids=False``); P-A6 settles whether the preference touches them.
    """
    lists = [t for t in await fetch_all_task_lists(lookup.sdk) if _list_name(t) != "flaggedEmails"]
    if filters.task_list:
        wanted = filters.task_list.strip()
        lists = [
            t
            for t in lists
            if t.id == wanted or (t.display_name or "").casefold() == wanted.casefold()
        ]
        if not lists:
            raise ValueError(
                f"No task list called '{sanitize_output(wanted)[:60]}'. Pass a list's name or id."
            )
    lists = lists[:TASK_LISTS]

    if _id_shaped(words):
        for task_list in lists:
            try:
                row = await lookup.graph.get(
                    f"/me/todo/lists/{quote(task_list.id, safe='')}/tasks/{quote(words, safe='')}",
                    ids=False,
                )
            except GraphAPIError as error:
                if not _missing(error):
                    raise
                continue
            return _Found(
                ranked=[_task_candidate(row, task_list.id, words, filters, lookup)],
                exact=True,
                sources=["your task lists"],
            )

    ranked: list[Candidate] = []
    for task_list in lists:
        path = f"/me/todo/lists/{quote(task_list.id, safe='')}/tasks"
        try:
            # P-H1a: whether To Do takes this filter on a personal account. On a 400 the
            # list is read unfiltered; open tasks are picked out locally either way.
            body = await lookup.graph.get(
                path, {"$filter": "status ne 'completed'", "$top": TASK_TOP}, ids=False
            )
        except GraphAPIError as error:
            if error.status_code != 400:
                raise
            body = await lookup.graph.get(path, {"$top": TASK_TOP}, ids=False)
        for row in body.get("value") or []:
            if row.get("status") == "completed" or not label_score(words, row.get("title")):
                continue
            if filters.window and not _due_in(row, filters.window, lookup.zone):
                continue
            ranked.append(_task_candidate(row, task_list.id, words, filters, lookup))
    return _Found(ranked=ranked, sources=["your task lists"])


def _due(row: dict) -> date | None:
    raw = (
        (row.get("dueDateTime") or {}).get("dateTime")
        if isinstance(row.get("dueDateTime"), dict)
        else None
    )
    try:
        return date.fromisoformat(raw[:10]) if isinstance(raw, str) else None
    except ValueError:
        return None


def _due_in(row: dict, window: tuple[datetime, datetime], zone_name: str) -> bool:
    """Whether a task's due day starts inside ``window`` (To Do due dates are days)."""
    due = _due(row)
    if due is None:
        return False
    zone = ZoneInfo(zone_name) if zone_name != "UTC" else timezone.utc
    starts = datetime.combine(due, datetime.min.time(), tzinfo=zone)
    return window[0] <= starts < window[1]


def _task_candidate(
    row: dict, list_id: str, words: str, filters: Filters, lookup: Lookup
) -> Candidate:
    due = _due(row)
    zone = ZoneInfo(lookup.zone) if lookup.zone != "UTC" else timezone.utc
    today = lookup.now.astimezone(zone).date()
    reasons: list[dict] = []
    if due is not None:
        if due == today:
            reasons.append(why("due_today"))
        elif due < today and row.get("status") != "completed":
            reasons.append(why("overdue", due=due.isoformat()))
        if filters.window:
            reasons.append(why("due_in_window"))
    score = label_score(words, row.get("title"))
    item = {
        "kind": "task",
        "id": row.get("id"),
        "list_id": list_id,
        "due": due.isoformat() if due else None,
        "status": _enum(row.get("status"), _TASK_STATUS, "notStarted"),
        "importance": _enum(row.get("importance"), _IMPORTANCE, "normal"),
        "linked_message_id": None,
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": reasons,
        "warnings": [],
        "others_wrote": others_wrote(title=row.get("title")),
    }
    return Candidate(
        item=item,
        known=True,
        sort_key=(-score, due.isoformat() if due else "9999-12-31", str(row.get("id"))),
    )


# ── Rules and folders ─────────────────────────────────────────────────────


async def _rules(lookup: Lookup, words: str) -> _Found:
    """Inbox rules, matched on their names. Some personal mailboxes refuse rules (403):
    that is reported as not supported, not as a failure."""
    try:
        body = await lookup.graph.get("/me/mailFolders/inbox/messageRules", ids=False)
    except GraphAPIError as error:
        if error.status_code == 403:
            raise NotSupported(error) from None
        raise
    rows = [r for r in body.get("value") or [] if isinstance(r.get("id"), str)]
    exact = [r for r in rows if r["id"] == words]
    chosen = exact or [r for r in rows if label_score(words, r.get("displayName"))]
    ranked = []
    for row in chosen:
        score = 100 if exact else label_score(words, row.get("displayName"))
        sequence = row.get("sequence")
        item = {
            "kind": "rule",
            "id": row["id"],
            "enabled": row.get("isEnabled") is True,
            "sequence": sequence
            if isinstance(sequence, int) and not isinstance(sequence, bool)
            else None,
            "signals": {"score": score, "name_match": None, "saved_contact": False},
            "why": [],
            "warnings": [],
            "others_wrote": others_wrote(title=row.get("displayName")),
        }
        ranked.append(Candidate(item=item, known=True, sort_key=(-score, row["id"])))
    return _Found(ranked=ranked, exact=bool(exact), sources=["inbox rules"])


def _folder_item(folder_id: str, path: str, score: int) -> Candidate:
    item = {
        "kind": "folder",
        "id": folder_id,
        "path": sanitize_output(path),
        "signals": {"score": score, "name_match": None, "saved_contact": False},
        "why": [],
        "warnings": [],
        "others_wrote": {},
    }
    return Candidate(item=item, known=True, sort_key=(-score, path.casefold(), folder_id))


async def _folders(lookup: Lookup, words: str) -> _Found:
    """A folder by well-known name, path, id or display name (``resolve_folder_id``,
    reuse). Where that refuses (an ambiguous or unknown name), the folder tree is
    walked and every folder whose name matches the words is a candidate."""
    try:
        folder_id = await resolve_folder_id(lookup.sdk, words)
    except ValueError:
        folder_id = None
    if folder_id is not None:
        if folder_id == words.strip() and _id_shaped(folder_id):
            # An id is passed through unchecked by the resolver: confirm it exists.
            try:
                row = await lookup.graph.get(
                    f"/me/mailFolders/{quote(folder_id, safe='')}",
                    {"$select": "id,displayName"},
                    ids=False,
                )
            except GraphAPIError as error:
                if not _missing(error):
                    raise
            else:
                name = (
                    row.get("displayName") if isinstance(row.get("displayName"), str) else folder_id
                )
                return _Found(
                    ranked=[_folder_item(folder_id, name, 100)], exact=True, sources=["folders"]
                )
        else:
            return _Found(
                ranked=[_folder_item(folder_id, words.strip(), 100)],
                exact=True,
                sources=["folders"],
            )

    ranked = []
    top = await fetch_all_top_level_folders(lookup.sdk)
    async for folder, path in walk_folder_tree(lookup.sdk, top):
        score = label_score(words, folder.display_name)
        if score and isinstance(folder.id, str):
            ranked.append(_folder_item(folder.id, path, score))
    return _Found(ranked=ranked, sources=["folders"])
