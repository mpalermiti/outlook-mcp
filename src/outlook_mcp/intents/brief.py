"""``outlook_brief``: what's going on with a person, a meeting or a topic (contracts §12.5).

Exactly one of ``person``, ``meeting`` or ``topic``.

- **person** (a name or an address; the full resolution of §2): who they are, the
  recent mail both ways, the conversations waiting on either side, upcoming
  meetings with them, files they sent, and tasks linked to that mail or naming them.
- **meeting** (an event id, or words matched against your events from 30 days back
  to 30 days ahead): the event, who was invited and how each answered, recent mail
  among those people about it, and its files. A brief lists only what reached the
  mailbox: notes or decisions exist here only if someone emailed them.
- **topic** (search words): matching mail, events, files and tasks.

An ambiguous person or meeting answers ``candidates`` and gathers nothing else. A
person or meeting nothing matches raises ``NoMatchError``. ``window`` (30 days
unless given, at most 365) is the look-back for mail.

Requests: person 38 at most (resolution 22, then 16), meeting 25, topic 22, under
a ceiling of 40.
"""

from __future__ import annotations

import time
from datetime import datetime, timedelta, timezone
from typing import Any
from urllib.parse import quote

from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import budget
from outlook_mcp.intents.budget import INTENT_DEADLINE_S, RequestBudget, run_sections
from outlook_mcp.intents.envelope import (
    build_envelope,
    build_next,
    has_blocking_warning,
    others_wrote,
    section,
    sender_auth_result,
    suggestion,
    utc,
    warn,
    why,
)
from outlook_mcp.intents.people import (
    SentTable,
    is_meeting_type,
    is_valid_address,
    norm_name,
    parse_time,
    person_warnings,
    recipients,
    sender_of,
)
from outlook_mcp.intents.reading import (
    Once,
    event_item,
    event_span,
    folder_ids,
    header_reads,
    message_item,
    open_tasks,
    outcome_of,
    read_view,
    task_item,
    thread_item,
    window_echo,
    zone_info,
)
from outlook_mcp.intents.resolver import (
    _MIME_TYPE,
    ATTACHMENT_SELECT,
    MIME_TYPE_MAX,
    SENT_PATH,
    Filters,
    Lookup,
    Resolution,
    _immutable,
    _literal,
    check_words,
    label_score,
    open_lookup,
    resolve,
)
from outlook_mcp.intents.selection import MailSelection, kql_words, matches_locally, search_kql
from outlook_mcp.validation import resolve_window, sanitize_output

TOOL = "outlook_brief"
WINDOW_MAX_DAYS = 365
MEETING_SPAN = timedelta(days=30)
UPCOMING = timedelta(days=30)
CANDIDATES = 5
CAPS = {
    "candidates": CANDIDATES,
    "who": 1,
    "recent": 10,
    "open_loops": 10,
    "upcoming": 10,
    "files": 10,
    "tasks": 10,
    "meeting": 1,
    "attendees": 50,
    "mail": 15,
    "events": 10,
}
PERSON_SECTIONS = ("who", "recent", "open_loops", "upcoming", "files", "tasks")
MEETING_SECTIONS = ("meeting", "attendees", "recent", "files")
TOPIC_SECTIONS = ("mail", "events", "files", "tasks")
MAIL_TOP = 10  # P2, P3, M3
TOPIC_TOP = 15  # T1
PERSON_FILES = 3  # P5 (5 in §12.5: the Junk check and the header reads take the rest)
MEETING_FILES = 3  # M5
TOPIC_FILES = 5  # T3
TASK_LISTS = 6  # T4
PERSON_TASK_LISTS = 5  # P6 (6 in §12.5)
# Where a person's mail never comes from, for a brief (review E1).
NOT_THEIR_MAIL = ("junkemail", "deleteditems")
TOPIC_CALENDARS = 5  # T2
SUBJECT_WORDS = 6  # M3
MESSAGE_SELECT = (
    "id,conversationId,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,"
    "subject,bodyPreview,hasAttachments,isRead,flag,importance,parentFolderId,"
    "inferenceClassification"
)
LETTER_SELECT = MESSAGE_SELECT + ",bccRecipients"
CONTACT_SELECT = (
    "id,displayName,companyName,jobTitle,businessPhones,mobilePhone,homePhones,birthday"
)
VIEW_SELECT = (
    "id,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,isCancelled,"
    "showAs,location,attendees"
)
EVENT_SELECT = (
    "id,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,isCancelled,"
    "showAs,location,attendees,bodyPreview,hasAttachments,seriesMasterId,type"
)
TOPIC_EVENT_SELECT = (
    "id,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,isCancelled,showAs,location"
)
# P6: To Do tasks with the links back to where they came from (§12.5).
TASK_EXPAND = {"$expand": "linkedResources"}
_ATTENDEE_TYPES = frozenset({"required", "optional", "resource"})


async def read_brief(
    graph_client: Any,
    *,
    person: str | None = None,
    meeting: str | None = None,
    topic: str | None = None,
    window: str = "30d",
    zone: str = "UTC",
    signed_in: str | None = None,
    loaded_tools: frozenset[str] = frozenset(),
    now: datetime | None = None,
) -> dict:
    """A brief on one person, meeting or topic (§12.5), in the envelope.

    ``zone`` (the config's time zone) reads the window; ``signed_in`` is one of the
    owner's addresses; ``loaded_tools`` is what ``next`` may name. ``now`` is fixed
    by tests.
    """
    moment = now or datetime.now(timezone.utc)
    started = time.monotonic()

    # Every argument is checked before any request goes out.
    given = {k: v for k, v in (("person", person), ("meeting", meeting), ("topic", topic)) if v}
    if len(given) != 1 or any(v is not None and not v for v in (person, meeting, topic)):
        raise ValueError("Pass exactly one of person, meeting or topic.")
    form, words = next(iter(given.items()))
    words = check_words(words, form)
    if form == "topic" and not kql_words(words):
        raise ValueError("topic must hold at least one search word.")
    tz = zone_info(zone)
    start, end = resolve_window(window, zone, now=moment, max_days=WINDOW_MAX_DAYS)
    if start >= moment or end < moment:
        raise ValueError(
            "window is the look-back for mail: pass one that starts in the past and reaches "
            "now, such as 30d."
        )

    spent = RequestBudget(budget.CEILINGS[TOOL])
    async with open_lookup(
        graph_client, spent, now=moment, zone=zone, signed_in=signed_in
    ) as lookup:
        context = _Context(lookup, tz, zone, start, moment, started)
        if form == "person":
            found = await _person(context, words)
        elif form == "meeting":
            found = await _meeting(context, words)
        else:
            found = await _topic(context, words)
        if isinstance(found, Resolution):
            return _candidates(form, words, found, moment, spent, loaded_tools, lookup)
        if not isinstance(found, tuple):
            # The resolution ran out of time or requests (review C3).
            return _unfinished(form, words, found, moment, spent, lookup)
        sections, failed, understood, offers = found
        dropped = set(understood.pop("_drop_queries", []))
        understood.update(
            {
                form: sanitize_output(words) if form == "topic" else understood.get(form),
                "window": window_echo(window, start, end, zone),
                "queries": [q for q in lookup.graph.queries if q not in dropped],
            }
        )
    return build_envelope(
        tool=TOOL,
        outcome=outcome_of(sections),
        as_of=moment,
        understood=understood,
        sections=sections,
        failed_sections=failed,
        graph_requests=spent.spent,
        next=build_next(offers, loaded_tools),
        capped=bool(understood.get("not_covered")),
    )


class _Context:
    def __init__(self, lookup: Lookup, tz, zone: str, start: datetime, now, started: float):
        self.lookup, self.tz, self.zone, self.start, self.now = lookup, tz, zone, start, now
        self.started = started

    @property
    def remaining(self) -> float:
        return max(0.0, INTENT_DEADLINE_S - (time.monotonic() - self.started))


def _sections_from(
    names: tuple[str, ...], results: dict, build: dict
) -> tuple[list[dict], list[dict]]:
    """Each section from the run it depends on: built when that run finished, else marked
    with its status and the failure repeated under the section's own name."""
    sections, failed = [], []
    for name in names:
        run_name, make = build[name]
        result = results[run_name]
        if result.status == "ok":
            sections.append(make(result.value))
        else:
            sections.append(section(name, [], cap=CAPS[name], status=result.status))
            failed.append({**result.failure, "section": name})
    return sections, failed


# ── Shared reads ──────────────────────────────────────────────────────────


async def _files(
    lookup: Lookup, rows: list[dict], most: int, table: SentTable
) -> tuple[list[dict], int]:
    """The attachments (not inline ones) of up to ``most`` messages with files, in one
    batch, and how many of those messages' lists could not be read (review E3)."""
    with_files = [
        r for r in rows if r.get("hasAttachments") is True and isinstance(r.get("id"), str)
    ]
    with_files = with_files[:most]
    if not with_files:
        return [], 0
    replies = await lookup.graph.batch(
        [
            f"/me/messages/{quote(r['id'], safe='')}/attachments?$select={ATTACHMENT_SELECT}"
            for r in with_files
        ]
    )
    items, unread = [], 0
    for row, reply in zip(with_files, replies, strict=True):
        if not 200 <= (reply.get("status") or 0) < 300:
            unread += 1
            continue
        for attachment in (reply.get("body") or {}).get("value") or []:
            if attachment.get("isInline") is True:
                continue
            items.append(_file_item(row, attachment, table))
    return items, unread


def _file_item(row: dict, attachment: dict, table: SentTable) -> dict:
    """The file item of §1.4."""
    address, name = sender_of(row)
    entry = table.get(address or "") if address else None
    size = attachment.get("size")
    received = parse_time(row.get("receivedDateTime"))
    content_type = str(attachment.get("contentType") or "").strip().lower()
    return {
        "kind": "file",
        "id": attachment.get("id"),
        "message_id": row.get("id"),
        "size": size if isinstance(size, int) and not isinstance(size, bool) else None,
        # A plain MIME type of a sane length, as find checks it, or nothing: this field
        # sits outside `others_wrote` (review E2).
        "content_type": content_type
        if _MIME_TYPE.fullmatch(content_type) and len(content_type) <= MIME_TYPE_MAX
        else None,
        "from": {
            "address": address if address and is_valid_address(address) else None,
            "known": entry is not None and address not in table.me,
            "sender_auth": sender_auth_result(None),
        },
        "received": utc(received) if received else None,
        "why": [],
        "warnings": [],
        "others_wrote": others_wrote(
            file_name=attachment.get("name"), subject=row.get("subject"), name=name
        ),
    }


def _mine(row: dict, table: SentTable) -> bool:
    """The owner's own letter: in Sent Items (never decided by the From address)."""
    return table.sent_folder is not None and row.get("parentFolderId") == table.sent_folder


def _when(row: dict, table: SentTable) -> datetime | None:
    return parse_time(row.get("sentDateTime" if _mine(row, table) else "receivedDateTime"))


def _recent_items(rows: list[dict], table: SentTable, reads: dict | None = None) -> list[dict]:
    """Message items, newest first; the owner's letters marked as theirs. ``reads`` holds
    the header reads made for the messages ``next`` may point at."""
    ordered = sorted(
        {str(r.get("id")): r for r in rows if isinstance(r.get("id"), str)}.values(),
        key=lambda r: _when(r, table) or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True,
    )
    items = []
    for row in ordered:
        mine = _mine(row, table)
        item = message_item(
            row, table=table, read=(reads or {}).get(row.get("id")), folder=None, owner=mine
        )
        if mine:
            item["received"] = item["received"] or (
                utc(_when(row, table)) if _when(row, table) else None
            )
        items.append(item)
    return items


def _loops(rows: list[dict], table: SentTable, reads: dict | None = None) -> list[dict]:
    """Conversations among ``rows`` whose latest message is theirs to the owner (waits on
    the owner) or the owner's with nothing later from them (waits on them)."""
    by_conversation: dict[str, list[dict]] = {}
    for row in rows:
        if is_meeting_type(row):
            continue
        if not _mine(row, table) and sender_of(row)[0] in table.me:
            continue  # claims the owner's address, not in Sent Items: neither side
        by_conversation.setdefault(str(row.get("conversationId") or row.get("id")), []).append(row)
    items = []
    for rows_in in by_conversation.values():
        latest = max(
            rows_in, key=lambda r: _when(r, table) or datetime.min.replace(tzinfo=timezone.utc)
        )
        at = _when(latest, table)
        if _mine(latest, table):
            waiting_for = [
                a
                for a, _ in recipients(latest, ("toRecipients", "ccRecipients"))
                if a not in table.me
            ]
            if not waiting_for:
                continue
            items.append(
                thread_item(
                    latest,
                    waits_on="them",
                    at=at,
                    table=table,
                    read=None,
                    reasons=[],
                    waiting_for=waiting_for,
                )
            )
        elif any(a in table.me for a, _ in recipients(latest, ("toRecipients", "ccRecipients"))):
            items.append(
                thread_item(
                    latest,
                    waits_on="me",
                    at=at,
                    table=table,
                    read=(reads or {}).get(latest.get("id")),
                    reasons=[why("no_reply_from_you")],
                )
            )
    items.sort(key=lambda i: i["waiting_since"] or "")
    return items


def _linked_ids(row: dict) -> set[str]:
    found = set()
    for link in row.get("linkedResources") or []:
        if isinstance(link, dict) and isinstance(link.get("externalId"), str):
            found.add(link["externalId"])
    return found


def _task_items(pairs, *, today, message_ids: set[str], name_words: str | None, words: str | None):
    """Open tasks linked to one of ``message_ids``, or whose title holds ``name_words``
    (a person) or every one of ``words`` (a topic)."""
    items = []
    for row, list_id in pairs:
        linked = bool(_linked_ids(row) & message_ids)
        title = norm_name(str(row.get("title") or ""))
        named = bool(name_words) and f" {name_words} " in f" {title} "
        matched = bool(words) and label_score(words, row.get("title")) >= 30
        if not (linked or named or matched):
            continue
        item = task_item(row, list_id, today=today)
        if linked:
            item["why"].append(why("linked_to_message"))
            item["linked_message_id"] = sorted(_linked_ids(row) & message_ids)[0]
        items.append(item)
    return items


# ── Person (P0-P6) ────────────────────────────────────────────────────────


async def _resolved(ctx: _Context, kind: str, words: str, filters: Filters):
    """The resolution, under the call's one deadline (review C3): a ``Resolution``, or the
    ``SectionResult`` of a resolution that ran out of time or requests."""

    async def resolving():
        return await resolve(ctx.lookup, kind, words, filters, "read", CANDIDATES)

    (result,) = await run_sections([("resolution", resolving)], deadline_s=ctx.remaining)
    return result.value if result.status == "ok" else result


async def _person(ctx: _Context, words: str):
    lookup = ctx.lookup
    resolution = await _resolved(ctx, "person", words, Filters())
    if not isinstance(resolution, Resolution):
        return resolution
    chosen = resolution.proceed("read")  # nothing matched raises NoMatchError
    if chosen is None:
        return resolution
    who = chosen.item
    address = who["address"]
    table = await lookup.sent_table()

    async def contact():
        # P1: the saved contact's details, if there is one.
        if not who.get("contact_id"):
            return None
        return await lookup.graph.get(
            f"/me/contacts/{quote(who['contact_id'], safe='')}", {"$select": CONTACT_SELECT}
        )

    async def mail() -> tuple[list[dict], dict]:
        """P2: from them in the window, outside Junk and Deleted Items; P3: letters to
        them (P-SE1: only exact addresses). Then the sender check of the messages
        ``next`` may point at (review E1)."""
        received = await lookup.graph.get(
            "/me/messages",
            {
                "$filter": (
                    f"receivedDateTime ge {utc(ctx.start)} and "
                    f"from/emailAddress/address eq '{_literal(address)}'"
                ),
                "$orderby": "receivedDateTime desc",
                "$top": MAIL_TOP,
                "$select": MESSAGE_SELECT,
            },
        )
        letters = await lookup.graph.get(
            SENT_PATH,
            {"$search": f'"recipients:{address}"', "$select": LETTER_SELECT, "$top": MAIL_TOP},
        )
        excluded = await folder_ids(lookup, NOT_THEIR_MAIL)
        rows = [
            r
            for r in received.get("value") or []
            if excluded is None or r.get("parentFolderId") not in excluded.values()
        ]
        if excluded is None:
            not_covered.append("junk_not_told_apart")
        kept = []
        for row in letters.get("value") or []:
            sent = parse_time(row.get("sentDateTime"))
            if (
                not is_meeting_type(row)
                and sent is not None
                and sent >= ctx.start
                and any(a == address for a, _ in recipients(row))
            ):
                kept.append(row)
        # A search hands back regular ids (live round 1): one translation turns the kept
        # letters' ids into the ones P2's filter gave, so no message has two ids.
        rows += await _immutable(lookup, kept)
        targets = _targets(rows, table)
        reads = await header_reads(lookup, targets) if targets else {}
        return rows, reads

    not_covered: list[str] = []
    mail_once = Once(mail)

    async def files() -> list[dict]:
        rows, _ = await mail_once.get()
        theirs = [r for r in rows if not _mine(r, table)]
        items, unread = await _files(lookup, theirs, PERSON_FILES, table)
        if unread:
            not_covered.append("attachments_unread")
        return items

    async def upcoming() -> list[dict]:
        # P4: the primary calendar, now to 30 days ahead, in start order.
        rows, more = await read_view(
            lookup, "/me/calendarView", ctx.now, ctx.now + UPCOMING, VIEW_SELECT, pages=1
        )
        if more:
            not_covered.append("events_beyond_100")
        items = []
        for row in sorted(
            rows, key=lambda r: (event_span(r, ctx.tz)[0] or ctx.now, str(r.get("id")))
        ):
            people = {a for a, _ in recipients(row, ("attendees",))}
            organizer = sender_of({"from": row.get("organizer")})[0]
            if address in people or organizer == address:
                item = event_item(
                    row, calendar=None, tz=ctx.tz, zone_name=ctx.zone, now=ctx.now, table=table
                )
                item["why"].append(why("meeting_with_them"))
                items.append(item)
        return items

    async def tasks() -> list[dict]:
        # P6: open tasks linked to the recent mail, or naming the person.
        pairs, missed = await open_tasks(lookup, lists_max=PERSON_TASK_LISTS, params=TASK_EXPAND)
        not_covered.extend(missed)
        try:
            ids = {str(r.get("id")) for r in (await mail_once.get())[0]}
        except Exception:  # noqa: BLE001 - the mail section reports its own failure
            ids = set()
        name = norm_name(who["others_wrote"].get("name") or "")
        return _task_items(
            pairs,
            today=ctx.now.astimezone(ctx.tz).date(),
            message_ids=ids,
            name_words=name or None,
            words=None,
        )

    runs = [
        ("who", contact),
        ("mail", mail_once.get),
        ("files", files),
        ("upcoming", upcoming),
        ("tasks", tasks),
    ]
    results = {r.name: r for r in await run_sections(runs, deadline_s=ctx.remaining)}

    def who_section(contact_row) -> dict:
        item = dict(who)
        if contact_row:
            phones = [
                *(contact_row.get("businessPhones") or []),
                *([contact_row["mobilePhone"]] if contact_row.get("mobilePhone") else []),
                *(contact_row.get("homePhones") or []),
            ]
            item["others_wrote"] = {
                **who["others_wrote"],
                **others_wrote(
                    company=contact_row.get("companyName"),
                    job_title=contact_row.get("jobTitle"),
                    phones=", ".join(str(p) for p in phones) or None,
                    birthday=str(contact_row.get("birthday") or "")[:10] or None,
                ),
            }
        return section("who", [item], cap=1)

    build = {
        "who": ("who", who_section),
        "recent": (
            "mail",
            lambda value: section(
                "recent",
                _recent_items(value[0], table, value[1])[:50],
                cap=CAPS["recent"],
                total_exact=False,
            ),
        ),
        "open_loops": (
            "mail",
            lambda value: section(
                "open_loops",
                _loops(value[0], table, value[1]),
                cap=CAPS["open_loops"],
                total_exact=False,
            ),
        ),
        "upcoming": (
            "upcoming",
            lambda items: section(
                "upcoming",
                items,
                cap=CAPS["upcoming"],
                total_exact="events_beyond_100" not in not_covered,
            ),
        ),
        "files": (
            "files",
            lambda items: section(
                "files",
                items,
                cap=CAPS["files"],
                total_exact="attachments_unread" not in not_covered,
            ),
        ),
        "tasks": (
            "tasks",
            lambda items: section(
                "tasks",
                items,
                cap=CAPS["tasks"],
                total_exact=not any(c.startswith(("task_lists", "tasks_")) for c in not_covered),
            ),
        ),
    }
    sections, failed = _sections_from(PERSON_SECTIONS, results, build)
    understood = {"person": address, "matched": resolution.outcome, "not_covered": not_covered}
    offers = []
    loops = next(s for s in sections if s["name"] == "open_loops")["items"]
    waiting = [i for i in loops if i["waits_on"] == "me" and not has_blocking_warning(i)]
    if waiting:
        latest = max(waiting, key=lambda i: i["last_at"] or "")
        # Only where the sender check of that message ran and did not fail (review E1).
        if latest["sender_auth"]["result"] in ("verified", "unverified"):
            offers.append(
                suggestion(
                    "draft_reply",
                    "outlook_compose",
                    {"mode": "reply", "message_id": latest["last_message_id"]},
                    ["body"],
                )
            )
    recent = next(s for s in sections if s["name"] == "recent")["items"]
    rows = results["mail"].value[0] if results["mail"].status == "ok" else []
    if (
        recent
        and not has_blocking_warning(recent[0])
        and (
            _mine_id(recent[0]["id"], rows, table)
            or recent[0]["from"]["sender_auth"]["result"] in ("verified", "unverified")
        )
    ):
        offers.append(suggestion("open", "outlook_read_message", {"message_id": recent[0]["id"]}))
    return sections, failed, understood, offers


def _targets(rows: list[dict], table: SentTable) -> list[str]:
    """The messages a person brief's ``next`` may point at, when they aren't the
    owner's: the newest one (``open``) and the latest one waiting on the owner
    (``draft_reply``)."""
    targets: list[str] = []
    recent = _recent_items(rows, table)
    if recent and not _mine_id(recent[0]["id"], rows, table):
        targets.append(recent[0]["id"])
    waiting = [i for i in _loops(rows, table) if i["waits_on"] == "me"]
    if waiting:
        latest = max(waiting, key=lambda i: i["last_at"] or "")["last_message_id"]
        if latest not in targets:
            targets.append(latest)
    return [t for t in targets if isinstance(t, str)]


def _mine_id(message_id: str, rows: list[dict], table: SentTable) -> bool:
    return any(r.get("id") == message_id and _mine(r, table) for r in rows)


# ── Meeting (M0-M5) ───────────────────────────────────────────────────────


async def _meeting(ctx: _Context, words: str):
    lookup = ctx.lookup
    span = (ctx.now - MEETING_SPAN, ctx.now + MEETING_SPAN)
    resolution = await _resolved(ctx, "event", words, Filters(window=span))
    if not isinstance(resolution, Resolution):
        return resolution
    chosen = resolution.proceed("read")
    if chosen is None:
        return resolution
    event_id = chosen.item["id"]
    from_others: set[str] = set()
    not_covered: list[str] = []

    async def details():
        # M1: the event, with its attendees and body preview.
        row = await lookup.graph.get(
            f"/me/events/{quote(event_id, safe='')}", {"$select": EVENT_SELECT}
        )
        table = await lookup.sent_table()  # M2: for whether each attendee is known
        return row, table

    details_once = Once(details)

    async def recent() -> list[dict]:
        # M3: mail about it among its people.
        row, table = await details_once.get()
        subject_words = " ".join(kql_words(str(row.get("subject") or "")).split()[:SUBJECT_WORDS])
        if not subject_words:
            return []
        query = f'"{subject_words}"'
        # Built from the subject, which someone else may have written: it is searched,
        # and not echoed in `understood.queries` (only the agent's words and the
        # server's own strings are).
        from_others.add(query)
        body = await lookup.graph.get(
            "/me/messages", {"$search": query, "$select": MESSAGE_SELECT, "$top": MAIL_TOP}
        )
        people = {a for a, _ in recipients(row, ("attendees",))}
        organizer = sender_of({"from": row.get("organizer")})[0]
        if organizer:
            people.add(organizer)
        # The owner is on every message they get: a match must be one of the others.
        people -= table.me
        found = []
        for message in body.get("value") or []:
            when = parse_time(message.get("receivedDateTime"))
            participants = {a for a, _ in recipients(message, ("toRecipients", "ccRecipients"))}
            participants.add(sender_of(message)[0] or "")
            if when is not None and when >= ctx.start and participants & people:
                found.append(message)
        # A search's regular ids, translated (live round 1).
        return await _immutable(lookup, found)

    recent_once = Once(recent)

    async def files() -> list[dict]:
        row, table = await details_once.get()
        items = []
        if row.get("hasAttachments") is True:
            # M4: the event's own attachments.
            body = await lookup.graph.get(
                f"/me/events/{quote(event_id, safe='')}/attachments",
                {"$select": ATTACHMENT_SELECT},
            )
            for attachment in body.get("value") or []:
                if attachment.get("isInline") is not True:
                    item = _file_item({"id": None}, attachment, table)
                    item["message_id"] = None
                    item["event_id"] = event_id
                    items.append(item)
        # M5: the attachments of up to three recent messages.
        try:
            messages = await recent_once.get()
        except GraphAPIError:
            messages = []
        found, unread = await _files(lookup, messages, MEETING_FILES, table)
        if unread:
            not_covered.append("attachments_unread")
        return items + found

    runs = [("details", details_once.get), ("recent", recent_once.get), ("files", files)]
    results = {r.name: r for r in await run_sections(runs, deadline_s=ctx.remaining)}

    def meeting_section(value) -> dict:
        row, table = value
        item = event_item(
            row, calendar=None, tz=ctx.tz, zone_name=ctx.zone, now=ctx.now, table=table
        )
        item["others_wrote"] = {
            **item["others_wrote"],
            **others_wrote(body_preview=row.get("bodyPreview")),
        }
        return section("meeting", [item], cap=1)

    def attendees_section(value) -> dict:
        row, table = value
        return section("attendees", _attendees(row, table), cap=CAPS["attendees"])

    def recent_section(rows) -> dict:
        table = results["details"].value[1] if results["details"].status == "ok" else SentTable()
        return section("recent", _recent_items(rows, table), cap=CAPS["recent"], total_exact=False)

    build = {
        "meeting": ("details", meeting_section),
        "attendees": ("details", attendees_section),
        "recent": ("recent", recent_section),
        "files": ("files", lambda items: section("files", items, cap=CAPS["files"])),
    }
    sections, failed = _sections_from(MEETING_SECTIONS, results, build)
    offers = []
    recent_items = next(s for s in sections if s["name"] == "recent")["items"]
    if recent_items and not has_blocking_warning(recent_items[0]):
        offers.append(
            suggestion("open", "outlook_read_message", {"message_id": recent_items[0]["id"]})
        )
    owner = results["details"].value[1].me if results["details"].status == "ok" else frozenset()
    for attendee in next(s for s in sections if s["name"] == "attendees")["items"]:
        # Never a brief on the owner themself (review nit).
        if (
            attendee["written_to"] is None
            and attendee["address"] not in owner
            and not has_blocking_warning(attendee)
        ):
            offers.append(suggestion("brief_person", TOOL, {"person": attendee["address"]}))
            break
    understood = {
        "meeting": event_id,
        "matched": resolution.outcome,
        # How many searches were built from the meeting's subject, and so not echoed.
        "queries_not_echoed": len(from_others),
        "not_covered": not_covered,
    }
    understood["_drop_queries"] = sorted(from_others)
    return sections, failed, understood, offers


def _attendees(row: dict, table: SentTable) -> list[dict]:
    """Who was invited, as people (§2.4's ``Who``), each with their answer and type."""
    items = []
    for attendee in row.get("attendees") or []:
        if not isinstance(attendee, dict):
            continue
        email = attendee.get("emailAddress") or {}
        raw = str(email.get("address") or "").strip().lower()
        name = email.get("name") if isinstance(email.get("name"), str) else None
        entry = table.get(raw) if raw else None
        valid = raw if raw and is_valid_address(raw) else None
        warnings = (
            person_warnings(raw, name, table, written=entry is not None, contacts=[])
            if valid
            else [warn("invalid_address")]
        )
        status = attendee.get("status") if isinstance(attendee.get("status"), dict) else {}
        response = status.get("response")
        kind = attendee.get("type")
        written_to = None
        reasons = []
        if entry is not None and raw not in table.me:
            written_to = {
                "count": entry.count,
                "last": utc(entry.last) if entry.last else None,
                "basis": "scan",
            }
            reasons.append(
                why("written_to", **{k: v for k, v in written_to.items() if v is not None})
            )
        items.append(
            {
                "kind": "person",
                "id": f"person:{valid}" if valid else None,
                "address": valid,
                "written_to": written_to,
                "response": response
                if response
                in (
                    "none",
                    "notResponded",
                    "accepted",
                    "tentativelyAccepted",
                    "declined",
                    "organizer",
                )
                else "none",
                "type": kind if kind in _ATTENDEE_TYPES else "required",
                "why": reasons,
                "warnings": warnings,
                "others_wrote": others_wrote(
                    name=name, address_text=None if valid else raw or None
                ),
            }
        )
    return items


# ── Topic (T1-T4) ─────────────────────────────────────────────────────────


async def _topic(ctx: _Context, words: str):
    lookup = ctx.lookup
    selection = MailSelection(words=words, newer_than=ctx.start)

    async def mail():
        # T1: the words, over all mail, within the look-back.
        body = await lookup.graph.get(
            "/me/messages",
            {"$search": search_kql(selection), "$select": MESSAGE_SELECT, "$top": TOPIC_TOP},
        )
        rows = [r for r in body.get("value") or [] if matches_locally(r, selection)]
        # As built: the sent table (Q1), so a sender's `known` and the lookalike checks
        # are real, beyond §12.5's arithmetic. Its list reply also sets the id family
        # before the search's regular ids are translated (live round 1).
        table = await lookup.sent_table()
        return await _immutable(lookup, rows), table

    mail_once = Once(mail)

    async def events() -> tuple[list[dict], list[str]]:
        # T2: your calendars, 30 days back to 30 ahead, one page each.
        calendars = await fetch_all_calendars(lookup.sdk)
        calendars.sort(key=lambda c: 0 if getattr(c, "is_default_calendar", False) else 1)
        missed = ["calendars_beyond_5"] if len(calendars) > TOPIC_CALENDARS else []
        found: dict[str, tuple[dict, str]] = {}
        for calendar in calendars[:TOPIC_CALENDARS]:
            rows, more = await read_view(
                lookup,
                f"/me/calendars/{quote(calendar.id, safe='')}/calendarView",
                ctx.now - MEETING_SPAN,
                ctx.now + MEETING_SPAN,
                TOPIC_EVENT_SELECT,
                pages=1,
            )
            if more and "events_beyond_100" not in missed:
                missed.append("events_beyond_100")
            for row in rows:
                if isinstance(row.get("id"), str) and label_score(words, row.get("subject")) >= 30:
                    found.setdefault(row["id"], (row, calendar.id))
        items = [
            event_item(row, calendar=cal, tz=ctx.tz, zone_name=ctx.zone, now=ctx.now)
            for row, cal in sorted(
                found.values(), key=lambda p: (event_span(p[0], ctx.tz)[0] or ctx.now, p[0]["id"])
            )
        ]
        return items, missed

    async def files() -> list[dict]:
        rows, table = await mail_once.get()
        items, unread = await _files(lookup, rows, TOPIC_FILES, table)
        if unread:
            topic_missed.append("attachments_unread")
        return items

    async def tasks() -> list[dict]:
        pairs, missed = await open_tasks(lookup, lists_max=TASK_LISTS)
        topic_missed.extend(missed)
        return _task_items(
            pairs,
            today=ctx.now.astimezone(ctx.tz).date(),
            message_ids=set(),
            name_words=None,
            words=words,
        )

    topic_missed: list[str] = []
    runs = [("mail", mail_once.get), ("events", events), ("files", files), ("tasks", tasks)]
    results = {r.name: r for r in await run_sections(runs, deadline_s=ctx.remaining)}

    def mail_section(value) -> dict:
        rows, table = value
        items = _recent_items(rows, table)
        for item in items:
            item["why"] = [why("search_match"), *item["why"]]
            item["warnings"] = [
                warn("search_ranked"),
                warn("index_lag_possible"),
                *item["warnings"],
            ]
        return section(
            "mail",
            items,
            cap=CAPS["mail"],
            total_exact=False,
            total_display=f"{len(items)}+" if len(items) >= TOPIC_TOP else None,
        )

    not_covered: list[str] = list(topic_missed)

    def events_section(value) -> dict:
        items, missed = value
        not_covered.extend(missed)
        return section("events", items, cap=CAPS["events"], total_exact=not missed)

    build = {
        "mail": ("mail", mail_section),
        "events": ("events", events_section),
        "files": (
            "files",
            lambda items: section(
                "files",
                items,
                cap=CAPS["files"],
                total_exact="attachments_unread" not in topic_missed,
            ),
        ),
        "tasks": (
            "tasks",
            lambda items: section(
                "tasks",
                items,
                cap=CAPS["tasks"],
                total_exact=not any(c.startswith(("task_lists", "tasks_")) for c in topic_missed),
            ),
        ),
    }
    sections, failed = _sections_from(TOPIC_SECTIONS, results, build)
    offers = []
    mail_items = next(s for s in sections if s["name"] == "mail")["items"]
    if mail_items and not has_blocking_warning(mail_items[0]):
        offers.append(
            suggestion("open", "outlook_read_message", {"message_id": mail_items[0]["id"]})
        )
    understood = {"not_covered": not_covered}
    return sections, failed, understood, offers


# ── Candidates ────────────────────────────────────────────────────────────


_SECTIONS = {"person": PERSON_SECTIONS, "meeting": MEETING_SECTIONS, "topic": TOPIC_SECTIONS}


def _unfinished(form, words, result, moment, spent, lookup) -> dict:
    """A resolution that ran out of time or requests: every section says so, and the
    answer is ``incomplete`` (review C3)."""
    names = _SECTIONS[form]
    return build_envelope(
        tool=TOOL,
        outcome="incomplete",
        as_of=moment,
        understood={form: sanitize_output(words), "queries": list(lookup.graph.queries)},
        sections=[section(n, [], cap=CAPS[n], status=result.status) for n in names],
        failed_sections=[{**result.failure, "section": n} for n in names],
        graph_requests=spent.spent,
        next=[],
    )


def _candidates(form, words, resolution: Resolution, moment, spent, loaded_tools, lookup) -> dict:
    """An ambiguous person or meeting: the candidates, and nothing else gathered (§12.5).
    A search that could not finish answers ``incomplete`` with what it did not search,
    never an empty ``candidates`` (review E4)."""
    argument, field = ("person", "address") if form == "person" else ("meeting", "id")
    offers = [
        suggestion("choose", TOOL, {argument: c.item[field]})
        for c in resolution.candidates
        if c.known and not c.blocking and c.item.get(field)
    ]
    understood = {
        form: sanitize_output(words),
        "matched_kind": resolution.kind,
        "queries": list(lookup.graph.queries),
    }
    if resolution.outcome == "incomplete":
        understood["not_searched"] = list(resolution.not_searched)
    return build_envelope(
        tool=TOOL,
        outcome="incomplete" if resolution.outcome == "incomplete" else "candidates",
        as_of=moment,
        understood=understood,
        sections=[
            section(
                "candidates",
                [c.item for c in resolution.candidates],
                cap=CANDIDATES,
                total=resolution.total,
            )
        ],
        failed_sections=[],
        graph_requests=spent.spent,
        next=build_next(offers, loaded_tools),
    )
