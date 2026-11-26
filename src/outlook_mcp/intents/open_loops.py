"""``outlook_open_loops``: conversations waiting on you, or on someone else (contracts §12.3).

A conversation **waits on the owner** when its latest message is someone else's,
to them (To or Cc), and **waits on them** when its latest message is the owner's
own letter to someone else. These are candidates for the agent to judge, never
verdicts: nothing here knows whether a message needs an answer.

The data is a look-back window (14 days unless given) of Sent Items (L1), the
Inbox (L2) and the Archive (L3). Flagged letters older than the window are added
when no reply came (§8.2: F1 and the F2 reply checks, L4 and L5). Unknown senders
of what waits on the owner get a header read (L6): mailing-list and machine mail
is dropped, and the sender check is shown. ``person`` narrows to conversations a
person is in (the light path, L7); ``words`` to the conversations a search finds
(L8), reading the latest message of those outside the window (L9). 38 requests at
most, under a ceiling of 40.

The owner's mail is what is in Sent Items. A message elsewhere that claims the
owner's address (a copy of a note to self, or a forgery) is left out of the
question of who wrote last, so it can neither end a wait nor start one.
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import budget
from outlook_mcp.intents.budget import INTENT_DEADLINE_S, RequestBudget, run_sections
from outlook_mcp.intents.envelope import (
    build_envelope,
    build_next,
    has_blocking_warning,
    section,
    suggestion,
    utc,
    why,
)
from outlook_mcp.intents.flags import (
    Flags,
    flag_reasons,
    flagged_mail,
    is_flagged,
    reply_checks,
)
from outlook_mcp.intents.people import (
    SentTable,
    build_sent_table,
    is_meeting_type,
    parse_time,
    recipients,
    sender_of,
)
from outlook_mcp.intents.reading import (
    header_reads,
    outcome_of,
    recipients_of,
    thread_item,
    to_me,
    window_echo,
    written_reason,
    zone_info,
)
from outlook_mcp.intents.resolver import (
    SENT_PATH,
    Filters,
    Lookup,
    _conversation_filter,
    _finish,
    _people,
    check_words,
    open_lookup,
)
from outlook_mcp.intents.selection import kql_words
from outlook_mcp.intents.transport import build_url
from outlook_mcp.validation import resolve_window, sanitize_output

TOOL = "outlook_open_loops"
DIRECTIONS = ("on_me", "on_them", "both")
CAPS = {"waiting_on_me": 25, "waiting_on_them": 25, "candidates": 10}
WINDOW_MAX_DAYS = 90
OLDER_THAN_MAX = timedelta(days=365)
PAGE_TOP = 100
SENT_PAGES = 3  # L1
INBOX_PAGES = 3  # L2
ARCHIVE_PAGES = 1  # L3
FLAG_CHECKS = 5  # L5
HEADER_READS = 8  # L6: the top-ranked waits, whoever they claim to be from (review C1)
PERSON_LIMIT = 5  # L7
SEARCH_TOP = 25  # L8
LATEST_READS = 4  # L9 (9 in §12.3: the reply checks' confirmations took the rest)
LATEST_TOP = 3  # the newest three, so a draft on top doesn't hide the conversation
# L1's list, plus `parentFolderId` (the Sent Items folder's id, for L9 and the flags)
# and the subject and preview the item shows. The eventMessage cast asks for the
# meeting-message type (P-SI1), and Sent Items refusing it is read without it.
SENT_SELECT_PLAIN = (
    "id,conversationId,from,toRecipients,ccRecipients,bccRecipients,sentDateTime,flag,"
    "parentFolderId,subject,bodyPreview"
)
SENT_SELECT = SENT_SELECT_PLAIN + ",microsoft.graph.eventMessage/meetingMessageType"
# L2 and L3's list, plus `parentFolderId`.
RECEIVED_SELECT = (
    "id,conversationId,from,toRecipients,ccRecipients,receivedDateTime,"
    "inferenceClassification,importance,subject,bodyPreview,parentFolderId"
)
LATEST_SELECT = RECEIVED_SELECT + ",sentDateTime,flag,isDraft"
SEARCH_SELECT = "id,conversationId"
NOT_COVERED = ["rule_filed_folders", "junk", "older_than_window_unless_flagged_or_searched"]
_OFFSET = re.compile(r"^-?(\d{1,5})([mhdw])$", re.IGNORECASE)
_UNITS = {"m": "minutes", "h": "hours", "d": "days", "w": "weeks"}


def _older_than(value: str | None) -> timedelta:
    """``older_than`` as a duration: an offset such as ``3d`` (units m, h, d, w)."""
    if value is None:
        return timedelta(0)
    match = _OFFSET.fullmatch(value.strip()) if isinstance(value, str) else None
    if match is None:
        raise ValueError(
            f"older_than must be an offset such as 3d or 12h (units m, h, d, w), not "
            f"{sanitize_output(str(value)[:30])!r}."
        )
    span = timedelta(**{_UNITS[match.group(2).lower()]: int(match.group(1))})
    if span > OLDER_THAN_MAX:
        raise ValueError("older_than is at most 365 days.")
    return span


@dataclass
class _Mail:
    """One message in the data: whose it is, and when."""

    row: dict
    mine: bool
    at: datetime | None


@dataclass
class _Core:
    """What both sections read: the window's mail by conversation."""

    table: SentTable
    sent_folder: str | None
    conversations: dict[str, list[_Mail]] = field(default_factory=dict)
    capped: list[str] = field(default_factory=list)
    # When a window read was cut at its cap, only conversations whose latest message is
    # at least this new can be judged: an older one may have a later message in the part
    # not read (review C2).
    judged_since: datetime | None = None
    not_judged: int = 0
    searched_unread: int = 0
    searched_unjudged: int = 0


def _conversation_of(row: dict) -> str:
    return str(row.get("conversationId") or row.get("id"))


def _people_in(mails: list[_Mail]) -> set[str]:
    found: set[str] = set()
    for mail in mails:
        address, _ = sender_of(mail.row)
        if address:
            found.add(address)
        found |= {a for a, _ in recipients(mail.row, ("toRecipients", "ccRecipients"))}
    return found


async def read_open_loops(
    graph_client: Any,
    *,
    direction: str = "both",
    person: str | None = None,
    words: str | None = None,
    older_than: str | None = None,
    window: str = "14d",
    zone: str = "UTC",
    signed_in: str | None = None,
    loaded_tools: frozenset[str] = frozenset(),
    now: datetime | None = None,
) -> dict:
    """Conversations waiting on the owner or on others (§12.3), in the envelope.

    ``zone`` (the config's time zone) reads the window; ``signed_in`` is one of the
    owner's own addresses; ``loaded_tools`` is what ``next`` may name. ``now`` is
    fixed by tests.
    """
    moment = now or datetime.now(timezone.utc)
    started = time.monotonic()

    # Every argument is checked before any request goes out.
    if direction not in DIRECTIONS:
        raise ValueError(f"direction must be one of: {', '.join(DIRECTIONS)}.")
    if person is not None:
        person = check_words(person, "person")
    if words is not None:
        words = check_words(words)
        if not kql_words(words):
            raise ValueError("words must hold at least one search word.")
    waited = _older_than(older_than)
    zone_info(zone)
    start, end = resolve_window(window, zone, now=moment, max_days=WINDOW_MAX_DAYS)
    if start >= moment or end < moment:
        raise ValueError(
            "open_loops looks back from now: pass a window such as 14d or this_week, or "
            "A/now, that starts in the past and reaches now."
        )

    spent = RequestBudget(budget.CEILINGS[TOOL])
    async with open_lookup(
        graph_client, spent, now=moment, zone=zone, signed_in=signed_in
    ) as lookup:
        names = [
            n
            for n, d in (("waiting_on_me", "on_me"), ("waiting_on_them", "on_them"))
            if direction in (d, "both")
        ]

        async def lookups():
            """L7 and L8: who and which words narrow the answer."""
            found: str | None = None
            if person is not None:
                # L7: the light path. A name nobody matches raises; several are
                # candidates, and then nothing else is gathered.
                light = await _people(lookup, person, PERSON_LIMIT, light=True)
                people = _finish(lookup, "person", person, Filters(), PERSON_LIMIT, light)
                chosen = people.proceed("read")
                if chosen is None:
                    return people, None, None
                found = chosen.item["address"]
            hits: set[str] | None = None
            if words is not None:
                # L8: one search over all mail; its conversations are the ones kept.
                body = await lookup.graph.get(
                    "/me/messages",
                    {
                        "$search": f'"{kql_words(words)}"',
                        "$select": SEARCH_SELECT,
                        "$top": SEARCH_TOP,
                    },
                )
                hits = {_conversation_of(r) for r in body.get("value") or []}
            return None, found, hits

        # One deadline for the whole call, from its start: the lookups and the core
        # read too, not only the sections (review C3).
        (pre,) = await run_sections([("lookups", lookups)], deadline_s=_remaining(started))
        resolved: str | None = None
        searched: set[str] | None = None
        if pre.status == "ok":
            people, resolved, searched = pre.value
            if people is not None:
                return _candidates(people, person, moment, spent, loaded_tools, lookup)

        async def gather() -> _Core:
            return await _gather(lookup, start, searched)

        if pre.status == "ok":
            (core_result,) = await run_sections([("core", gather)], deadline_s=_remaining(started))
        else:
            core_result = pre
        remaining = _remaining(started)
        sections: list[dict] = []
        failed: list[dict] = []
        understood_extra: dict = {}
        exact = True
        if core_result.status != "ok":
            for name in names:
                sections.append(section(name, [], cap=CAPS[name], status=core_result.status))
                failed.append({**core_result.failure, "section": name})
        else:
            core: _Core = core_result.value

            def keep(conversation: str, mails: list[_Mail]) -> bool:
                if searched is not None and conversation not in searched:
                    return False
                return resolved is None or resolved in _people_in(mails)

            async def on_me() -> list[dict]:
                return await _waiting_on_me(lookup, core, keep, waited)

            async def on_them() -> tuple[list[dict], dict, bool]:
                return await _waiting_on_them(lookup, core, keep, waited, start, resolved, searched)

            runs = []
            if "waiting_on_me" in names:
                runs.append(("waiting_on_me", on_me))
            if "waiting_on_them" in names:
                runs.append(("waiting_on_them", on_them))
            results = {r.name: r for r in await run_sections(runs, deadline_s=remaining)}
            for name in names:
                result = results[name]
                if result.status != "ok":
                    sections.append(section(name, [], cap=CAPS[name], status=result.status))
                    failed.append(result.failure)
                    continue
                if name == "waiting_on_me":
                    items = result.value
                    sections.append(
                        section(name, items, cap=CAPS[name], total_exact=not core.capped)
                    )
                else:
                    items, notes, flags_exact = result.value
                    understood_extra.update(notes)
                    exact = flags_exact and not core.capped
                    sections.append(section(name, items, cap=CAPS[name], total_exact=exact))
        queries = list(lookup.graph.queries)

    not_covered = list(NOT_COVERED)
    outcome = outcome_of(sections)
    if core_result.status == "ok":
        core = core_result.value
        not_covered += core.capped
        if core.judged_since is not None:
            # A window read was cut: older conversations were not judged (review C2).
            understood_extra["judged_since"] = utc(core.judged_since)
            understood_extra["conversations_not_judged"] = core.not_judged
            outcome = "incomplete"
        if core.searched_unread:
            not_covered.append("searched_conversations_unread")
            understood_extra["searched_conversations_unread"] = core.searched_unread
        if core.searched_unjudged:
            understood_extra["searched_conversations_unjudged"] = core.searched_unjudged
    not_covered += understood_extra.pop("not_covered", [])
    understood = {
        "window": window_echo(window, start, end, zone),
        "direction": direction,
        "person": resolved,
        "words": sanitize_output(words) if words else None,
        "older_than": older_than,
        "queries": queries,
        "not_covered": not_covered,
        **understood_extra,
    }
    return build_envelope(
        tool=TOOL,
        outcome=outcome,
        as_of=moment,
        understood=understood,
        sections=sections,
        failed_sections=failed,
        graph_requests=spent.spent,
        next=build_next(_suggestions(sections), loaded_tools),
        capped=len(not_covered) > len(NOT_COVERED)
        or bool(understood.get("reply_checks_failed"))
        or bool(understood.get("replies_unconfirmed"))
        or not exact,
    )


async def _pages(lookup: Lookup, path: str, params: dict, pages: int) -> tuple[list[dict], bool]:
    body = await lookup.graph.get(path, params)
    rows = list(body.get("value") or [])
    for _ in range(pages - 1):
        link = body.get("@odata.nextLink")
        if not isinstance(link, str) or not link:
            return rows, False
        body = await lookup.graph.follow(link)
        rows += body.get("value") or []
    return rows, bool(body.get("@odata.nextLink"))


async def _sent_rows(lookup: Lookup, start: datetime) -> tuple[list[dict], bool]:
    """L1: the owner's letters in the window, newest first."""
    query = {
        "$filter": f"sentDateTime ge {utc(start)}",
        "$orderby": "sentDateTime desc",
        "$select": SENT_SELECT,
        "$top": PAGE_TOP,
    }
    try:
        return await _pages(lookup, SENT_PATH, query, SENT_PAGES)
    except GraphAPIError as error:
        # P-SI1: Sent Items refusing the eventMessage cast is read without it; the
        # type annotation and the subject still mark meeting messages.
        if error.status_code != 400:
            raise
        return await _pages(lookup, SENT_PATH, {**query, "$select": SENT_SELECT_PLAIN}, SENT_PAGES)


async def _sent_folder_id(lookup: Lookup, table: SentTable) -> str | None:
    """The Sent Items folder's id: off the window's letters, the sent table a person
    lookup read, or one read of the folder when neither had any."""
    if table.sent_folder is not None:
        return table.sent_folder
    if lookup._sent is not None and lookup._sent.sent_folder is not None:
        return lookup._sent.sent_folder
    body = await lookup.graph.get("/me/mailFolders/sentitems", {"$select": "id"}, ids=False)
    folder = body.get("id")
    return folder if isinstance(folder, str) else None


async def _gather(lookup: Lookup, start: datetime, searched: set[str] | None) -> _Core:
    """L1, L2, L3 and L9: the window's mail, and the searched conversations' latest."""
    window_filter = {
        "$filter": f"receivedDateTime ge {utc(start)}",
        "$orderby": "receivedDateTime desc",
        "$select": RECEIVED_SELECT,
        "$top": PAGE_TOP,
    }
    sent_rows, sent_more = await _sent_rows(lookup, start)
    inbox, inbox_more = await _pages(
        lookup, "/me/mailFolders/inbox/messages", window_filter, INBOX_PAGES
    )
    capped = []
    if sent_more:
        capped.append("sent_beyond_300")
    if inbox_more:
        capped.append("inbox_beyond_300")
    try:
        archive, archive_more = await _pages(
            lookup, "/me/mailFolders/archive/messages", window_filter, ARCHIVE_PAGES
        )
    except GraphAPIError as error:
        # An account without an Archive folder: said, not assumed empty.
        if error.status_code != 404:
            raise
        archive, archive_more = [], False
        capped.append("archive")
    if archive_more:
        capped.append("archive_beyond_100")

    table = build_sent_table(sent_rows, signed_in=lookup.signed_in)
    core = _Core(table=table, sent_folder=table.sent_folder, capped=capped)
    cut_at = [
        _oldest(rows, key)
        for rows, more, key in (
            (sent_rows, sent_more, "sentDateTime"),
            (inbox, inbox_more, "receivedDateTime"),
            (archive, archive_more, "receivedDateTime"),
        )
        if more
    ]
    cut_at = [t for t in cut_at if t is not None]
    if cut_at:
        core.judged_since = max(cut_at)
    for row in sent_rows:
        core.conversations.setdefault(_conversation_of(row), []).append(
            _Mail(row, mine=True, at=parse_time(row.get("sentDateTime")))
        )
    for row in inbox + archive:
        _add_received(core, row)

    if searched:
        outside = sorted(searched - set(core.conversations))
        if len(outside) > LATEST_READS:
            capped.append(f"searched_conversations_beyond_{LATEST_READS}")
        outside = outside[:LATEST_READS]
        if outside:
            core.sent_folder = await _sent_folder_id(lookup, table)
            replies = await lookup.graph.batch(
                [
                    build_url(
                        "/me/messages",
                        {
                            "$filter": _conversation_filter(c),
                            "$orderby": "receivedDateTime desc",
                            "$top": LATEST_TOP,
                            "$select": LATEST_SELECT,
                        },
                    )
                    for c in outside
                ]
            )
            for conversation, reply in zip(outside, replies, strict=True):
                rows = (reply.get("body") or {}).get("value") or []
                if not 200 <= (reply.get("status") or 0) < 300:
                    core.searched_unread += 1
                    continue
                # A draft is not a message anyone sent: the conversation is judged by its
                # latest message that is not one (review nit).
                sent = [r for r in rows if r.get("isDraft") is not True]
                if not sent:
                    core.searched_unjudged += 1
                    continue
                for row in sent[:1]:
                    if core.sent_folder is not None and row.get("parentFolderId") == (
                        core.sent_folder
                    ):
                        core.conversations.setdefault(conversation, []).append(
                            _Mail(row, mine=True, at=parse_time(row.get("sentDateTime")))
                        )
                    else:
                        _add_received(core, row)
    return core


def _remaining(started: float) -> float:
    """What is left of the call's deadline (read at call time, so tests can shorten it)."""
    return max(0.0, INTENT_DEADLINE_S - (time.monotonic() - started))


def _oldest(rows: list[dict], key: str) -> datetime | None:
    times = [t for t in (parse_time(r.get(key)) for r in rows) if t is not None]
    return min(times) if times else None


def _judged(core: _Core, latest: _Mail) -> bool:
    """Whether a conversation can be judged on the mail read (review C2)."""
    if core.judged_since is None or (latest.at is not None and latest.at >= core.judged_since):
        return True
    core.not_judged += 1
    return False


def _add_received(core: _Core, row: dict) -> None:
    """A message from outside Sent Items. One that claims the owner's address is left
    out of who wrote last: it is not theirs (not in Sent Items), and not someone
    else's either."""
    address, _ = sender_of(row)
    if address in core.table.me:
        return
    core.conversations.setdefault(_conversation_of(row), []).append(
        _Mail(row, mine=False, at=parse_time(row.get("receivedDateTime")))
    )


def _latest(mails: list[_Mail]) -> _Mail | None:
    dated = [m for m in mails if m.at is not None]
    return max(dated, key=lambda m: m.at) if dated else None  # type: ignore[arg-type,return-value]


async def _waiting_on_me(lookup: Lookup, core: _Core, keep, waited: timedelta) -> list[dict]:
    """Conversations whose latest message is someone else's, to the owner (§12.3)."""
    me = core.table.me
    found: list[tuple[_Mail, str]] = []
    for conversation, mails in core.conversations.items():
        latest = _latest(mails)
        if latest is None or latest.mine or not keep(conversation, mails):
            continue
        if not _judged(core, latest):
            continue
        row = latest.row
        line = to_me(row, me)
        if (
            is_meeting_type(row)
            or line == "none"
            or row.get("inferenceClassification") == "other"
            or (latest.at is not None and lookup.now - latest.at < waited)
        ):
            continue
        found.append((latest, line))

    def rank(entry: tuple[_Mail, str]) -> tuple:
        latest, line = entry
        entry_in_table = core.table.get(sender_of(latest.row)[0] or "")
        count = entry_in_table.count if entry_in_table else 0
        return (0 if line == "to" else 1, -count, latest.at, _conversation_of(latest.row))

    found.sort(key=rank)
    # L6: the sender check and the bulk check for the waits shown first, whoever they
    # claim to be from: a spoof of someone the owner writes to ranks high, so it is the
    # one that most needs reading (review C1).
    top = [latest.row["id"] for latest, _ in found if isinstance(latest.row.get("id"), str)][
        :HEADER_READS
    ]
    reads = await header_reads(lookup, top)

    items = []
    for latest, line in found:
        read = reads.get(str(latest.row.get("id")))
        if read is not None and (read.bulk or read.auto):
            continue  # mailing-list or machine mail does not wait on anyone
        address = sender_of(latest.row)[0]
        reasons = [why("you_are_on_to" if line == "to" else "you_are_on_cc")]
        written = written_reason(address, core.table, "window")
        if written is not None:
            reasons.append(written)
        reasons.append(why("no_reply_from_you"))
        items.append(
            thread_item(
                latest.row,
                waits_on="me",
                at=latest.at,
                table=core.table,
                read=read,
                reasons=reasons,
            )
        )
    return items


async def _waiting_on_them(
    lookup: Lookup,
    core: _Core,
    keep,
    waited: timedelta,
    start: datetime,
    resolved: str | None,
    searched: set[str] | None,
) -> tuple[list[dict], dict, bool]:
    """Conversations whose latest message is the owner's letter to someone else, and
    flagged letters older than the window that no reply followed (§12.3, §8.2)."""
    me = core.table.me
    found: list[tuple[dict, datetime | None]] = []
    for conversation, mails in core.conversations.items():
        latest = _latest(mails)
        if latest is None or not latest.mine or not keep(conversation, mails):
            continue
        if not _judged(core, latest):
            continue
        if is_meeting_type(latest.row):
            continue
        if not recipients_of(latest.row, ("toRecipients", "ccRecipients"), me):
            continue  # a letter only to oneself waits on no one
        if latest.at is not None and lookup.now - latest.at < waited:
            continue
        found.append((latest.row, latest.at))

    # L4, L5: flagged letters older than the window, when no reply followed.
    flags: Flags = await flagged_mail(lookup, fallback_folders=("sentitems",))
    sent_folder = core.sent_folder or await _sent_folder_id(lookup, core.table)
    older: list[tuple[dict, datetime]] = []
    for row in flags.rows:
        at = parse_time(row.get("sentDateTime"))
        conversation = _conversation_of(row)
        if (
            row.get("parentFolderId") != sent_folder
            or at is None
            or at >= start
            or conversation in core.conversations
            or is_meeting_type(row)
            or lookup.now - at < waited
            or (searched is not None and conversation not in searched)
            or (resolved is not None and resolved not in _people_in([_Mail(row, True, at)]))
            or not recipients_of(row, ("toRecipients", "ccRecipients"), me)
        ):
            continue
        older.append((row, at))
    older.sort(key=lambda pair: (_due_key(pair[0]), pair[1]))
    notes: dict = {"not_covered": []}
    if flags.beyond_first_page:
        notes["not_covered"].append("flags_beyond_first_page")
    if len(older) > FLAG_CHECKS:
        notes["not_covered"].append(f"flagged_letters_beyond_{FLAG_CHECKS}")
    checked = older[:FLAG_CHECKS]
    checks = await reply_checks(lookup, checked, sent_folder=sent_folder, me=me)
    failed_checks, unconfirmed = [], []
    for row, at in checked:
        check = checks[str(row.get("id"))]
        if not check.checked:
            # Not checked is not "no reply": the letter stays, and the answer says so.
            failed_checks.append(row.get("id"))
        elif check.reply is not None:
            continue  # they answered: a checked reply from whom it went to (review B1)
        elif check.unconfirmed:
            # Something came back that could not be confirmed as their reply: the
            # follow-up stays open, and the answer says so.
            unconfirmed.append(row.get("id"))
        found.append((row, at))
    if failed_checks:
        notes["reply_checks_failed"] = failed_checks
    if unconfirmed:
        notes["replies_unconfirmed"] = unconfirmed

    items = []
    for row, at in sorted(
        found, key=lambda pair: (pair[1] or lookup.now, _conversation_of(pair[0]))
    ):
        reasons = flag_reasons(row, lookup.now) if is_flagged(row) else []
        waiting_for = recipients_of(row, ("toRecipients",), me) + [
            a
            for a in recipients_of(row, ("ccRecipients",), me)
            if a not in recipients_of(row, ("toRecipients",), me)
        ]
        items.append(
            thread_item(
                row,
                waits_on="them",
                at=at,
                table=core.table,
                read=None,
                reasons=reasons,
                waiting_for=waiting_for,
            )
        )
    return items, notes, flags.exact


def _due_key(row: dict) -> float:
    from outlook_mcp.intents.flags import flag_due

    due = flag_due(row)
    return due.timestamp() if due else float("inf")


def _suggestions(sections: list[dict]) -> list[dict]:
    """``open_thread`` for the top two; ``draft_reply`` for the top waiting on the owner;
    ``set_follow_up`` for the top waiting on someone else (§12.3)."""
    by_name = {s["name"]: [i for i in s["items"] if not has_blocking_warning(i)] for s in sections}
    on_me, on_them = by_name.get("waiting_on_me", []), by_name.get("waiting_on_them", [])
    offers = [
        suggestion("open_thread", "outlook_list_thread", {"conversation_id": i["conversation_id"]})
        for i in (on_me + on_them)[:2]
        if i.get("conversation_id")
    ]
    # A reply is offered only where the sender check ran and did not fail (review C1).
    checked = [
        i
        for i in on_me
        if i["sender_auth"]["result"] in ("verified", "unverified") and i.get("last_message_id")
    ]
    if checked:
        offers.append(
            suggestion(
                "draft_reply",
                "outlook_compose",
                {"mode": "reply", "message_id": checked[0]["last_message_id"]},
                ["body"],
            )
        )
    if on_them and on_them[0].get("last_message_id"):
        offers.append(
            suggestion(
                "set_follow_up",
                "outlook_follow_up",
                {"message_id": on_them[0]["last_message_id"]},
                ["by"],
            )
        )
    return offers


def _candidates(
    people, words: str, moment: datetime, spent: RequestBudget, loaded_tools, lookup: Lookup
) -> dict:
    """An ambiguous person: the candidates, and nothing gathered (§12.3, §2.2). A search
    that could not finish answers ``incomplete``, never an empty ``candidates``."""
    items = [c.item for c in people.candidates]
    offers = [
        suggestion("choose", TOOL, {"person": c.item["address"]})
        for c in people.candidates
        if c.known and not c.blocking
    ]
    understood = {
        "person": sanitize_output(words),
        "matched_kind": "person",
        "queries": list(lookup.graph.queries),
    }
    if people.outcome == "incomplete":
        understood["not_searched"] = list(people.not_searched)
    return build_envelope(
        tool=TOOL,
        outcome="incomplete" if people.outcome == "incomplete" else "candidates",
        as_of=moment,
        understood=understood,
        sections=[section("candidates", items, cap=CAPS["candidates"], total=people.total)],
        failed_sections=[],
        graph_requests=spent.spent,
        next=build_next(offers, loaded_tools),
    )
