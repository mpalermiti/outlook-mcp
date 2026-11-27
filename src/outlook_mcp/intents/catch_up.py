"""``outlook_catch_up``: what arrived and what needs you (contracts §12.2).

The first call reads a look-back window (``24h`` unless given). Every answer
carries a ``cursor``; the next call passes it back, replacing the old one, and
hears only what arrived since (§7). Arrivals are new mail in the Inbox (§7.3:
created, received after the last call, not from the owner, not a meeting
message), each in the first section it qualifies for:

1. ``follow_ups``: the reply on a flagged conversation, shown on the flag's item;
2. ``replies_to_you``: a conversation the owner wrote in before it arrived;
3. ``needs_reply``: to the owner on the To line, not Focused Inbox's "other", not
   answered since, not mailing-list or machine mail;
4. ``other_mail``: everything else, counted only, with the top senders.

The state sections are read fresh each call: flags due by the look-ahead
(``follow_ups``), open invitations, today's meetings, To Do tasks due and, with a
cursor, what changed on the primary calendar (``calendar_changes``).

Requests (§12.2): the sent table (G1, 2), the Inbox delta (G2, 4), header reads
(G3, 10), flags and their reply checks (G4, G5, 9), the calendar delta (G6, 2),
the calendar view (G7, 2) and To Do (G8, 7): 36, under a ceiling of 40. A
fallback's extra request is taken from its own step where it can be.

Blind spots, every call: mail outside the Inbox, Junk, and secondary calendars.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import budget
from outlook_mcp.intents import cursor as cursors
from outlook_mcp.intents.budget import RequestBudget, run_sections
from outlook_mcp.intents.cursor import CalState, CursorState, MailState
from outlook_mcp.intents.envelope import (
    build_envelope,
    build_next,
    has_blocking_warning,
    section,
    suggestion,
    utc,
    warn,
    why,
)
from outlook_mcp.intents.flags import flag_due, flag_reasons, flagged_mail, reply_checks
from outlook_mcp.intents.people import (
    SentTable,
    is_meeting_type,
    is_valid_address,
    parse_time,
    recipients,
    sender_of,
)
from outlook_mcp.intents.reading import (
    Once,
    event_item,
    event_span,
    header_reads,
    is_open_invite,
    local_midnight,
    message_item,
    open_tasks,
    outcome_of,
    read_view,
    response_of,
    task_item,
    task_order,
    window_echo,
    zone_info,
)
from outlook_mcp.intents.resolver import Lookup, open_lookup
from outlook_mcp.tools._delta import require_graph_url
from outlook_mcp.tools.digest import is_stale_delta
from outlook_mcp.validation import resolve_window

TOOL = "outlook_catch_up"
FOCUSES = ("all", "mail", "calendar", "tasks")
MAX_DAYS = 31
LOOK_AHEAD_MAX = timedelta(days=7)
OVERLAP = timedelta(minutes=2)  # delivery that lands just before the last call (§7.3)
RESYNC_MAX = timedelta(days=31)
SENT_PAGES = 2  # G1
MAIL_PAGES = 4  # G2, attempts included
MAIL_TOP = 50
# Graph ignores `$top` on message delta and pages by 10 (live, step-6 eval), so the page
# size goes in `Prefer`, combined with the id preference (ids.combine_prefer).
MAIL_PREFER = f"odata.maxpagesize={MAIL_TOP}"
HEADER_READS = 8  # G3 (10 in §12.2)
FLAG_CHECKS = 5  # G5 (8 in §12.2): the confirmations of replies take the rest (review B1)
CAL_PAGES = 2  # G6, attempts included
CAL_PAGE_SIZE = 100
VIEW_PAGES = 2  # G7
TASK_LISTS = 6  # G8
CAL_BACK = timedelta(days=1)
CAL_AHEAD = timedelta(days=30)
ROLL_WHEN = timedelta(days=14)
TOP_SENDERS = 5
CAPS = {
    "needs_reply": 10,
    "replies_to_you": 10,
    "follow_ups": 10,
    "invites": 10,
    "meetings_today": 15,
    "tasks_due": 15,
    "calendar_changes": 10,
    "other_mail": 0,
}
MAIL_SECTIONS = ("needs_reply", "replies_to_you", "follow_ups", "other_mail")
CALENDAR_SECTIONS = ("invites", "meetings_today", "calendar_changes")
NOT_COVERED = ["mail_outside_inbox", "junk", "secondary_calendars"]
MAIL_DELTA_PATH = "/me/mailFolders/inbox/messages/delta"
# G2's `$select` (§12.2).
MAIL_SELECT = (
    "id,conversationId,from,toRecipients,ccRecipients,receivedDateTime,isRead,importance,"
    "inferenceClassification,flag,hasAttachments,subject,bodyPreview"
)
CAL_DELTA_PATH = "/me/calendarView/delta"
# G7's `$select` (§12.2).
VIEW_SELECT = (
    "id,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,isCancelled,"
    "showAs,attendees"
)


@dataclass
class _MailRead:
    """What G2 found, and the Inbox state for the next cursor."""

    arrivals: list[dict]
    state: MailState
    rule: str  # "created" or "received_time"
    since: datetime
    has_more: bool
    resynced: bool = False
    gap: dict | None = None
    from_you: int = 0
    repeat_before: datetime | None = None  # arrivals up to here may repeat (§7.3 fallback)


async def read_catch_up(
    graph_client: Any,
    *,
    cursor: str | None = None,
    window: str | None = None,
    focus: str = "all",
    zone: str = "UTC",
    signed_in: str | None = None,
    loaded_tools: frozenset[str] = frozenset(),
    now: datetime | None = None,
) -> dict:
    """What arrived and what needs the owner (§12.2), in the envelope, with a cursor.

    ``zone`` (the config's time zone) reads the window and decides "today";
    ``signed_in`` is the signed-in username, one of the owner's addresses and the
    cursor's account; ``loaded_tools`` is what ``next`` may name. ``now`` is fixed
    by tests.
    """
    moment = now or datetime.now(timezone.utc)

    # Every argument is checked before any request goes out.
    if focus not in FOCUSES:
        raise ValueError(f"focus must be one of: {', '.join(FOCUSES)}.")
    tz = zone_info(zone)
    state, cursor_status = cursors.read(cursor, username=signed_in, now=moment)
    window_value = window if window is not None else ("24h" if state is None else None)
    span = None
    if window_value is not None:
        span = resolve_window(window_value, zone, now=moment, max_days=MAX_DAYS)
        # With a cursor the window sets only the look-ahead, so it may lie ahead (§12.2).
        if span[0] >= moment and state is None:
            raise ValueError(
                "catch_up reads what arrived, so the window must start in the past: 24h is "
                "the last day, while +24h would be the next one."
            )
    end_of_today = local_midnight(moment.astimezone(tz).date() + timedelta(days=1), tz)
    horizon = min(max(end_of_today, span[1] if span else end_of_today), moment + LOOK_AHEAD_MAX)

    wants_mail = focus in ("all", "mail")
    wants_calendar = focus in ("all", "calendar")
    wants_tasks = focus in ("all", "tasks")
    spent = RequestBudget(budget.CEILINGS[TOOL])

    async with open_lookup(
        graph_client, spent, now=moment, zone=zone, signed_in=signed_in
    ) as lookup:
        sent_table = Once(lambda: lookup.sent_table(pages=SENT_PAGES))

        async def mail() -> dict:
            table: SentTable = await sent_table.get()
            read = await _read_mail(lookup, state, span, table)
            return await _mail_sections(lookup, read, table, horizon)

        async def calendar() -> dict:
            table = await sent_table.get()
            return await _calendar_sections(lookup, state, table, tz, zone, end_of_today)

        async def tasks() -> list[dict]:
            pairs, missed = await open_tasks(lookup, lists_max=TASK_LISTS)
            today = moment.astimezone(tz).date()
            last = (horizon - timedelta(microseconds=1)).astimezone(tz).date()
            items = []
            for row, list_id in pairs:
                due = _due(row)
                if due is not None and due <= last:
                    items.append(task_item(row, list_id, today=today, in_window=True))
            return {"items": sorted(items, key=task_order), "not_covered": missed}

        runs = []
        if wants_mail:
            runs.append(("mail", mail))
        if wants_calendar:
            runs.append(("calendar", calendar))
        if wants_tasks:
            runs.append(("tasks", tasks))
        results = {r.name: r for r in await run_sections(runs)}

    sections: list[dict] = []
    failed: list[dict] = []
    understood: dict = {"not_covered": list(NOT_COVERED)}
    offers: list[dict] = []
    mail_state = None
    cal_state = None

    if wants_mail:
        result = results["mail"]
        if result.status == "ok":
            found = result.value
            sections += found["sections"]
            understood.update(found["understood"])
            understood["not_covered"] += found.get("not_covered", [])
            mail_state = found["state"]
            offers += found["offers"]
        else:
            for name in MAIL_SECTIONS:
                sections.append(section(name, [], cap=CAPS[name], status=result.status))
                failed.append({**result.failure, "section": name})
    if wants_calendar:
        result = results["calendar"]
        names = [n for n in CALENDAR_SECTIONS if n != "calendar_changes" or state is not None]
        if result.status == "ok":
            found = result.value
            sections += [s for s in found["sections"] if s["name"] in names]
            understood.update(found["understood"])
            understood["not_covered"] += found["not_covered"]
            cal_state = found["state"]
            offers += found["offers"]
        else:
            for name in names:
                sections.append(section(name, [], cap=CAPS[name], status=result.status))
                failed.append({**result.failure, "section": name})
    if wants_tasks:
        result = results["tasks"]
        if result.status == "ok":
            sections.append(section("tasks_due", result.value["items"], cap=CAPS["tasks_due"]))
            understood["not_covered"] += result.value["not_covered"]
        else:
            sections.append(section("tasks_due", [], cap=CAPS["tasks_due"], status=result.status))
            failed.append(result.failure)
    sections.sort(key=lambda s: list(CAPS).index(s["name"]))

    new_state = CursorState(
        acct=cursors.account_tag(signed_in),
        made=utc(moment),
        mail=mail_state or _kept_mail(state, span, moment),
        cal=cal_state or _kept_calendar(state, moment),
    )
    token = cursors.encode(new_state)
    understood.update(
        {
            "window": window_echo(window_value, span[0], span[1], zone) if span else None,
            "horizon": utc(horizon),
            "cursor": cursor_status if token is not None else cursors.TOO_LARGE,
            "focus": focus,
        }
    )
    outcome = outcome_of(sections)
    if understood.pop("_incomplete", False):
        # Arrivals left unread, a flag scan that could not reach back, or a calendar
        # view cut before today ended: the answer is not the whole of it (review D2, D4).
        outcome = "incomplete"
    env = build_envelope(
        tool=TOOL,
        outcome=outcome,
        as_of=moment,
        understood=understood,
        sections=sections,
        failed_sections=failed,
        graph_requests=spent.spent,
        next=build_next(_order_offers(offers), loaded_tools),
        capped=len(understood["not_covered"]) > len(NOT_COVERED)
        or bool(understood.get("flags_beyond_first_page"))
        or bool(understood.get("replies_unconfirmed"))
        or bool((understood.get("arrivals") or {}).get("has_more")),
    )
    env["cursor"] = token
    return env


def _due(row: dict):
    from outlook_mcp.intents.reading import due_date

    return due_date(row)


def _kept_mail(state: CursorState | None, span, moment: datetime) -> MailState:
    """The Inbox state when this call did not read mail: the old link, with the arrival
    threshold kept, so the next call reports everything since."""
    if state is None:
        since = span[0] if span else moment - timedelta(hours=24)
        return MailState(link="", created_only=True, bootstrap_from=utc(since))
    threshold = state.mail.bootstrap_from or utc(parse_time(state.made) - OVERLAP)
    return MailState(
        link=state.mail.link, created_only=state.mail.created_only, bootstrap_from=threshold
    )


def _kept_calendar(state: CursorState | None, moment: datetime) -> CalState:
    """The calendar state when this call did not read it (focus, a failure, the time
    limit): the old link, and the time changes are owed from, so a later snapshot (a
    resync, a roll) still reports what changed in between (review D1)."""
    if state is None:
        return CalState(
            link="", start=utc(moment - CAL_BACK), end=utc(moment + CAL_AHEAD), since=utc(moment)
        )
    return CalState(
        link=state.cal.link,
        start=state.cal.start,
        end=state.cal.end,
        since=state.cal.since or state.made,
    )


# ── Mail: the Inbox delta (G2) ────────────────────────────────────────────


async def _delta_pages(
    lookup: Lookup, first: str, *, resource: str, pages: int, prefer: str | None = None
) -> tuple[list[dict], str | None, bool]:
    """Follow a delta round from ``first`` for at most ``pages`` requests: the items,
    the link to keep (the deltaLink, or the nextLink when unfinished) and whether the
    round is unfinished. ``first`` must already be checked."""
    items: list[dict] = []
    url = first
    for _ in range(pages):
        body = await lookup.graph.follow(url, prefer=prefer)
        items += body.get("value") or []
        delta = body.get("@odata.deltaLink")
        if isinstance(delta, str) and delta:
            return (
                items,
                require_graph_url(delta, source="@odata.deltaLink", resource=resource),
                False,
            )
        link = body.get("@odata.nextLink")
        if not isinstance(link, str) or not link:
            return items, None, False
        url = require_graph_url(link, source="@odata.nextLink", resource=resource)
    return items, url, True


def _mail_bootstrap_url(since: datetime, *, created_only: bool) -> str:
    from outlook_mcp.intents.transport import GRAPH_BASE, build_url

    params: dict = {}
    if created_only:
        params["changeType"] = "created"
    params.update(
        {"$filter": f"receivedDateTime ge {utc(since)}", "$select": MAIL_SELECT, "$top": MAIL_TOP}
    )
    return GRAPH_BASE + build_url(MAIL_DELTA_PATH, params)


async def _bootstrap_mail(
    lookup: Lookup, since: datetime, pages: int
) -> tuple[list[dict], str | None, bool, bool, int]:
    """A new created-only Inbox round from ``since`` (§7.3), or, when Graph refuses
    ``changeType`` (P-DL1), an ordinary one. Also whether it is created-only, and the
    requests it took."""
    try:
        items, link, more = await _delta_pages(
            lookup,
            _mail_bootstrap_url(since, created_only=True),
            resource="mail",
            pages=pages,
            prefer=MAIL_PREFER,
        )
        return items, link, more, True, pages
    except GraphAPIError as error:
        # P-DL1: whether a personal mailbox takes changeType=created on message delta.
        if error.status_code != 400 or pages < 2:
            raise
    items, link, more = await _delta_pages(
        lookup,
        _mail_bootstrap_url(since, created_only=False),
        resource="mail",
        pages=pages - 1,
        prefer=MAIL_PREFER,
    )
    return items, link, more, False, pages


async def _read_mail(
    lookup: Lookup, state: CursorState | None, span, table: SentTable
) -> _MailRead:
    """G2: the Inbox's new mail since the window's start or the last call (§7.3, §7.4)."""
    moment = lookup.now
    resynced, gap = False, None
    if state is None:
        threshold = span[0]
    else:
        threshold = (
            parse_time(state.mail.bootstrap_from)
            if state.mail.bootstrap_from
            else (parse_time(state.made) - OVERLAP)
        )
    repeat_before = None
    if state is not None and state.mail.link:
        created_only = state.mail.created_only
        try:
            items, link, more = await _delta_pages(
                lookup, state.mail.link, resource="mail", pages=MAIL_PAGES, prefer=MAIL_PREFER
            )
        except GraphAPIError as error:
            if not is_stale_delta(error.status_code):
                raise
            # A stale link: this source starts again from the last call, never from
            # the default window (§7.4), and says so.
            resynced = True
            start = threshold
            if moment - start > RESYNC_MAX:
                start = moment - RESYNC_MAX
                gap = {"from": utc(threshold), "to": utc(start)}
                threshold = start
            items, link, more, created_only, _ = await _bootstrap_mail(
                lookup, start, MAIL_PAGES - 1
            )
        if not created_only:
            repeat_before = parse_time(state.made)
    else:
        start = threshold
        if moment - start > RESYNC_MAX:
            start = moment - RESYNC_MAX
            gap = {"from": utc(threshold), "to": utc(start)}
            threshold = start
        items, link, more, created_only, _ = await _bootstrap_mail(lookup, start, MAIL_PAGES)
        if state is not None and not created_only:
            repeat_before = parse_time(state.made)

    arrivals, from_you = [], 0
    for row in items:
        if "@removed" in row or not isinstance(row.get("id"), str):
            continue
        received = parse_time(row.get("receivedDateTime"))
        if received is None or received <= threshold:
            continue  # not an arrival: read, flagged, moved in, or older (§7.3)
        if is_meeting_type(row):
            continue  # meeting requests are reported through invites (by type: review B2)
        if sender_of(row)[0] in table.me:
            from_you += 1
            continue
        arrivals.append(row)
    return _MailRead(
        arrivals=arrivals,
        state=MailState(
            link=link or (state.mail.link if state else ""),
            created_only=created_only,
            bootstrap_from=utc(threshold) if more else None,
        ),
        rule="created" if created_only else "received_time",
        since=threshold,
        has_more=more,
        resynced=resynced,
        gap=gap,
        from_you=from_you,
        repeat_before=repeat_before,
    )


# ── Mail sections ─────────────────────────────────────────────────────────


def _only_recipient(row: dict, me: frozenset[str]) -> bool:
    to = [a for a, _ in recipients(row, ("toRecipients",))]
    cc = [a for a, _ in recipients(row, ("ccRecipients",))]
    return len(to) == 1 and to[0] in me and not cc


async def _mail_sections(
    lookup: Lookup, read: _MailRead, table: SentTable, horizon: datetime
) -> dict:
    """needs_reply, replies_to_you, follow_ups and other_mail, from the arrivals and the
    flags (§12.2, §8.2)."""
    me = table.me
    now = lookup.now

    # G4, G5: flags due by the look-ahead, or overdue, and replies to the earliest due.
    flags = await flagged_mail(lookup, fallback_folders=("inbox", "sentitems", "archive"))
    due_flags = []
    for row in flags.rows:
        due = flag_due(row)
        if due is not None and due <= horizon and isinstance(row.get("id"), str):
            due_flags.append((due, row))
    due_flags.sort(key=lambda pair: (pair[0], str(pair[1].get("id"))))
    checked = due_flags[:FLAG_CHECKS]
    sent_folder = table.sent_folder

    def flag_time(row: dict) -> datetime:
        mine = sent_folder is not None and row.get("parentFolderId") == sent_folder
        moment = parse_time(row.get("sentDateTime" if mine else "receivedDateTime"))
        return moment or now

    checks = await reply_checks(
        lookup, [(row, flag_time(row)) for _, row in checked], sent_folder=sent_folder, me=me
    )
    answered: set[str] = set()  # arrivals shown on a flag's item, and nowhere else
    clearable: set[str] = set()  # flags a known person's checked reply settled
    unconfirmed: list[str] = []
    follow_rows: list[tuple[dict, list[dict]]] = []
    for due, row in due_flags:
        reasons = flag_reasons(row, now)
        check = checks.get(str(row.get("id")))
        if check is not None and check.checked:
            if check.reply is not None:
                # A checked reply from whom the flag waits on (review B1): who, and
                # whether the owner writes to them, as §8.2 asks.
                at = parse_time(check.reply.get("receivedDateTime"))
                address = check.reply_from
                known = bool(address and address not in me and table.get(address) is not None)
                params: dict = {"at": utc(at)} if at else {}
                if address and is_valid_address(address):
                    params["address"] = address
                params["known"] = known
                reasons.append(why("reply_from_them", **params))
                answered.add(str(check.reply.get("id")))
                if known:
                    clearable.add(str(row.get("id")))
            elif check.you_replied is not None:
                reasons.append(why("you_replied", at=utc(check.you_replied)))
            if check.unconfirmed:
                unconfirmed.append(str(row.get("id")))
        follow_rows.append((row, reasons))

    # Each arrival goes into the first section it qualifies for.
    letters = table.letters
    replies, needs, other = [], [], []
    for row in read.arrivals:
        if str(row.get("id")) in answered:
            continue
        received = parse_time(row.get("receivedDateTime"))
        mine_times = letters.get(str(row.get("conversationId")), [])
        before = [t for t in mine_times if received and t < received]
        after = [t for t in mine_times if received and t > received]
        if before:
            replies.append((row, max(before)))
        elif (
            any(a in me for a, _ in recipients(row, ("toRecipients",)))
            and row.get("inferenceClassification") != "other"
            and not after
        ):
            needs.append(row)
        else:
            other.append(row)

    replies.sort(key=lambda pair: pair[0].get("receivedDateTime") or "", reverse=True)

    def needs_rank(row: dict) -> tuple:
        entry = table.get(sender_of(row)[0] or "")
        received = parse_time(row.get("receivedDateTime"))
        return (
            -(entry.count if entry else 0),
            0 if _only_recipient(row, me) else 1,
            0 if row.get("importance") == "high" else 1,
            -(received.timestamp() if received else 0),
            str(row.get("id")),
        )

    needs.sort(key=needs_rank)
    # G3: header reads for the items shown first, replies then needs_reply, 10 in all.
    shown = [row for row, _ in replies[: CAPS["replies_to_you"]]] + needs[: CAPS["needs_reply"]]
    reads = await header_reads(lookup, [r["id"] for r in shown[:HEADER_READS]])

    def bulk(row: dict) -> bool:
        found = reads.get(row["id"])
        return found is not None and (found.bulk or found.auto)

    def marks(row: dict) -> list[dict]:
        received = parse_time(row.get("receivedDateTime"))
        if read.repeat_before is not None and received is not None:
            if received <= read.repeat_before:
                return [warn("may_repeat")]
        return []

    reply_items = []
    for row, at in replies:
        if bulk(row):
            other.append(row)
            continue
        item = message_item(row, table=table, read=reads.get(row["id"]), folder="Inbox")
        item["why"] = [why("you_wrote_in_thread", at=utc(at)), *item["why"]]
        item["warnings"] += marks(row)
        reply_items.append(item)
    need_items = []
    for row in needs:
        if bulk(row):
            other.append(row)
            continue
        item = message_item(row, table=table, read=reads.get(row["id"]), folder="Inbox")
        reasons = [why("only_recipient" if _only_recipient(row, me) else "you_are_on_to")]
        entry = table.get(sender_of(row)[0] or "")
        if entry is not None:
            params = {"count": entry.count}
            if entry.last is not None:
                params["last"] = utc(entry.last)
            reasons.append(why("sender_written_to", **params))
        if row.get("importance") == "high":
            reasons.append(why("high_importance"))
        item["why"] = reasons
        item["warnings"] += marks(row)
        need_items.append(item)

    follow_items = []
    for row, reasons in follow_rows:
        mine = sent_folder is not None and row.get("parentFolderId") == sent_folder
        item = message_item(row, table=table, read=None, folder=None, owner=mine)
        item["why"] = reasons
        follow_items.append(item)

    senders = Counter(
        address
        for address in (sender_of(r)[0] for r in other)
        if address and is_valid_address(address)
    )
    by_sender = [{"address": a, "count": n} for a, n in senders.most_common(TOP_SENDERS)]

    # A reply is offered only where the sender check ran and did not fail (review C1).
    offers = [
        suggestion(
            "draft_reply", "outlook_compose", {"mode": "reply", "message_id": i["id"]}, ["body"]
        )
        for i in [
            i
            for i in need_items
            if not has_blocking_warning(i)
            and i["from"]["sender_auth"]["result"] in ("verified", "unverified")
        ][:2]
    ]
    # Clearing is offered only on a checked reply from someone the owner writes to:
    # never on a message that merely sits in the conversation (review B1).
    offers += [
        suggestion("clear_follow_up", "outlook_follow_up", {"message_id": i["id"], "clear": True})
        for i in follow_items
        if i["id"] in clearable and not has_blocking_warning(i)
    ]
    tops = [i for i in reply_items if not has_blocking_warning(i)][:1]
    offers += [
        suggestion("open_thread", "outlook_list_thread", {"conversation_id": i["conversation_id"]})
        for i in tops
        if i.get("conversation_id")
    ]

    arrivals = {
        "rule": read.rule,
        "since": utc(read.since),
        "total": len(read.arrivals),
        "has_more": read.has_more,
    }
    if read.from_you:
        arrivals["from_your_address"] = read.from_you
    understood: dict = {
        "arrivals": arrivals,
        "other_mail_total": len(other),
        "other_mail_by_sender": by_sender,
        "flags_beyond_first_page": flags.beyond_first_page,
        "resynced": ["mail"] if read.resynced else [],
    }
    if read.gap:
        understood["gap"] = read.gap
    if len(due_flags) > FLAG_CHECKS:
        understood["flags_not_checked"] = len(due_flags) - FLAG_CHECKS
    if unconfirmed:
        understood["replies_unconfirmed"] = unconfirmed
    not_covered: list[str] = []
    if not flags.exact:
        # P-D3a's fallback ran: flags older than its scan are not seen (review D4).
        not_covered.append("flags_fallback_scan")
        if flags.scan_capped:
            not_covered.append("flags_older_than_scan")
            understood["_incomplete"] = True
    if read.has_more:
        # Arrivals were left for the next call: no mail count is whole (review D2).
        understood["_incomplete"] = True
    mail_exact = not read.has_more
    return {
        "sections": [
            section("needs_reply", need_items, cap=CAPS["needs_reply"], total_exact=mail_exact),
            section(
                "replies_to_you", reply_items, cap=CAPS["replies_to_you"], total_exact=mail_exact
            ),
            section(
                "follow_ups",
                [i for i in follow_items],
                cap=CAPS["follow_ups"],
                total_exact=flags.exact and not flags.beyond_first_page,
            ),
            section("other_mail", [], cap=0, total=len(other), total_exact=mail_exact),
        ],
        "not_covered": not_covered,
        "understood": understood,
        "state": read.state,
        "offers": offers,
    }


# ── Calendar: the delta (G6) and the view (G7) ────────────────────────────


def _cal_bootstrap_url(start: datetime, end: datetime) -> str:
    from outlook_mcp.intents.transport import GRAPH_BASE, build_url

    return GRAPH_BASE + build_url(
        CAL_DELTA_PATH, {"startDateTime": utc(start), "endDateTime": utc(end)}
    )


async def _calendar_sections(
    lookup: Lookup, state: CursorState | None, table: SentTable, tz, zone: str, end_of_today
) -> dict:
    """invites, meetings_today and calendar_changes (§12.2, §7.4)."""
    now = lookup.now
    prefer = f"odata.maxpagesize={CAL_PAGE_SIZE}"
    understood: dict = {}
    not_covered: list[str] = []
    since: datetime | None = None
    rolled = resynced = False
    start, end = now - CAL_BACK, now + CAL_AHEAD

    if state is not None and state.cal.link:
        cal_end = parse_time(state.cal.end)
        if cal_end is not None and cal_end - now < ROLL_WHEN:
            # Fewer than 14 days of look-ahead left: a new snapshot, kept only for what
            # changed since the last call. Deletions in between can't be told (§7.4).
            rolled, since = True, _owed_since(state)
            not_covered.append("deletions_at_roll")
            items, link, more = await _delta_pages(
                lookup,
                _cal_bootstrap_url(start, end),
                resource="calendar",
                pages=CAL_PAGES,
                prefer=prefer,
            )
        else:
            start, end = parse_time(state.cal.start) or start, cal_end or end
            since = parse_time(state.cal.since) if state.cal.since else None
            try:
                items, link, more = await _delta_pages(
                    lookup, state.cal.link, resource="calendar", pages=CAL_PAGES, prefer=prefer
                )
            except GraphAPIError as error:
                if not is_stale_delta(error.status_code):
                    raise
                resynced, since = True, _owed_since(state)
                start, end = now - CAL_BACK, now + CAL_AHEAD
                items, link, more = await _delta_pages(
                    lookup,
                    _cal_bootstrap_url(start, end),
                    resource="calendar",
                    pages=CAL_PAGES - 1,
                    prefer=prefer,
                )
    else:
        # A first read: the snapshot only starts the delta; nothing in it is a change.
        since = _owed_since(state) if state is not None else now
        items, link, more = await _delta_pages(
            lookup,
            _cal_bootstrap_url(start, end),
            resource="calendar",
            pages=CAL_PAGES,
            prefer=prefer,
        )

    changes, removed = [], []
    if state is not None:
        for row in items:
            if "@removed" in row:
                if isinstance(row.get("id"), str):
                    removed.append(row["id"])
                continue
            if row.get("type") == "seriesMaster":
                # The delta returns each series master beside its occurrences (P-C2, live),
                # dated at the series' first start: a series' changes are its occurrences.
                continue
            modified = parse_time(row.get("lastModifiedDateTime"))
            if since is not None and (modified is None or modified <= since):
                continue
            begins = event_span(row, tz)[0]
            if begins is None or begins <= now - CAL_BACK:
                continue
            item = event_item(row, calendar=None, tz=tz, zone_name=zone, now=now, table=table)
            if row.get("isCancelled") is not True:
                item["why"] = [why("changed_since_cursor"), *item["why"]]
            changes.append(item)
        changes.sort(key=lambda i: (i["start"] or {}).get("utc") or "")
    if removed:
        understood["calendar_removed"] = {"count": len(removed), "ids": removed[:20]}
    understood["calendar_rolled"] = rolled
    if resynced:
        understood["calendar_resynced"] = True

    # G7: the primary calendar from the start of today to 30 days ahead.
    today_start = local_midnight(now.astimezone(tz).date(), tz)
    rows, view_more = await read_view(
        lookup, "/me/calendarView", today_start, now + CAL_AHEAD, VIEW_SELECT, pages=VIEW_PAGES
    )
    read_until = None
    if view_more:
        not_covered.append("events_beyond_200")
        # In start order (reading.read_view), so the cut is the view's tail (review A1).
        starts = [s for s in (event_span(r, tz)[0] for r in rows) if s]
        read_until = max(starts) if starts else today_start
        if read_until < end_of_today:
            understood["_incomplete"] = True
    events = sorted(
        (r for r in rows if isinstance(r.get("id"), str)),
        key=lambda r: (event_span(r, tz)[0] or now, r["id"]),
    )
    invites = [
        event_item(r, calendar=None, tz=tz, zone_name=zone, now=now, table=table)
        for r in events
        if is_open_invite(r) and (event_span(r, tz)[0] or now) > now
    ]
    today = []
    for r in events:
        begins, ends = event_span(r, tz)
        if (
            begins is not None
            and ends is not None
            and begins < end_of_today
            and ends > now
            and r.get("isCancelled") is not True
            and response_of(r) != "declined"
        ):
            today.append(event_item(r, calendar=None, tz=tz, zone_name=zone, now=now, table=table))
    offers = [
        suggestion("answer", "outlook_answer_invites", {"event_ids": [i["id"]]}, ["answer"])
        for i in [i for i in invites if not has_blocking_warning(i)][:1]
    ]
    state_out = CalState(
        link=link or (state.cal.link if state is not None else ""),
        start=utc(start),
        end=utc(end),
        since=utc(since) if (more and since is not None) else None,
    )
    return {
        "sections": [
            section("invites", invites, cap=CAPS["invites"], total_exact=read_until is None),
            section(
                "meetings_today",
                today,
                cap=CAPS["meetings_today"],
                total_exact=read_until is None or read_until >= end_of_today,
            ),
            # An unfinished round has more changes to come (review nit).
            section(
                "calendar_changes", changes, cap=CAPS["calendar_changes"], total_exact=not more
            ),
        ],
        "understood": understood,
        "not_covered": not_covered,
        "state": state_out,
        "offers": offers,
    }


def _owed_since(state: CursorState) -> datetime | None:
    """The time calendar changes are owed from: the last call that read the calendar
    (``cal.since``, carried across calls that did not), else the last call (review D1)."""
    return parse_time(state.cal.since) if state.cal.since else parse_time(state.made)


def _order_offers(offers: list[dict]) -> list[dict]:
    """§12.2's order: draft_reply, clear_follow_up, answer, open_thread."""
    order = ["draft_reply", "clear_follow_up", "answer", "open_thread"]
    return sorted(offers, key=lambda s: order.index(s["template"]))
