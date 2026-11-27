"""The follow-up marker, read side (contracts §8.2): flagged mail, and whether a reply came.

A follow-up is a flag on one message, with a due date. ``catch_up`` and
``open_loops`` read them the same way:

- **F1**, flagged mail: one page of the newest 100 flags across all mail. The
  received-time floor comes first so Graph takes the ``$orderby``. P-D3a probes
  this exact shape; if Graph refuses it, the given folders are scanned instead
  (90 days, one page each) and the flags picked out locally, and the answer says
  its count is not exact.
- **F2**, a reply check per flag: the later messages in its conversation, in one
  ``$batch``. A later letter in Sent Items is the owner's reply. A later message
  settles the follow-up only when it is from someone the flagged message waits on,
  is outside Junk and Deleted Items, and its headers show a verified sender and no
  automatic or bulk mail (the review of the read tools, B1). Anything less leaves
  it open: a bounce, an out-of-office reply, a stranger's message in Junk.

The owner's mail is what is in Sent Items. A later message that only claims the
owner's address counts as neither (the lesson of find's review, B1). A reply
check that failed, or could not be confirmed, reads as not checked, never as "no
reply" or "answered".
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents.envelope import sender_auth_result, utc, why
from outlook_mcp.intents.people import (
    is_meeting_type,
    parse_time,
    recipients,
    sender_of,
)
from outlook_mcp.intents.reading import header_reads
from outlook_mcp.intents.resolver import Lookup, _graph_time, _literal
from outlook_mcp.intents.transport import build_url
from outlook_mcp.tools.mail_read import RECEIVED_FLOOR

F1_FILTER = f"receivedDateTime ge {RECEIVED_FLOOR} and flag/flagStatus eq 'flagged'"
# §8.2's list, plus the recipients and subject: open_loops says whom a flagged letter
# waits on, and the subject is shown (under `others_wrote`) and backs up P-SI1.
F1_SELECT = (
    "id,conversationId,receivedDateTime,sentDateTime,parentFolderId,from,flag,"
    "toRecipients,ccRecipients,subject"
)
F1_TOP = 100
# `subject` beyond §8.2's list: the meeting-message check reads it (P-SI1).
F2_SELECT = "id,from,receivedDateTime,parentFolderId,subject"
F2_TOP = 10
FALLBACK_DAYS = 90


def is_flagged(row: dict) -> bool:
    flag = row.get("flag") if isinstance(row.get("flag"), dict) else {}
    return flag.get("flagStatus") == "flagged"


def flag_due(row: dict) -> datetime | None:
    flag = row.get("flag") if isinstance(row.get("flag"), dict) else {}
    return _graph_time(flag.get("dueDateTime"))


def flag_reasons(row: dict, now: datetime) -> list[dict]:
    """``flag_overdue`` or ``flag_due``, with the due time, for a flagged row."""
    due = flag_due(row)
    if not is_flagged(row) or due is None:
        return []
    return [why("flag_overdue" if due < now else "flag_due", due=utc(due))]


@dataclass
class Flags:
    rows: list[dict]
    # False when the P-D3a fallback scanned folders instead of asking Graph to filter.
    exact: bool = True
    # True when more flags exist than the one page read.
    beyond_first_page: bool = False


async def flagged_mail(lookup: Lookup, *, fallback_folders: tuple[str, ...]) -> Flags:
    """F1: the newest flagged messages (§8.2), or the P-D3a fallback scan."""
    try:
        body = await lookup.graph.get(
            "/me/messages",
            {
                "$filter": F1_FILTER,
                "$orderby": "receivedDateTime desc",
                "$select": F1_SELECT,
                "$top": F1_TOP,
            },
        )
    except GraphAPIError as error:
        # P-D3a: whether a personal mailbox takes the flag filter with the order-by.
        if error.status_code != 400:
            raise
        rows: list[dict] = []
        since = utc(lookup.now - timedelta(days=FALLBACK_DAYS))
        for folder in fallback_folders:
            try:
                scanned = await lookup.graph.get(
                    f"/me/mailFolders/{folder}/messages",
                    {
                        "$filter": f"receivedDateTime ge {since}",
                        "$orderby": "receivedDateTime desc",
                        "$select": F1_SELECT,
                        "$top": F1_TOP,
                    },
                )
            except GraphAPIError as missing:
                if missing.status_code == 404 and folder == "archive":
                    continue
                raise
            rows += [r for r in scanned.get("value") or [] if is_flagged(r)]
        return Flags(rows=rows, exact=False)
    return Flags(
        rows=[r for r in body.get("value") or [] if is_flagged(r)],
        beyond_first_page=bool(body.get("@odata.nextLink")),
    )


@dataclass
class ReplyCheck:
    """What F2 found after one flagged message.

    ``reply`` is set only for a reply that settles the follow-up: from someone the
    flagged message was waiting on, not in Junk or Deleted Items, whose sender check
    passed, and that is not automatic or bulk mail (review B1). ``unconfirmed`` is set
    when a later message from them was found but could not be confirmed: its check
    failed or did not run. Neither settles anything; the follow-up stays.
    """

    checked: bool
    you_replied: datetime | None = None
    reply: dict | None = None
    reply_from: str | None = None
    unconfirmed: bool = False


# Well-known folders whose mail never settles a follow-up.
_NOT_REPLIES = ("junkemail", "deleteditems")


def _waited_on(row: dict, mine: bool, me: frozenset[str]) -> set[str]:
    """Whom a flagged message waits on: a letter's recipients, or a message's sender."""
    if mine:
        found = {a for a, _ in recipients(row, ("toRecipients", "ccRecipients", "bccRecipients"))}
    else:
        found = {sender_of(row)[0] or ""}
    return {a for a in found if a and a not in me}


def reply_urls(flags: list[tuple[dict, datetime]]) -> list[str]:
    """F2's query for each ``(flagged row, its time)``: later mail in its conversation."""
    return [
        build_url(
            "/me/messages",
            {
                "$filter": (
                    f"receivedDateTime gt {utc(moment)} and "
                    f"conversationId eq '{_literal(str(row.get('conversationId')))}'"
                ),
                "$orderby": "receivedDateTime asc",
                "$select": F2_SELECT,
                "$top": F2_TOP,
            },
        )
        for row, moment in flags
    ]


async def reply_checks(
    lookup: Lookup,
    flags: list[tuple[dict, datetime]],
    *,
    sent_folder: str | None,
    me: frozenset[str],
) -> dict[str, ReplyCheck]:
    """F2 for each ``(flagged row, the time it was sent or received)``, by message id.

    A later message counts as their reply only when it is from one of the people the
    flagged message waits on, outside Junk and Deleted Items, and its headers say
    Microsoft verified the sender and it is not automatic or bulk mail. Anything
    short of that leaves the follow-up open (review B1). A meeting message is told by
    its type only (review B2).
    """
    wanted = [(row, moment) for row, moment in flags if isinstance(row.get("conversationId"), str)]
    results: dict[str, ReplyCheck] = {
        str(row.get("id")): ReplyCheck(checked=False) for row, _ in flags
    }
    if not wanted:
        return results
    replies = await lookup.graph.batch(reply_urls(wanted))
    candidates: dict[str, list[dict]] = {}
    for (row, _), reply in zip(wanted, replies, strict=True):
        if not 200 <= (reply.get("status") or 0) < 300:
            continue  # not checked: never read as "no reply"
        check = ReplyCheck(checked=True)
        mine = sent_folder is not None and row.get("parentFolderId") == sent_folder
        waited_on = _waited_on(row, mine, me)
        for later in (reply.get("body") or {}).get("value") or []:
            if later.get("id") == row.get("id") or is_meeting_type(later):
                continue
            at = parse_time(later.get("receivedDateTime"))
            if sent_folder is not None and later.get("parentFolderId") == sent_folder:
                if at is not None and (check.you_replied is None or at > check.you_replied):
                    check.you_replied = at
                continue
            if sender_of(later)[0] in waited_on and isinstance(later.get("id"), str):
                candidates.setdefault(str(row.get("id")), []).append(later)
        results[str(row.get("id"))] = check
    if candidates:
        await _confirm(lookup, candidates, results)
    return results


async def _confirm(
    lookup: Lookup, candidates: dict[str, list[dict]], results: dict[str, ReplyCheck]
) -> None:
    """Confirm the first candidate reply of each flag: where it is, and its headers."""
    folders = await lookup.graph.batch(
        [f"/me/mailFolders/{name}?$select=id" for name in _NOT_REPLIES], ids=False
    )
    excluded: set[str] = set()
    known_folders = True
    for reply in folders:
        folder_id = (reply.get("body") or {}).get("id")
        if 200 <= (reply.get("status") or 0) < 300 and isinstance(folder_id, str):
            excluded.add(folder_id)
        else:
            known_folders = False  # can't tell Junk apart: confirm nothing
    chosen: dict[str, dict] = {}
    for flag_id, rows in candidates.items():
        usable = [r for r in rows if r.get("parentFolderId") not in excluded]
        if usable and known_folders:
            chosen[flag_id] = usable[0]
        elif usable or not known_folders:
            results[flag_id].unconfirmed = True
    if not chosen:
        return
    reads = await header_reads(lookup, [r["id"] for r in chosen.values()])
    for flag_id, row in chosen.items():
        read = reads.get(row["id"])
        if read is None or read.bulk or read.auto:
            results[flag_id].unconfirmed = True
            continue
        if sender_auth_result(read.verdict)["result"] != "verified":
            results[flag_id].unconfirmed = True
            continue
        results[flag_id].reply = row
        results[flag_id].reply_from = sender_of(row)[0]
