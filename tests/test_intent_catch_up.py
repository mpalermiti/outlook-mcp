"""``outlook_catch_up`` (contracts §12.2, §7): what arrived and what needs you.

The first call reads a look-back window; each answer carries a cursor that the
next call passes back. Pinned here: each section's rule as a table, the first
section an arrival qualifies for winning; what is and isn't an arrival (§7.3);
stale cursors resyncing from the last call, with the gap said (§7.4); the
created-only fallback (P-DL1); the calendar delta, its roll and removed events
(P-C2); the flags (P-D3a); ``focus``; the request count; and where other
people's text may appear, the cursor included.
"""

from __future__ import annotations

import base64
import json
import re
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import ForeignCursorError
from outlook_mcp.intents import budget, catch_up, cursor, flags, ids, transport
from outlook_mcp.intents.catch_up import read_catch_up
from outlook_mcp.intents.envelope import WARNING_CODES, WHY_CODES
from outlook_mcp.intents.ids import IdFamily
from tests.intent_fakes import (
    ME,
    SENT_FOLDER,
    FakeGraph,
    Mailbox,
    addr,
    auth_header,
    error,
    event,
    page,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
LATER = datetime(2026, 11, 26, 15, 0, tzinfo=timezone.utc)
MARK = " ZQXMARK "
SAM, DANA, PAT, LEE, KIM, BILL, JO = (
    "sam@example.com",
    "dana@unknown.example",
    "pat@example.org",
    "lee@example.com",
    "kim@example.com",
    "bill@example.net",
    "jo@example.com",
)
STRANGER = "stranger@unknown.example"
MAIL_LINK = "https://graph.microsoft.com/v1.0/me/mailFolders('inbox')/messages/delta"
CAL_LINK = "https://graph.microsoft.com/v1.0/me/calendarView/delta"


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def sent(msg_id, to, when, *, conversation, subject="Letter", flag=None, odata_type=None):
    row = {
        "id": msg_id,
        "subject": subject,
        "bodyPreview": f"preview of {subject}",
        "from": addr(ME, "Me"),
        "toRecipients": [addr(a) for a in to],
        "ccRecipients": [],
        "bccRecipients": [],
        "sentDateTime": when,
        "receivedDateTime": when,
        "conversationId": conversation,
        "parentFolderId": SENT_FOLDER,
        "flag": flag or {"flagStatus": "notFlagged"},
    }
    if odata_type:
        row["@odata.type"] = odata_type
    return row


def flagged(due: str) -> dict:
    return {
        "flagStatus": "flagged",
        "startDateTime": {"dateTime": "2026-11-01T00:00:00.0000000", "timeZone": "UTC"},
        "dueDateTime": {"dateTime": f"{due}T00:00:00.0000000", "timeZone": "UTC"},
    }


class InboxDelta:
    """The Inbox's message delta: a first round of what the filter's bound lets in, then
    whatever ``changes`` holds on each later round."""

    def __init__(self, inbox: list[dict]):
        self.inbox = inbox
        self.changes: list[dict] = []
        self.refuse_created = False
        self.stale = False
        self.boot: list[dict] = []

    def __call__(self, seen, match):
        params = seen.params
        if "$deltatoken" in params:
            if self.stale:
                return error(410, "SyncStateNotFound", "The sync state is not found.")
            body = page(self.changes)
            body["@odata.deltaLink"] = f"{MAIL_LINK}?$deltatoken=round2"
            return body
        if "$skiptoken" in params:
            start = int(params["$skiptoken"])
        else:
            if "changeType" in params and self.refuse_created:
                return error(400, "BadRequest", "Unsupported query option changeType")
            bound = (params.get("$filter") or " ge 0000").split(" ge ")[-1][:19]
            self.boot = [r for r in self.inbox if r["receivedDateTime"][:19] >= bound]
            if "$top" in params:
                # Live (step-6 eval): `$top` caps the whole round, not a page.
                self.boot = self.boot[: int(params["$top"])]
            start = 0
        # As Graph does (live, step-6 eval): pages are 10 unless `Prefer:
        # odata.maxpagesize` says otherwise.
        sized = re.search(r"odata\.maxpagesize=(\d+)", seen.headers.get("prefer", ""))
        top = int(sized.group(1)) if sized else 10
        chunk = self.boot[start : start + top]
        body = page(chunk)
        if start + top < len(self.boot):
            body["@odata.nextLink"] = f"{MAIL_LINK}?$skiptoken={start + top}"
        else:
            body["@odata.deltaLink"] = f"{MAIL_LINK}?$deltatoken=round1"
        return body


class CalendarDelta:
    """The primary calendar's view delta: a snapshot, then ``changes`` each round."""

    def __init__(self, snapshot: list[dict]):
        self.snapshot = snapshot
        self.changes: list[dict] = []
        self.stale = False
        self.boots = 0

    def __call__(self, seen, match):
        if "$deltatoken" in seen.params:
            if self.stale:
                return error(410, "SyncStateNotFound")
            body = page(self.changes)
        else:
            self.boots += 1
            body = page(self.snapshot)
        body["@odata.deltaLink"] = f"{CAL_LINK}?$deltatoken=c{self.boots}"
        return body


def ev(event_id, start, end, *, modified="2026-11-20T00:00:00Z", m="", **kwargs) -> dict:
    row = event(event_id, f"Meeting {event_id}{m}", start, end, **kwargs)
    row["lastModifiedDateTime"] = modified
    return row


def mailbox(m: str = "") -> tuple[Mailbox, InboxDelta, CalendarDelta]:
    """A day of mail since 2026-11-25 12:00. Sam answers a letter (a reply to you); Dana
    and Jo write to you alone (needs a reply; Jo is written to); Pat, Lee, a newsletter
    and Alex's "other" mail are other mail; Kim answers a flagged letter (a follow-up);
    a forged "from me" and an invitation are not arrivals."""
    inbox = [
        received(
            "a1",
            SAM,
            f"Re: Plan{m}",
            name=f"Sam{m}",
            conversation="conv-sam",
            when="2026-11-25T20:00:00Z",
            headers=auth_header("pass", "example.com"),
        ),
        {
            **received(
                "a2",
                DANA,
                f"Question{m}",
                name=f"Dana{m}",
                conversation="conv-dana",
                when="2026-11-26T09:00:00Z",
                headers=auth_header("pass", "unknown.example"),
            ),
            "importance": "high",
        },
        received("a3", PAT, f"Quote{m}", conversation="conv-pat", when="2026-11-26T11:00:00Z"),
        received(
            "a4",
            LEE,
            f"FYI{m}",
            conversation="conv-lee",
            when="2026-11-26T10:00:00Z",
            to=["other@example.com"],
            cc=[ME],
        ),
        received(
            "a5",
            "news@list.example",
            f"Weekly{m}",
            conversation="conv-news",
            when="2026-11-26T07:00:00Z",
            headers=[{"name": "List-Unsubscribe", "value": "<mailto:x@list.example>"}],
        ),
        {
            **received(
                "a6",
                "alex@example.net",
                f"Offer{m}",
                conversation="conv-alex",
                when="2026-11-26T06:00:00Z",
            ),
            "inferenceClassification": "other",
        },
        received("a7", ME, f"Note{m}", conversation="conv-self", when="2026-11-26T05:00:00Z"),
        {
            **received(
                "a8", STRANGER, f"Invite{m}", conversation="conv-inv", when="2026-11-26T04:00:00Z"
            ),
            "@odata.type": "#microsoft.graph.eventMessageRequest",
        },
        received(
            "a9",
            KIM,
            f"Re: Visit{m}",
            conversation="conv-kim",
            when="2026-11-26T08:00:00Z",
            # A checked reply: only that settles a follow-up (review B1).
            headers=auth_header("pass", "example.com"),
        ),
        received(
            "a10",
            JO,
            f"Lunch{m}",
            name=f"Jo{m}",
            conversation="conv-jo-new",
            when="2026-11-26T03:00:00Z",
        ),
        received("old", SAM, f"Older{m}", conversation="conv-old", when="2026-11-20T09:00:00Z"),
        {
            **received(
                "f-bill", BILL, f"Invoice{m}", conversation="conv-bill", when="2026-11-10T09:00:00Z"
            ),
            "flag": flagged("2026-11-20"),
        },
        {
            **received(
                "f-later",
                BILL,
                f"Renewal{m}",
                conversation="conv-renew",
                when="2026-11-10T09:00:00Z",
            ),
            "flag": flagged("2026-12-15"),
        },
    ]
    box = Mailbox(
        sent=[
            sent("s1", [SAM], "2026-11-25T10:00:00Z", conversation="conv-sam"),
            sent("s2", [PAT], "2026-11-26T11:30:00Z", conversation="conv-pat"),
            sent(
                "s3",
                ["org@example.com"],
                "2026-11-25T13:00:00Z",
                conversation="conv-org",
                subject="Accepted: Sync",
                odata_type="#microsoft.graph.eventMessageResponse",
            ),
            sent("s4", [JO], "2026-11-01T10:00:00Z", conversation="conv-jo-old"),
            sent(
                "f-kim",
                [KIM],
                "2026-11-20T10:00:00Z",
                conversation="conv-kim",
                subject=f"Visit{m}",
                flag=flagged("2026-11-27"),
            ),
        ],
        received=inbox,
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                ev("e1", "2026-11-26T14:00:00", "2026-11-26T15:00:00", m=m),
                ev(
                    "e2",
                    "2026-11-27T10:00:00",
                    "2026-11-27T11:00:00",
                    m=m,
                    organizer=STRANGER,
                    organizer_name=f"Stranger{m}",
                    is_organizer=False,
                    response="notResponded",
                ),
                ev(
                    "e3",
                    "2026-11-28T10:00:00",
                    "2026-11-28T11:00:00",
                    m=m,
                    organizer=SAM,
                    organizer_name=f"Sam{m}",
                    is_organizer=False,
                    response="notResponded",
                ),
                ev("e4", "2026-11-26T09:00:00", "2026-11-26T10:00:00", m=m),
                ev(
                    "e5",
                    "2026-11-26T16:00:00",
                    "2026-11-26T17:00:00",
                    m=m,
                    is_organizer=False,
                    response="declined",
                ),
            ]
        },
        task_lists=[
            {"id": "list-1", "displayName": "Tasks", "wellknownListName": "defaultList"},
            {"id": "list-f", "displayName": "Flagged email", "wellknownListName": "flaggedEmails"},
        ],
        tasks={
            "list-1": [
                {
                    "id": "t-today",
                    "title": f"Send{m}",
                    "status": "notStarted",
                    "dueDateTime": {"dateTime": "2026-11-26T00:00:00", "timeZone": "UTC"},
                },
                {
                    "id": "t-late",
                    "title": f"Pay{m}",
                    "status": "notStarted",
                    "dueDateTime": {"dateTime": "2026-11-20T00:00:00", "timeZone": "UTC"},
                },
                {
                    "id": "t-later",
                    "title": f"Book{m}",
                    "status": "notStarted",
                    "dueDateTime": {"dateTime": "2026-11-30T00:00:00", "timeZone": "UTC"},
                },
            ],
            "list-f": [
                {
                    "id": "t-flag",
                    "title": "Flag",
                    "status": "notStarted",
                    "dueDateTime": {"dateTime": "2026-11-26T00:00:00", "timeZone": "UTC"},
                }
            ],
        },
    )
    delta = InboxDelta(inbox)
    cal = CalendarDelta(box.events["cal-1"])
    return box, delta, cal


def install(fake, parts):
    box, delta, cal = parts
    box.install(fake)
    fake.on(r"/me/mailFolders(?:/inbox|\('inbox'\))/messages/delta", delta)
    fake.on(r"/me/calendarView/delta", cal)
    return parts


async def run(fake, parts=None, *, tools=frozenset(), now=NOW, install_it=True, **args):
    if install_it:
        install(fake, parts or mailbox())
    return await read_catch_up(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=tools, now=now, **args
    )


def section(env, name) -> dict:
    return next(s for s in env["sections"] if s["name"] == name)


def ids_in(env, name) -> list[str]:
    return [i["id"] for i in section(env, name)["items"]]


def decoded(token: str) -> dict:
    body = token[len("cu1.") :]
    return json.loads(base64.urlsafe_b64decode(body + "=" * (-len(body) % 4)))


# ── Arguments ─────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "args",
    [
        {"focus": "everything"},
        {"window": "+24h"},
        {"window": "tomorrow"},
        {"window": "32d"},
        {"window": "yesterday at noon"},
        {"cursor": "junk"},
        {"cursor": "cu1.!!"},
        {"cursor": "cu1." + base64.urlsafe_b64encode(b"[1]").decode().rstrip("=")},
    ],
)
async def test_bad_arguments_are_refused_before_any_request(fake, args):
    with pytest.raises(ValueError):
        await run(fake, **args)
    assert fake.counted == 0


async def test_a_foreign_link_in_a_cursor_raises_before_any_request(fake):
    state = cursor.CursorState(
        acct=cursor.account_tag(ME),
        made="2026-11-26T09:00:00Z",
        mail=cursor.MailState(
            link="https://graph.microsoft.com/v1.0/me/messages",
            created_only=True,
            bootstrap_from=None,
        ),
        cal=cursor.CalState(link="", start="2026-11-25T09:00:00Z", end="2026-12-26T09:00:00Z"),
    )
    with pytest.raises(ForeignCursorError):
        await run(fake, cursor=cursor.encode(state))
    assert fake.counted == 0


# ── The first call ────────────────────────────────────────────────────────


async def test_the_first_call_reads_the_window_and_returns_a_cursor(fake):
    env = await run(fake)
    assert env["tool"] == "outlook_catch_up"
    assert env["outcome"] == "answered"
    assert [s["name"] for s in env["sections"]] == [
        "needs_reply",
        "replies_to_you",
        "follow_ups",
        "invites",
        "meetings_today",
        "tasks_due",
        "other_mail",
    ]
    understood = env["understood"]
    assert understood["cursor"] == "none"
    assert understood["arrivals"]["rule"] == "created"
    assert understood["arrivals"]["since"] == "2026-11-25T12:00:00Z"
    assert understood["horizon"] == "2026-11-27T00:00:00Z"
    assert understood["not_covered"][:3] == ["mail_outside_inbox", "junk", "secondary_calendars"]
    assert env["cursor"].startswith("cu1.")
    state = decoded(env["cursor"])
    assert state["made"] == "2026-11-26T12:00:00Z"
    assert state["mail"]["link"] == f"{MAIL_LINK}?$deltatoken=round1"
    assert state["mail"]["created_only"] is True
    assert state["cal"]["link"] == f"{CAL_LINK}?$deltatoken=c1"
    boot = next(s for s in fake.operations if s.path.endswith("/messages/delta"))
    assert boot.params["changeType"] == "created"
    assert boot.params["$filter"] == "receivedDateTime ge 2026-11-25T12:00:00Z"
    assert boot.params["$select"] == catch_up.MAIL_SELECT
    assert "$top" not in boot.params


async def test_each_arrival_goes_to_the_first_section_it_qualifies_for(fake):
    env = await run(fake)
    assert ids_in(env, "replies_to_you") == ["a1"]
    # Jo is written to (one letter); Dana is not: Jo ranks first.
    assert ids_in(env, "needs_reply") == ["a10", "a2"]
    # Kim's reply is on the flag's item, and nowhere else.
    assert ids_in(env, "follow_ups") == ["f-bill", "f-kim"]
    shown = set(ids_in(env, "needs_reply")) | set(ids_in(env, "replies_to_you"))
    assert "a9" not in shown
    # Pat (already answered), Lee (Cc only), the newsletter, and Alex ("other").
    other = section(env, "other_mail")
    assert other["total"] == 4 and other["items"] == []
    by_sender = {s["address"] for s in env["understood"]["other_mail_by_sender"]}
    assert by_sender == {PAT, LEE, "news@list.example", "alex@example.net"}
    arrivals = env["understood"]["arrivals"]
    assert arrivals["from_your_address"] == 1
    assert arrivals["total"] == 8  # a1-a6, a9, a10: not the forgery, not the invitation


async def test_the_reasons_on_each_kind_of_arrival(fake):
    env = await run(fake)
    reply = section(env, "replies_to_you")["items"][0]
    assert reply["why"][0] == {"code": "you_wrote_in_thread", "at": "2026-11-25T10:00:00Z"}
    assert reply["from"]["known"] is True
    assert reply["from"]["sender_auth"]["result"] == "verified"
    needs = {i["id"]: i for i in section(env, "needs_reply")["items"]}
    assert {r["code"] for r in needs["a2"]["why"]} == {"only_recipient", "high_importance"}
    assert needs["a10"]["why"][1]["code"] == "sender_written_to"
    assert needs["a10"]["why"][1]["count"] == 1
    follow = {i["id"]: i for i in section(env, "follow_ups")["items"]}
    assert {"code": "flag_overdue", "due": "2026-11-20T00:00:00Z"} in follow["f-bill"]["why"]
    assert {
        "code": "reply_from_them",
        "at": "2026-11-26T08:00:00Z",
        "address": KIM,
        "known": True,
    } in follow["f-kim"]["why"]
    assert follow["f-kim"]["from"]["known"] is True  # the owner's own letter


async def test_cc_only_mail_never_needs_a_reply_and_bulk_mail_drops_to_other(fake):
    env = await run(fake)
    assert "a4" not in ids_in(env, "needs_reply")
    assert "a5" not in ids_in(env, "needs_reply")
    reads = {s.path for s in fake.operations if s.in_batch and s.path.startswith("/me/messages/")}
    assert "/me/messages/a5" in reads  # it was read, found bulk, and dropped


async def test_an_arrival_already_answered_is_not_needs_reply(fake):
    env = await run(fake)
    assert "a3" not in ids_in(env, "needs_reply") + ids_in(env, "replies_to_you")


async def test_the_state_sections(fake):
    env = await run(fake)
    invites = {i["id"]: i for i in section(env, "invites")["items"]}
    assert list(invites) == ["e2", "e3"]
    assert {"code": "stranger_invite"} in invites["e2"]["warnings"]
    assert "stranger_invite" not in {w["code"] for w in invites["e3"]["warnings"]}
    assert invites["e3"]["organizer"] == {"address": SAM, "known": True}
    # e4 ended, e5 was declined: only e1 is still ahead today.
    assert ids_in(env, "meetings_today") == ["e1"]
    assert ids_in(env, "tasks_due") == ["t-late", "t-today"]


# ── Later calls: the cursor ───────────────────────────────────────────────


async def test_read_flags_categories_and_moves_are_not_arrivals(fake):
    parts = install(fake, mailbox())
    first = await run(fake, install_it=False)
    _, delta, _ = parts
    moved_in = received("moved", "zed@example.com", "Old news", when="2026-10-01T09:00:00Z")
    read_now = {**received("old", SAM, "Older", when="2026-11-20T09:00:00Z"), "isRead": True}
    flag_now = {
        **received("a2", DANA, "Question", when="2026-11-26T09:00:00Z"),
        "flag": flagged("2026-12-01"),
    }
    categorized = {
        **received("a4", LEE, "FYI", when="2026-11-26T10:00:00Z"),
        "categories": ["Blue"],
    }
    new = received("new", "nia@example.com", "Hello", when="2026-11-26T14:00:00Z")
    delta.changes = [
        moved_in,
        read_now,
        flag_now,
        categorized,
        new,
        {"id": "gone", "@removed": {"reason": "deleted"}},
    ]
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    assert env["understood"]["cursor"] == "ok"
    assert env["understood"]["arrivals"]["total"] == 1
    assert env["understood"]["arrivals"]["since"] == "2026-11-26T11:58:00Z"
    assert ids_in(env, "needs_reply") == ["new"]
    followed = [s for s in fake.operations if "$deltatoken" in s.params]
    assert followed and followed[0].params["$deltatoken"] == "round1"


async def test_a_stale_cursor_resyncs_from_the_last_call(fake):
    parts = install(fake, mailbox())
    first = await run(fake, install_it=False)
    _, delta, _ = parts
    delta.stale = True
    delta.inbox.append(received("new", "nia@example.com", "Hello", when="2026-11-26T14:00:00Z"))
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    assert env["understood"]["resynced"] == ["mail"]
    boots = [
        s for s in fake.operations if s.path.endswith("/messages/delta") and "$filter" in s.params
    ]
    # From the last call (less the two-minute overlap), never from the default window.
    assert boots[-1].params["$filter"] == "receivedDateTime ge 2026-11-26T11:58:00Z"
    assert "$top" not in boots[-1].params  # it caps the whole round (step-6 eval)
    assert ids_in(env, "needs_reply") == ["new"]
    assert "gap" not in env["understood"]


async def test_a_stale_cursor_older_than_31_days_says_what_it_skipped(fake):
    parts = install(fake, mailbox())
    old_state = cursor.CursorState(
        acct=cursor.account_tag(ME),
        made="2026-10-01T12:00:00Z",
        mail=cursor.MailState(
            link=f"{MAIL_LINK}?$deltatoken=round1", created_only=True, bootstrap_from=None
        ),
        cal=cursor.CalState(
            link=f"{CAL_LINK}?$deltatoken=c1",
            start="2026-11-25T12:00:00Z",
            end="2026-12-26T12:00:00Z",
        ),
    )
    parts[1].stale = True
    env = await run(fake, install_it=False, cursor=cursor.encode(old_state))
    assert env["understood"]["gap"] == {
        "from": "2026-10-01T11:58:00Z",
        "to": "2026-10-26T12:00:00Z",
    }
    assert env["understood"]["arrivals"]["since"] == "2026-10-26T12:00:00Z"


async def test_p_dl1_without_created_only_the_received_time_rule_holds(fake):
    parts = mailbox()
    parts[1].refuse_created = True
    install(fake, parts)
    first = await run(fake, install_it=False)
    assert first["understood"]["arrivals"]["rule"] == "received_time"
    assert decoded(first["cursor"])["mail"]["created_only"] is False
    assert ids_in(first, "replies_to_you") == ["a1"]
    # The next round: a message received inside the two-minute overlap that also
    # changed may come back; it is shown once more, marked.
    edge = received("edge", "nia@example.com", "Edge", when="2026-11-26T11:59:00Z")
    parts[1].changes = [edge]
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    (item,) = section(env, "needs_reply")["items"]
    assert item["id"] == "edge"
    assert {"code": "may_repeat"} in item["warnings"]


async def test_another_accounts_cursor_starts_from_the_window(fake):
    other = cursor.CursorState(
        acct=cursor.account_tag("someone@example.org"),
        made="2026-11-26T09:00:00Z",
        mail=cursor.MailState(link="", created_only=True, bootstrap_from=None),
        cal=cursor.CalState(link="", start="2026-11-25T09:00:00Z", end="2026-12-26T09:00:00Z"),
    )
    env = await run(fake, cursor=cursor.encode(other))
    assert env["understood"]["cursor"] == "reset_other_account"
    assert env["understood"]["arrivals"]["since"] == "2026-11-25T12:00:00Z"
    assert "calendar_changes" not in [s["name"] for s in env["sections"]]


async def test_calendar_changes_and_removed_events_with_a_cursor(fake):
    parts = install(fake, mailbox())
    first = await run(fake, install_it=False)
    moved = ev("e1", "2026-11-26T16:00:00", "2026-11-26T17:00:00", modified="2026-11-26T13:00:00Z")
    cancelled = {
        **ev(
            "e3",
            "2026-11-28T10:00:00",
            "2026-11-28T11:00:00",
            modified="2026-11-26T13:30:00Z",
            organizer=SAM,
            is_organizer=False,
        ),
        "isCancelled": True,
    }
    long_ago = ev(
        "e0", "2026-11-01T10:00:00", "2026-11-01T11:00:00", modified="2026-11-26T13:00:00Z"
    )
    parts[2].changes = [moved, cancelled, long_ago, {"id": "e9", "@removed": {"reason": "x"}}]
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    changes = {i["id"]: i for i in section(env, "calendar_changes")["items"]}
    assert list(changes) == ["e1", "e3"]
    assert changes["e1"]["why"][0] == {"code": "changed_since_cursor"}
    assert {"code": "cancelled"} in changes["e3"]["why"]
    assert env["understood"]["calendar_removed"] == {"count": 1, "ids": ["e9"]}
    assert env["understood"]["calendar_rolled"] is False
    delta_reads = [s for s in fake.operations if s.path == "/me/calendarView/delta"]
    assert all("odata.maxpagesize=100" in s.headers["prefer"] for s in delta_reads)


async def test_the_calendar_rolls_when_its_look_ahead_runs_short(fake):
    parts = install(fake, mailbox())
    state = cursor.CursorState(
        acct=cursor.account_tag(ME),
        made="2026-11-26T09:00:00Z",
        mail=cursor.MailState(
            link=f"{MAIL_LINK}?$deltatoken=round1", created_only=True, bootstrap_from=None
        ),
        cal=cursor.CalState(
            link=f"{CAL_LINK}?$deltatoken=c1",
            start="2026-11-01T09:00:00Z",
            end="2026-12-05T09:00:00Z",
        ),
    )
    parts[2].snapshot[2]["lastModifiedDateTime"] = "2026-11-26T10:00:00Z"  # e3 changed
    env = await run(fake, install_it=False, cursor=cursor.encode(state))
    assert env["understood"]["calendar_rolled"] is True
    assert "deletions_at_roll" in env["understood"]["not_covered"]
    assert ids_in(env, "calendar_changes") == ["e3"]
    boot = next(s for s in fake.operations if s.path == "/me/calendarView/delta")
    assert boot.params["startDateTime"] == "2026-11-25T12:00:00Z"
    assert boot.params["endDateTime"] == "2026-12-26T12:00:00Z"
    assert decoded(env["cursor"])["cal"]["end"] == "2026-12-26T12:00:00Z"


async def test_series_masters_are_not_calendar_changes(fake):
    """P-C2, live (step-6 eval): calendarView delta returns each series master beside the
    occurrences, dated at the series' first start. A change to a series is its
    occurrences, so the master is not listed again."""
    parts = install(fake, mailbox())
    first = await run(fake, install_it=False)
    changed = "2026-11-26T13:00:00Z"
    master = {
        **ev("s1", "2026-11-27T09:00:00", "2026-11-27T09:30:00", modified=changed),
        "type": "seriesMaster",
    }
    occurrence = {
        **ev("s1-o1", "2026-11-27T09:00:00", "2026-11-27T09:30:00", modified=changed),
        "type": "occurrence",
        "seriesMasterId": "s1",
    }
    parts[2].changes = [master, occurrence]
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    assert ids_in(env, "calendar_changes") == ["s1-o1"]


# ── focus ─────────────────────────────────────────────────────────────────


async def test_focus_narrows_the_sections_and_reads(fake):
    env = await run(fake, focus="tasks")
    assert [s["name"] for s in env["sections"]] == ["tasks_due"]
    assert not [s for s in fake.operations if "/messages" in s.path or "calendar" in s.path]
    # The cursor keeps the window's start as the threshold, so mail is not lost.
    state = decoded(env["cursor"])
    assert state["mail"] == {
        "link": "",
        "created_only": True,
        "bootstrap_from": "2026-11-25T12:00:00Z",
    }
    fake.routes.clear()
    fake.seen.clear()
    env = await run(fake, focus="calendar")
    assert [s["name"] for s in env["sections"]] == ["invites", "meetings_today"]
    fake.routes.clear()
    env = await run(fake, focus="mail")
    assert [s["name"] for s in env["sections"]] == [
        "needs_reply",
        "replies_to_you",
        "follow_ups",
        "other_mail",
    ]


async def test_a_call_that_skipped_mail_hands_its_threshold_on(fake):
    install(fake, mailbox())
    first = await run(fake, install_it=False)
    env = await run(fake, install_it=False, now=LATER, cursor=first["cursor"], focus="tasks")
    state = decoded(env["cursor"])
    assert state["mail"]["link"] == f"{MAIL_LINK}?$deltatoken=round1"
    assert state["mail"]["bootstrap_from"] == "2026-11-26T11:58:00Z"
    # The calendar wasn't read: its link is kept, and the changes it owes date from the
    # first call (review D1).
    kept = decoded(first["cursor"])["cal"]
    assert state["cal"] == {**kept, "since": "2026-11-26T12:00:00Z"}


# ── Probes, failures and the budget ───────────────────────────────────────


async def test_p_d3a_a_refused_flag_filter_scans_three_folders(fake):
    parts = install(fake, mailbox())
    box = parts[0]

    def messages(seen, match):
        if seen.filter == flags.F1_FILTER:
            return error(400, "BadRequest", "invalid nodes")
        return box._messages(seen, match)

    fake.on("/me/messages", messages)
    env = await run(fake, install_it=False)
    assert ids_in(env, "follow_ups") == ["f-bill", "f-kim"]
    assert section(env, "follow_ups")["total_exact"] is False
    scanned = {
        s.path
        for s in fake.operations
        if s.path.startswith("/me/mailFolders/") and s.params.get("$select") == flags.F1_SELECT
    }
    assert scanned == {
        "/me/mailFolders/inbox/messages",
        "/me/mailFolders/sentitems/messages",
        "/me/mailFolders/archive/messages",
    }


@pytest.mark.parametrize("refuse_created", [False, True], ids=["created_only", "p_dl1_fallback"])
async def test_a_day_of_mail_reaches_its_delta_link_within_g2(fake, refuse_created):
    """Live (step-6 eval): message delta pages by 10 whatever `$top` says, so a day of 50
    arrivals never reached its deltaLink in G2's four reads and the answer was
    `incomplete`; and `$top` caps the whole round, so it is never sent.
    `odata.maxpagesize` sizes the pages, sent beside the id preference on every read of
    the round (D2: one id family per process)."""
    parts = mailbox()
    parts[1].refuse_created = refuse_created
    parts[0].received += [
        received(
            f"d{n}",
            f"p{n}@example.com",
            "Hi",
            conversation=f"cd{n}",
            when=f"2026-11-26T0{n % 10}:30:00Z",
        )
        for n in range(50)
    ]
    install(fake, parts)
    first = await run(fake, install_it=False)
    assert first["understood"]["arrivals"]["has_more"] is False
    assert first["outcome"] == "answered"
    assert first["understood"]["arrivals"]["total"] == 58  # all of them: none cut by `$top`
    assert decoded(first["cursor"])["mail"]["link"] == f"{MAIL_LINK}?$deltatoken=round1"
    await run(fake, install_it=False, now=LATER, cursor=first["cursor"])
    reads = [s for s in fake.operations if s.path.endswith("/messages/delta")]
    assert any("$deltatoken" in s.params for s in reads)  # the cursor's round too
    for seen in reads:
        assert "$top" not in seen.params
        prefer = [p.strip() for p in seen.headers["prefer"].split(",")]
        assert f"odata.maxpagesize={catch_up.MAIL_PAGE_SIZE}" in prefer
        assert ids.IMMUTABLE_ID_PREFERENCE in prefer


async def test_a_failed_calendar_read_leaves_the_mail(fake):
    install(fake, mailbox())
    fake.on("/me/calendarView", lambda s, m: error(500, "InternalServerError"))
    env = await run(fake, install_it=False)
    assert env["outcome"] == "incomplete"
    statuses = {s["name"]: s["status"] for s in env["sections"]}
    assert statuses["invites"] == statuses["meetings_today"] == "failed"
    assert statuses["needs_reply"] == "ok"
    # The calendar's state is kept for the next call, not lost.
    assert decoded(env["cursor"])["cal"]["link"] == ""


async def test_a_spent_budget_is_incomplete_not_empty(fake, monkeypatch):
    monkeypatch.setitem(budget.CEILINGS, "outlook_catch_up", 3)
    env = await run(fake)
    assert env["outcome"] == "incomplete"
    assert "skipped_budget" in {s["status"] for s in env["sections"]}
    assert env["cursor"].startswith("cu1.")


def worst_case() -> tuple[Mailbox, InboxDelta, CalendarDelta]:
    """Every read at its most: 300 letters, 250 arrivals, 12 flags due, 250 events and
    seven task lists."""
    box, delta, cal = mailbox()
    box.sent += [
        sent(f"w{n}", [f"p{n}@example.com"], "2026-11-24T10:00:00Z", conversation=f"cw{n}")
        for n in range(300)
    ]
    box.received += [
        received(
            f"n{n}", f"p{n}@example.com", "Hi", conversation=f"cn{n}", when="2026-11-26T02:00:00Z"
        )
        for n in range(250)
    ]
    box.received += [
        {
            **received(f"g{n}", BILL, "Due", conversation=f"cg{n}", when="2026-11-01T09:00:00Z"),
            "flag": flagged("2026-11-25"),
        }
        for n in range(12)
    ]
    delta.inbox[:] = box.received
    box.events["cal-1"] += [
        ev(f"x{n}", "2026-11-29T10:00:00", "2026-11-29T11:00:00") for n in range(250)
    ]
    box.task_lists += [
        {"id": f"l{n}", "displayName": f"L{n}", "wellknownListName": "none"} for n in range(7)
    ]
    return box, delta, cal


async def test_the_worst_case_stays_within_thirty_six_requests(fake):
    parts = worst_case()
    env = await run(fake, parts)
    assert env["graph_requests"] == fake.counted
    assert fake.counted <= 36 <= budget.CEILINGS["outlook_catch_up"] == 40
    assert env["understood"]["arrivals"]["has_more"] is True
    assert "task_lists_beyond_6" in env["understood"]["not_covered"]
    assert env["understood"]["flags_not_checked"] > 0
    # The unfinished round goes on next time from where it stopped, with its threshold.
    state = decoded(env["cursor"])
    assert "$skiptoken" in state["mail"]["link"]
    assert state["mail"]["bootstrap_from"] == "2026-11-25T12:00:00Z"


# ── Text, vocabularies and next ───────────────────────────────────────────


def _marker_paths(node, path=()):
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _marker_paths(value, (*path, key))
    elif isinstance(node, list):
        for value in node:
            yield from _marker_paths(value, path)
    elif isinstance(node, str) and MARK.strip() in node:
        yield path


async def test_text_from_graph_appears_only_under_others_wrote(fake):
    env = await run(fake, mailbox(MARK))
    found = list(_marker_paths(env))
    assert found
    for path in found:
        assert "others_wrote" in path, f"Graph text outside others_wrote at {path}"
        assert "why" not in path and "warnings" not in path and path[0] != "next"
    text = json.dumps(decoded(env["cursor"]))
    assert MARK.strip() not in text and "@" not in text


async def test_codes_are_in_the_vocabulary(fake):
    env = await run(fake)
    for item in [i for s in env["sections"] for i in s["items"]]:
        assert {r["code"] for r in item["why"]} <= WHY_CODES
        assert {w["code"] for w in item["warnings"]} <= WARNING_CODES


async def test_next_drafts_clears_answers_and_opens(fake):
    tools = frozenset(
        {"outlook_compose", "outlook_follow_up", "outlook_answer_invites", "outlook_list_thread"}
    )
    env = await run(fake, tools=tools)
    assert [(s["template"], s["args"]) for s in env["next"]] == [
        ("draft_reply", {"mode": "reply", "message_id": "a10"}),
        ("draft_reply", {"mode": "reply", "message_id": "a2"}),
        ("clear_follow_up", {"message_id": "f-kim", "clear": True}),
        # The first invitation is a stranger's: answer goes to the first unblocked one.
        ("answer", {"event_ids": ["e3"]}),
        ("open_thread", {"conversation_id": "conv-sam"}),
    ]
    fake.routes.clear()
    env = await run(fake, tools=frozenset({"outlook_list_thread"}))
    assert [s["template"] for s in env["next"]] == ["open_thread"]
