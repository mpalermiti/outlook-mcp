"""Regression tests for the review of the intent read tools (#281-#284).

Each began as a repro of the reviewer's, asserting the faulty behaviour; each is
written here the right way round and was run red before its fix. The ids are the
review's: A for agenda, B for the blockers, C for open_loops, D for catch_up, E
for brief, and "nit" for the small ones fixed with them.
"""

from __future__ import annotations

import time
from datetime import datetime, timedelta, timezone

import pytest

from outlook_mcp.intents import ids, open_loops, transport
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.intents.open_loops import read_open_loops
from tests import test_intent_agenda as ag
from tests import test_intent_open_loops as ol
from tests.intent_fakes import ME, FakeGraph, Mailbox, auth_header, error, received

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
EVERY_TOOL = frozenset(
    {
        "outlook_compose",
        "outlook_follow_up",
        "outlook_list_thread",
        "outlook_read_message",
        "outlook_brief",
        "outlook_answer_invites",
    }
)


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def sec(env, name):
    return next(s for s in env["sections"] if s["name"] == name)


# ── #281 agenda, and the shared task read ─────────────────────────────────


def _worst_agenda() -> Mailbox:
    """Five calendars of 300 events (three pages each) and eight task lists."""
    calendars = [
        {"id": f"cal-{n}", "name": f"Calendar {n}", "isDefaultCalendar": n == 0} for n in range(5)
    ]
    events = {
        c["id"]: [
            ag.ev(f"{c['id']}-{n}", "2026-11-26T09:00:00", "2026-11-26T09:30:00")
            for n in range(300)
        ]
        for c in calendars
    }
    lists = [
        {"id": f"list-{n}", "displayName": f"L{n}", "wellknownListName": "none"} for n in range(8)
    ]
    tasks = {lst["id"]: [ag.task(f"{lst['id']}-t", "t", "2026-11-26")] for lst in lists}
    return Mailbox(calendars=calendars, events=events, task_lists=lists, tasks=tasks)


async def test_a2_a_refused_task_filter_costs_one_request_not_one_per_list(fake):
    """P-H1a's fallback: the first 400 is remembered, so the other lists are read
    unfiltered at once. The worst case is 26, within agenda's 30."""
    box = _worst_agenda()
    box.task_filter_supported = False
    env = await ag.run(fake, box)
    assert env["outcome"] == "answered"
    assert fake.counted == 26
    assert len(sec(env, "tasks_due")["items"]) == 8


async def test_a1_a_capped_calendar_read_never_shows_unread_time_as_free(fake):
    """The view is read in start order, so a cap cuts the tail; past the last event
    read, nothing is called free; and the answer says it is incomplete."""
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                ag.ev(f"f{n}", "2026-11-27T09:00:00", "2026-11-27T09:30:00", show_as="free")
                for n in range(300)
            ]
            + [ag.ev("board", "2026-11-26T14:00:00", "2026-11-26T17:00:00")]
            + [ag.ev("late", "2026-11-28T10:00:00", "2026-11-28T11:00:00")]
        },
    )
    env = await ag.run(fake, box, window="this_week", include_tasks=False)
    first = next(s for s in fake.operations if s.path.endswith("/calendarView"))
    assert first.params["$orderby"] == "start/dateTime"
    days = {d["date"]: d for d in sec(env, "days")["items"]}
    today = days["2026-11-26"]
    assert "board" in [e["id"] for e in today["events"]]
    assert today["free"] == [
        {"start": "2026-11-26T12:00:00+00:00", "end": "2026-11-26T14:00:00+00:00", "minutes": 120},
        {"start": "2026-11-26T17:00:00+00:00", "end": "2026-11-26T18:00:00+00:00", "minutes": 60},
    ]
    # The read stopped inside 11-27; nothing after that is known to be free.
    assert days["2026-11-28"]["free"] == [] and days["2026-11-29"]["free"] == []
    assert env["outcome"] == "incomplete"
    assert sec(env, "days")["total_exact"] is False
    assert env["understood"]["events_read_until"] == "2026-11-27T09:00:00Z"


async def test_a4_next_never_points_at_a_strangers_unanswered_invitation(fake):
    """agenda reads no sent mail, so it can't tell a stranger: brief_meeting goes only to
    a meeting the owner organizes or has accepted."""
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                ag.ev(
                    "bait",
                    "2026-11-26T12:30:00",
                    "2026-11-26T13:00:00",
                    organizer=ag.STRANGER,
                    organizer_name="Stranger",
                    is_organizer=False,
                    response="notResponded",
                    show_as="tentative",
                ),
                ag.ev("mine", "2026-11-26T15:00:00", "2026-11-26T16:00:00"),
            ]
        },
    )
    env = await ag.run(fake, box, tools=EVERY_TOOL, include_tasks=False)
    briefs = [n["args"] for n in env["next"] if n["template"] == "brief_meeting"]
    assert briefs == [{"meeting": "mine"}]


async def test_a3_a_plain_month_keeps_every_day(fake):
    """31 days of three meetings with ids of Graph's length: every day stays, with its
    free time and busy minutes; events are cut from the last days first, and counted."""
    long_id = "AAMkAGI2" + "x" * 140
    cal_id = "AQMkADAw" + "y" * 110
    events = []
    for day in range(31):
        date = (datetime(2026, 11, 26) + timedelta(days=day)).date().isoformat()
        for n, hour in enumerate((9, 11, 14)):
            events.append(
                ag.ev(
                    f"{long_id}{day:02d}{n}",
                    f"{date}T{hour:02d}:00:00",
                    f"{date}T{hour + 1:02d}:00:00",
                )
            )
    box = Mailbox(
        calendars=[{"id": cal_id, "name": "Calendar", "isDefaultCalendar": True}],
        events={cal_id: events},
    )
    env = await ag.run(fake, box, window="2026-11-26T00:00:00Z/2026-12-27T00:00:00Z")
    days = sec(env, "days")
    assert days["total"] == 31 and len(days["items"]) == 31
    assert all(d["busy_minutes"] == 180 for d in days["items"])
    assert all(d["events_total"] == 3 for d in days["items"])
    assert len(days["items"][0]["events"]) == 3  # the first days keep their events
    # A day's events name their calendar by its place in understood.calendars.
    assert days["items"][0]["events"][0]["calendar"] == 0
    assert env["understood"]["calendars"][0]["id"] == cal_id
    assert env["trimmed"] is True and env["partial"] is True


def test_nit_the_view_is_read_in_start_order():
    from outlook_mcp.intents import reading

    assert reading.VIEW_ORDER == "start/dateTime"


async def test_nit_an_overlap_names_at_most_five_others(fake):
    """A day of 30 meetings at once: each overlap item names five, not 29, others."""
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                ag.ev(f"m{n:02}", "2026-11-26T13:00:00", "2026-11-26T14:00:00") for n in range(30)
            ]
        },
    )
    env = await ag.run(fake, box, include_tasks=False)
    for item in sec(env, "overlaps")["items"]:
        assert len([r for r in item["why"] if r["code"] == "overlaps"]) == 5


# ── #282 open_loops, and the reply checks it shares with catch_up ────────


def _later(
    msg_id, sender, subject, when, *, folder="inbox-id", headers=None, conversation="conv-lee"
):
    return received(
        msg_id,
        sender,
        subject,
        conversation=conversation,
        when=when,
        folder=folder,
        headers=headers if headers is not None else [],
    )


@pytest.mark.parametrize(
    "later",
    [
        # A bounce, in Junk (the review's repro).
        _later(
            "junk-1",
            "mailer-daemon@bounce.example",
            "Undeliverable: Contract",
            "2026-10-01T10:05:00Z",
            folder="junk-id",
        ),
        # A bounce in the Inbox: not from anyone the letter went to.
        _later(
            "bounce", "postmaster@example.com", "Undeliverable: Contract", "2026-10-01T10:05:00Z"
        ),
        # Lee's out-of-office reply.
        _later(
            "oof",
            ol.LEE,
            "Automatic reply: Contract",
            "2026-10-01T10:06:00Z",
            headers=[
                *auth_header("pass", "example.com"),
                {"name": "Auto-Submitted", "value": "auto-replied"},
            ],
        ),
        # Mail claiming Lee's address that fails Microsoft's check.
        _later(
            "spoof",
            ol.LEE,
            "Re: Contract",
            "2026-10-02T10:00:00Z",
            headers=auth_header("fail", "example.com"),
        ),
        # A checked message from someone the letter did not go to.
        _later(
            "third",
            "colleague@example.net",
            "Re: Contract",
            "2026-10-02T09:00:00Z",
            headers=auth_header("pass", "example.net"),
        ),
        # Lee's real reply, but in Deleted Items.
        _later(
            "deleted",
            ol.LEE,
            "Re: Contract",
            "2026-10-02T10:00:00Z",
            folder="deleted-id",
            headers=auth_header("pass", "example.com"),
        ),
    ],
    ids=[
        "bounce_in_junk",
        "bounce_in_inbox",
        "out_of_office",
        "failed_check",
        "not_a_recipient",
        "deleted",
    ],
)
async def test_b1_only_a_checked_reply_from_whom_the_letter_went_to_ends_a_wait(fake, later):
    box = ol.mailbox()
    box.received.append(later)
    env = await ol.run(fake, box)
    assert "conv-lee" in [i["id"] for i in sec(env, "waiting_on_them")["items"]]


async def test_b1_a_checked_reply_from_the_recipient_does_end_it(fake):
    box = ol.mailbox()
    box.received.append(
        _later(
            "real",
            ol.LEE,
            "Re: Contract",
            "2026-10-02T10:00:00Z",
            headers=auth_header("pass", "example.com"),
        )
    )
    env = await ol.run(fake, box)
    assert "conv-lee" not in [i["id"] for i in sec(env, "waiting_on_them")["items"]]


async def test_b1_a_reply_whose_check_could_not_run_keeps_the_follow_up(fake):
    box = ol.mailbox()
    box.received.append(
        _later(
            "real",
            ol.LEE,
            "Re: Contract",
            "2026-10-02T10:00:00Z",
            headers=auth_header("pass", "example.com"),
        )
    )
    box.install(fake)
    fake.on(r"/me/messages/real", lambda s, m: error(500, "InternalServerError"))
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=frozenset(), now=NOW
    )
    assert "conv-lee" in [i["id"] for i in sec(env, "waiting_on_them")["items"]]
    assert env["understood"]["replies_unconfirmed"] == ["s5"]


async def test_b2_a_received_letter_that_starts_cancelled_is_listed(fake):
    """The subject prefix only guards the owner's sent table; received mail is a meeting
    message only by its type."""
    box = ol.mailbox()
    box.received.append(
        received(
            "flight",
            "noreply@airline.example",
            "Cancelled: your flight tomorrow",
            conversation="conv-flight",
            when="2026-11-25T10:30:00Z",
            headers=auth_header("pass", "airline.example"),
        )
    )
    env = await ol.run(fake, box)
    assert "conv-flight" in [i["id"] for i in sec(env, "waiting_on_me")["items"]]


async def test_c1_a_spoof_of_a_known_sender_is_checked_and_never_offered_a_reply(fake):
    box = ol.mailbox()
    box.received.append(
        received(
            "spoof",
            ol.PAT,
            "Wire instructions changed",
            conversation="conv-spoof",
            when="2026-11-13T09:00:00Z",
            headers=auth_header("fail", "example.org"),
        )
    )
    box.install(fake)
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=EVERY_TOOL, now=NOW
    )
    spoof = next(i for i in sec(env, "waiting_on_me")["items"] if i["id"] == "conv-spoof")
    assert spoof["sender_auth"]["result"] == "failed"
    assert {"code": "sender_failed"} in spoof["warnings"]
    drafts = [n["args"]["message_id"] for n in env["next"] if n["template"] == "draft_reply"]
    assert drafts and "spoof" not in drafts


async def test_c1_draft_reply_needs_a_check_that_ran(fake):
    box = ol.mailbox()
    box.install(fake)
    for msg in ("a1", "r1", "r7", "r2"):
        fake.on(rf"/me/messages/{msg}", lambda s, m: error(500, "InternalServerError"))
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=EVERY_TOOL, now=NOW
    )
    assert not [n for n in env["next"] if n["template"] == "draft_reply"]


async def test_c2_a_capped_inbox_never_reports_an_answered_letter_as_waiting(fake):
    box = ol.mailbox()
    box.received.append(
        received(
            "pat-reply", ol.PAT, "Re: Quote", conversation="conv-pat", when="2026-11-19T09:00:00Z"
        )
    )
    box.received += [
        received(f"n{n}", f"list{n}@example.net", "News", when="2026-11-25T09:00:00Z")
        for n in range(300)
    ]
    env = await ol.run(fake, box)
    assert "conv-pat" not in [i["id"] for i in sec(env, "waiting_on_them")["items"]]
    assert env["outcome"] == "incomplete"
    assert env["understood"]["conversations_not_judged"] >= 1
    assert env["understood"]["judged_since"] == "2026-11-25T09:00:00Z"


async def test_c3_one_deadline_covers_the_lookups_before_the_sections(fake, monkeypatch):
    monkeypatch.setattr(open_loops, "INTENT_DEADLINE_S", 0.2)
    fake.delays.append((lambda s: bool(s.search), 1.0))
    began = time.monotonic()
    env = await ol.run(fake, words="plan")
    assert time.monotonic() - began < 0.9
    assert env["outcome"] == "incomplete"
    assert {s["status"] for s in env["sections"]} == {"skipped_time"}


async def test_nit_a_searched_thread_ending_in_a_draft_is_judged_by_its_last_message(fake):
    box = ol.mailbox()
    box.received += [
        received(
            "old-q", ol.BOB, "Zebra budget", conversation="conv-z", when="2026-09-01T09:00:00Z"
        ),
        received(
            "old-draft",
            ME,
            "Re: Zebra budget",
            conversation="conv-z",
            when="2026-09-02T09:00:00Z",
            folder="drafts-id",
            draft=True,
        ),
    ]
    env = await ol.run(fake, box, words="zebra")
    assert "conv-z" in [i["id"] for i in sec(env, "waiting_on_me")["items"]]


async def test_nit_an_unread_searched_conversation_is_counted_not_named(fake):
    box = ol.mailbox()
    box.received.append(
        received(
            "old",
            "zed@example.com",
            "Invoice 42",
            conversation="conv-old",
            when="2026-09-01T09:00:00Z",
        )
    )
    box.install(fake)
    real = box._messages

    def messages(seen, match):
        if seen.in_batch and "conv-old" in (seen.filter or ""):
            return error(500, "InternalServerError")
        return real(seen, match)

    fake.on("/me/messages", messages)
    env = await read_open_loops(
        fake.graph_client(),
        zone="UTC",
        signed_in=ME,
        loaded_tools=frozenset(),
        now=NOW,
        words="invoice",
    )
    assert env["understood"]["searched_conversations_unread"] == 1
    assert not [c for c in env["understood"]["not_covered"] if ":" in c]
