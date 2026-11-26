"""Regression tests for the review of the intent read tools (#281-#284).

Each began as a repro of the reviewer's, asserting the faulty behaviour; each is
written here the right way round and was run red before its fix. The ids are the
review's: A for agenda, B for the blockers, C for open_loops, D for catch_up, E
for brief, and "nit" for the small ones fixed with them.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from outlook_mcp.intents import ids, transport
from outlook_mcp.intents.ids import IdFamily
from tests import test_intent_agenda as ag
from tests.intent_fakes import ME, FakeGraph, Mailbox

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


# ── Live finding (probe day 1): Graph's isOrganizer is not the owner's word ──


def _owned(event_id: str, start: str, end: str, **kwargs) -> dict:
    """The owner's own meeting as a personal mailbox returns it: organizer the owner's
    address, isOrganizer false, and no response of its own."""
    row = ag.ev(event_id, start, end, organizer=ME, organizer_name="Me", **kwargs)
    row["isOrganizer"] = False
    row["responseStatus"] = {"response": "none"}
    return row


async def test_live_agenda_knows_the_owners_meeting_by_its_organizer_address(fake):
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={"cal-1": [_owned("mine", "2026-11-26T15:00:00", "2026-11-26T16:00:00")]},
    )
    env = await ag.run(fake, box, tools=EVERY_TOOL, include_tasks=False, signed_in=ME)
    assert sec(env, "invites")["items"] == []
    (shown,) = next(d for d in sec(env, "days")["items"] if d["date"] == "2026-11-26")["events"]
    assert shown["is_organizer"] is True
    assert {"code": "organizer_you"} in shown["why"]
    assert [n["args"] for n in env["next"] if n["template"] == "brief_meeting"] == [
        {"meeting": "mine"}
    ]


async def test_live_agenda_never_trusts_is_organizer_on_a_strangers_invitation(fake):
    row = ag.ev(
        "bait",
        "2026-11-26T15:00:00",
        "2026-11-26T16:00:00",
        organizer=ag.STRANGER,
        organizer_name="Stranger",
        response="notResponded",
    )
    row["isOrganizer"] = True  # whatever Graph says, the organizer is a stranger
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={"cal-1": [row]},
    )
    env = await ag.run(fake, box, tools=EVERY_TOOL, include_tasks=False, signed_in=ME)
    (invite,) = sec(env, "invites")["items"]
    assert invite["is_organizer"] is False
    assert not [n for n in env["next"] if n["template"] == "brief_meeting"]
