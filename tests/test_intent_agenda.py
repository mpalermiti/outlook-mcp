"""``outlook_agenda`` (contracts §12.4): your days in a window.

Days with their meetings and free blocks, overlapping meetings, open invitations
and To Do tasks due. Pinned here: which events count as busy (``showAs``, the
response, cancellation, all-day events), local days across a daylight-saving
change, one meeting on two calendars kept once, free time that starts at now,
the 31-day limit, the request count, and where other people's text may appear.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest

from outlook_mcp.intents import agenda, budget, ids, transport
from outlook_mcp.intents.agenda import read_agenda
from outlook_mcp.intents.envelope import WARNING_CODES, WHY_CODES
from outlook_mcp.intents.ids import IdFamily
from tests.intent_fakes import FakeGraph, Mailbox, error, event, page

# A Thursday, noon in UTC.
NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
STRANGER = "stranger@unknown.example"
MARK = " ZQXMARK "


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def ev(event_id, start, end, *, uid=None, show_as="busy", m="", **kwargs) -> dict:
    row = event(event_id, f"Meeting {event_id}{m}", start, end, show_as=show_as, **kwargs)
    row["iCalUId"] = uid or f"uid-{event_id}"
    row["location"] = {"displayName": f"Room {event_id}{m}"}
    return row


def all_day(event_id, day, next_day, *, show_as="oof", m="") -> dict:
    return ev(
        event_id,
        f"{day}T00:00:00.0000000",
        f"{next_day}T00:00:00.0000000",
        show_as=show_as,
        all_day=True,
        m=m,
    )


def task(task_id, title, due=None, *, status="notStarted", importance="normal") -> dict:
    row = {"id": task_id, "title": title, "status": status, "importance": importance}
    if due:
        row["dueDateTime"] = {"dateTime": f"{due}T00:00:00.0000000", "timeZone": "UTC"}
    return row


def mailbox(m: str = "") -> Mailbox:
    """Today (2026-11-26, UTC): a busy 13:00-14:00 on the default calendar, a tentative
    invitation 13:30-14:30 from a stranger on the second, a meeting shown as free, a
    declined one, a cancelled one, and a copy of the first on the second calendar.
    Tomorrow: an all-day out-of-office."""
    declined = ev("ev-d", "2026-11-26T16:00:00", "2026-11-26T17:00:00", m=m, response="declined")
    declined["isOrganizer"] = False
    cancelled = ev("ev-e", "2026-11-26T10:00:00", "2026-11-26T11:00:00", m=m)
    cancelled["isCancelled"] = True
    return Mailbox(
        calendars=[
            {"id": "cal-2", "name": f"Shared{m}", "isDefaultCalendar": False},
            {"id": "cal-1", "name": f"Calendar{m}", "isDefaultCalendar": True},
        ],
        events={
            "cal-1": [
                ev("ev-a", "2026-11-26T13:00:00", "2026-11-26T14:00:00", m=m),
                ev("ev-c", "2026-11-26T15:00:00", "2026-11-26T16:00:00", show_as="free", m=m),
                declined,
                cancelled,
                all_day("ev-oof", "2026-11-27", "2026-11-28", m=m),
            ],
            "cal-2": [
                ev(
                    "ev-b",
                    "2026-11-26T13:30:00",
                    "2026-11-26T14:30:00",
                    show_as="tentative",
                    organizer=STRANGER,
                    organizer_name=f"Stranger{m}",
                    is_organizer=False,
                    response="notResponded",
                    m=m,
                ),
                ev("ev-a-copy", "2026-11-26T13:00:00", "2026-11-26T14:00:00", uid="uid-ev-a", m=m),
            ],
        },
        task_lists=[
            {"id": "list-1", "displayName": f"Tasks{m}", "wellknownListName": "defaultList"},
            {"id": "list-f", "displayName": "Flagged email", "wellknownListName": "flaggedEmails"},
        ],
        tasks={
            "list-1": [
                task("t-today", f"Send the deck{m}", "2026-11-26", importance="high"),
                task("t-late", f"Pay the invoice{m}", "2026-11-20"),
                task("t-later", f"Book travel{m}", "2026-12-15"),
                task("t-done", f"Done already{m}", "2026-11-26", status="completed"),
                task("t-none", f"No date{m}"),
            ],
            "list-f": [task("t-flag", f"Flagged mail{m}", "2026-11-26")],
        },
    )


async def run(fake, box=None, *, tools=frozenset({"outlook_brief"}), **args):
    (box or mailbox()).install(fake)
    return await read_agenda(
        fake.graph_client(), config_zone="UTC", loaded_tools=tools, now=NOW, **args
    )


def section(env, name) -> dict:
    return next(s for s in env["sections"] if s["name"] == name)


def ids_in(env, name) -> list[str]:
    return [i["id"] for i in section(env, name)["items"]]


def day(env, date) -> dict:
    return next(d for d in section(env, "days")["items"] if d["date"] == date)


# ── Arguments ─────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "args",
    [
        {"window": "tomorrow at 3"},
        {"window": "+32d"},
        {"window": "2026-11-01/2026-12-15"},
        {"zone": "PST"},
        {"day_start": "9:00"},
        {"day_start": "24:00"},
        {"day_end": "25:00"},
        {"day_start": "18:00", "day_end": "09:00"},
        {"day_start": "09:00", "day_end": "09:00"},
        {"min_free": 4},
        {"min_free": 481},
        {"min_free": True},
        {"calendars": ["a", "b", "c", "d", "e", "f"]},
        {"calendars": []},
        {"calendars": ["bad\nname"]},
        {"include_tasks": "yes"},
    ],
)
async def test_bad_arguments_are_refused_before_any_request(fake, args):
    with pytest.raises(ValueError):
        await run(fake, **args)
    assert fake.counted == 0


async def test_the_envelope_and_its_sections(fake):
    env = await run(fake)
    assert env["tool"] == "outlook_agenda"
    assert env["outcome"] == "answered"
    assert [s["name"] for s in env["sections"]] == ["days", "overlaps", "invites", "tasks_due"]
    assert {s["cap"] for s in env["sections"]} <= {31, 20, 30}
    assert env["partial"] is False and env["failed_sections"] == []
    understood = env["understood"]
    assert understood["window"]["start"] == "2026-11-26T00:00:00Z"
    assert understood["window"]["end"] == "2026-11-27T00:00:00Z"
    assert understood["zone"] == "UTC"
    assert understood["day_start"] == "09:00" and understood["day_end"] == "18:00"
    assert understood["min_free"] == 30
    assert [c["id"] for c in understood["calendars"]] == ["cal-1", "cal-2"]
    assert understood["calendars"][0]["is_default"] is True


# ── Busy, free and the days ───────────────────────────────────────────────


async def test_busy_follows_show_as_the_response_and_cancellation(fake):
    today = day(await run(fake), "2026-11-26")
    # Busy: ev-a 13:00-14:00 and the tentative ev-b 13:30-14:30. Not busy: ev-c (free),
    # ev-d (declined), ev-e (cancelled).
    assert today["busy_minutes"] == 90
    # Free from now (12:00) to 18:00, around the busy 13:00-14:30.
    assert today["free"] == [
        {"start": "2026-11-26T12:00:00+00:00", "end": "2026-11-26T13:00:00+00:00", "minutes": 60},
        {"start": "2026-11-26T14:30:00+00:00", "end": "2026-11-26T18:00:00+00:00", "minutes": 210},
    ]
    shown = {e["id"]: e for e in today["events"]}
    assert set(shown) == {"ev-e", "ev-a", "ev-b", "ev-c", "ev-d"}
    assert [e["id"] for e in today["events"]] == ["ev-e", "ev-a", "ev-b", "ev-c", "ev-d"]
    assert {w["code"] for w in shown["ev-c"]["warnings"]} == {"shows_free"}
    assert {w["code"] for w in shown["ev-d"]["warnings"]} == {"declined"}
    assert any(r["code"] == "cancelled" for r in shown["ev-e"]["why"])
    assert today["why"] == [{"code": "today"}]


@pytest.mark.parametrize(
    ("show_as", "blocks"), [("oof", True), ("busy", True), ("tentative", False)]
)
async def test_an_all_day_event_blocks_the_day_only_when_busy_or_away(fake, show_as, blocks):
    box = mailbox()
    box.events["cal-1"][-1]["showAs"] = show_as
    tomorrow = day(await run(fake, box, window="tomorrow"), "2026-11-27")
    assert tomorrow["busy_minutes"] == (540 if blocks else 0)
    assert (tomorrow["free"] == []) is blocks
    (shown,) = tomorrow["events"]
    assert shown["all_day"] is True
    assert {"code": "all_day"} in shown["warnings"]


@pytest.mark.parametrize(("date", "minutes"), [("2026-11-01", 1500), ("2026-03-08", 1380)])
async def test_a_local_day_across_a_daylight_saving_change(fake, date, minutes):
    """An all-day busy event over the whole local day, 00:00 to 24:00, in Los Angeles: the
    day is 25 hours long when the clocks go back and 23 when they go forward."""
    next_day = {"2026-11-01": "2026-11-02", "2026-03-08": "2026-03-09"}[date]
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={"cal-1": [all_day("ev-x", date, next_day, show_as="busy")]},
    )
    env = await run(
        fake,
        box,
        window=date,
        zone="America/Los_Angeles",
        day_start="00:00",
        day_end="24:00",
        include_tasks=False,
    )
    (only,) = section(env, "days")["items"]
    assert only["date"] == date
    assert only["busy_minutes"] == minutes
    (shown,) = only["events"]
    assert shown["start"] == f"{date}T00:00:00" + ("-07:00" if date == "2026-11-01" else "-08:00")


async def test_one_meeting_on_two_calendars_is_kept_once_from_the_default(fake):
    env = await run(fake)
    shown = day(env, "2026-11-26")["events"]
    assert [e["id"] for e in shown].count("ev-a") == 1
    assert "ev-a-copy" not in [e["id"] for e in shown]
    assert next(e for e in shown if e["id"] == "ev-a")["calendar"] == "cal-1"


async def test_free_time_today_starts_at_now(fake):
    box = Mailbox(calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}])
    env = await run(fake, box, include_tasks=False)
    assert day(env, "2026-11-26")["free"] == [
        {"start": "2026-11-26T12:00:00+00:00", "end": "2026-11-26T18:00:00+00:00", "minutes": 360}
    ]
    env = await run(fake, box, window="tomorrow", include_tasks=False)
    assert day(env, "2026-11-27")["free"][0]["start"] == "2026-11-27T09:00:00+00:00"


async def test_short_gaps_are_not_free_blocks(fake):
    env = await run(fake, min_free=90)
    assert [b["minutes"] for b in day(env, "2026-11-26")["free"]] == [210]


async def test_a_day_holds_at_most_thirty_events_and_counts_the_rest(fake):
    box = Mailbox(
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                ev(
                    f"ev-{n:02}",
                    f"2026-11-26T{9 + n // 4:02}:{(n % 4) * 15:02}:00",
                    f"2026-11-26T{9 + n // 4:02}:{(n % 4) * 15 + 10:02}:00",
                )
                for n in range(35)
            ]
        },
    )
    today = day(await run(fake, box, include_tasks=False), "2026-11-26")
    assert len(today["events"]) == 30
    assert today["events_total"] == 35


async def test_a_week_has_one_item_per_local_day(fake):
    env = await run(fake, window="this_week")
    days = section(env, "days")
    assert [d["date"] for d in days["items"]] == [f"2026-11-{n}" for n in range(23, 30)]
    assert days["total"] == 7


# ── Overlaps and invites ──────────────────────────────────────────────────


async def test_overlaps_pair_busy_events_across_calendars(fake):
    env = await run(fake)
    overlaps = {i["id"]: i for i in section(env, "overlaps")["items"]}
    assert set(overlaps) == {"ev-a", "ev-b"}
    assert {"code": "overlaps", "with": "ev-b"} in overlaps["ev-a"]["why"]
    assert {"code": "overlaps", "with": "ev-a"} in overlaps["ev-b"]["why"]
    assert overlaps["ev-b"]["calendar"] == "cal-2"


async def test_invites_are_open_invitations_with_no_known_flag(fake):
    env = await run(fake)
    (invite,) = section(env, "invites")["items"]
    assert invite["id"] == "ev-b"
    assert invite["organizer"] == {"address": STRANGER, "known": None}
    assert invite["response"] == "notResponded"
    # agenda reads no sent mail, so it cannot say whether the organizer is a stranger.
    assert "stranger_invite" not in {w["code"] for w in invite["warnings"]}
    assert {"code": "not_responded"} in invite["why"]
    assert invite["start"] == {
        "utc": "2026-11-26T13:30:00Z",
        "local": "2026-11-26T13:30:00+00:00",
        "zone": "UTC",
    }


async def test_a_cancelled_or_answered_invitation_is_not_open(fake):
    box = mailbox()
    box.events["cal-2"][0]["responseStatus"] = {"response": "tentativelyAccepted"}
    assert section(await run(fake, box), "invites")["items"] == []
    box = mailbox()
    box.events["cal-2"][0]["isCancelled"] = True
    fake.routes.clear()
    assert section(await run(fake, box), "invites")["items"] == []


# ── Tasks ─────────────────────────────────────────────────────────────────


async def test_tasks_due_in_the_window_and_overdue_ones(fake):
    env = await run(fake)
    assert ids_in(env, "tasks_due") == ["t-late", "t-today"]
    items = {i["id"]: i for i in section(env, "tasks_due")["items"]}
    assert {"code": "overdue", "due": "2026-11-20"} in items["t-late"]["why"]
    assert {"code": "due_today"} in items["t-today"]["why"]
    assert items["t-today"]["list_id"] == "list-1"
    assert items["t-today"]["importance"] == "high"


async def test_a_window_without_today_leaves_overdue_tasks_out(fake):
    env = await run(fake, window="2026-12-01/2026-12-20")
    assert ids_in(env, "tasks_due") == ["t-later"]
    assert {"code": "due_in_window"} in section(env, "tasks_due")["items"][0]["why"]


async def test_tasks_can_be_left_out_and_then_cost_nothing(fake):
    env = await run(fake, include_tasks=False)
    assert "tasks_due" not in [s["name"] for s in env["sections"]]
    assert not [s for s in fake.operations if "/todo/" in s.path]


async def test_p_h1a_a_refused_task_filter_falls_back_to_a_local_one(fake):
    box = mailbox()
    box.task_filter_supported = False
    env = await run(fake, box)
    assert ids_in(env, "tasks_due") == ["t-late", "t-today"]


# ── Calendars ─────────────────────────────────────────────────────────────


async def test_calendars_by_name_or_id(fake):
    env = await run(fake, calendars=["shared"])
    assert [c["id"] for c in env["understood"]["calendars"]] == ["cal-2"]
    assert "ev-a" not in [e["id"] for e in day(env, "2026-11-26")["events"]]
    fake.routes.clear()
    env = await run(fake, calendars=["cal-1"])
    assert [c["id"] for c in env["understood"]["calendars"]] == ["cal-1"]


async def test_an_unknown_calendar_is_refused(fake):
    with pytest.raises(ValueError, match="No calendar"):
        await run(fake, calendars=["Holidays"])


async def test_more_than_five_calendars_are_reported_not_read(fake):
    box = mailbox()
    box.calendars += [
        {"id": f"cal-{n}", "name": f"Extra {n}", "isDefaultCalendar": False} for n in range(3, 8)
    ]
    env = await run(fake, box)
    assert len(env["understood"]["calendars"]) == 5
    assert env["understood"]["calendars_beyond_5"] == 2
    assert "calendars_beyond_5" in env["understood"]["not_covered"]
    assert env["partial"] is True


async def test_the_zone_reads_the_window_and_shows_the_times(fake):
    env = await run(fake, zone="Asia/Kolkata")
    assert env["understood"]["zone"] == "Asia/Kolkata"
    assert env["understood"]["window"]["start"] == "2026-11-25T18:30:00Z"
    (invite,) = section(env, "invites")["items"]
    assert invite["start"]["local"] == "2026-11-26T19:00:00+05:30"
    assert invite["start"]["zone"] == "Asia/Kolkata"


# ── Failures, budget and requests ─────────────────────────────────────────


async def test_a_failed_calendar_read_fails_its_sections_and_keeps_the_tasks(fake):
    box = mailbox()
    box.install(fake)
    fake.on(r"/me/calendars/cal-2/calendarView", lambda s, m: error(503, "ServiceUnavailable"))
    env = await read_agenda(
        fake.graph_client(), config_zone="UTC", loaded_tools=frozenset(), now=NOW
    )
    assert env["outcome"] == "incomplete"
    assert {s["name"]: s["status"] for s in env["sections"]} == {
        "days": "failed",
        "overlaps": "failed",
        "invites": "failed",
        "tasks_due": "ok",
    }
    assert {f["section"] for f in env["failed_sections"]} == {"days", "overlaps", "invites"}
    assert env["partial"] is True


async def test_every_section_failing_raises(fake):
    from outlook_mcp.errors import GraphAPIError

    mailbox().install(fake)
    fake.on(r"/me/calendars/cal-1/calendarView", lambda s, m: error(503, "ServiceUnavailable"))
    with pytest.raises(GraphAPIError):
        await read_agenda(
            fake.graph_client(),
            config_zone="UTC",
            loaded_tools=frozenset(),
            now=NOW,
            include_tasks=False,
        )


async def test_a_spent_budget_is_incomplete_not_empty(fake, monkeypatch):
    monkeypatch.setitem(budget.CEILINGS, "outlook_agenda", 2)
    env = await run(fake)
    assert env["outcome"] == "incomplete"
    statuses = {s["name"]: s["status"] for s in env["sections"]}
    assert "skipped_budget" in statuses.values()
    assert env["graph_requests"] <= 2


def worst_case() -> Mailbox:
    """Seven calendars of 301 events each, and nine task lists of 101 open tasks."""
    calendars = [
        {"id": f"cal-{n}", "name": f"Calendar {n}", "isDefaultCalendar": n == 0} for n in range(7)
    ]
    events = {
        c["id"]: [
            ev(f"{c['id']}-{n}", "2026-11-26T09:00:00", "2026-11-26T09:30:00") for n in range(301)
        ]
        for c in calendars
    }
    lists = [
        {"id": f"list-{n}", "displayName": f"List {n}", "wellknownListName": "none"}
        for n in range(9)
    ]
    tasks = {
        lst["id"]: [task(f"{lst['id']}-{n}", "t", "2026-11-26") for n in range(101)]
        for lst in lists
    }
    return Mailbox(calendars=calendars, events=events, task_lists=lists, tasks=tasks)


async def test_the_worst_case_stays_within_twenty_five_requests(fake):
    env = await run(fake, worst_case(), window="this_week")
    assert fake.counted <= 25
    assert env["graph_requests"] == fake.counted
    assert env["graph_requests"] <= budget.CEILINGS["outlook_agenda"] == 30
    not_covered = env["understood"]["not_covered"]
    assert "calendars_beyond_5" in not_covered
    assert "events_beyond_300" in not_covered
    assert "task_lists_beyond_8" in not_covered
    assert env["partial"] is True


async def test_the_calendar_view_asks_for_what_the_contract_names(fake):
    await run(fake)
    views = [s for s in fake.operations if s.path.endswith("/calendarView")]
    assert {s.path for s in views} == {
        "/me/calendars/cal-1/calendarView",
        "/me/calendars/cal-2/calendarView",
    }
    for view in views:
        assert view.params["$select"] == agenda.EVENT_SELECT
        assert view.params["$top"] == "100"
        # A day wider on each side, so an all-day event, which Graph shows as a floating
        # date, is never missed at the edge; what is shown is decided locally.
        assert view.params["startDateTime"] == "2026-11-25T00:00:00Z"
        assert view.params["endDateTime"] == "2026-11-28T00:00:00Z"
    assert "iCalUId" in agenda.EVENT_SELECT and "location" in agenda.EVENT_SELECT


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
    env = await run(fake, mailbox(MARK), window="this_week")
    found = list(_marker_paths(env))
    assert found
    for path in found:
        assert "others_wrote" in path, f"Graph text outside others_wrote at {path}"
        assert "why" not in path and "warnings" not in path and path[0] != "next"
    # Calendar names are their owners' words too.
    assert any(p[:2] == ("understood", "calendars") for p in found)


async def test_codes_are_in_the_vocabulary(fake):
    env = await run(fake, window="this_week")
    text = json.dumps(env)
    for item in [i for s in env["sections"] for i in s["items"]]:
        for nested in [item, *item.get("events", [])]:
            assert {r["code"] for r in nested["why"]} <= WHY_CODES
            assert {w["code"] for w in nested["warnings"]} <= WARNING_CODES
    assert "ZQX" not in text


async def test_next_briefs_the_next_meeting_when_brief_is_loaded(fake):
    env = await run(fake)
    assert env["next"] == [
        {
            "template": "brief_meeting",
            "tool": "outlook_brief",
            "args": {"meeting": "ev-a"},
            "needs": [],
        }
    ]
    fake.routes.clear()
    env = await run(fake, tools=frozenset())
    assert env["next"] == []


async def test_next_answers_the_first_invite_when_answer_invites_is_loaded(fake):
    env = await run(fake, tools=frozenset({"outlook_answer_invites"}))
    assert env["next"] == [
        {
            "template": "answer",
            "tool": "outlook_answer_invites",
            "args": {"event_ids": ["ev-b"]},
            "needs": ["answer"],
        }
    ]


async def test_an_empty_calendar_answers_with_empty_days(fake):
    box = Mailbox(calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}])
    env = await run(fake, box)
    assert env["outcome"] == "answered"
    assert day(env, "2026-11-26")["events"] == []
    assert section(env, "overlaps")["items"] == [] and section(env, "invites")["items"] == []


async def test_a_calendar_list_of_one_page_is_one_request(fake):
    """A1 is the calendar list; the fake answers it in one page as Graph does for a few."""
    box = mailbox()
    box.install(fake)
    fake.on("/me/calendars", lambda s, m: page(box.calendars))
    await read_agenda(fake.graph_client(), config_zone="UTC", loaded_tools=frozenset(), now=NOW)
    assert len([s for s in fake.operations if s.path == "/me/calendars"]) == 1
