"""``outlook_find``'s implementation (contracts §12.1), before it is registered as a tool.

``find`` is the resolver behind a tool's arguments: it checks every argument
before a request goes out, refuses one that doesn't apply to the kind, and
answers in the envelope with one section, ``matches``. These pin the argument
matrix, the envelope for every kind, where other people's text may appear, the
suggested next calls, the per-kind request counts, and the conservative choices
made where Graph's behaviour on a personal mailbox is not yet probed.
"""

import json
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import budget, find, ids, transport
from outlook_mcp.intents.envelope import WARNING_CODES, WHY_CODES
from outlook_mcp.intents.find import APPLIES_TO, find_items
from outlook_mcp.intents.ids import IdFamily
from tests.intent_fakes import (
    ME,
    SENT_FOLDER,
    FakeGraph,
    Mailbox,
    auth_header,
    contact,
    error,
    event,
    letter,
    meeting_response,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
LONG_ID = "AAMkAD" + "q" * 40
ALL_TOOLS = frozenset({"outlook_read_message", "outlook_list_thread", "outlook_brief"})


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def mailbox(marker: str = "") -> Mailbox:
    """A mailbox with something of every kind. ``marker`` goes into every free-text
    field Graph returns, for the others_wrote test."""
    m = marker
    return Mailbox(
        sent=[
            letter([("sam.jones@example.com", f"Sam {m}Jones")], sent="2026-11-20T10:00:00Z"),
            letter(["sam.jones@example.com"], sent="2026-11-10T10:00:00Z"),
            letter([("pat@example.org", f"Pat {m}Lee")], sent="2026-10-01T10:00:00Z"),
        ],
        contacts=[contact("c1", f"Sammy {m}Davis", "sammy@example.net")],
        received=[
            received(
                "m1",
                "sam.jones@example.com",
                f"Budget review {m}",
                name=f"Sam {m}Jones",
                conversation="conv-budget",
                attachments=True,
                headers=auth_header("pass", "example.com"),
                when="2026-11-21T09:00:00Z",
            ),
            received(
                "m2",
                "stranger@unknown.example",
                f"Budget offer {m}",
                name=f"Stranger {m}",
                conversation="conv-offer",
                when="2026-11-22T09:00:00Z",
            ),
            received(
                "m3",
                ME,
                f"Re: Budget review {m}",
                conversation="conv-budget",
                when="2026-11-23T09:00:00Z",
                to=["sam.jones@example.com"],
                folder=SENT_FOLDER,  # the owner's reply, in Sent Items
            ),
            received(
                "d1",
                ME,
                f"Budget draft {m}",
                draft=True,
                folder="drafts-id",
                to=["sam.jones@example.com"],
            ),
        ],
        attachments={
            "m1": [
                {
                    "id": "att-1",
                    "name": f"budget-2027 {m}.xlsx",
                    "size": 2048,
                    "contentType": "application/vnd.ms-excel",
                    "isInline": False,
                },
                {
                    "id": "att-2",
                    "name": "logo.png",
                    "size": 10,
                    "contentType": "image/png",
                    "isInline": True,
                },
            ]
        },
        calendars=[{"id": "cal-1", "name": f"Calendar {m}", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                event(
                    "ev-1",
                    f"Budget sync {m}",
                    "2026-11-27T17:00:00",
                    "2026-11-27T18:00:00",
                    attendees=[("sam.jones@example.com", f"Sam {m}Jones")],
                ),
                event(
                    "ev-2",
                    f"Budget pitch {m}",
                    "2026-11-28T17:00:00",
                    "2026-11-28T18:00:00",
                    organizer="stranger@unknown.example",
                    organizer_name=f"Stranger {m}",
                    is_organizer=False,
                    response="notResponded",
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
                    "id": "t1",
                    "title": f"File budget {m}",
                    "status": "notStarted",
                    "importance": "high",
                    "dueDateTime": {"dateTime": "2026-11-27T00:00:00.0000000", "timeZone": "UTC"},
                },
                {"id": "t2", "title": f"Old budget {m}", "status": "completed"},
            ],
            "list-f": [{"id": "tf", "title": f"Budget flagged {m}", "status": "notStarted"}],
        },
        rules=[
            {"id": "rule-1", "displayName": f"Budget mail {m}", "isEnabled": True, "sequence": 1}
        ],
        folders=[
            {"id": "f-proj", "displayName": "Projects", "childFolderCount": 1},
            {"id": "f-old", "displayName": "Old", "childFolderCount": 1},
        ],
        child_folders={
            "f-proj": [{"id": "f-budget-2", "displayName": "Budget", "childFolderCount": 0}],
            "f-old": [{"id": "f-budget-1", "displayName": "Budget", "childFolderCount": 0}],
        },
    )


async def run(fake, box=None, *, tools=ALL_TOOLS, **args):
    (box or mailbox()).install(fake)
    args.setdefault("words", "budget")
    return await find_items(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=tools, now=NOW, **args
    )


def ids_of(env) -> list[str]:
    return [i["id"] for i in env["sections"][0]["items"]]


# ── Arguments ─────────────────────────────────────────────────────────────


def test_the_argument_matrix_is_the_contracts():
    assert APPLIES_TO == {
        "sender": ("message", "thread", "file"),
        "after": ("message", "thread", "file", "draft"),
        "before": ("message", "thread", "file", "draft"),
        "window": ("event", "task"),
        "has_file": ("message", "thread"),
        "folder": ("message", "thread", "file"),
        "task_list": ("task",),
    }


_VALUES = {
    "sender": "sam",
    "after": "7d",
    "before": "now",
    "window": "+7d",
    "has_file": True,
    "folder": "inbox",
    "task_list": "Tasks",
}


@pytest.mark.parametrize(
    ("argument", "kind"),
    [(a, k) for a, kinds in APPLIES_TO.items() for k in find.KINDS if k not in kinds],
)
async def test_an_argument_that_does_not_apply_is_refused_before_any_request(fake, argument, kind):
    with pytest.raises(ValueError) as refused:
        await run(fake, kind=kind, **{argument: _VALUES[argument]})
    assert argument in str(refused.value)
    assert all(k in str(refused.value) for k in APPLIES_TO[argument])
    assert fake.seen == []


@pytest.mark.parametrize(
    "args",
    [
        {"kind": "people"},
        {"kind": "person", "words": ""},
        {"kind": "person", "words": "x" * 201},
        {"kind": "person", "words": "sam\x07"},
        {"kind": "person", "limit": 0},
        {"kind": "person", "limit": 11},
        {"kind": "message", "after": "last tuesday"},
        {"kind": "message", "after": "now", "before": "7d"},
        {"kind": "event", "window": "+401d"},
        {"kind": "event", "window": "tomorrow at 3"},
        {"kind": "message", "sender": "sam\nx"},
        {"kind": "message", "folder": " "},
    ],
)
async def test_bad_arguments_are_refused_before_any_request(fake, args):
    with pytest.raises(ValueError):
        await run(fake, **args)
    assert fake.seen == []


# ── The envelope, per kind ────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("kind", "words", "outcome", "matched"),
    [
        ("person", "jones", "single", ["person:sam.jones@example.com"]),
        ("message", "budget", "candidates", ["m3", "m1", "m2"]),
        ("thread", "budget", "candidates", ["conv-budget", "conv-offer"]),
        ("file", "budget", "single", ["att-1"]),
        ("event", "budget", "candidates", ["ev-1", "ev-2"]),
        ("task", "budget", "single", ["t1"]),
        ("draft", "budget", "single", ["d1"]),
        ("rule", "budget", "single", ["rule-1"]),
        ("folder", "projects", "exact", ["f-proj"]),
    ],
)
async def test_every_kind_returns_the_envelope(fake, kind, words, outcome, matched):
    env = await run(fake, kind=kind, words=words)
    assert env["tool"] == "outlook_find"
    assert env["v"] == 1
    assert env["outcome"] == outcome
    assert ids_of(env) == matched
    sec = env["sections"][0]
    assert sec["name"] == "matches" and sec["status"] == "ok" and sec["cap"] == 5
    assert sec["total"] == len(matched)
    assert env["understood"]["kind"] == kind
    assert env["understood"]["words"] == words
    assert isinstance(env["understood"]["queries"], list)
    assert env["graph_requests"] == fake.counted
    assert env["partial"] is False and env["failed_sections"] == []
    json.dumps(env)  # the result is plain JSON
    for item in sec["items"]:
        assert item["kind"] == kind
        assert {"id", "why", "warnings", "others_wrote"} <= set(item)
        assert {w["code"] for w in item["why"]} <= WHY_CODES
        assert {w["code"] for w in item["warnings"]} <= WARNING_CODES


async def test_a_message_item_has_its_fields(fake):
    env = await run(fake, kind="message", words="budget review")
    assert ids_of(env)[:2] == ["m3", "m1"]  # equal scores: the newer first
    first = env["sections"][0]["items"][1]
    assert first["from"] == {
        "address": "sam.jones@example.com",
        "known": True,
        "sender_auth": {"result": "verified", "from_domain": "example.com"},
    }
    assert first["to_me"] == "to"
    assert first["received"] == "2026-11-21T09:00:00Z"
    assert first["has_attachments"] is True
    assert first["others_wrote"]["subject"] == "Budget review "
    assert {"search_ranked", "index_lag_possible"} <= {w["code"] for w in first["warnings"]}


async def test_an_unknown_sender_is_never_a_single_message(fake):
    env = await run(fake, kind="message", words="offer")
    assert ids_of(env) == ["m2"]
    assert env["outcome"] == "candidates"
    assert env["sections"][0]["items"][0]["from"]["known"] is False


async def test_a_known_sender_with_one_message_is_single(fake):
    env = await run(fake, kind="message", words="budget", sender="sam.jones@example.com")
    assert ids_of(env) == ["m1"]
    assert env["outcome"] == "single"
    assert env["understood"]["filters"]["sender"] == "sam.jones@example.com"


async def test_a_thread_waits_on_them_when_the_owner_wrote_last(fake):
    env = await run(fake, kind="thread", words="budget review")
    thread = env["sections"][0]["items"][0]
    assert thread["conversation_id"] == "conv-budget"
    assert thread["last_message_id"] == "m3"
    assert thread["waits_on"] == "them"
    assert "you_wrote_in_thread" in {w["code"] for w in thread["why"]}


async def test_an_exact_thread_id_counts_its_messages(fake):
    box = mailbox()
    for r in box.received:
        if r["conversationId"] == "conv-budget":
            r["conversationId"] = LONG_ID
    env = await run(fake, box, kind="thread", words=LONG_ID)
    assert env["outcome"] == "exact"
    assert env["sections"][0]["items"][0]["message_count"] == 2


async def test_a_message_id_header_is_exact(fake):
    env = await run(fake, kind="message", words="<m2@mail.example.com>")
    assert env["outcome"] == "exact"
    assert ids_of(env) == ["m2"]
    assert fake.searches() == []


async def test_an_event_from_a_stranger_carries_the_warning(fake):
    env = await run(fake, kind="event", words="pitch")
    item = env["sections"][0]["items"][0]
    assert item["id"] == "ev-2"
    assert "stranger_invite" in {w["code"] for w in item["warnings"]}
    assert item["start"] == {
        "utc": "2026-11-28T17:00:00Z",
        "local": "2026-11-28T17:00:00+00:00",
        "zone": "UTC",
    }
    assert env["outcome"] == "candidates"
    assert env["next"] == []


async def test_the_event_window_defaults_and_is_echoed(fake):
    env = await run(fake, kind="event", words="budget")
    view = next(s for s in fake.operations if s.path.endswith("/calendarView"))
    assert view.params["startDateTime"] == "2026-10-27T12:00:00Z"
    assert view.params["endDateTime"] == "2027-05-25T12:00:00Z"
    assert env["understood"]["window"]["start"] == "2026-10-27T12:00:00Z"
    fake.routes.clear()
    env = await run(fake, kind="event", words="budget", window="2026-11-28")
    assert ids_of(env) == ["ev-2"]
    assert env["understood"]["window"] == {
        "value": "2026-11-28",
        "start": "2026-11-28T00:00:00Z",
        "end": "2026-11-29T00:00:00Z",
        "zone": "UTC",
        "start_local": "2026-11-28T00:00:00+00:00",
        "end_local": "2026-11-29T00:00:00+00:00",
    }


async def test_tasks_skip_the_flagged_list_and_completed_tasks(fake):
    env = await run(fake, kind="task", words="budget")
    assert ids_of(env) == ["t1"]
    task = env["sections"][0]["items"][0]
    assert task["due"] == "2026-11-27" and task["list_id"] == "list-1"
    assert task["importance"] == "high"
    assert not any("list-f" in s.path for s in fake.operations)
    fake.routes.clear()
    env = await run(fake, kind="task", words="budget", window="2026-12-01")
    assert ids_of(env) == []
    assert env["outcome"] == "nothing_found"


async def test_a_folder_name_that_resolves_is_exact_and_an_ambiguous_one_is_candidates(fake):
    env = await run(fake, kind="folder", words="Projects")
    assert env["outcome"] == "exact" and ids_of(env) == ["f-proj"]
    fake.routes.clear()
    env = await run(fake, kind="folder", words="inbox")
    assert env["outcome"] == "exact" and ids_of(env) == ["inbox"]
    fake.routes.clear()
    env = await run(fake, kind="folder", words="Budget")
    assert env["outcome"] == "candidates"
    assert [i["path"] for i in env["sections"][0]["items"]] == ["Old/Budget", "Projects/Budget"]
    assert ids_of(env) == ["f-budget-1", "f-budget-2"]


async def test_rules_a_mailbox_refuses_are_not_supported(fake):
    box = mailbox()
    box.rules = error(403, "ErrorAccessDenied", "Access is denied.")
    env = await run(fake, box, kind="rule")
    assert env["outcome"] == "incomplete"  # not searched is not "nothing there"
    assert env["understood"]["not_searched"] == ["Access is denied."]
    assert env["partial"] is True
    assert env["sections"][0]["status"] == "failed"
    assert env["failed_sections"] == [
        {
            "section": "matches",
            "reason": "not_supported",
            "code": "ErrorAccessDenied",
            "message": "Access is denied.",
        }
    ]


async def test_nothing_found_is_an_answer_not_an_error(fake):
    env = await run(fake, kind="person", words="nobody")
    assert env["outcome"] == "nothing_found"
    assert ids_of(env) == []


async def test_a_graph_failure_is_raised(fake):
    box = mailbox()
    box.install(fake)
    fake.on("/me/messages", lambda seen, m: error(500, "ServerError", "boom"))
    with pytest.raises(GraphAPIError, match="boom"):
        await find_items(
            fake.graph_client(),
            kind="message",
            words="budget",
            zone="UTC",
            signed_in=ME,
            loaded_tools=ALL_TOOLS,
            now=NOW,
        )


async def test_a_spent_budget_is_reported(fake, monkeypatch):
    monkeypatch.setitem(budget.CEILINGS, "outlook_find", 2)
    env = await run(fake, kind="person", words="sam")
    assert env["outcome"] == "incomplete"
    assert env["sections"][0]["status"] == "skipped_budget"
    assert env["failed_sections"][0]["reason"] == "budget"
    assert env["graph_requests"] <= 2


async def test_limit_caps_the_section_and_keeps_the_total(fake):
    env = await run(fake, kind="message", words="budget", limit=2)
    assert ids_of(env) == ["m3", "m1"]
    assert env["sections"][0]["total"] == 3
    assert env["sections"][0]["cap"] == 2


async def test_after_and_before_narrow_by_the_exact_instant(fake):
    env = await run(
        fake,
        kind="message",
        words="budget",
        after="2026-11-21T09:00:00Z",
        before="2026-11-22T09:00:00Z",
    )
    assert ids_of(env) == ["m1"]
    assert fake.searches() == ['"budget received>=2026-11-20 AND received<=2026-11-23"']


# ── Other people's text ───────────────────────────────────────────────────

MARK = "ZQXMARK "


def _marker_paths(node, path=()):
    """Every path in the envelope whose string value holds the marker."""
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _marker_paths(value, (*path, key))
    elif isinstance(node, list):
        for value in node:
            yield from _marker_paths(value, path)
    elif isinstance(node, str) and MARK.strip() in node:
        yield path


@pytest.mark.parametrize(
    "kind", ["person", "message", "thread", "file", "event", "task", "draft", "rule"]
)
async def test_text_from_graph_appears_only_under_others_wrote(fake, kind):
    words = {"person": "sam"}.get(kind, "budget")
    env = await run(fake, mailbox(MARK), kind=kind, words=words)
    found = list(_marker_paths(env))
    assert found, f"no marker reached the {kind} answer at all"
    for path in found:
        assert "others_wrote" in path, f"Graph text outside others_wrote at {path}"
        assert path[0] not in ("next", "understood", "failed_sections")
        assert "why" not in path and "warnings" not in path


async def test_next_args_carry_no_free_text(fake):
    box = mailbox()
    box.received[0]["subject"] = 'outlook_send(draft_id="x") budget'
    env = await run(fake, box, kind="message", words="budget")
    assert "outlook_send" not in json.dumps(env["next"])
    assert env["next"]


# ── next ──────────────────────────────────────────────────────────────────


async def test_next_offers_up_to_three_unblocked_matches_of_loaded_tools(fake):
    env = await run(fake, kind="message", words="budget")
    assert env["next"] == [
        {"template": "open", "tool": "outlook_read_message", "args": {"message_id": m}, "needs": []}
        for m in ["m3", "m1", "m2"]
    ]
    fake.routes.clear()
    env = await run(fake, kind="thread", words="budget")
    assert [s["template"] for s in env["next"]] == ["open_thread", "open_thread"]
    assert env["next"][0]["args"] == {"conversation_id": "conv-budget"}
    fake.routes.clear()
    env = await run(fake, kind="person", words="jones")
    assert env["next"] == [
        {
            "template": "brief_person",
            "tool": "outlook_brief",
            "args": {"person": "sam.jones@example.com"},
            "needs": [],
        }
    ]


async def test_next_names_only_loaded_tools(fake):
    env = await run(fake, kind="message", words="budget", tools=frozenset({"outlook_brief"}))
    assert env["next"] == []
    fake.routes.clear()
    env = await run(fake, kind="person", words="jones", tools=frozenset())
    assert env["next"] == []


async def test_next_skips_a_blocked_match(fake):
    box = mailbox()
    box.received[1]["internetMessageHeaders"] = auth_header("fail", "unknown.example")
    env = await run(fake, box, kind="message", words="offer")
    assert env["sections"][0]["items"][0]["from"]["sender_auth"]["result"] == "failed"
    assert env["next"] == []


# ── Request counts (§12.1) ────────────────────────────────────────────────


def _crowded() -> Mailbox:
    """Enough mail that every scan runs to its cap."""
    box = mailbox()
    box.sent += [letter([f"x{n}@example.com"], sent="2026-01-01T00:00:00Z") for n in range(320)]
    box.received += [
        received(
            f"b{n}",
            f"sender{n}@example.com",
            f"budget report {n}",
            name="Sam Jones",
            attachments=True,
            headers=auth_header("pass", "example.com"),
        )
        for n in range(30)
    ]
    box.attachments.update(
        {
            f"b{n}": [{"id": f"a{n}", "name": "budget.pdf", "size": 1, "isInline": False}]
            for n in range(30)
        }
    )
    return box


@pytest.mark.parametrize(
    ("args", "most"),
    [
        ({"kind": "person", "words": "sam"}, 22),
        # +1 per search since live round 1: its ids are translated (translateExchangeIds).
        ({"kind": "message", "words": "budget", "sender": "jones"}, 13),
        ({"kind": "thread", "words": "budget", "sender": "jones"}, 13),
        ({"kind": "file", "words": "budget", "sender": "jones"}, 23),
        ({"kind": "event", "words": "budget"}, 1 + 15 + 3),
        ({"kind": "task", "words": "budget"}, 1 + 8),
        ({"kind": "draft", "words": "budget"}, 2),
        ({"kind": "rule", "words": "budget"}, 1),
        ({"kind": "folder", "words": "inbox"}, 2),
    ],
)
async def test_each_kind_stays_within_its_worst_case(fake, args, most):
    env = await run(fake, _crowded(), **args)
    assert env["graph_requests"] == fake.counted
    assert env["graph_requests"] <= most
    assert env["graph_requests"] <= budget.CEILINGS["outlook_find"]


# ── Where Graph's behaviour is not yet probed ─────────────────────────────


async def test_p_f5_headers_missing_from_a_list_are_read_one_by_one(fake):
    box = mailbox()
    box.headers_on_lists = False
    box.received.append(
        received(
            "r9",
            "newbie@example.net",
            "hello",
            name="New Bie",
            headers=auth_header("pass", "example.net"),
        )
    )
    env = await run(fake, box, kind="person", words="newbie@example.net")
    person = env["sections"][0]["items"][0]
    assert person["sender_auth"] == {"result": "verified", "from_domain": "example.net"}
    assert any(s.path == "/me/messages/r9" for s in fake.operations)


async def test_p_at1_a_word_inside_a_file_name_is_found_only_by_the_local_match(fake):
    """Live, plain words find a file by the start of its name, never by a word inside
    it. Here the message is found by its subject, and the local match on its fetched
    attachment names finds the inner word."""
    box = mailbox()
    box.received.append(
        received("n1", "sam.jones@example.com", "Forecast for you", attachments=True)
    )
    box.attachments["n1"] = [
        {"id": "att-n", "name": "Q3-Forecast-2027.xlsx", "size": 9, "isInline": False}
    ]
    env = await run(fake, box, kind="file", words="forecast")
    assert ids_of(env) == ["att-n"]
    assert fake.searches() == ['"forecast hasAttachments:true"']


async def test_p_h1a_a_refused_task_filter_falls_back_to_a_local_one(fake):
    box = mailbox()
    box.task_filter_supported = False
    env = await run(fake, box, kind="task", words="budget")
    assert ids_of(env) == ["t1"]
    reads = [s for s in fake.operations if s.path == "/me/todo/lists/list-1/tasks"]
    assert [s.filter for s in reads] == ["status ne 'completed'", None]


async def test_p_se1_only_letters_that_list_the_exact_address_count(fake):
    box = Mailbox(
        sent=[letter([f"x{n}@example.com"], sent="2026-11-01T00:00:00Z") for n in range(300)]
        + [
            letter(["probe@example.com.example"], sent="2025-01-01T00:00:00Z"),
            letter(["pro.be@example.com"], sent="2025-01-02T00:00:00Z"),
        ]
    )
    env = await run(fake, box, kind="person", words="probe@example.com")
    person = env["sections"][0]["items"][0]
    assert person["written_to"] is None
    assert "first_time" in {w["code"] for w in person["warnings"]}
    assert any('"recipients:probe@example.com"' == s for s in fake.searches())

    fake.routes.clear()
    box.sent.append(letter(["probe@example.com"], sent="2024-06-01T00:00:00Z"))
    env = await run(fake, box, kind="person", words="probe@example.com")
    person = env["sections"][0]["items"][0]
    assert person["written_to"] == {"count": 1, "last": "2024-06-01T00:00:00Z", "basis": "search"}


async def test_p_si1_an_accepted_invitation_does_not_make_the_organizer_known(fake):
    box = Mailbox(sent=[meeting_response("organizer@example.org")])
    env = await run(fake, box, kind="person", words="organizer@example.org")
    person = env["sections"][0]["items"][0]
    assert env["outcome"] == "exact"
    assert person["written_to"] is None
    assert "first_time" in {w["code"] for w in person["warnings"]}


async def test_every_request_asks_for_immutable_ids(fake):
    await run(fake, kind="message", words="budget", sender="jones")
    assert fake.operations
    assert all(s.headers.get("prefer") == 'IdType="ImmutableId"' for s in fake.operations)
    assert ids.ID_FAMILY.family == "immutable"


async def test_the_public_entry_point_is_find_items():
    assert find.TOOL == "outlook_find"
    assert callable(find.find_items)
