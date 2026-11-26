"""``outlook_brief`` (contracts §12.5): what's going on with a person, a meeting or a topic.

Pinned here: each path's sections and caps, an ambiguous person or meeting
answering ``candidates`` with nothing gathered, each path's request count, the
rule that links a task to a person, the owner's mail decided by Sent Items, and
where other people's text may appear, the contact's details included.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import NoMatchError
from outlook_mcp.intents import brief, budget, ids, transport
from outlook_mcp.intents.brief import read_brief
from outlook_mcp.intents.envelope import WARNING_CODES, WHY_CODES
from outlook_mcp.intents.ids import IdFamily
from tests.intent_fakes import (
    ME,
    SENT_FOLDER,
    FakeGraph,
    Mailbox,
    addr,
    auth_header,
    contact,
    event,
    letter,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
MARK = " ZQXMARK "
SAM = "sam@example.com"
NEW = "newbie@example.org"


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def sent_row(msg_id, to, when, *, conversation, subject="Letter") -> dict:
    """A letter in Sent Items with an id, as the mail queries select it."""
    return {
        **letter([to], sent=when, conversation=conversation),
        "id": msg_id,
        "subject": subject,
        "bodyPreview": f"preview of {subject}",
        "receivedDateTime": when,
        "hasAttachments": False,
        "isRead": True,
        "flag": {"flagStatus": "notFlagged"},
    }


def meeting_row(m: str = "", *, attachments: bool = True) -> dict:
    row = event(
        "ev-1",
        f"Budget sync{m}",
        "2026-11-27T17:00:00",
        "2026-11-27T18:00:00",
        attendees=[(SAM, f"Sam Jones{m}"), (NEW, f"New Person{m}")],
    )
    row["attendees"][0]["status"] = {"response": "accepted"}
    row["attendees"][1]["status"] = {"response": "none"}
    row["attendees"][1]["type"] = "optional"
    row["bodyPreview"] = f"Agenda: numbers{m}"
    row["hasAttachments"] = attachments
    row["location"] = {"displayName": f"Room 4{m}"}
    return row


def mailbox(m: str = "") -> Mailbox:
    """Sam: written to (a letter on 11-25, unanswered), saved as a contact, wrote twice
    this month (one with a file, one still waiting on you). A meeting with Sam and
    someone new. Tasks linked to Sam's mail and naming Sam."""
    return Mailbox(
        sent=[
            sent_row(
                "s1",
                (SAM, f"Sam Jones{m}"),
                "2026-11-25T10:00:00Z",
                conversation="conv-c",
                subject=f"Proposal{m}",
            ),
            sent_row(
                "s-old",
                (SAM, f"Sam Jones{m}"),
                "2026-09-01T10:00:00Z",
                conversation="conv-old",
                subject=f"Old{m}",
            ),
        ],
        received=[
            received(
                "r1",
                SAM,
                f"Plan{m}",
                name=f"Sam Jones{m}",
                conversation="conv-a",
                when="2026-11-20T09:00:00Z",
                attachments=True,
                headers=auth_header("pass", "example.com"),
            ),
            received(
                "r2",
                SAM,
                f"Budget sync notes{m}",
                name=f"Sam Jones{m}",
                conversation="conv-b",
                when="2026-11-24T09:00:00Z",
                headers=auth_header("pass", "example.com"),
            ),
            received(
                "r3",
                "other@example.net",
                f"Budget sync{m}",
                conversation="conv-x",
                when="2026-11-23T09:00:00Z",
                to=[ME],
            ),
            received(
                "forged", ME, f"Re: Plan{m}", conversation="conv-a", when="2026-11-21T09:00:00Z"
            ),
        ],
        contacts=[
            {
                **contact("c-sam", f"Sam Jones{m}", SAM),
                "companyName": f"Acme{m}",
                "jobTitle": f"Buyer{m}",
                "businessPhones": ["+1 555 0100"],
                "birthday": "1990-05-01T11:59:00Z",
            }
        ],
        attachments={
            "r1": [
                {
                    "id": "att-1",
                    "name": f"plan{m}.pdf",
                    "size": 1000,
                    "contentType": "application/pdf",
                    "isInline": False,
                }
            ],
            "ev-1": [
                {
                    "id": "att-e",
                    "name": f"agenda{m}.docx",
                    "size": 500,
                    "contentType": "application/msword",
                    "isInline": False,
                }
            ],
        },
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                meeting_row(m),
                event("ev-2", f"Dentist{m}", "2026-11-28T09:00:00", "2026-11-28T10:00:00"),
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
                    "title": f"Follow up the plan{m}",
                    "status": "notStarted",
                    "linkedResources": [{"externalId": "r1", "applicationName": "Outlook"}],
                },
                {"id": "t2", "title": f"Call Sam Jones{m}", "status": "notStarted"},
                {"id": "t3", "title": f"Pay the invoice{m}", "status": "notStarted"},
                {"id": "t4", "title": f"Sam Jones lunch{m}", "status": "completed"},
            ],
        },
    )


async def run(fake, box=None, *, tools=frozenset(), **args):
    (box or mailbox()).install(fake)
    return await read_brief(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=tools, now=NOW, **args
    )


def section(env, name) -> dict:
    return next(s for s in env["sections"] if s["name"] == name)


def ids_in(env, name) -> list:
    return [i["id"] for i in section(env, name)["items"]]


# ── Arguments ─────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "args",
    [
        {},
        {"person": "sam", "topic": "budget"},
        {"person": "sam", "meeting": "sync"},
        {"person": ""},
        {"topic": "x" * 201},
        {"topic": "AND OR"},
        {"person": "sam", "window": "366d"},
        {"person": "sam", "window": "+3d"},
        {"meeting": "a\nb"},
    ],
)
async def test_bad_arguments_are_refused_before_any_request(fake, args):
    with pytest.raises(ValueError):
        await run(fake, **args)
    assert fake.counted == 0


# ── Person ────────────────────────────────────────────────────────────────


async def test_a_person_brief_has_its_sections(fake):
    env = await run(fake, person="sam")
    assert env["tool"] == "outlook_brief"
    assert env["outcome"] == "answered"
    assert [s["name"] for s in env["sections"]] == list(brief.PERSON_SECTIONS)
    assert env["understood"]["person"] == SAM
    (who,) = section(env, "who")["items"]
    assert who["address"] == SAM and who["contact_id"] == "c-sam"
    assert who["others_wrote"]["company"] == "Acme"
    assert who["others_wrote"]["job_title"] == "Buyer"
    assert who["others_wrote"]["phones"] == "+1 555 0100"
    assert who["others_wrote"]["birthday"] == "1990-05-01"


async def test_recent_mail_both_ways_within_the_window(fake):
    box = mailbox()
    # P-SE1: KQL `recipients:` may match more than the exact address; only letters
    # that really list it count.
    box.sent.append(
        sent_row(
            "decoy", "sam@example.com.example", "2026-11-25T11:00:00Z", conversation="conv-decoy"
        )
    )
    env = await run(fake, box, person=SAM)
    assert env["understood"]["matched"] == "exact"
    # Newest first; the September letter is outside the 30-day window.
    assert ids_in(env, "recent") == ["s1", "r2", "r1"]
    recent = {i["id"]: i for i in section(env, "recent")["items"]}
    assert recent["s1"]["from"]["known"] is True
    assert recent["r2"]["from"]["address"] == SAM


async def test_open_loops_from_the_mail_both_ways(fake):
    env = await run(fake, person="sam")
    loops = {i["id"]: i for i in section(env, "open_loops")["items"]}
    # Oldest wait first: Sam's plan (11-20) and notes (11-24) wait on you, your
    # proposal (11-25) on Sam.
    assert list(loops) == ["conv-a", "conv-b", "conv-c"]
    assert loops["conv-a"]["waits_on"] == "me"
    assert loops["conv-b"]["waits_on"] == "me"
    assert loops["conv-c"]["waits_on"] == "them"
    assert loops["conv-c"]["waiting_for"] == [SAM]


async def test_a_forged_message_from_the_owner_is_no_ones_reply(fake):
    """conv-a's latest is a forged "from me" in the Inbox: the conversation is Sam's
    message, still waiting, never the owner's."""
    box = mailbox()
    box.received[3]["toRecipients"] = [addr(SAM)]
    env = await run(fake, box, person="sam", window="60d")
    loops = {i["id"]: i for i in section(env, "open_loops")["items"]}
    assert "conv-a" not in loops or loops["conv-a"]["waits_on"] == "me"


async def test_upcoming_meetings_with_them(fake):
    env = await run(fake, person="sam")
    (meeting,) = section(env, "upcoming")["items"]
    assert meeting["id"] == "ev-1"
    assert {"code": "meeting_with_them"} in meeting["why"]


async def test_files_they_sent(fake):
    env = await run(fake, person="sam")
    (found,) = section(env, "files")["items"]
    assert found["id"] == "att-1" and found["message_id"] == "r1"
    assert found["size"] == 1000 and found["content_type"] == "application/pdf"


async def test_tasks_linked_to_their_mail_or_naming_them(fake):
    env = await run(fake, person="sam")
    tasks = {i["id"]: i for i in section(env, "tasks")["items"]}
    assert set(tasks) == {"t1", "t2"}
    assert tasks["t1"]["linked_message_id"] == "r1"
    assert {"code": "linked_to_message"} in tasks["t1"]["why"]
    task_reads = [s for s in fake.operations if s.path.endswith("/tasks")]
    assert all(s.params.get("$expand") == "linkedResources" for s in task_reads)


async def test_the_person_queries(fake):
    await run(fake, person="sam")
    assert (
        "receivedDateTime ge 2026-10-27T12:00:00Z and from/emailAddress/address eq "
        "'sam@example.com'"
    ) in fake.filters()
    assert '"recipients:sam@example.com"' in fake.searches()
    view = next(s for s in fake.operations if s.path == "/me/calendarView")
    assert view.params["startDateTime"] == "2026-11-26T12:00:00Z"
    assert view.params["endDateTime"] == "2026-12-26T12:00:00Z"


async def test_an_ambiguous_person_returns_candidates_and_gathers_nothing(fake):
    box = mailbox()
    box.sent.append(
        sent_row(
            "s9",
            ("sam.other@example.net", "Sam Other"),
            "2026-11-01T10:00:00Z",
            conversation="conv-z",
        )
    )
    env = await run(fake, box, person="sam", tools=frozenset({"outlook_brief"}))
    assert env["outcome"] == "candidates"
    assert [s["name"] for s in env["sections"]] == ["candidates"]
    assert not [s for s in fake.operations if s.path == "/me/calendarView"]
    assert not [s for s in fake.operations if "/todo/" in s.path]
    assert {o["template"] for o in env["next"]} == {"choose"}
    assert {o["args"]["person"] for o in env["next"]} == {SAM, "sam.other@example.net"}


async def test_a_person_nobody_matches_raises(fake):
    with pytest.raises(NoMatchError):
        await run(fake, person="nobody-at-all")


async def test_the_person_path_stays_within_thirty_eight(fake):
    box = mailbox()
    box.task_lists += [
        {"id": f"l{n}", "displayName": f"L{n}", "wellknownListName": "none"} for n in range(8)
    ]
    for n in range(7):
        box.received.append(
            received(
                f"f{n}",
                SAM,
                "File",
                conversation=f"cf{n}",
                when="2026-11-22T09:00:00Z",
                attachments=True,
            )
        )
    env = await run(fake, box, person="sam")
    assert env["graph_requests"] == fake.counted <= 38
    attachment_reads = [s for s in fake.operations if s.path.endswith("/attachments")]
    assert len(attachment_reads) == brief.PERSON_FILES
    tasks = [s for s in fake.operations if s.path.endswith("/tasks")]
    assert len(tasks) == brief.PERSON_TASK_LISTS
    # The Junk and Deleted Items ids, and the sender checks of what `next` may name.
    folders = [s for s in fake.operations if s.path.startswith("/me/mailFolders/") and s.in_batch]
    assert {s.path for s in folders} == {
        "/me/mailFolders/junkemail",
        "/me/mailFolders/deleteditems",
    }


# ── Meeting ───────────────────────────────────────────────────────────────


async def test_a_meeting_brief_has_its_sections(fake):
    env = await run(
        fake, meeting="budget sync", tools=frozenset({"outlook_brief", "outlook_read_message"})
    )
    assert env["outcome"] == "answered"
    assert [s["name"] for s in env["sections"]] == list(brief.MEETING_SECTIONS)
    (meeting,) = section(env, "meeting")["items"]
    assert meeting["id"] == "ev-1"
    assert meeting["others_wrote"]["body_preview"] == "Agenda: numbers"
    # The search built from the subject ran, and is not echoed as if it were the agent's.
    assert '"Budget sync"' in fake.searches()
    assert '"Budget sync"' not in env["understood"]["queries"]
    assert env["understood"]["queries_not_echoed"] == 1
    attendees = {i["address"]: i for i in section(env, "attendees")["items"]}
    assert attendees[SAM]["response"] == "accepted" and attendees[SAM]["type"] == "required"
    assert attendees[SAM]["written_to"]["count"] == 2
    assert attendees[NEW]["written_to"] is None
    assert attendees[NEW]["type"] == "optional"
    assert {"code": "first_time"} in attendees[NEW]["warnings"]
    # Mail about it among its people: Sam's notes, not the stranger's "Budget sync".
    assert ids_in(env, "recent") == ["r2"]
    files = {i["id"] for i in section(env, "files")["items"]}
    assert files == {"att-e"}
    assert ("brief_person", {"person": NEW}) in [(o["template"], o["args"]) for o in env["next"]]
    assert ("open", {"message_id": "r2"}) in [(o["template"], o["args"]) for o in env["next"]]


async def test_a_meeting_by_its_id_is_exact(fake):
    long_id = "AAMkAD" + "e" * 40
    box = mailbox()
    box.events["cal-1"][0]["id"] = long_id
    box.attachments[long_id] = box.attachments.pop("ev-1")
    env = await run(fake, box, meeting=long_id)
    assert env["outcome"] == "answered"
    assert env["understood"]["matched"] == "exact"
    assert ids_in(env, "meeting") == [long_id]
    # An exact id is read directly: no calendar is listed or viewed to find it.
    assert not [s for s in fake.operations if s.path.endswith("/calendarView")]


async def test_an_ambiguous_meeting_returns_candidates(fake):
    box = mailbox()
    box.events["cal-1"].append(
        event("ev-3", "Budget sync again", "2026-11-29T17:00:00", "2026-11-29T18:00:00")
    )
    env = await run(fake, box, meeting="budget sync")
    assert env["outcome"] == "candidates"
    assert {i["id"] for i in env["sections"][0]["items"]} == {"ev-1", "ev-3"}
    assert not [s for s in fake.operations if s.path.startswith("/me/events/")]


async def test_the_meeting_path_stays_within_twenty_five(fake):
    box = mailbox()
    box.calendars += [
        {"id": f"c{n}", "name": f"C{n}", "isDefaultCalendar": False} for n in range(5)
    ]
    box.events.update(
        {
            f"c{n}": [
                event(f"x{n}-{k}", "Other", "2026-11-27T09:00:00", "2026-11-27T09:30:00")
                for k in range(301)
            ]
            for n in range(5)
        }
    )
    env = await run(fake, box, meeting="budget sync")
    assert env["graph_requests"] == fake.counted <= 25


# ── Topic ─────────────────────────────────────────────────────────────────


async def test_a_topic_brief_has_its_sections(fake):
    box = mailbox()
    box.events["cal-1"].append(
        event("ev-inv", "Invoice review", "2026-11-30T09:00:00", "2026-11-30T10:00:00")
    )
    box.received.append(
        received(
            "inv",
            "billing@example.net",
            "Invoice 42",
            when="2026-11-25T09:00:00Z",
            attachments=True,
        )
    )
    box.attachments["inv"] = [
        {
            "id": "att-inv",
            "name": "invoice.pdf",
            "size": 9,
            "contentType": "application/pdf",
            "isInline": False,
        }
    ]
    env = await run(fake, box, topic="invoice")
    assert [s["name"] for s in env["sections"]] == list(brief.TOPIC_SECTIONS)
    assert ids_in(env, "mail") == ["inv"]
    mail = section(env, "mail")
    assert mail["total_exact"] is False
    assert {"code": "search_match"} in mail["items"][0]["why"]
    assert ids_in(env, "events") == ["ev-inv"]
    assert ids_in(env, "files") == ["att-inv"]
    assert ids_in(env, "tasks") == ["t3"]
    assert env["understood"]["topic"] == "invoice"


async def test_a_topic_nothing_matches_is_an_empty_answer(fake):
    env = await run(fake, topic="zebra")
    assert env["outcome"] == "answered"
    assert all(s["items"] == [] for s in env["sections"])


async def test_the_topic_path_stays_within_twenty_two(fake):
    box = mailbox()
    box.task_lists += [
        {"id": f"l{n}", "displayName": f"L{n}", "wellknownListName": "none"} for n in range(8)
    ]
    box.calendars += [
        {"id": f"c{n}", "name": f"C{n}", "isDefaultCalendar": False} for n in range(6)
    ]
    for n in range(20):
        box.received.append(
            received(
                f"t{n}",
                "x@example.net",
                "zebra stripes",
                when="2026-11-25T09:00:00Z",
                attachments=True,
            )
        )
    env = await run(fake, box, topic="zebra")
    assert env["graph_requests"] == fake.counted <= 22
    assert "calendars_beyond_5" in env["understood"]["not_covered"]


# ── Failures, text and vocabularies ───────────────────────────────────────


async def test_a_spent_budget_is_incomplete_not_empty(fake, monkeypatch):
    monkeypatch.setitem(budget.CEILINGS, "outlook_brief", 9)
    env = await run(fake, person="sam")
    assert env["outcome"] == "incomplete"
    assert "skipped_budget" in {s["status"] for s in env["sections"]}


def _marker_paths(node, path=()):
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _marker_paths(value, (*path, key))
    elif isinstance(node, list):
        for value in node:
            yield from _marker_paths(value, path)
    elif isinstance(node, str) and MARK.strip() in node:
        yield path


@pytest.mark.parametrize(
    "args", [{"person": "sam"}, {"meeting": "budget sync"}, {"topic": "budget"}]
)
async def test_text_from_graph_appears_only_under_others_wrote(fake, args):
    env = await run(fake, mailbox(MARK), **args)
    found = list(_marker_paths(env))
    assert found
    for path in found:
        assert "others_wrote" in path, f"Graph text outside others_wrote at {path}"
        assert "why" not in path and "warnings" not in path and path[0] != "next"


@pytest.mark.parametrize(
    "args", [{"person": "sam"}, {"meeting": "budget sync"}, {"topic": "budget"}]
)
async def test_codes_are_in_the_vocabulary(fake, args):
    env = await run(fake, **args)
    for item in [i for s in env["sections"] for i in s["items"]]:
        assert {r["code"] for r in item["why"]} <= WHY_CODES
        assert {w["code"] for w in item["warnings"]} <= WARNING_CODES
    assert "ZQX" not in json.dumps(env)


async def test_next_drafts_a_reply_and_opens_the_latest(fake):
    tools = frozenset({"outlook_compose", "outlook_read_message"})
    env = await run(fake, person="sam", tools=tools)
    assert [(o["template"], o["args"]) for o in env["next"]] == [
        ("draft_reply", {"mode": "reply", "message_id": "r2"}),
        ("open", {"message_id": "s1"}),
    ]
    assert SENT_FOLDER  # the owner's letter is opened as theirs
