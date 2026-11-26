"""Regressions from the independent review of PR #280, one test per finding.

Each started as the reviewer's repro, which asserted the faulty behaviour, and is
written here the right way round. The finding ids (B1, S1-S6, N1) are the review's.
"""

from datetime import datetime, timezone

import pytest

from outlook_mcp.intents import budget, ids, people, transport
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.find import find_items
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.intents.resolver import Filters, open_lookup, resolve
from tests.intent_fakes import (
    ME,
    SENT_FOLDER,
    FakeGraph,
    Mailbox,
    auth_header,
    error,
    event,
    letter,
    page,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
TOOLS = frozenset({"outlook_read_message", "outlook_list_thread", "outlook_brief"})


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


async def res(fake, box, kind, words, purpose="act", filters=None):
    box.install(fake)
    spent = RequestBudget(30)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
        return await resolve(lk, kind, words, filters or Filters(), purpose, 5)


async def find(fake, box, **args):
    box.install(fake)
    return await find_items(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=TOOLS, now=NOW, **args
    )


def codes(candidate):
    return [w["code"] for w in candidate.item["warnings"]]


def header_reads(fake, message_id):
    return [s for s in fake.operations if s.path == f"/me/messages/{message_id}"]


# ── B1: mail that claims to be from the owner ─────────────────────────────


async def test_b1_a_received_message_from_the_owners_address_is_checked_and_not_known(fake):
    box = Mailbox(
        sent=[letter([("sam@example.com", "Sam")])],
        received=[
            received(
                "spoof",
                ME,
                "Wire instructions",
                folder="junk-id",
                headers=auth_header("fail", "example.com"),
            )
        ],
    )
    resolution = await res(fake, box, "message", "wire instructions")
    candidate = resolution.candidates[0]
    assert resolution.outcome == "candidates"
    assert candidate.known is False
    assert candidate.item["from"]["known"] is False
    assert candidate.item["from"]["sender_auth"]["result"] == "failed"
    assert "sender_failed" in codes(candidate)
    # The check ran on the headers the search returned (P-F5), so no separate read.
    search = next(s for s in fake.operations if s.search)
    assert "internetMessageHeaders" in search.params["$select"]


async def test_b1_the_owners_real_sent_copy_is_known_by_its_folder(fake):
    box = Mailbox(
        sent=[letter([("sam@example.com", "Sam")])],
        received=[received("real", ME, "Wire instructions", folder=SENT_FOLDER)],
    )
    resolution = await res(fake, box, "message", "wire instructions")
    assert resolution.outcome == "single"
    assert resolution.candidates[0].item["from"]["known"] is True
    assert not header_reads(fake, "real"), "nothing to check on mail the owner sent"


async def test_b1_a_forged_reply_from_the_owner_does_not_make_a_thread_known(fake):
    box = Mailbox(
        sent=[letter([("sam@example.com", "Sam")])],
        received=[
            received(
                "a",
                "stranger@evil.example",
                "Offer",
                conversation="cv",
                when="2026-11-20T09:00:00Z",
            ),
            received(
                "b",
                ME,
                "Re: Offer",
                conversation="cv",
                when="2026-11-21T09:00:00Z",
                folder="inbox-id",
                headers=auth_header("fail", "example.com"),
            ),
        ],
    )
    resolution = await res(fake, box, "thread", "offer")
    candidate = resolution.candidates[0]
    assert resolution.outcome == "candidates"
    assert candidate.known is False
    assert candidate.item["waits_on"] == "me"
    assert "you_wrote_in_thread" not in [w["code"] for w in candidate.item["why"]]
    assert "sender_failed" in codes(candidate)


async def test_b1_an_event_organized_under_the_owners_address_but_not_by_them_is_a_stranger(fake):
    box = Mailbox(
        sent=[letter([("sam@example.com", "Sam")])],
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={
            "cal-1": [
                event(
                    "ev",
                    "Quarterly payout",
                    "2026-11-27T17:00:00",
                    "2026-11-27T18:00:00",
                    organizer=ME,
                    organizer_name="Me",
                    is_organizer=False,
                    response="notResponded",
                )
            ]
        },
    )
    resolution = await res(fake, box, "event", "payout")
    candidate = resolution.candidates[0]
    assert resolution.outcome == "candidates"
    assert "stranger_invite" in codes(candidate)
    assert "organizer_you" not in [w["code"] for w in candidate.item["why"]]
    assert candidate.item["is_organizer"] is False


# ── S1: names come from sent mail and contacts only ───────────────────────


async def test_s1_a_received_display_name_never_makes_a_written_to_address_single(fake):
    box = Mailbox(
        sent=[
            letter([("billing@vendor-mail.example", "Vendor Billing")], sent="2026-03-01T10:00:00Z")
        ],
        received=[
            received(
                "x1",
                "billing@vendor-mail.example",
                "Quick favour",
                name="Dana Whitfield",
                headers=auth_header("pass", "vendor-mail.example"),
            )
        ],
    )
    resolution = await res(fake, box, "person", "Dana Whitfield")
    candidate = resolution.candidates[0]
    assert resolution.outcome == "candidates"
    assert candidate.known is False
    assert "received_only" in codes(candidate)
    # Still found under the name the owner wrote to it with.
    fake.routes.clear()
    resolution = await res(fake, box, "person", "Vendor Billing")
    assert resolution.outcome == "single"


# ── S2: a sender check that didn't run ────────────────────────────────────


async def test_s2_a_failed_header_read_is_not_checked_and_never_single(fake):
    box = Mailbox(
        sent=[letter([("pat@paypal.com", "Pat Dev")])],
        received=[
            received(
                "spoof",
                "pat@paypal.com",
                "Invoice overdue",
                headers=auth_header("fail", "paypal.com"),
            )
        ],
    )
    # No headers on the search reply, so the single-message fallback read is made,
    # and that read fails.
    box.headers_on_lists = False
    box.install(fake)
    fake.on(r"/me/messages/(\S+)", lambda seen, m: error(503, "ServiceUnavailable"))
    spent = RequestBudget(30)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
        resolution = await resolve(lk, "message", "invoice overdue", Filters(), "act", 5)
    candidate = resolution.candidates[0]
    assert candidate.item["from"]["sender_auth"]["result"] == "not_checked"
    assert resolution.outcome == "candidates"


async def test_s2_a_person_whose_check_errored_is_never_single(fake):
    box = Mailbox(
        sent=[letter([("pat@paypal.com", "Pat Dev")])],
        received=[received("r1", "pat@paypal.com", "Hello", name="Pat Dev")],
    )
    box.install(fake)
    # Q6, the latest message with its headers, errors inside the batch.
    fake.on(
        "/me/messages",
        lambda seen, m: (
            error(503, "ServiceUnavailable")
            if "internetMessageHeaders" in (seen.params.get("$select") or "")
            else page([])
        ),
    )
    spent = RequestBudget(30)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
        resolution = await resolve(lk, "person", "pat dev", Filters(), "act", 5)
    assert resolution.candidates[0].item["sender_auth"]["result"] == "not_checked"
    assert resolution.outcome == "candidates"


async def test_s2_a_person_with_no_mail_to_check_can_still_be_single(fake):
    box = Mailbox(sent=[letter([("pat@example.org", "Pat Lee")])])
    resolution = await res(fake, box, "person", "pat lee")
    assert resolution.outcome == "single"


# ── S3: P-SI1, meeting responses that carry no type ───────────────────────


async def test_s3_an_untyped_accept_is_caught_by_its_subject(fake):
    accept = letter([("organizer@stranger.example", "Olga Stranger")], sent="2026-11-02T10:00:00Z")
    accept["subject"] = "Accepted: Free consultation"
    resolution = await res(fake, Mailbox(sent=[accept]), "person", "Olga Stranger")
    assert resolution.outcome != "single"
    assert not any(c.known for c in resolution.candidates)


async def test_s3_a_row_carrying_a_meeting_message_type_is_dropped():
    row = letter(["organizer@stranger.example"])
    row["meetingMessageType"] = "meetingAccepted"
    table = people.build_sent_table([row], signed_in=ME)
    assert table.entries == {}
    assert table.meeting_messages_dropped == 1


@pytest.mark.parametrize(
    "subject",
    ["Accepted: x", "Declined: x", "Tentative: x", "tentatively accepted: x", "Canceled: x"],
)
def test_s3_the_standard_response_prefixes_are_meeting_messages(subject):
    assert people.is_meeting_message({"subject": subject})


def test_s3_an_ordinary_subject_is_a_letter():
    assert not people.is_meeting_message({"subject": "Accept the offer?"})
    assert not people.is_meeting_message({"subject": None})


async def test_s3_the_sent_table_asks_for_the_type_and_falls_back_if_refused(fake):
    box = Mailbox(sent=[letter([("sam@example.com", "Sam")])], cast_select_supported=False)
    resolution = await res(fake, box, "person", "sam")
    q1 = [s for s in fake.operations if s.path == "/me/mailFolders/sentitems/messages"]
    assert "microsoft.graph.eventMessage/meetingMessageType" in q1[0].params["$select"]
    assert "meetingMessageType" not in q1[1].params["$select"]
    assert resolution.outcome == "single"


# ── S4: a search that didn't finish says so ───────────────────────────────


def _big_tree() -> dict:
    folders = [
        {"id": f"f{i}", "displayName": f"Client {i}", "childFolderCount": 1} for i in range(40)
    ]
    children = {
        f"f{i}": [{"id": f"c{i}", "displayName": f"Sub {i}", "childFolderCount": 0}]
        for i in range(40)
    }
    children["f39"] = [{"id": "target", "displayName": "Receipts", "childFolderCount": 0}]
    return {"folders": folders, "child_folders": children}


async def test_s4_a_folder_walk_that_runs_out_of_budget_is_incomplete(fake):
    env = await find(fake, Mailbox(**_big_tree()), kind="folder", words="Receipts")
    assert env["outcome"] == "incomplete"
    assert env["sections"][0]["status"] == "skipped_budget"
    assert env["partial"] is True
    assert env["understood"]["not_searched"]


async def test_s4_a_folder_filter_that_runs_out_of_budget_is_incomplete(fake):
    box = Mailbox(
        **_big_tree(), received=[received("m", "a@b.example", "Hotel receipt", folder="target")]
    )
    env = await find(fake, box, kind="message", words="receipt", folder="Receipts")
    assert env["outcome"] == "incomplete"


async def test_s4_tasks_past_the_cap_are_reported_as_not_searched(fake):
    tasks = [{"id": f"t{i}", "title": f"Chore {i}", "status": "notStarted"} for i in range(100)]
    tasks.append({"id": "t-last", "title": "Renew passport", "status": "notStarted"})
    box = Mailbox(
        task_lists=[{"id": "l1", "displayName": "Tasks", "wellknownListName": "defaultList"}],
        tasks={"l1": tasks},
    )
    env = await find(fake, box, kind="task", words="passport")
    section = env["sections"][0]
    assert env["outcome"] == "incomplete"
    assert env["partial"] is True
    assert section["total_exact"] is False
    assert any("l1" in line for line in env["understood"]["not_searched"])


async def test_s4_a_capped_scan_that_found_one_match_is_not_single(fake):
    tasks = [{"id": f"t{i}", "title": f"Chore {i}", "status": "notStarted"} for i in range(99)]
    tasks.append({"id": "t-hit", "title": "Renew passport", "status": "notStarted"})
    tasks.append({"id": "t-late", "title": "Passport photos", "status": "notStarted"})
    box = Mailbox(
        task_lists=[{"id": "l1", "displayName": "Tasks", "wellknownListName": "defaultList"}],
        tasks={"l1": tasks},
    )
    env = await find(fake, box, kind="task", words="passport")
    assert env["outcome"] == "candidates"
    assert env["sections"][0]["total_display"] == "1+"
    assert env["partial"] is True


async def test_s4_the_deadline_is_incomplete(fake, monkeypatch):
    monkeypatch.setattr(budget, "INTENT_DEADLINE_S", 0.05)
    fake.delays.append((lambda seen: True, 0.5))
    env = await find(
        fake, Mailbox(sent=[letter([("sam@example.com", "Sam")])]), kind="person", words="sam"
    )
    assert env["outcome"] == "incomplete"
    assert env["sections"][0]["status"] == "skipped_time"


async def test_s4_a_finished_search_that_matched_nothing_is_still_nothing_found(fake):
    env = await find(fake, Mailbox(), kind="person", words="nobody")
    assert env["outcome"] == "nothing_found"
    assert env["partial"] is False


# ── S5: the id family hears only replies that carry Outlook ids ───────────


@pytest.mark.parametrize(
    "url",
    [
        "https://graph.microsoft.com/v1.0/me/todo/lists/L/tasks/T/attachments",
        "https://graph.microsoft.com/v1.0/me/messages/M/send",
        "https://graph.microsoft.com/v1.0/me/messages/M/attachments/createUploadSession",
        "https://graph.microsoft.com/v1.0/me/contacts/C/photo/$value",
        "https://graph.microsoft.com/v1.0/me/events/E/accept",
        "https://graph.microsoft.com/v1.0/me/messages/M/$value",
    ],
)
def test_s5_replies_without_ids_do_not_vote(url):
    assert ids.carries_outlook_ids(url) is False


@pytest.mark.parametrize(
    "url",
    [
        "https://graph.microsoft.com/v1.0/me/messages/M/move",
        "https://graph.microsoft.com/v1.0/me/messages/M/copy",
        "https://graph.microsoft.com/v1.0/me/messages/M/createReply",
        "https://graph.microsoft.com/v1.0/me/messages",
    ],
)
def test_s5_replies_that_return_a_message_vote(url):
    assert ids.carries_outlook_ids(url) is True


async def test_s5_a_202_without_a_body_does_not_make_the_family_mixed(monkeypatch):
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    ids.ID_FAMILY.observe('IdType="ImmutableId"')

    class Reply:
        status_code = 202
        headers: dict = {}
        request = type("R", (), {"url": "https://graph.microsoft.com/v1.0/me/messages/M/forward"})

    await ids._hear(Reply())
    assert ids.ID_FAMILY.family == "immutable"


async def test_s5_the_classic_raw_paths_report_the_family(monkeypatch):
    import httpx

    from outlook_mcp.tools._delta import fetch_delta_pages
    from outlook_mcp.tools.mail_read import read_messages

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", True)
    fake = FakeGraph()
    fake.preference_applied = None  # Graph ignores the preference here
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1), "subject": "s"})
    delta = "https://graph.microsoft.com/v1.0/me/calendarView/delta"
    fake.on(
        "/me/calendarView/delta",
        lambda seen, m: {"value": [], "@odata.deltaLink": f"{delta}?$deltatoken=t"},
    )
    real = httpx.AsyncClient
    for module in ("outlook_mcp.tools._delta", "outlook_mcp.tools.mail_read"):
        monkeypatch.setattr(
            f"{module}.httpx.AsyncClient",
            lambda *a, **k: real(transport=httpx.MockTransport(fake.handle)),
        )

    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    await read_messages(fake.graph_client(), ["m1"], format="text")
    assert ids.ID_FAMILY.family == "rest"

    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    await fetch_delta_pages(
        fake.graph_client().credential,
        initial_url=f"{delta}?startDateTime=2026-11-01T00:00:00Z&endDateTime=2026-12-01T00:00:00Z",
        delta_token=None,
        page_size=50,
        resource="calendar",
        headers={"Prefer": "odata.maxpagesize=50"},
    )
    assert ids.ID_FAMILY.family == "rest"


async def test_s5_batch_triage_reports_the_family(monkeypatch):
    from outlook_mcp.config import Config
    from outlook_mcp.graph import GraphClient
    from outlook_mcp.tools.batch import batch_triage

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", True)
    fake = FakeGraph()
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)}, method="PATCH")
    monkeypatch.setattr(ids, "base_client", fake.http_client)
    client = GraphClient(fake.graph_client().credential)
    await batch_triage(client.sdk_client, ["m1"], "mark_read", "true", config=Config())
    assert ids.ID_FAMILY.family == "immutable"


def test_s5_the_classic_raw_paths_do_not_listen_when_the_switch_is_off(monkeypatch):
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", False)
    ids.hear_reply("https://graph.microsoft.com/v1.0/me/messages", 200, {})
    assert ids.ID_FAMILY.family is None


# ── N1: name tricks ───────────────────────────────────────────────────────


def test_n1_name_tricks_are_caught():
    table = people.build_sent_table(
        [letter([("john.smith@example.com", "John Smith")])], signed_in=ME
    )

    def warnings(address, name):
        found = people.person_warnings(address, name, table, written=False, contacts=[])
        return [w["code"] for w in found]

    assert "name_contains_other_address" in warnings("x@evil.example", "john.smith＠example.com")
    armenian = warnings("x@evil.example", "Jօhn Smith")
    assert "lookalike_name" in armenian or "mixed_script_name" in armenian
    assert "lookalike_name" in warnings("x@evil.example", "Jo‍hn Smith")


# ── Re-check of e26f981 ───────────────────────────────────────────────────


def _older_than_the_scan(old_letter: dict) -> Mailbox:
    filler = [
        letter(
            [("sam@example.com", "Sam")],
            sent=f"2026-{1 + i // 28 % 10:02d}-{1 + i % 28:02d}T10:00:00Z",
        )
        for i in range(300)
    ]
    return Mailbox(
        sent=filler + [old_letter],
        received=[
            received(
                "x1",
                "billing@vendor-mail.example",
                "Quick favour",
                name="Dana Whitfield",
                headers=auth_header("pass", "vendor-mail.example"),
            )
        ],
    )


async def test_s1b_a_received_name_never_counts_even_when_the_letters_are_older_than_the_scan(fake):
    """V7: Q7 proves the owner wrote to the address, not the name. The rule is about
    where the name came from, whichever query found the address."""
    old = letter([("billing@vendor-mail.example", "Vendor Billing")], sent="2024-03-01T10:00:00Z")
    resolution = await res(fake, _older_than_the_scan(old), "person", "Dana Whitfield")
    assert resolution.outcome != "single"
    candidate = next(
        c for c in resolution.candidates if c.item["address"] == "billing@vendor-mail.example"
    )
    assert candidate.item["written_to"]["basis"] == "search"  # Q7 found the old letter
    assert candidate.known is False
    assert "received_only" in codes(candidate)


async def test_s1b_the_name_on_an_old_letter_does_count(fake):
    """The owner's own letter names the person: that name may make them single."""
    old = letter([("billing@vendor-mail.example", "Dana Whitfield")], sent="2024-03-01T10:00:00Z")
    resolution = await res(fake, _older_than_the_scan(old), "person", "Dana Whitfield")
    assert resolution.outcome == "single"
    assert resolution.target.item["address"] == "billing@vendor-mail.example"


@pytest.mark.xfail(
    strict=True,
    reason="P-SI1 release gate: an accept with no type annotation, no meetingMessageType and "
    "a localized subject is still read as a letter; the probe must settle it before release",
)
async def test_p_si1_gate_a_localized_untyped_accept_does_not_make_a_stranger_known(fake):
    """V9. Fails today on purpose; when P-SI1's fix lands this passes, and `strict`
    makes the run fail until the marker comes off."""
    accept = letter([("organizer@stranger.example", "Olga Stranger")], sent="2026-11-02T10:00:00Z")
    accept["subject"] = "Zugesagt: Kostenlose Beratung"
    resolution = await res(fake, Mailbox(sent=[accept]), "person", "Olga Stranger")
    assert resolution.outcome != "single"


async def test_nit_a_note_to_self_in_the_inbox_is_no_lookalike_of_a_correspondent(fake):
    box = Mailbox(
        sent=[letter([("me1@example.com", "Me One")])],
        received=[received("n", ME, "Note to self", headers=auth_header("pass", "example.com"))],
    )
    resolution = await res(fake, box, "message", "note to self")
    candidate = resolution.candidates[0]
    assert "lookalike_address" not in codes(candidate)
    assert candidate.known is False  # still not the owner's without Sent Items (B1)
