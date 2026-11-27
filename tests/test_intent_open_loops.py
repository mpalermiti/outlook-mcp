"""``outlook_open_loops`` (contracts §12.3): conversations waiting on you, or on others.

A conversation waits on the owner when its latest message is someone else's, to
them; it waits on someone else when the latest is the owner's own letter. These
are candidates for the agent to judge, not verdicts. Pinned here: each rule
(meeting messages, letters to oneself, Cc, Focused Inbox's "other", bulk mail,
``older_than``), flagged sent mail older than the window and the reply check
behind it (§8.2), the ``person`` and ``words`` filters, the conservative choices
where Graph's behaviour is unprobed (P-D3a, P-SI1), the request count, and where
other people's text may appear.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import NoMatchError
from outlook_mcp.intents import budget, flags, ids, open_loops, transport
from outlook_mcp.intents.envelope import WARNING_CODES, WHY_CODES
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.intents.open_loops import read_open_loops
from tests.intent_fakes import (
    ME,
    SENT_FOLDER,
    FakeGraph,
    Mailbox,
    addr,
    auth_header,
    error,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
MARK = " ZQXMARK "
SAM = "sam@example.com"
PAT = "pat@example.org"
DANA = "dana@unknown.example"
LEE = "lee@example.com"
KIM = "kim@example.com"
BOB = "bob@example.com"


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def sent(
    msg_id: str,
    to: list,
    when: str,
    *,
    conversation: str,
    subject: str = "Letter",
    cc: list | None = None,
    flag: dict | None = None,
    odata_type: str | None = None,
) -> dict:
    """A letter in Sent Items, as the list and message queries select it."""
    row = {
        "id": msg_id,
        "subject": subject,
        "bodyPreview": f"preview of {subject}",
        "from": addr(ME, "Me"),
        "toRecipients": [addr(*t) if isinstance(t, tuple) else addr(t) for t in to],
        "ccRecipients": [addr(c) for c in cc or []],
        "bccRecipients": [],
        "sentDateTime": when,
        "receivedDateTime": when,
        "conversationId": conversation,
        "parentFolderId": SENT_FOLDER,
        "inferenceClassification": "focused",
        "importance": "normal",
        "isRead": True,
        "hasAttachments": False,
        "isDraft": False,
        "flag": flag or {"flagStatus": "notFlagged"},
    }
    if odata_type:
        row["@odata.type"] = odata_type
    return row


def flagged(due: str | None) -> dict:
    flag = {
        "flagStatus": "flagged",
        "startDateTime": {"dateTime": "2026-09-30T00:00:00.0000000", "timeZone": "UTC"},
    }
    if due:
        flag["dueDateTime"] = {"dateTime": f"{due}T00:00:00.0000000", "timeZone": "UTC"}
    return flag


def mailbox(m: str = "") -> Mailbox:
    """Two weeks of mail. Waiting on the owner: Sam (To, written to, also in Archive),
    Bob (To, an hour ago), Dana (Cc). Waiting on others: Pat (sent last week) and Lee
    (a flagged letter from October, never answered). Left out: a newsletter, mail
    Focused Inbox called "other", a letter to oneself, meeting messages, and Kim's
    flagged thread, which Kim answered."""
    return Mailbox(
        sent=[
            sent("s1", [SAM], "2026-11-20T10:00:00Z", conversation="conv-sam", subject=f"Plan{m}"),
            sent(
                "s2",
                [(PAT, f"Pat{m}")],
                "2026-11-18T10:00:00Z",
                conversation="conv-pat",
                subject=f"Quote{m}",
            ),
            sent("s3", [ME], "2026-11-19T10:00:00Z", conversation="conv-self", subject=f"Note{m}"),
            sent(
                "s4",
                ["org@example.com"],
                "2026-11-22T10:00:00Z",
                conversation="conv-meet",
                subject=f"Accepted: Sync{m}",
                odata_type="#microsoft.graph.eventMessageResponse",
            ),
            sent(
                "s5",
                [LEE],
                "2026-10-01T10:00:00Z",
                conversation="conv-lee",
                subject=f"Contract{m}",
                flag=flagged("2026-11-20"),
            ),
            sent(
                "s6",
                [KIM],
                "2026-10-02T10:00:00Z",
                conversation="conv-kim",
                subject=f"Visit{m}",
                flag=flagged("2026-12-01"),
            ),
        ],
        received=[
            received(
                "r1",
                SAM,
                f"Re: Plan{m}",
                name=f"Sam{m}",
                conversation="conv-sam",
                when="2026-11-21T09:00:00Z",
                headers=auth_header("pass", "example.com"),
            ),
            received(
                "r2",
                DANA,
                f"Hello{m}",
                name=f"Dana{m}",
                conversation="conv-dana",
                when="2026-11-15T09:00:00Z",
                to=["other@example.com"],
                cc=[ME],
                headers=auth_header("pass", "unknown.example"),
            ),
            {
                **received(
                    "r3",
                    "news@list.example",
                    f"Weekly{m}",
                    conversation="conv-news",
                    when="2026-11-24T09:00:00Z",
                ),
                "internetMessageHeaders": [
                    {"name": "List-Unsubscribe", "value": "<mailto:stop@list.example>"}
                ],
            },
            {
                **received(
                    "r4",
                    "alex@example.net",
                    f"Offer{m}",
                    conversation="conv-alex",
                    when="2026-11-23T09:00:00Z",
                ),
                "inferenceClassification": "other",
            },
            received("r5", ME, f"Note{m}", conversation="conv-self", when="2026-11-19T10:00:01Z"),
            {
                **received(
                    "r6",
                    "org@example.com",
                    f"Sync{m}",
                    conversation="conv-meet2",
                    when="2026-11-22T09:00:00Z",
                ),
                "@odata.type": "#microsoft.graph.eventMessageRequest",
            },
            received(
                "r7",
                BOB,
                f"Quick one{m}",
                name=f"Bob{m}",
                conversation="conv-bob",
                when="2026-11-26T11:00:00Z",
            ),
            received(
                "a1",
                SAM,
                f"Archived{m}",
                name=f"Sam{m}",
                conversation="conv-arch",
                when="2026-11-14T09:00:00Z",
                folder="archive-id",
            ),
            # Kim's reply, its sender checked: only that settles a follow-up (review B1).
            received(
                "k1",
                KIM,
                f"Re: Visit{m}",
                conversation="conv-kim",
                when="2026-10-05T09:00:00Z",
                headers=auth_header("pass", "example.com"),
            ),
        ],
    )


async def run(fake, box=None, *, tools=frozenset(), **args):
    (box or mailbox()).install(fake)
    return await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=tools, now=NOW, **args
    )


def section(env, name) -> dict:
    return next(s for s in env["sections"] if s["name"] == name)


def ids_in(env, name) -> list[str]:
    return [i["id"] for i in section(env, name)["items"]]


# ── Arguments ─────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "args",
    [
        {"direction": "sideways"},
        {"window": "91d"},
        {"window": "+3d"},
        {"window": "2026-11-01/2026-11-10"},
        {"older_than": "3 days"},
        {"older_than": "+3d"},
        {"older_than": "366d"},
        {"older_than": "today"},
        {"words": "x" * 201},
        {"words": "a\nb"},
        {"person": "  "},
    ],
)
async def test_bad_arguments_are_refused_before_any_request(fake, args):
    with pytest.raises(ValueError):
        await run(fake, **args)
    assert fake.counted == 0


async def test_the_envelope_and_its_sections(fake):
    env = await run(fake)
    assert env["tool"] == "outlook_open_loops"
    assert env["outcome"] == "answered"
    assert [s["name"] for s in env["sections"]] == ["waiting_on_me", "waiting_on_them"]
    assert {s["cap"] for s in env["sections"]} == {25}
    understood = env["understood"]
    assert understood["direction"] == "both"
    assert understood["window"]["start"] == "2026-11-12T12:00:00Z"
    assert understood["person"] is None
    assert understood["not_covered"][:3] == [
        "rule_filed_folders",
        "junk",
        "older_than_window_unless_flagged_or_searched",
    ]


@pytest.mark.parametrize(
    ("direction", "names"), [("on_me", ["waiting_on_me"]), ("on_them", ["waiting_on_them"])]
)
async def test_a_direction_reads_only_its_own_section(fake, direction, names):
    env = await run(fake, direction=direction)
    assert [s["name"] for s in env["sections"]] == names
    reads = {s.path for s in fake.operations}
    if direction == "on_me":
        # No flags are read for a question about what waits on the owner.
        assert not [s for s in fake.operations if "flag/flagStatus" in (s.filter or "")]
    else:
        # No sender checks on what waits on the owner; only Kim's reply to a flagged
        # letter is read, to confirm it (review B1).
        checked = {
            s.path
            for s in fake.operations
            if s.in_batch
            and "$select" in s.params
            and "internetMessageHeaders" in s.params["$select"]
        }
        assert checked == {"/me/messages/k1"}
    assert reads


# ── Waiting on me ─────────────────────────────────────────────────────────


async def test_waiting_on_me_ranks_to_before_cc_then_known_then_longest(fake):
    env = await run(fake)
    assert ids_in(env, "waiting_on_me") == ["conv-arch", "conv-sam", "conv-bob", "conv-dana"]
    items = {i["id"]: i for i in section(env, "waiting_on_me")["items"]}
    sam = items["conv-sam"]
    assert sam["kind"] == "thread" and sam["waits_on"] == "me"
    assert sam["last_message_id"] == "r1"
    assert sam["last_from"] == {"address": SAM, "known": True}
    assert sam["waiting_since"] == "2026-11-21T09:00:00Z"
    codes = {r["code"]: r for r in sam["why"]}
    assert set(codes) == {"you_are_on_to", "sender_written_to", "no_reply_from_you"}
    assert codes["sender_written_to"]["basis"] == "window"
    assert codes["sender_written_to"]["count"] == 1
    assert {r["code"] for r in items["conv-dana"]["why"]} >= {"you_are_on_cc"}


async def test_left_out_bulk_other_meetings_and_letters_to_oneself(fake):
    env = await run(fake)
    shown = set(ids_in(env, "waiting_on_me")) | set(ids_in(env, "waiting_on_them"))
    for left_out in ("conv-news", "conv-alex", "conv-self", "conv-meet", "conv-meet2"):
        assert left_out not in shown, left_out


async def test_the_waits_shown_first_get_a_sender_check_whoever_they_claim_to_be(fake):
    """L6 goes to the top-ranked waits, known senders included: a spoof of someone the
    owner writes to ranks high, so it most needs the check (review C1)."""
    env = await run(fake)
    items = {i["id"]: i for i in section(env, "waiting_on_me")["items"]}
    assert items["conv-dana"]["sender_auth"]["result"] == "verified"
    assert items["conv-bob"]["sender_auth"]["result"] == "unverified"
    assert {"code": "sender_unverified"} in items["conv-bob"]["warnings"]
    assert items["conv-sam"]["sender_auth"]["result"] == "verified"
    reads = {
        s.path
        for s in fake.operations
        if s.in_batch and s.path.startswith("/me/messages/") and "Headers" in s.params["$select"]
    }
    assert {"/me/messages/r1", "/me/messages/a1", "/me/messages/r2", "/me/messages/r7"} <= reads


async def test_a_failed_sender_check_reads_as_not_checked_and_keeps_the_item(fake):
    box = mailbox()
    box.install(fake)
    fake.on(r"/me/messages/r2", lambda s, m: error(500, "InternalServerError"))
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=frozenset(), now=NOW
    )
    items = {i["id"]: i for i in section(env, "waiting_on_me")["items"]}
    assert items["conv-dana"]["sender_auth"]["result"] == "not_checked"


async def test_a_forged_message_from_the_owners_address_does_not_end_a_wait(fake):
    """The owner's mail is what is in Sent Items, never what claims their address."""
    box = mailbox()
    box.received.append(
        received("forged", ME, "Re: Plan", conversation="conv-sam", when="2026-11-22T09:00:00Z")
    )
    env = await run(fake, box)
    assert "conv-sam" in ids_in(env, "waiting_on_me")
    assert "conv-sam" not in ids_in(env, "waiting_on_them")


async def test_older_than_keeps_only_long_waits(fake):
    env = await run(fake, older_than="3d")
    assert "conv-bob" not in ids_in(env, "waiting_on_me")
    assert "conv-pat" in ids_in(env, "waiting_on_them")
    fake.routes.clear()
    env = await run(fake, older_than="10d")
    assert ids_in(env, "waiting_on_them") == ["conv-lee"]
    assert ids_in(env, "waiting_on_me") == ["conv-arch", "conv-dana"]


# ── Waiting on them, and flags ────────────────────────────────────────────


async def test_waiting_on_them_oldest_first_with_who_and_the_flag(fake):
    env = await run(fake)
    assert ids_in(env, "waiting_on_them") == ["conv-lee", "conv-pat"]
    items = {i["id"]: i for i in section(env, "waiting_on_them")["items"]}
    lee = items["conv-lee"]
    assert lee["waits_on"] == "them"
    assert lee["waiting_for"] == [LEE]
    assert lee["waiting_since"] == "2026-10-01T10:00:00Z"
    assert {"code": "flag_overdue", "due": "2026-11-20T00:00:00Z"} in lee["why"]
    assert items["conv-pat"]["waiting_for"] == [PAT]


async def test_a_flagged_letter_that_was_answered_is_not_waiting(fake):
    env = await run(fake)
    assert "conv-kim" not in ids_in(env, "waiting_on_them")


async def test_the_flag_queries_are_the_contracts(fake):
    await run(fake)
    assert flags.F1_FILTER in fake.filters()
    assert fake.filters().count(flags.F1_FILTER) == 1
    replies = [f for f in fake.filters() if "conversationId eq" in f and "gt" in f]
    assert sorted(replies) == [
        "receivedDateTime gt 2026-10-01T10:00:00Z and conversationId eq 'conv-lee'",
        "receivedDateTime gt 2026-10-02T10:00:00Z and conversationId eq 'conv-kim'",
    ]
    f1 = next(s for s in fake.operations if s.filter == flags.F1_FILTER)
    assert f1.params["$orderby"] == "receivedDateTime desc"
    assert f1.params["$top"] == "100"


async def test_p_d3a_a_refused_flag_filter_falls_back_to_a_sent_items_scan(fake):
    box = mailbox()
    box.install(fake)

    real = fake.routes[-1]  # keep the mailbox's own routes; refuse only F1

    def messages(seen, match):
        if seen.filter == flags.F1_FILTER:
            return error(400, "BadRequest", "The query filter contains one or more invalid nodes")
        return box._messages(seen, match)

    fake.on("/me/messages", messages)
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=frozenset(), now=NOW
    )
    assert real is not None
    assert ids_in(env, "waiting_on_them") == ["conv-lee", "conv-pat"]
    assert section(env, "waiting_on_them")["total_exact"] is False
    assert env["outcome"] == "answered"


async def test_a_failed_reply_check_keeps_the_letter_and_says_so(fake):
    box = mailbox()
    box.install(fake)

    def messages(seen, match):
        if "conversationId eq 'conv-lee'" in (seen.filter or "") and seen.in_batch:
            return error(500, "InternalServerError")
        return box._messages(seen, match)

    fake.on("/me/messages", messages)
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=frozenset(), now=NOW
    )
    assert "conv-lee" in ids_in(env, "waiting_on_them")
    assert env["understood"]["reply_checks_failed"] == ["s5"]
    assert env["partial"] is True


# ── person and words ──────────────────────────────────────────────────────


async def test_person_keeps_conversations_they_are_in(fake):
    env = await run(fake, person=SAM)
    assert env["understood"]["person"] == SAM
    assert ids_in(env, "waiting_on_me") == ["conv-arch", "conv-sam"]
    assert ids_in(env, "waiting_on_them") == []


async def test_a_name_goes_through_the_light_path(fake):
    env = await run(fake, person="pat")
    assert env["understood"]["person"] == PAT
    assert ids_in(env, "waiting_on_them") == ["conv-pat"]
    assert ids_in(env, "waiting_on_me") == []


async def test_an_ambiguous_person_returns_candidates_and_gathers_nothing(fake):
    box = mailbox()
    box.sent.append(
        sent(
            "s9",
            [("sam.other@example.net", "Sam Other")],
            "2026-11-01T10:00:00Z",
            conversation="conv-x",
        )
    )
    env = await run(fake, box, person="sam", tools=frozenset({"outlook_open_loops"}))
    assert env["outcome"] == "candidates"
    assert [s["name"] for s in env["sections"]] == ["candidates"]
    addresses = {i["address"] for i in env["sections"][0]["items"]}
    assert {SAM, "sam.other@example.net"} <= addresses
    # Nothing was gathered: no Inbox or Archive read.
    assert not [s for s in fake.operations if "/mailFolders/inbox" in s.path]
    for offer in env["next"]:
        assert offer["template"] == "choose" and offer["tool"] == "outlook_open_loops"
        assert set(offer["args"]) == {"person"}


async def test_a_person_nobody_matches_raises(fake):
    with pytest.raises(NoMatchError):
        await run(fake, person="nobody-at-all")


async def test_words_keep_searched_conversations_even_outside_the_window(fake):
    box = mailbox()
    box.received.append(
        received(
            "old",
            "zed@example.com",
            "Invoice 42",
            conversation="conv-old",
            when="2026-09-01T09:00:00Z",
        )
    )
    box.received[0]["subject"] = "Invoice 41"  # r1, conv-sam
    env = await run(fake, box, words="invoice")
    # Sam is written to, so ranks above the older wait from someone unknown.
    assert ids_in(env, "waiting_on_me") == ["conv-sam", "conv-old"]
    assert ids_in(env, "waiting_on_them") == []
    latest = [s for s in fake.operations if s.in_batch and "conv-old" in (s.filter or "")]
    # The newest three, so a draft on top doesn't hide the conversation.
    assert len(latest) == 1 and latest[0].params["$top"] == "3"


# ── Conservative fallbacks ────────────────────────────────────────────────


async def test_p_si1_sent_items_refusing_the_cast_is_read_without_it(fake):
    box = mailbox()
    box.cast_select_supported = False
    env = await run(fake, box)
    assert "conv-meet" not in ids_in(env, "waiting_on_them")
    assert ids_in(env, "waiting_on_them") == ["conv-lee", "conv-pat"]


async def test_an_account_without_an_archive_folder_says_so(fake):
    box = mailbox()
    box.install(fake)
    fake.on(r"/me/mailFolders/archive/messages", lambda s, m: error(404, "ErrorItemNotFound"))
    env = await read_open_loops(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=frozenset(), now=NOW
    )
    assert "archive" in env["understood"]["not_covered"]
    assert "conv-arch" not in ids_in(env, "waiting_on_me")
    assert env["outcome"] == "answered"


async def test_a_spent_budget_is_incomplete_not_empty(fake, monkeypatch):
    monkeypatch.setitem(budget.CEILINGS, "outlook_open_loops", 3)
    env = await run(fake)
    assert env["outcome"] == "incomplete"
    assert {s["status"] for s in env["sections"]} == {"skipped_budget"}


P = "p@example.com"


def worst_case() -> Mailbox:
    """Every read at its most. Sent Items, the Inbox and the Archive past their caps; ten
    searched waits from unknown senders (ten header reads); nine searched conversations
    from before the window (nine latest reads); six flagged letters from before the
    window that the search also finds (five reply checks). P is in every one of them."""
    box = Mailbox()
    for n in range(320):
        box.sent.append(
            sent(
                f"s{n}",
                [P],
                f"2026-11-2{n % 5}T08:00:00Z",
                conversation=f"cs{n:03}",
                subject="filler",
            )
        )
    for n in range(6):
        box.sent.append(
            sent(
                f"f{n}",
                [P],
                "2026-10-01T10:00:00Z",
                conversation=f"zf{n}",
                subject="word",
                flag=flagged("2026-11-01"),
            )
        )
    for n in range(10):
        box.received.append(
            received(
                f"r{n}",
                f"u{n}@stranger.example",
                "word",
                cc=[P],
                conversation=f"cr{n}",
                when="2026-11-25T09:00:00Z",
            )
        )
    for n in range(320):
        box.received.append(
            received(
                f"x{n}",
                f"x{n}@stranger.example",
                "filler",
                conversation=f"cx{n}",
                when="2026-11-24T09:00:00Z",
            )
        )
    for n in range(120):
        box.received.append(
            received(
                f"a{n}",
                f"v{n}@stranger.example",
                "filler",
                conversation=f"ca{n}",
                when="2026-11-24T09:00:00Z",
                folder="archive-id",
            )
        )
    for n in range(6):
        # A checked reply from P to each flagged letter: the confirmations at their most.
        box.received.append(
            received(
                f"zr{n}",
                P,
                "reply",
                conversation=f"zf{n}",
                when="2026-10-02T09:00:00Z",
                headers=auth_header("pass", "example.com"),
            )
        )
    for n in range(9):
        box.received.append(
            received(
                f"o{n}",
                f"w{n}@stranger.example",
                "word",
                cc=[P],
                conversation=f"co{n}",
                when="2026-10-20T09:00:00Z",
            )
        )
    return box


async def test_the_worst_case_stays_within_thirty_eight_requests(fake):
    env = await run(fake, worst_case(), person=P, words="word")
    batched = [s for s in fake.operations if s.in_batch]
    # L6 (8) and the five replies' confirmations, L5 and L9, each at their most.
    assert len([s for s in batched if s.path.startswith("/me/messages/")]) == 13
    assert len([s for s in batched if "gt" in (s.filter or "")]) == 5
    assert len([s for s in batched if s.params.get("$top") == "3"]) == 4
    assert len([s for s in batched if s.path.startswith("/me/mailFolders/")]) == 2
    # L7 (3: an address reads only the sent table), L8 1, L1 3, L2 3, L3 1, L9 4, L4 1,
    # L5 5, the Junk and Deleted Items ids 2, the replies' headers 5, L6 8.
    assert fake.counted == 36
    assert env["graph_requests"] == fake.counted
    assert budget.CEILINGS["outlook_open_loops"] == 40
    # Every window read was cut, so the older conversations were not judged.
    assert env["outcome"] == "incomplete"
    assert env["understood"]["conversations_not_judged"] > 0
    assert len(section(env, "waiting_on_me")["items"]) == 10
    # Both P-SI1 fallbacks (Q1's and L1's) cost one request each, and stay under 40.
    fake.routes.clear()
    fake.seen.clear()
    box = worst_case()
    box.cast_select_supported = False
    env = await run(fake, box, person=P, words="word")
    assert fake.counted == 38 <= 40


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


async def test_codes_are_in_the_vocabulary(fake):
    env = await run(fake)
    for item in [i for s in env["sections"] for i in s["items"]]:
        assert {r["code"] for r in item["why"]} <= WHY_CODES
        assert {w["code"] for w in item["warnings"]} <= WARNING_CODES
    assert "ZQX" not in json.dumps(env)


async def test_next_offers_threads_a_reply_and_a_follow_up_when_loaded(fake):
    tools = frozenset({"outlook_list_thread", "outlook_compose", "outlook_follow_up"})
    env = await run(fake, tools=tools)
    assert [(s["template"], s["tool"]) for s in env["next"]] == [
        ("open_thread", "outlook_list_thread"),
        ("open_thread", "outlook_list_thread"),
        ("draft_reply", "outlook_compose"),
        ("set_follow_up", "outlook_follow_up"),
    ]
    assert env["next"][0]["args"] == {"conversation_id": "conv-arch"}
    assert env["next"][2] == {
        "template": "draft_reply",
        "tool": "outlook_compose",
        "args": {"mode": "reply", "message_id": "a1"},
        "needs": ["body"],
    }
    assert env["next"][3]["args"] == {"message_id": "s5"} and env["next"][3]["needs"] == ["by"]
    fake.routes.clear()
    env = await run(fake)
    assert env["next"] == []


async def test_the_window_queries_are_the_contracts(fake):
    await run(fake)
    sent_reads = [
        s for s in fake.operations if s.path == "/me/mailFolders/sentitems/messages" and s.filter
    ]
    assert sent_reads[0].filter == "sentDateTime ge 2026-11-12T12:00:00Z"
    assert sent_reads[0].params["$orderby"] == "sentDateTime desc"
    inbox = next(s for s in fake.operations if s.path == "/me/mailFolders/inbox/messages")
    assert inbox.filter == "receivedDateTime ge 2026-11-12T12:00:00Z"
    assert inbox.params["$orderby"] == "receivedDateTime desc"
    assert inbox.params["$select"] == open_loops.RECEIVED_SELECT
    archive = next(s for s in fake.operations if s.path == "/me/mailFolders/archive/messages")
    assert archive.filter == inbox.filter
