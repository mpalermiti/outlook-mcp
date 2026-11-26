"""The resolver (contracts §2): names and words to things, and when one is safe to act on.

These run the real resolver over the recorded fake (``intent_fakes``). The rules
pinned here are the ones that decide whether a name is acted on without asking:

- ranking uses only mail the owner sent and contacts they saved, never mail they
  received;
- a single match is acted on only if the owner has sent that person a letter
  (decision D4): a saved contact, a meeting response or the owner's own address
  doesn't count;
- a lookalike or spoofed sender is never a single match.
"""

import asyncio
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import NoMatchError
from outlook_mcp.intents import ids, resolver, transport
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.intents.resolver import Filters, open_lookup, resolve
from tests.intent_fakes import (
    ME,
    FakeGraph,
    Mailbox,
    auth_header,
    contact,
    event,
    letter,
    meeting_response,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def mailbox() -> Mailbox:
    return Mailbox(
        sent=[
            letter([("sam.jones@example.com", "Sam Jones")], sent="2026-11-20T10:00:00Z"),
            letter(["sam.jones@example.com"], sent="2026-11-10T10:00:00Z"),
            letter(["sam.jones@example.com"], sent="2026-10-30T10:00:00Z"),
            letter(["sam.jones@example.com"], sent="2026-10-01T10:00:00Z"),
            letter([("sam.lee@example.org", "Sam Lee")], sent="2025-06-01T10:00:00Z"),
            meeting_response("organizer@stranger.example"),
        ],
        contacts=[contact("c1", "Sammy Davis", "sammy@example.net")],
        received=[
            received(
                "r1",
                "sam.received@example.net",
                "hello",
                name="Sam Received",
                headers=auth_header("pass", "example.net"),
            )
        ],
    )


async def run(fake, box, kind, words, *, filters=None, purpose="read", limit=5, budget=30):
    box.install(fake)
    spent = RequestBudget(budget)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lookup:
        resolution = await resolve(lookup, kind, words, filters or Filters(), purpose, limit)
    return resolution, spent


def addresses(resolution) -> list[str]:
    return [c.item["address"] for c in resolution.candidates]


def scores(resolution) -> dict[str, int]:
    return {c.item["address"]: c.item["signals"]["score"] for c in resolution.candidates}


def warning_codes(candidate) -> list[str]:
    return [w["code"] for w in candidate.item["warnings"]]


# ── Ranking ───────────────────────────────────────────────────────────────


async def test_ranking_is_exact_for_a_fixture(fake):
    resolution, _ = await run(fake, mailbox(), "person", "sam")
    assert resolution.outcome == "candidates"
    assert addresses(resolution) == [
        "sam.jones@example.com",  # first name 30 + 4 letters 15 + recent 10
        "sam.lee@example.org",  # first name 30 + 1 letter 5, over a year ago
        "sam.received@example.net",  # first name 30, received only
        "sammy@example.net",  # prefix 20 + saved contact 10
    ]
    assert scores(resolution) == {
        "sam.jones@example.com": 55,
        "sam.lee@example.org": 35,
        "sam.received@example.net": 30,
        "sammy@example.net": 30,
    }
    jones = resolution.candidates[0].item
    assert jones["written_to"] == {"count": 4, "last": "2026-11-20T10:00:00Z", "basis": "scan"}
    assert jones["signals"] == {"score": 55, "name_match": "first_or_last", "saved_contact": False}
    assert jones["id"] == "person:sam.jones@example.com"
    assert jones["others_wrote"] == {"name": "Sam Jones"}
    received_only = resolution.candidates[2]
    assert "received_only" in warning_codes(received_only)
    assert received_only.item["written_to"] is None
    assert received_only.item["first_seen"] == "2026-11-20T09:00:00Z"
    assert received_only.item["first_seen_basis"] == "received"
    assert received_only.item["sender_auth"] == {"result": "verified", "from_domain": "example.net"}
    sammy = resolution.candidates[3].item
    assert sammy["contact_id"] == "c1"
    assert sammy["signals"]["saved_contact"] is True


async def test_received_mail_never_changes_rank(fake):
    before, _ = await run(fake, mailbox(), "person", "sam")
    flooded = mailbox()
    flooded.received += [
        received(f"lee{n}", "sam.lee@example.org", f"note {n}", name="Sam Lee") for n in range(50)
    ]
    fake.seen.clear()
    fake.routes.clear()
    after, _ = await run(fake, flooded, "person", "sam")
    assert scores(after) == scores(before)
    assert addresses(after) == addresses(before)


async def test_a_saved_contact_alone_is_not_known(fake):
    resolution, _ = await run(fake, mailbox(), "person", "sammy")
    assert addresses(resolution) == ["sammy@example.net"]
    assert resolution.outcome == "candidates"
    assert resolution.candidates[0].known is False
    assert "first_time" in warning_codes(resolution.candidates[0])


# ── The act-on-match rule ─────────────────────────────────────────────────


async def test_single_known_match_is_acted_on(fake):
    resolution, _ = await run(fake, mailbox(), "person", "jones", purpose="act")
    assert resolution.outcome == "single"
    assert resolution.target.item["address"] == "sam.jones@example.com"
    assert resolution.proceed("act") is resolution.target


async def test_single_unknown_match_returns_candidates(fake):
    resolution, _ = await run(fake, mailbox(), "person", "received")
    assert resolution.outcome == "candidates"
    assert addresses(resolution) == ["sam.received@example.net"]
    # A read proceeds with the one candidate, carrying its warnings; an act writes nothing.
    assert resolution.proceed("read") is resolution.candidates[0]
    assert resolution.proceed("act") is None


async def test_two_matches_return_candidates(fake):
    resolution, _ = await run(fake, mailbox(), "person", "sam", purpose="act")
    assert resolution.outcome == "candidates"
    assert resolution.target is None
    assert resolution.proceed("act") is None
    assert resolution.proceed("read") is None


async def test_nothing_found_raises_for_act_and_for_reads_other_than_find(fake):
    resolution, _ = await run(fake, mailbox(), "person", "nobody")
    assert resolution.outcome == "nothing_found"
    with pytest.raises(NoMatchError, match="person 'nobody'"):
        resolution.proceed("act")
    with pytest.raises(NoMatchError):
        resolution.proceed("read")


async def test_sent_table_drops_meeting_messages_and_own_addresses(fake):
    box = mailbox()
    box.sent.append(letter(["me.alias@example.com"], sender="me.alias@example.com"))
    # An accept sent to a stranger's invitation does not make the stranger known.
    resolution, _ = await run(fake, box, "person", "organizer@stranger.example", purpose="act")
    assert resolution.outcome == "exact"
    assert resolution.target.known is False
    assert "first_time" in warning_codes(resolution.target)
    # The owner's own addresses never rank.
    fake.routes.clear()
    resolution, _ = await run(fake, box, "person", "me")
    assert "me.alias@example.com" not in addresses(resolution)
    assert ME not in addresses(resolution)


# ── Spoofing ──────────────────────────────────────────────────────────────


def _spoof_box(name: str, address: str) -> Mailbox:
    box = mailbox()
    box.sent.append(letter([(address, name)], sent="2026-11-21T10:00:00Z"))
    return box


@pytest.mark.parametrize(
    ("name", "address", "code"),
    [
        ("Sаm Smith", "sam.smith@example.com", "mixed_script_name"),
        ("Pat Xn", "pat@xn--pypal-4ve.com", "idn_domain"),
        ("billing@bank.example.com", "pat@evil.example", "name_contains_other_address"),
    ],
)
async def test_a_blocking_warning_keeps_even_a_known_match_from_single(fake, name, address, code):
    resolution, _ = await run(fake, _spoof_box(name, address), "person", address)
    assert resolution.outcome == "exact"  # an exact address is still acted on...
    assert code in warning_codes(resolution.target)  # ...and the warning travels
    fake.routes.clear()
    words = {
        "mixed_script_name": "smith",
        "idn_domain": "xn",
        "name_contains_other_address": "billing",
    }
    resolution, _ = await run(fake, _spoof_box(name, address), "person", words[code], purpose="act")
    assert len(resolution.candidates) == 1
    assert resolution.candidates[0].known is True
    assert resolution.outcome == "candidates"


@pytest.mark.parametrize(
    ("known", "lookalike"),
    [("pat@paypal.com", "pat@paypa1.com"), ("dev@modern.com", "dev@rnodern.com")],
)
async def test_blocking_warnings_prevent_single(fake, known, lookalike):
    box = mailbox()
    box.sent.append(letter([(known, "Pat Dev")]))
    box.received.append(received("spoof", lookalike, "invoice", name="Pat Dev"))
    resolution, _ = await run(fake, box, "person", "pat dev", purpose="act")
    by_address = {c.item["address"]: c for c in resolution.candidates}
    assert set(by_address) == {known, lookalike}
    assert "lookalike_domain" in warning_codes(by_address[lookalike])
    assert "lookalike_name" in warning_codes(by_address[lookalike])
    assert by_address[lookalike].blocking
    assert resolution.outcome == "candidates"


async def test_a_spoofed_sender_alone_is_never_single(fake):
    box = mailbox()
    box.sent.append(letter([("pat@paypal.com", "Pat Dev")]))
    box.received.append(
        received(
            "spoof",
            "pat@paypal.com",
            "invoice",
            name="Pat Dev",
            headers=auth_header("fail", "paypal.com"),
        )
    )
    resolution, _ = await run(fake, box, "person", "pat dev", purpose="act")
    assert addresses(resolution) == ["pat@paypal.com"]
    target = resolution.candidates[0]
    assert target.known is True
    assert target.item["sender_auth"]["result"] == "failed"
    assert "sender_failed" in warning_codes(target)
    assert resolution.outcome == "candidates"


async def test_stranger_invite_flag_and_bulk_exclusion(fake):
    box = mailbox()
    box.calendars = [{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}]
    box.events = {
        "cal-1": [
            event(
                "e-stranger",
                "Coffee chat",
                "2026-11-27T10:00:00",
                "2026-11-27T11:00:00",
                organizer="organizer@stranger.example",
                organizer_name="Org",
                is_organizer=False,
                response="notResponded",
            ),
            event(
                "e-known",
                "Coffee with Sam",
                "2026-11-28T10:00:00",
                "2026-11-28T11:00:00",
                organizer="sam.jones@example.com",
                organizer_name="Sam Jones",
                is_organizer=False,
                response="notResponded",
            ),
        ]
    }
    resolution, _ = await run(fake, box, "event", "coffee")
    by_id = {c.item["id"]: c for c in resolution.candidates}
    assert "stranger_invite" in warning_codes(by_id["e-stranger"])
    assert by_id["e-stranger"].blocking and not by_id["e-stranger"].known
    assert "stranger_invite" not in warning_codes(by_id["e-known"])
    assert by_id["e-known"].known
    fake.routes.clear()
    only_stranger, _ = await run(fake, box, "event", "chat", purpose="act")
    assert only_stranger.outcome == "candidates"


# ── Exact references and queries ──────────────────────────────────────────


async def test_exact_id_and_address_skip_search(fake):
    resolution, _ = await run(fake, mailbox(), "person", "sam.jones@example.com")
    assert resolution.outcome == "exact"
    assert resolution.target.known
    # No name search: not the contacts search (Q2), nor from: (Q4) or to: (Q8).
    assert not any(s.startswith('"from:') or s.startswith('"to:') for s in fake.searches())
    assert not any(s.path == "/me/contacts" and s.search for s in fake.operations)

    fake.seen.clear()
    long_id = "AAMkAD" + "r" * 40
    box = mailbox()
    box.received.append(received(long_id, "sam.jones@example.com", "budget"))
    fake.routes.clear()
    resolution, _ = await run(fake, box, "message", long_id)
    assert resolution.outcome == "exact"
    assert resolution.target.item["id"] == long_id
    assert fake.searches() == []


async def test_id_like_words_fall_through_to_search(fake):
    box = mailbox()
    box.received.append(
        received("m1", "sam.jones@example.com", "budget-spreadsheet-final-version-2026")
    )
    resolution, _ = await run(fake, box, "message", "budget-spreadsheet-final-version-2026")
    assert any(s.path.startswith("/me/messages/budget") for s in fake.operations)
    assert fake.searches()
    assert [c.item["id"] for c in resolution.candidates] == ["m1"]


async def test_person_queries_are_exactly_these(fake):
    await run(fake, mailbox(), "person", "organizer@stranger.example")
    sent = {(s.path, tuple(sorted(s.params.items()))) for s in fake.operations}
    q1 = (
        "/me/mailFolders/sentitems/messages",
        (
            ("$orderby", "sentDateTime desc"),
            (
                "$select",
                "from,toRecipients,ccRecipients,bccRecipients,sentDateTime,conversationId",
            ),
            ("$top", "100"),
        ),
    )
    q3 = (
        "/me/contacts",
        (
            ("$filter", "emailAddresses/any(a:a/address eq 'organizer@stranger.example')"),
            ("$select", "id,displayName,emailAddresses"),
            ("$top", "5"),
        ),
    )
    q5 = (
        "/me/messages",
        (
            (
                "$filter",
                "receivedDateTime ge 1900-01-01T00:00:00Z and "
                "from/emailAddress/address eq 'organizer@stranger.example'",
            ),
            ("$orderby", "receivedDateTime asc"),
            ("$select", "id,receivedDateTime"),
            ("$top", "1"),
        ),
    )
    q6 = (
        "/me/messages",
        (
            (
                "$filter",
                "receivedDateTime ge 1900-01-01T00:00:00Z and "
                "from/emailAddress/address eq 'organizer@stranger.example'",
            ),
            ("$orderby", "receivedDateTime desc"),
            ("$select", "id,from,internetMessageHeaders"),
            ("$top", "1"),
        ),
    )
    q7 = (
        "/me/mailFolders/sentitems/messages",
        (
            ("$search", '"recipients:organizer@stranger.example"'),
            ("$select", "toRecipients,ccRecipients,bccRecipients,sentDateTime"),
            ("$top", "25"),
        ),
    )
    assert {q1, q3, q5, q6, q7} <= sent

    fake.seen.clear()
    fake.routes.clear()
    await run(fake, Mailbox(), "person", "zed")
    sent = {(s.path, tuple(sorted(s.params.items()))) for s in fake.operations}
    q2 = (
        "/me/contacts",
        (
            ("$search", '"zed"'),
            ("$select", "id,displayName,givenName,surname,emailAddresses"),
            ("$top", "25"),
        ),
    )
    q4 = (
        "/me/messages",
        (("$search", '"from:zed"'), ("$select", "from,receivedDateTime"), ("$top", "25")),
    )
    q8 = (
        "/me/mailFolders/sentitems/messages",
        (
            ("$search", '"to:zed"'),
            ("$select", "toRecipients,ccRecipients,bccRecipients,sentDateTime"),
            ("$top", "25"),
        ),
    )
    assert {q1, q2, q4, q8} <= sent


async def test_kql_from_filters_is_sanitized(fake):
    await run(fake, mailbox(), "person", 'zed" OR from:evil@example.com (x')
    for search in fake.searches():
        inner = search[1:-1]
        assert '"' not in inner
        assert " OR " not in inner
        assert "(" not in inner and ")" not in inner
    fake.seen.clear()
    fake.routes.clear()
    await run(fake, mailbox(), "message", "budget", filters=Filters(sender="sam.jones@example.com"))
    assert fake.searches() == ['"budget from:sam.jones@example.com"']


async def test_resolution_stays_within_budget(fake):
    box = mailbox()
    box.sent += [letter([(f"sam{n}@example.{n}.com", f"Sam {n}")]) for n in range(400)]
    box.contacts += [contact(f"c{n}", f"Sam C{n}", f"samc{n}@example.org") for n in range(30)]
    resolution, spent = await run(fake, box, "person", "sam")
    assert spent.spent <= 22
    assert fake.counted == spent.spent
    assert len(resolution.candidates) == 5


async def test_first_seen_times_out_to_null(fake, monkeypatch):
    monkeypatch.setattr(resolver, "FIRST_SEEN_TIMEOUT_S", 0.05)
    fake.delays.append((lambda seen: (seen.params.get("$orderby") or "").endswith("asc"), 1.0))
    resolution, _ = await run(fake, mailbox(), "person", "received")
    target = resolution.candidates[0].item
    assert target["first_seen"] is None
    assert target["first_seen_basis"] == "none"
    assert target["sender_auth"]["result"] == "verified"


async def test_light_path_issues_at_most_five_requests(fake):
    box = mailbox()
    box.sent += [letter([f"x{n}@example.com"]) for n in range(250)]
    box.received.append(received("m1", "sam.jones@example.com", "budget"))
    resolution, spent = await run(fake, box, "message", "budget", filters=Filters(sender="jones"))
    light = [
        s
        for s in fake.operations
        if s.path in ("/me/mailFolders/sentitems/messages", "/me/contacts")
    ]
    assert len(light) <= 5
    assert not any(s.path == "/me/messages" and s.filter for s in fake.operations)
    assert [c.item["id"] for c in resolution.candidates] == ["m1"]


async def test_an_ambiguous_sender_returns_the_people(fake):
    box = mailbox()
    box.received.append(received("m1", "sam.jones@example.com", "budget"))
    resolution, _ = await run(fake, box, "message", "budget", filters=Filters(sender="sam"))
    assert resolution.outcome == "candidates"
    assert {c.item["kind"] for c in resolution.candidates} == {"person"}
    assert not any(s.search == '"budget"' for s in fake.operations)


async def test_the_same_function_serves_every_purpose(fake):
    for purpose in ("read", "act"):
        fake.routes.clear()
        resolution, _ = await run(fake, mailbox(), "person", "jones", purpose=purpose)
        assert resolution.outcome == "single"


def test_kinds_and_purposes_are_the_contracts():
    assert resolver.KINDS == (
        "person",
        "message",
        "thread",
        "file",
        "event",
        "task",
        "draft",
        "rule",
        "folder",
    )
    assert resolver.PURPOSES == ("read", "act")


async def test_concurrent_resolutions_share_the_limiter(fake):
    box = mailbox()
    box.install(fake)

    async def one():
        spent = RequestBudget(30)
        async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
            return await resolve(lk, "person", "sam", Filters(), "read", 5)

    first, second = await asyncio.gather(one(), one())
    assert addresses(first) == addresses(second)
