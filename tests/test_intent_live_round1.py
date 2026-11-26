"""What the first live run against a test account showed, held offline.

The fake answers as Graph did on 2026-11-26 (live-day1 report, probes p20-p28):

- **`$search` ids.** A `$search` hands back regular (REST) ids even when its
  `Preference-Applied` says immutable, while lists and filters hand back immutable
  ones. `POST /me/translateExchangeIds` (restId to restImmutableEntryId) returns the
  immutable id. So every id `find` takes from a search is translated, one request per
  search, and no search reply votes on the id family.
- **P-AT1.** KQL `attachment:` matches only the start of a file's name, which the
  plain words search finds too, so the second search is gone.
- **P-F5.** `internetMessageHeaders` comes back on lists, `$search` and delta, so the
  sender check reads it there; a single-message read is only the fallback.
- **P-SE1.** `recipients:` also returns a longer lookalike domain, so Q7's exact
  local filter stays (pinned in test_intent_find.py).
- **P-SI1 (b).** `isOrganizer` is true on an event the owner organizes and false on a
  received meeting that names the owner as organizer, so it stays the test.
"""

from datetime import datetime, timezone

import pytest

from outlook_mcp.intents import ids, transport
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.find import find_items
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.intents.resolver import Filters, open_lookup, resolve
from outlook_mcp.intents.transport import IntentGraph
from tests.intent_fakes import (
    ME,
    REST,
    FakeGraph,
    Mailbox,
    auth_header,
    contact,
    event,
    letter,
    page,
    received,
)

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)
TOOLS = frozenset({"outlook_read_message", "outlook_list_thread", "outlook_brief"})
IMMUTABLE = 'IdType="ImmutableId"'


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


def box() -> Mailbox:
    return Mailbox(
        sent=[letter([("sam.jones@example.com", "Sam Jones")])],
        contacts=[contact("c1", "Sammy Davis", "sammy@example.net")],
        received=[
            received(
                "m1",
                "sam.jones@example.com",
                "Budget review",
                headers=auth_header("pass", "example.com"),
                attachments=True,
            ),
            received("m2", "stranger@unknown.example", "Budget offer"),
        ],
        attachments={
            "m1": [{"id": "a1", "name": "Forecast-2027.xlsx", "size": 9, "isInline": False}]
        },
    )


async def run(fake, mailbox, **args):
    mailbox.install(fake)
    return await find_items(
        fake.graph_client(), zone="UTC", signed_in=ME, loaded_tools=TOOLS, now=NOW, **args
    )


def translations(fake):
    return [s for s in fake.operations if s.path == "/me/translateExchangeIds"]


def item_ids(env):
    return [i["id"] for i in env["sections"][0]["items"]]


# ── $search ids ───────────────────────────────────────────────────────────


async def test_search_hits_are_handed_out_as_immutable_ids(fake):
    env = await run(fake, box(), kind="message", words="budget")
    assert sorted(item_ids(env)) == ["m1", "m2"]
    assert not any(str(i).startswith(REST) for i in item_ids(env))
    (call,) = translations(fake)
    assert call.method == "POST"
    assert sorted(call.body["inputIds"]) == [f"{REST}m1", f"{REST}m2"]
    assert env["graph_requests"] == fake.counted


async def test_later_reads_use_the_translated_id(fake):
    """The attachment lists of a file search go to the immutable id, not the search's."""
    env = await run(fake, box(), kind="file", words="forecast")
    assert item_ids(env) == ["a1"]
    reads = [s.path for s in fake.operations if s.path.endswith("/attachments")]
    assert reads == ["/me/messages/m1/attachments"]


async def test_contacts_from_a_search_are_translated(fake):
    env = await run(fake, box(), kind="person", words="sammy")
    (person,) = env["sections"][0]["items"]
    assert person["contact_id"] == "c1"


async def test_a_failed_translation_keeps_the_ids_and_the_family_says_mixed(fake):
    mailbox = box()
    mailbox.translate_supported = False
    env = await run(fake, mailbox, kind="message", words="budget")
    assert sorted(item_ids(env)) == [f"{REST}m1", f"{REST}m2"]
    assert ids.ID_FAMILY.family == "mixed"


async def test_no_translation_when_graph_ignores_the_preference(fake):
    fake.preference_applied = None
    await run(fake, box(), kind="message", words="budget")
    assert ids.ID_FAMILY.family == "rest"
    assert translations(fake) == []


async def test_a_search_reply_on_the_intent_path_never_votes(fake):
    fake.on("/me/messages", lambda seen, m: page([{"id": "x"}]))
    family = IdFamily()
    family.observe(None)  # the lists said regular ids
    async with IntentGraph(
        fake.graph_client().credential, RequestBudget(5), id_family=family
    ) as graph:
        await graph.get("/me/messages", {"$search": '"x"'})
    assert family.family == "rest", "a search reply's header is not evidence"


def test_a_search_reply_elsewhere_does_not_vote(monkeypatch):
    """The SDK hook and the classic raw paths: a `$search` reply's header is no
    evidence, and its ids are translated before they're handed out
    (test_classic_search_ids.py), so it doesn't vote; a failed translation does."""
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", True)
    ids.hear_reply(
        "https://graph.microsoft.com/v1.0/me/messages", 200, {"Preference-Applied": IMMUTABLE}
    )
    ids.hear_reply(
        "https://graph.microsoft.com/v1.0/me/messages?$search=%22budget%22",
        200,
        {"Preference-Applied": None},
    )
    assert ids.ID_FAMILY.family == "immutable"


# ── P-AT1 ─────────────────────────────────────────────────────────────────


async def test_the_file_kind_sends_one_search(fake):
    env = await run(fake, box(), kind="file", words="forecast")
    assert item_ids(env) == ["a1"]  # the words search finds a name by its first word
    assert fake.searches() == ['"forecast hasAttachments:true"']


# ── P-F5 ──────────────────────────────────────────────────────────────────


async def test_the_sender_check_reads_the_headers_on_the_search_reply(fake):
    env = await run(fake, box(), kind="message", words="budget review")
    first = next(i for i in env["sections"][0]["items"] if i["id"] == "m1")
    assert first["from"]["sender_auth"]["result"] == "verified"
    assert (
        "internetMessageHeaders" in next(s for s in fake.operations if s.search).params["$select"]
    )
    assert not [s for s in fake.operations if s.path.startswith("/me/messages/")]


async def test_a_row_without_the_headers_is_read_on_its_own(fake):
    mailbox = box()
    mailbox.headers_on_lists = False
    env = await run(fake, mailbox, kind="message", words="budget review")
    first = next(i for i in env["sections"][0]["items"] if i["id"] == "m1")
    assert first["from"]["sender_auth"]["result"] == "verified"
    assert any(s.path == "/me/messages/m1" for s in fake.operations)


# ── P-SI1 (b) ─────────────────────────────────────────────────────────────


def _calendar(*events) -> Mailbox:
    return Mailbox(
        sent=[letter([("sam@example.com", "Sam")])],
        calendars=[{"id": "cal-1", "name": "Calendar", "isDefaultCalendar": True}],
        events={"cal-1": list(events)},
    )


async def test_p_si1_b_an_event_the_owner_organizes_reads_is_organizer_true(fake):
    """Recorded: the owner's own "Yoga at Blue Door" reads isOrganizer true."""
    mailbox = _calendar(
        event("own", "Yoga at Blue Door", "2026-11-27T17:00:00", "2026-11-27T18:00:00")
    )
    mailbox.install(fake)
    spent = RequestBudget(30)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
        resolution = await resolve(lk, "event", "yoga", Filters(), "act", 5)
    assert resolution.outcome == "single"
    assert "organizer_you" in [w["code"] for w in resolution.target.item["why"]]


async def test_p_si1_b_a_received_meeting_naming_the_owner_is_not_theirs(fake):
    """Recorded: a received meeting naming the owner as organizer reads isOrganizer
    false with responseStatus `organizer`; the address must not override that."""
    mailbox = _calendar(
        event(
            "forged",
            "Quarterly payout",
            "2026-11-27T17:00:00",
            "2026-11-27T18:00:00",
            organizer=ME,
            is_organizer=False,
            response="organizer",
        )
    )
    mailbox.install(fake)
    spent = RequestBudget(30)
    async with open_lookup(fake.graph_client(), spent, now=NOW, zone="UTC", signed_in=ME) as lk:
        resolution = await resolve(lk, "event", "payout", Filters(), "act", 5)
    candidate = resolution.candidates[0]
    assert candidate.item["is_organizer"] is False
    assert candidate.known is False
    assert "stranger_invite" in [w["code"] for w in candidate.item["warnings"]]
    assert resolution.outcome == "candidates"
