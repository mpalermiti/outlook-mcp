"""Live query-shape guards for the intent resolver behind outlook_find. Read-only.

Run with: scripts/live-tests.sh live (RELEASING.md 1b), against the project's
test accounts. Deselected by default, so the offline suite and CI never run them.

The offline tests (``test_intent_resolver.py``, ``test_intent_find.py``) assert
what the resolver *sends*. These assert what Graph *does* with it, for every new
``$filter``, ``$orderby`` and ``$search`` shape the resolver builds (contracts
§2.8, §4.5): a loud 400 fails the call, and the quiet failure, a 200 with the
wrong rows, fails an assertion on the rows. Each query is built by the same
helper or constant the resolver uses, so the two cannot drift apart.

They harvest their own fixtures from whatever mailbox is signed in and skip when
it lacks the data an assertion needs. Every request is a read: the live guard
refuses anything else before it leaves (``tests/live_guard.py``).
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import pytest

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.intents import resolver
from outlook_mcp.intents.budget import RequestBudget
from outlook_mcp.intents.find import find_items
from outlook_mcp.intents.people import parse_time, recipients, sender_of
from outlook_mcp.intents.resolver import open_lookup
from outlook_mcp.intents.selection import MailSelection, attachment_kql, search_kql
from outlook_mcp.tools.todo import fetch_all_task_lists

pytestmark = [pytest.mark.live, pytest.mark.asyncio]

NOMATCH_ADDRESS = "zzqqxx-nomatch@example.invalid"
_WORD = re.compile(r"^[A-Za-z]{4,}$")


@pytest.fixture
async def lookup(real_graph_client):
    """A counted lookup over the signed-in mailbox, with room for harvesting."""
    async with open_lookup(
        real_graph_client,
        RequestBudget(200),
        now=datetime.now(timezone.utc),
        zone="UTC",
        signed_in=None,
    ) as found:
        yield found


async def _received(lookup, top: int = 50) -> list[dict]:
    body = await lookup.graph.get(
        "/me/messages",
        {"$select": "id,subject,from,receivedDateTime,conversationId,hasAttachments", "$top": top},
    )
    return [r for r in body.get("value") or [] if sender_of(r)[0]]


def _word_of(text: str | None) -> str | None:
    return next((w for w in (text or "").split() if _WORD.fullmatch(w)), None)


# ── Q1: the sent table ────────────────────────────────────────────────────


async def test_q1_sent_items_come_newest_first(lookup):
    body = await lookup.graph.get(
        resolver.SENT_PATH,
        {"$select": resolver.SENT_SELECT, "$orderby": "sentDateTime desc", "$top": 100},
    )
    rows = body.get("value") or []
    if not rows:
        pytest.skip("Sent Items is empty")
    sent = [r.get("sentDateTime") for r in rows if r.get("sentDateTime")]
    assert sent == sorted(sent, reverse=True), "the sent table must read the newest letters"
    assert all("toRecipients" in r for r in rows), "the $select must reach every row"
    table = await lookup.sent_table()
    assert table.scanned >= len(rows)


# ── Q2, Q3: contacts ──────────────────────────────────────────────────────


async def _contacts(lookup) -> list[dict]:
    body = await lookup.graph.get(
        "/me/contacts", {"$select": "id,displayName,emailAddresses", "$top": 50}
    )
    return body.get("value") or []


async def test_q2_contact_search_finds_a_contact_by_name(lookup):
    harvested = [(c, _word_of(c.get("displayName"))) for c in await _contacts(lookup)]
    harvested = [(c, w) for c, w in harvested if w]
    if not harvested:
        pytest.skip("no contact with a plain word in its name")
    contact, word = harvested[0]
    body = await lookup.graph.get(
        "/me/contacts",
        {
            "$search": f'"{word}"',
            "$select": resolver.CONTACT_SEARCH_SELECT,
            "$top": resolver.SEARCH_TOP,
        },
    )
    assert contact["id"] in {c.get("id") for c in body.get("value") or []}


async def test_q3_contact_filter_matches_only_the_exact_address(lookup):
    harvested = [
        (c, (c.get("emailAddresses") or [{}])[0].get("address"))
        for c in await _contacts(lookup)
        if c.get("emailAddresses")
    ]
    harvested = [(c, a) for c, a in harvested if a]
    if not harvested:
        pytest.skip("no contact with an address")
    contact, address = harvested[0]

    async def by_address(value: str) -> set[str]:
        body = await lookup.graph.get(
            "/me/contacts",
            {
                "$filter": f"emailAddresses/any(a:a/address eq '{resolver._literal(value)}')",
                "$select": resolver.CONTACT_EXACT_SELECT,
                "$top": 5,
            },
        )
        return {c.get("id") for c in body.get("value") or []}

    assert contact["id"] in await by_address(address)
    assert await by_address(NOMATCH_ADDRESS) == set()


# ── Q4: people who wrote to the owner ─────────────────────────────────────


async def test_q4_from_search_finds_the_sender(lookup):
    for row in await _received(lookup):
        address, name = sender_of(row)
        word = _word_of(name)
        if word:
            break
    else:
        pytest.skip("no received message whose sender has a plain-word name")
    body = await lookup.graph.get(
        "/me/messages",
        {
            "$search": resolver._q(word, "from:"),
            "$select": "from,receivedDateTime",
            "$top": resolver.SEARCH_TOP,
        },
    )
    hits = body.get("value") or []
    assert hits, f"from:{word} found nothing, though a message from that sender exists"
    assert any(word.casefold() in " ".join(filter(None, sender_of(r))).casefold() for r in hits)


# ── Q5, Q6: first and latest message from an address ──────────────────────


async def test_q5_and_q6_filter_on_the_sender_and_order_both_ways(lookup):
    rows = await _received(lookup)
    if not rows:
        pytest.skip("no received mail")
    address = sender_of(rows[0])[0]

    async def one(order: str, select: str) -> dict:
        body = await lookup.graph.get(
            "/me/messages",
            {
                "$filter": resolver._from_filter(address),
                "$orderby": f"receivedDateTime {order}",
                "$top": 1,
                "$select": select,
            },
        )
        found = body.get("value") or []
        assert found, f"the {order} lookup found no message from a sender that has one"
        return found[0]

    first = await one("asc", "id,receivedDateTime")
    latest = await one("desc", resolver.HEADER_SELECT)
    assert sender_of(latest)[0] == address
    # The earliest is no later than any message from them, the harvested one included.
    assert parse_time(first["receivedDateTime"]) <= parse_time(rows[0]["receivedDateTime"])
    if "internetMessageHeaders" not in latest:
        # P-F5: a list left the headers out, so the resolver reads the message itself.
        read = await lookup.graph.get(
            f"/me/messages/{latest['id']}", {"$select": resolver.HEADER_SELECT}
        )
        assert "internetMessageHeaders" in read, "the single-message fallback must return headers"


# ── Q7, Q8: older sent mail ───────────────────────────────────────────────


async def test_q7_recipients_search_finds_letters_that_list_the_address(lookup):
    table = await lookup.sent_table()
    if not table.entries:
        pytest.skip("no letters in Sent Items")
    address = max(table.entries.values(), key=lambda e: e.count).address

    async def exact_hits(value: str) -> int:
        body = await lookup.graph.get(
            resolver.SENT_PATH,
            {
                "$search": f'"recipients:{value}"',
                "$select": resolver.RECIPIENTS_SELECT,
                "$top": resolver.SEARCH_TOP,
            },
        )
        return sum(1 for r in body.get("value") or [] if any(a == value for a, _ in recipients(r)))

    assert await exact_hits(address) >= 1
    assert await exact_hits(NOMATCH_ADDRESS) == 0


async def test_q8_to_search_finds_a_recipient_by_name(lookup):
    table = await lookup.sent_table()
    named = [(e, _word_of(e.name)) for e in table.entries.values()]
    named = [(e, w) for e, w in named if w]
    if not named:
        pytest.skip("no recipient with a plain-word name")
    entry, word = named[0]
    body = await lookup.graph.get(
        resolver.SENT_PATH,
        {
            "$search": resolver._q(word, "to:"),
            "$select": resolver.RECIPIENTS_SELECT,
            "$top": resolver.SEARCH_TOP,
        },
    )
    hits = body.get("value") or []
    assert hits, f"to:{word} found no letter, though one was sent to that name"
    assert any(
        word.casefold() in f"{a} {n or ''}".casefold()
        for r in hits
        for a, n in recipients(r, ("toRecipients",))
    )


# ── Message search: the KQL a selection sends (§4.2) ──────────────────────


async def _search(lookup, selection: MailSelection) -> list[dict]:
    body = await lookup.graph.get(
        "/me/messages",
        {
            "$search": search_kql(selection),
            "$select": resolver.MESSAGE_SELECT,
            "$top": resolver.SEARCH_TOP,
        },
    )
    return body.get("value") or []


async def test_a_search_with_a_sender_returns_only_their_mail(lookup):
    for row in await _received(lookup):
        word = _word_of(row.get("subject"))
        if word:
            break
    else:
        pytest.skip("no received message with a plain word in its subject")
    address = sender_of(row)[0]
    hits = await _search(lookup, MailSelection(words=word, from_=[address]))
    assert row["id"] in {h["id"] for h in hits}, "the harvested message must match itself"
    assert all(sender_of(h)[0] == address for h in hits)


async def test_several_senders_are_one_or_group(lookup):
    rows = await _received(lookup)
    senders = list(dict.fromkeys(sender_of(r)[0] for r in rows))[:2]
    word = next((w for w in (_word_of(r.get("subject")) for r in rows) if w), None)
    if len(senders) < 2 or not word:
        pytest.skip("fewer than two senders, or no plain subject word")
    hits = await _search(lookup, MailSelection(words=word, from_=senders))
    assert all(sender_of(h)[0] in senders for h in hits)


async def test_received_dates_narrow_a_search_to_the_day(lookup):
    for row in await _received(lookup):
        word = _word_of(row.get("subject"))
        when = parse_time(row.get("receivedDateTime"))
        if word and when:
            break
    else:
        pytest.skip("no received message with a subject word and a time")
    hits = await _search(
        lookup, MailSelection(words=word, newer_than=when, older_than=when + timedelta(seconds=1))
    )
    assert row["id"] in {h["id"] for h in hits}, "a message must fall inside its own day"
    for hit in hits:
        received = parse_time(hit.get("receivedDateTime"))
        # The KQL bound is a whole day, widened by one on each side.
        assert received is not None
        assert when - timedelta(days=2) <= received <= when + timedelta(days=2)


async def test_has_attachments_narrows_a_search(lookup):
    for row in await _received(lookup):
        word = _word_of(row.get("subject"))
        if row.get("hasAttachments") and word:
            break
    else:
        pytest.skip("no received message with attachments and a plain subject word")
    hits = await _search(lookup, MailSelection(words=word, has_attachments=True))
    assert row["id"] in {h["id"] for h in hits}
    assert all(h.get("hasAttachments") is True for h in hits)


async def test_p_at1_attachment_search_finds_a_file_by_its_name(lookup):
    """KQL `attachment:` must reach attachment names, or find falls back without it.

    Graph refusing it (400) is P-AT1's other answer: the resolver then searches the
    words alone, so that skips with the reason rather than failing.
    """
    for row in await _received(lookup):
        if not row.get("hasAttachments"):
            continue
        body = await lookup.graph.get(
            f"/me/messages/{quote(row['id'], safe='')}/attachments",
            {"$select": resolver.ATTACHMENT_SELECT},
        )
        names = [a.get("name") for a in body.get("value") or [] if a.get("isInline") is False]
        word = next(
            (w for n in names for w in re.split(r"[^A-Za-z]+", n or "") if _WORD.fullmatch(w)),
            None,
        )
        if word:
            break
    else:
        pytest.skip("no received message with a file whose name has a plain word")
    query = attachment_kql(MailSelection(words=word, has_attachments=True))
    try:
        hits = (
            await lookup.graph.get(
                "/me/messages",
                {"$search": query, "$select": resolver.MESSAGE_SELECT, "$top": resolver.SEARCH_TOP},
            )
        ).get("value") or []
    except GraphAPIError as error:
        if error.status_code == 400:
            pytest.skip("P-AT1: Graph refuses `attachment:` here; find searches the words alone")
        raise
    assert hits, f"attachment:{word} found nothing, though a file of that name exists"
    assert all(h.get("hasAttachments") is True for h in hits)


# ── A conversation, newest first ──────────────────────────────────────────


async def test_a_conversation_reads_newest_first(lookup):
    rows = await _received(lookup)
    conversation = next((r.get("conversationId") for r in rows if r.get("conversationId")), None)
    if not conversation:
        pytest.skip("no message with a conversation id")
    body = await lookup.graph.get(
        "/me/messages",
        {
            "$filter": resolver._conversation_filter(conversation),
            "$orderby": "receivedDateTime desc",
            "$top": 10,
            "$select": resolver.MESSAGE_SELECT,
        },
    )
    found = body.get("value") or []
    assert found, "the harvested message must be in its own conversation"
    assert all(r.get("conversationId") == conversation for r in found)
    received = [r["receivedDateTime"] for r in found if r.get("receivedDateTime")]
    assert received == sorted(received, reverse=True)


# ── To Do: open tasks (P-H1a) ─────────────────────────────────────────────


async def test_the_open_task_filter_returns_no_completed_task(lookup):
    lists = [
        t
        for t in await fetch_all_task_lists(lookup.sdk)
        if resolver._list_name(t) != "flaggedEmails"
    ]
    if not lists:
        pytest.skip("no task list")
    try:
        body = await lookup.graph.get(
            f"/me/todo/lists/{quote(lists[0].id, safe='')}/tasks",
            {"$filter": "status ne 'completed'", "$top": resolver.TASK_TOP},
            ids=False,
        )
    except GraphAPIError as error:
        if error.status_code == 400:
            pytest.skip("P-H1a: Graph refuses this filter here; the resolver filters locally")
        raise
    assert all(t.get("status") != "completed" for t in body.get("value") or [])


# ── find, end to end, every kind ──────────────────────────────────────────


@pytest.mark.parametrize(
    "kind", ["person", "message", "thread", "file", "event", "task", "draft", "rule", "folder"]
)
async def test_find_answers_every_kind_within_its_ceiling(real_graph_client, kind):
    words = {"folder": "inbox"}.get(kind, "the")
    env = await find_items(
        real_graph_client,
        kind=kind,
        words=words,
        zone="UTC",
        signed_in=None,
        loaded_tools=frozenset(),
    )
    assert env["tool"] == "outlook_find"
    assert env["outcome"] in ("exact", "single", "candidates", "nothing_found")
    assert env["graph_requests"] <= 30
    for failure in env["failed_sections"]:
        # Only a mailbox that refuses inbox rules may leave the section unanswered.
        assert kind == "rule" and failure["reason"] == "not_supported", failure
