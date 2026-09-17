"""Live query-shape regression guards. Requires a cached token.

Run with: uv run pytest -m live -v
Auto-skipped without credentials (so CI, which has none, skips them).

WHY THIS TIER EXISTS
--------------------
The default suite mocks the Graph client, so it asserts what string we *send*.
It cannot observe what Graph *does* with it. Three bugs shipped in 1.12.0 under
558 green tests because of that gap:

* ``$orderby`` + a filter on any other property returned 400 InefficientFilter,
  breaking ``from_address`` and ``classification`` on every call (#31).
* ``list_thread`` hit the same rule and returned 400 on *every* call — the tool
  had never worked in a released version.
* ``sanitize_kql`` stripped ``:``, so every documented KQL property restriction
  was sent as a literal phrase and returned 200 with zero results (#30).

Note the two distinct failure modes: a loud 400, and a silent 200-with-wrong-
results. Only the second is genuinely dangerous, and only a live call sees it.
So these tests assert on *returned data*, not just on absence of an exception.

``scripts/preflight.py`` does not cover this: it probes endpoint reachability
and treats a 400 as a non-blocking SKIP, whereas every bug above is a
query-shape failure against an endpoint that exists and responds.

READ-ONLY
---------
Every call in this module must be a read. No sends, drafts, category or folder
mutations, no deletes. A regression guard must never be able to damage the
mailbox it runs against.

MAILBOX-INDEPENDENT
-------------------
These run against whatever mailbox is authenticated, so they harvest their own
fixtures (a real sender, subject term, conversation id) and skip cleanly when
the mailbox lacks the data a given assertion needs. They must never assume this
maintainer's mailbox.
"""

import pytest

from outlook_mcp.tools.contacts import list_contacts, search_contacts
from outlook_mcp.tools.mail_read import list_inbox, search_mail
from outlook_mcp.tools.mail_thread import list_thread

pytestmark = [pytest.mark.live, pytest.mark.asyncio]


@pytest.fixture
async def sample(real_graph_client):
    """Harvest real fixture data from the authenticated mailbox (read-only)."""
    result = await list_inbox(real_graph_client.sdk_client, count=25)
    if not result["messages"]:
        pytest.skip("Mailbox inbox is empty — nothing to build live assertions from")
    messages = result["messages"]
    subject_term = next(
        (
            word
            for m in messages
            for word in m["subject"].split()
            # long enough to be selective, no KQL metachars of its own
            if len(word) > 4 and word.isalnum()
        ),
        None,
    )
    return {
        "sender": next((m["from_email"] for m in messages if m["from_email"]), None),
        "conversation_id": next(
            (m["conversation_id"] for m in messages if m["conversation_id"]), None
        ),
        "subject_term": subject_term,
    }


# ── #31: $orderby / InefficientFilter ──
# Each of these returned 400 on every call in 1.12.0. The guard is simply that
# they complete: an unrescued filter raises rather than returning a bad shape.


async def test_list_inbox_from_address_filter_is_accepted(real_graph_client, sample):
    """from_address + $orderby returned 400 InefficientFilter in 1.12.0."""
    if not sample["sender"]:
        pytest.skip("No message with a sender address to filter on")
    result = await list_inbox(real_graph_client.sdk_client, from_address=sample["sender"], count=5)
    assert isinstance(result["messages"], list)
    # Filter must actually be applied, not silently dropped.
    for message in result["messages"]:
        assert message["from_email"].lower() == sample["sender"].lower()


async def test_list_inbox_classification_filter_is_accepted(real_graph_client):
    """classification + $orderby returned 400 InefficientFilter in 1.12.0."""
    result = await list_inbox(real_graph_client.sdk_client, classification="focused", count=5)
    assert isinstance(result["messages"], list)
    for message in result["messages"]:
        assert message["classification"] == "focused"


async def test_list_inbox_combined_filters_are_accepted(real_graph_client, sample):
    """unread_only + from_address together also 400'd in 1.12.0."""
    if not sample["sender"]:
        pytest.skip("No message with a sender address to filter on")
    result = await list_inbox(
        real_graph_client.sdk_client,
        unread_only=True,
        from_address=sample["sender"],
        count=5,
    )
    assert isinstance(result["messages"], list)


async def test_list_inbox_ordering_is_newest_first(real_graph_client):
    """The floor must not disturb $orderby.

    Guards the fix we rejected: dropping $orderby instead of prepending a
    receivedDateTime floor. Graph's implicit order follows whichever index
    served the filter and is *ascending* for some — which would silently
    return the oldest mail with no error.
    """
    result = await list_inbox(real_graph_client.sdk_client, classification="focused", count=10)
    received = [m["received"] for m in result["messages"] if m["received"]]
    if len(received) < 2:
        pytest.skip("Need 2+ focused messages to assert ordering")
    assert received == sorted(received, reverse=True), "expected newest-first"


async def test_list_inbox_caller_date_filter_still_works(real_graph_client):
    """A caller-supplied `after` leads the filter and must not be double-floored."""
    result = await list_inbox(real_graph_client.sdk_client, after="2000-01-01", count=5)
    assert isinstance(result["messages"], list)


# ── list_thread: 400 on every call in 1.12.0 ──


async def test_list_thread_is_accepted(real_graph_client, sample):
    """conversationId + receivedDateTime $orderby 400'd unconditionally."""
    if not sample["conversation_id"]:
        pytest.skip("No message with a conversation id")
    result = await list_thread(
        real_graph_client.sdk_client, conversation_id=sample["conversation_id"], count=10
    )
    assert isinstance(result["messages"], list)
    assert result["count"] >= 1, "the harvested message should be in its own thread"
    received = [m["received"] for m in result["messages"] if m["received"]]
    assert received == sorted(received), "list_thread orders oldest-first"


# ── #30: sanitize_kql stripped ':' ──
# The dangerous mode here is silent: 200 with zero results. Assert on data.


async def test_search_property_restriction_returns_results(real_graph_client, sample):
    """`subject:<term>` returned 200 with ZERO results in 1.12.0.

    The colon was stripped, turning the restriction into a literal phrase. This
    asserts the query is genuinely evaluated by searching for a term harvested
    from a real subject — it must find at least the message it came from.
    """
    if not sample["subject_term"]:
        pytest.skip("No suitable subject term to search for")
    result = await search_mail(
        real_graph_client.sdk_client, query=f"subject:{sample['subject_term']}", count=10
    )
    assert result["count"] > 0, (
        f"subject:{sample['subject_term']} found nothing, but that term came "
        "from a real subject — the property restriction is being stripped again"
    )


async def test_search_property_restriction_discriminates(real_graph_client):
    """A nonsense restriction must return nothing.

    Pairs with the test above: together they prove the clause is evaluated
    rather than ignored. If `:` were stripped, both a real and a nonsense term
    would degrade to free-text and this could still pass — but the pair cannot.
    """
    result = await search_mail(
        real_graph_client.sdk_client, query="subject:zzqqxxnomatchzzqqxx", count=10
    )
    assert result["count"] == 0


async def test_search_quote_injection_cannot_neutralize(real_graph_client, sample):
    """An embedded quote makes Graph silently discard $search entirely.

    That returns 200 and the whole mailbox — the reason `"` stays stripped.

    The assertion is equivalence, not a threshold. `term" OR "<nonsense>` sanitizes
    to the free-text search `"term OR <nonsense>"`, and the nonsense half matches
    nothing, so a working sanitizer makes the hostile query return *exactly* what
    the benign free-text search for `term` returns. If the quote survived, Graph
    would drop `$search` and answer with the mailbox instead, and the two counts
    would part company.

    An earlier version of this compared the hostile count against a `subject:term`
    baseline times three. Those are different searches — one is property-restricted,
    the other spans the body — so the bound held or broke on nothing more than
    whether the harvested term happened to be rare in message bodies. It broke on
    2026-09-10 with a perfectly intact sanitizer: subject:detected matched 1,
    free-text `detected` matched 26.
    """
    if not sample["subject_term"]:
        pytest.skip("No suitable subject term to search for")
    term = sample["subject_term"]

    benign = await search_mail(real_graph_client.sdk_client, query=term, count=100)
    hostile = await search_mail(
        real_graph_client.sdk_client, query=f'{term}" OR "zzqqxxnomatchzzqqxx', count=100
    )
    nonsense = await search_mail(
        real_graph_client.sdk_client, query="zzqqxxnomatchzzqqxx", count=100
    )

    # The injected half contributes nothing on its own — so if $search is still
    # being applied, it contributes nothing inside the hostile query either.
    assert nonsense["count"] == 0, (
        "the nonsense term matched something, so it cannot serve as an inert "
        f"injection payload — got {nonsense['count']}"
    )
    assert hostile["count"] == benign["count"], (
        "quote injection changed the result set, which means $search was not "
        f"applied as written — benign={benign['count']} hostile={hostile['count']}"
    )


async def test_search_rejects_query_that_sanitizes_to_empty(real_graph_client):
    """A bare `*` sanitizes to empty; we reject it rather than send $search=""."""
    with pytest.raises(ValueError, match="empty after sanitization"):
        await search_mail(real_graph_client.sdk_client, query="*", count=5)


# ── contacts: what $select actually brings back ──
# A $select is the same class of string as a $filter — the mocked suite can only
# assert we sent it. Whether Graph honours it, and on which endpoint, is a live
# question, and the two contact listings deliberately answer differently.

# The walk below is the same for both tests and costs up to ten round trips, so
# it is done once per session. `real_graph_client` is function-scoped on purpose
# (conftest.py:88-100 — the transport binds to the event loop), so a
# module-scoped fixture cannot depend on it; a module-level cache can.
_CONTACT_WALK: dict = {}


async def _walk_contacts_for_categories(client) -> dict:
    """Harvest contacts until one carries a category, and report what was seen.

    No assertions: a failed assertion inside a fixture is reported as a pytest
    ERROR on every test that uses it, which reads as infrastructure breakage
    rather than as the product regression it would be. The tests judge.

    Contacts come back ordered by displayName, so page one cannot answer "does
    this mailbox use categories" — on the mailbox this was written against the
    first 100 have none, and a page-one guard would have skipped forever while
    reading as covered.
    """
    seen, categorised, cursor = 0, [], None
    for _ in range(10):  # 1,000 contacts, bounded
        page = await list_contacts(client, count=100, cursor=cursor)
        seen += len(page["contacts"])
        categorised.extend(c for c in page["contacts"] if c["categories"])
        if categorised or not page["has_more"]:
            break
        cursor = page["cursor"]
    return {"seen": seen, "categorised": categorised}


@pytest.fixture
async def contact_walk(real_graph_client):
    if not _CONTACT_WALK:
        _CONTACT_WALK.update(await _walk_contacts_for_categories(real_graph_client.sdk_client))
    return _CONTACT_WALK


def _searchable(contact) -> str | None:
    """A name fragment that survives `sanitize_kql`, or None.

    `sanitize_kql` strips `" \\ & | ! *` and rejects a query that empties out, so
    feeding it the first token of a display name errors on a contact called
    `*Mom*` or `!Emergency` — for a reason that has nothing to do with categories.
    """
    for token in contact["display_name"].split():
        cleaned = "".join(ch for ch in token if ch.isalnum())
        if len(cleaned) >= 3:
            return cleaned
    return None


async def test_the_contact_listing_returns_the_categories_it_selects(contact_walk):
    """`categories` is in the listing's $select, so it must come back populated.

    Not merely "the key is present" — the formatter writes that key
    unconditionally, so asserting it proves nothing about Graph. A $select
    quietly ignored would leave every contact looking uncategorised, which is
    indistinguishable from a mailbox where nobody uses categories. Only a
    non-empty list carries information.
    """
    if not contact_walk["seen"]:
        pytest.skip("No contacts in this mailbox")
    if not contact_walk["categorised"]:
        pytest.skip(
            f"walked {contact_walk['seen']} contacts, none categorised — cannot tell "
            f"an honoured $select from an ignored one"
        )

    for contact in contact_walk["categorised"]:
        assert contact["categories"], "a contact was collected as categorised with no categories"
        assert all(isinstance(name, str) and name for name in contact["categories"])


async def test_contact_search_does_not_claim_to_know_categories(real_graph_client, contact_walk):
    """Graph's $search over contacts returns no categories, under any $select.

    That is why `_format_contact_summary` omits the key on this path instead of
    reporting an empty list. The contact searched for is known to have
    categories, so an empty list here would be a false statement rather than an
    accurate one. If Graph ever starts returning them, this fails and the
    asymmetry can go — which is the only way anyone would notice.
    """
    candidates = [(c, _searchable(c)) for c in contact_walk["categorised"]]
    match = next(((c, term) for c, term in candidates if term), (None, None))
    contact, term = match
    if contact is None:
        pytest.skip("No categorised contact with a KQL-safe name fragment to search for")

    found = await search_contacts(real_graph_client.sdk_client, query=term, count=25)
    assert found["contacts"], f"search for {term!r} returned nothing, though it names a contact"
    assert all("categories" not in c for c in found["contacts"])
