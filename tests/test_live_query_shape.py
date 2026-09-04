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
    result = await list_inbox(
        real_graph_client.sdk_client, from_address=sample["sender"], count=5
    )
    assert isinstance(result["messages"], list)
    # Filter must actually be applied, not silently dropped.
    for message in result["messages"]:
        assert message["from_email"].lower() == sample["sender"].lower()


async def test_list_inbox_classification_filter_is_accepted(real_graph_client):
    """classification + $orderby returned 400 InefficientFilter in 1.12.0."""
    result = await list_inbox(
        real_graph_client.sdk_client, classification="focused", count=5
    )
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
    result = await list_inbox(
        real_graph_client.sdk_client, classification="focused", count=10
    )
    received = [m["received"] for m in result["messages"] if m["received"]]
    if len(received) < 2:
        pytest.skip("Need 2+ focused messages to assert ordering")
    assert received == sorted(received, reverse=True), "expected newest-first"


async def test_list_inbox_caller_date_filter_still_works(real_graph_client):
    """A caller-supplied `after` leads the filter and must not be double-floored."""
    result = await list_inbox(
        real_graph_client.sdk_client, after="2000-01-01", count=5
    )
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

    That returns 200 and the whole mailbox — the reason `"` stays stripped. The
    hostile query must stay bounded by the narrow one it is built from.
    """
    if not sample["subject_term"]:
        pytest.skip("No suitable subject term to search for")
    term = sample["subject_term"]
    narrow = await search_mail(
        real_graph_client.sdk_client, query=f"subject:{term}", count=25
    )
    hostile = await search_mail(
        real_graph_client.sdk_client, query=f'{term}" OR "zzqqxxnomatchzzqqxx', count=25
    )
    # If sanitization regresses, `hostile` becomes an unfiltered mailbox dump.
    assert hostile["count"] <= max(narrow["count"] * 3, 10), (
        "quote injection appears to have neutralized $search — "
        f"narrow={narrow['count']} hostile={hostile['count']}"
    )


async def test_search_rejects_query_that_sanitizes_to_empty(real_graph_client):
    """A bare `*` sanitizes to empty; we reject it rather than send $search=""."""
    with pytest.raises(ValueError, match="empty after sanitization"):
        await search_mail(real_graph_client.sdk_client, query="*", count=5)
