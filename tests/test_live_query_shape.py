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

import os
from datetime import datetime, timedelta
from datetime import timezone as dt_timezone

import pytest
from msgraph.generated.models.event_type import EventType
from msgraph.generated.models.free_busy_status import FreeBusyStatus

from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.sender_auth import NO_FROM, NO_HEADER, UNREADABLE
from outlook_mcp.tools._delta import parse_graph_datetime, require_graph_url
from outlook_mcp.tools.calendar_delta import list_events_delta
from outlook_mcp.tools.calendar_read import list_events
from outlook_mcp.tools.contacts import list_contacts, search_contacts
from outlook_mcp.tools.contacts_delta import list_contacts_delta
from outlook_mcp.tools.mail_delta import list_inbox_delta
from outlook_mcp.tools.mail_read import (
    list_folders,
    list_inbox,
    read_message,
    read_messages,
    search_mail,
)
from outlook_mcp.tools.mail_thread import list_thread
from outlook_mcp.tools.todo import get_task, list_tasks

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


# ── read_message by Message-ID: an exact $filter on internetMessageId ──
# A filter Graph rejects is a loud 400. The quiet failure is one that finds
# nothing: a Message-ID altered on the way, by quoting or by a `+` read as a
# space, matches no message and reads as "not in this mailbox".


async def test_read_message_finds_a_message_by_its_message_id(real_graph_client):
    """The lookup must land on the message the Message-ID was read from.

    Harvested with the SDK, because no tool returns `internetMessageId`. A
    Message-ID with `+`, `/` or `=` goes first: those mean something in a URL
    query, so finding its message shows the value reached Graph intact.
    """
    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    from outlook_mcp.pagination import build_request_config
    from outlook_mcp.tools.mail_read import read_message

    client = real_graph_client.sdk_client
    page = await client.me.messages.get(
        request_configuration=build_request_config(
            MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters,
            {"$select": "id,internetMessageId", "$top": 50},
        )
    )
    harvested = [m for m in page.value or [] if m.id and "@" in (m.internet_message_id or "")]
    if not harvested:
        pytest.skip("no message with a Message-ID among the first 50")
    harvested.sort(key=lambda m: not any(c in m.internet_message_id for c in "+/="))

    for message in harvested:
        try:
            result = await read_message(client, message.internet_message_id, concise=True)
        except ValueError as exc:
            if "matches" in str(exc):
                continue  # this one has a copy in another folder; try the next
            raise
        assert result["id"] == message.id, (
            "the Message-ID lookup returned a different message than the one it was read from"
        )
        return
    pytest.skip("every harvested Message-ID belongs to more than one copy")


async def test_an_unknown_message_id_is_not_found_rather_than_refused(real_graph_client):
    """An apostrophe, `&`, `#`, `%` and `+` must still make a filter Graph accepts.

    If the quoting or the encoding broke the literal, Graph answers 400, which
    arrives as GraphAPIError, not NotFoundError.
    """
    from outlook_mcp.errors import NotFoundError
    from outlook_mcp.tools.mail_read import read_message

    with pytest.raises(NotFoundError):
        await read_message(
            real_graph_client.sdk_client, "<zzqqxx'nomatch&a#b%20c+d/e=f@example.invalid>"
        )


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


# ── $search paging: Graph pages by $skiptoken, and stops a little past 250 ──
# Both failure modes are silent. No typed request builder has a `$skiptoken`
# field, so a cursor that loses the token fetches page one again with has_more
# still true. And Graph ends a long search with no nextLink and no error, which
# reads exactly like the end of a complete one.

# Matches every message, so the result size is the folder's size.
_SEARCH_EVERYTHING = "received>=2000-01-01"


@pytest.mark.parametrize("folder", [None, "inbox"])
async def test_search_cursor_moves_to_the_next_page(real_graph_client, folder):
    """Page two of a search must not be page one again, in all mail or in one folder.

    Until the cursor went out through `get_page`, it did: `build_request_config`
    dropped the `$skiptoken`, so every page was the first, with has_more true.
    Page two of a folder search goes to the folder the cursor carries.
    """
    client = real_graph_client.sdk_client
    first = await search_mail(client, query=_SEARCH_EVERYTHING, count=25, folder=folder)
    if not first["has_more"]:
        pytest.skip(f"the search fits on one page ({first['count']} messages)")

    second = await search_mail(client, query=_SEARCH_EVERYTHING, count=25, cursor=first["cursor"])

    assert second["count"] > 0, "Graph said there was more, and page two came back empty"
    seen = {m["id"] for m in first["messages"]}
    repeated = [m["id"] for m in second["messages"] if m["id"] in seen]
    assert not repeated, (
        f"page two repeated {len(repeated)} of page one's {len(seen)} messages: the "
        "$skiptoken did not reach Graph"
    )


async def test_a_search_graph_cut_short_says_so(real_graph_client):
    """Walk a populated folder's search and compare it with the folder's own count.

    Defaults to the largest folder in the first top-level listing (one GET).
    OUTLOOK_MCP_LIVE_SEARCH_FOLDER selects a name or ID instead, including a
    folder beyond that page. No mailbox writes or fixtures are needed.

    The folder's `totalItemCount` is the answer key: Graph counts it without
    searching. Exchange Online and Outlook.com have been observed stopping at 275
    (255-275 by page size), so a populated folder can end short and
    the last page must say so. The other direction is a claim about the
    constant: if Graph returns all of a folder past 300 and the tool still says
    `incomplete`, this account type's limit is not just past 250, and
    `SEARCH_STOPS_AFTER` needs its number. The 300 margin keeps a few messages
    the index has not caught up with from deciding either way.
    """
    from outlook_mcp.tools.mail_read import SEARCH_STOPS_AFTER

    client = real_graph_client.sdk_client
    requested = os.environ.get("OUTLOOK_MCP_LIVE_SEARCH_FOLDER", "").strip()
    if requested:
        folder_id = await resolve_folder_id(client, requested)
        folder = await client.me.mail_folders.by_mail_folder_id(folder_id).get()
        total = folder.total_item_count or 0
        folder_label = getattr(folder, "display_name", None) or requested
    else:
        listing = await list_folders(client)
        folders = listing["folders"]
        if not folders:
            pytest.skip(
                "No top-level mail folders; set OUTLOOK_MCP_LIVE_SEARCH_FOLDER to a name or ID"
            )
        folder = max(folders, key=lambda item: item["total"])
        folder_id, total = folder["id"], folder["total"]
        folder_label = folder.get("name") or folder_id
    if total <= 300:
        pytest.skip(
            f"folder {folder_label!r} holds {total} messages, "
            "too few to reach Graph's search limit; "
            "set OUTLOOK_MCP_LIVE_SEARCH_FOLDER to a folder holding more than 300"
        )

    ids: set[str] = set()
    walked, cursor = 0, None
    for _ in range(15):  # 1,500 results, bounded; Graph's documentation says up to 1,000
        page = await search_mail(
            client, query=_SEARCH_EVERYTHING, count=100, folder=folder_id, cursor=cursor
        )
        walked += page["count"]
        ids.update(m["id"] for m in page["messages"])
        cursor = page["cursor"]
        if cursor is None:
            break
    else:
        pytest.skip(
            f"folder {folder_label!r}: Graph was still paging after {walked} results; "
            "no limit to observe here"
        )

    assert len(ids) == walked, (
        f"folder {folder_label!r}: the walk returned {walked} results but only {len(ids)} ids"
    )
    if walked < total:
        assert "incomplete" in page, (
            f"Graph ended the search in folder {folder_label!r} at {walked} of {total} messages "
            "and the "
            "tool said nothing, so a caller takes the result as complete"
        )
    else:
        assert "incomplete" not in page, (
            f"Graph returned all {total} messages in folder {folder_label!r} "
            "and the tool still called the "
            f"search incomplete: this account type's limit is not just past "
            f"{SEARCH_STOPS_AFTER}, so SEARCH_STOPS_AFTER needs its number"
        )


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


async def test_a_contact_read_carries_its_change_key_where_the_merge_looks(real_graph_client):
    """`update_contact`'s `add_emails` / `remove_emails` pin their PATCH to the read.

    The offline tests mock a contact whose change key sits in
    `additional_data["@odata.etag"]`, which is where the merge reads it. A mock
    cannot say the SDK still puts it there: if msgraph moved or dropped it, the
    pin would silently become None, the patch would go out unpinned (#87's
    fallback), and every offline test would still pass. Read-only: one listing
    row and one plain GET of that contact, the GET the merge itself sends.
    """
    page = await list_contacts(real_graph_client.sdk_client, count=1)
    if not page["contacts"]:
        pytest.skip("No contacts in this mailbox")
    contact_id = page["contacts"][0]["id"]

    contact = await real_graph_client.sdk_client.me.contacts.by_contact_id(contact_id).get()

    extras = getattr(contact, "additional_data", None) or {}
    assert extras.get("@odata.etag"), (
        "the SDK no longer exposes a contact's change key at "
        'additional_data["@odata.etag"], so add_emails / remove_emails send an '
        f"unpinned patch. additional_data keys: {sorted(extras)}"
    )


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


# ── calendar ──────────────────────────────────────────────────────────


def _wide_window() -> tuple[str, str]:
    """±180 days around now, computed rather than hardcoded.

    A literal date range would keep passing for a while and then start skipping
    for a reason that has nothing to do with the code under test.
    """
    now = datetime.now(dt_timezone.utc)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    return (
        (now - timedelta(days=180)).strftime(fmt),
        (now + timedelta(days=180)).strftime(fmt),
    )


# One walk feeds both listing tests below, for the same reason the contacts one
# does: it costs up to five round trips and `real_graph_client` is
# function-scoped, so the cache is module-level rather than a fixture scope.
_EVENT_WALK: dict = {}


def _sdk_enum_values(enum) -> set[str]:
    """The values of an SDK enum, read off the SDK rather than typed here.

    A hand-typed list beside a derived one drifts — #73's review made the point
    about `FreeBusyStatus`, and `test_calendar_write.py` reads that same enum
    for the same reason. Deriving both sets here means the listing guard, the
    delta guard and the write path cannot disagree about what a valid value is,
    and a value Microsoft adds to an enum stops being a spurious failure.

    Call it once per test and keep the result: it is a fixed set, and rebuilding
    it inside a 500-event loop says the loop depends on something it does not.
    """
    return {member.value for member in enum}


async def _walk_events_for_select_fields(client):
    """Events from a wide window, so a quiet next fortnight doesn't skip this.

    A ±180-day window rather than the default 7: `list_events` is a calendarView
    and returns nothing at all for a caller with no upcoming meetings, which
    would let this skip forever while reading as covered — #63's page-one
    lesson, one resource over.

    Harvests `show_as` and `type` in the *same* walk. They are two fields of one
    `$select` and the failure they guard against is one failure, so a second
    walk would double the round trips to learn the same thing. It keeps going
    until both are populated rather than stopping at the first hit, or a
    mailbox that answered one field would mask the other.

    No assertions here: an assertion in a fixture surfaces a product regression
    as a pytest ERROR on unrelated tests, which reads as infrastructure
    breakage. The tests judge.
    """
    after, before = _wide_window()
    seen, cursor = 0, None
    with_status, with_type = [], []
    for _ in range(5):  # 500 events, bounded
        page = await list_events(client, after=after, before=before, count=100, cursor=cursor)
        seen += len(page["events"])
        with_status.extend(e for e in page["events"] if e.get("show_as"))
        with_type.extend(e for e in page["events"] if e.get("type"))
        if (with_status and with_type) or not page["has_more"]:
            break
        cursor = page["cursor"]
    return {"seen": seen, "with_status": with_status, "with_type": with_type}


@pytest.fixture
async def event_walk(real_graph_client):
    if not _EVENT_WALK:
        _EVENT_WALK.update(await _walk_events_for_select_fields(real_graph_client.sdk_client))
    return _EVENT_WALK


async def test_the_calendar_listing_says_which_calendars_can_be_written(real_graph_client):
    """`outlook_create_event(calendar=...)` refuses a calendar listed with
    `canEdit: false`, and lets one through when the listing says nothing.

    So a listing that stopped carrying `canEdit` would turn the refusal off
    without failing anything offline: every write would go to Graph and come
    back as the 403 the refusal exists to explain. Read-only: one listing.
    """
    from outlook_mcp.calendar_resolver import fetch_all_calendars

    calendars = await fetch_all_calendars(real_graph_client.sdk_client)

    assert calendars, "the account lists no calendars at all"
    assert all(isinstance(cal.can_edit, bool) for cal in calendars), [
        (cal.name, cal.can_edit) for cal in calendars
    ]
    defaults = [cal for cal in calendars if cal.is_default_calendar]
    assert defaults and all(cal.can_edit for cal in defaults)


async def test_the_event_key_lookup_is_accepted(real_graph_client):
    """`find_events_by_key` sends an `Any()` lambda over extended properties.

    A shape Graph refused would be a 400 on every lookup, and a lookup that
    fails is read by an agent as "no such booking", which is when it creates a
    duplicate. A key no event can carry must come back as an empty answer.
    Read-only: one filtered listing of the default calendar.
    """
    from outlook_mcp.tools.calendar_read import find_events_by_key

    result = await find_events_by_key(
        real_graph_client.sdk_client, "outlook-mcp:live-guard:no-event-has-this-key"
    )

    assert result["events"] == [] and result["has_more"] is False


async def test_the_event_key_expand_is_accepted_on_a_listing_and_a_read(real_graph_client):
    """The listing and `get_event` expand the key property with an inner `$filter`.

    `mail_read` builds its one expand of this kind as a raw URL because kiota
    once dropped the inner filter; these go through the typed builders. So
    this asks Graph to accept both shapes and checks a listed event's key reads
    back the same way on both paths. Read-only.
    """
    from outlook_mcp.tools.calendar_read import get_event

    after, before = _wide_window()
    page = await list_events(real_graph_client.sdk_client, after=after, before=before, count=10)
    if not page["events"]:
        pytest.skip("No events in the window to read back")
    listed = page["events"][0]
    detail = await get_event(real_graph_client.sdk_client, listed["id"])

    assert detail["key"] == listed.get("key")


async def test_an_events_attachments_are_listed_without_their_bytes(real_graph_client):
    """`get_event` lists attachments with a `$select` on `/events/{id}/attachments`.

    Without that `$select` Graph sends every file's `contentBytes`. A shape Graph
    refused would fail every read of an event with files; one it ignored would
    pass the offline suite while pulling whole files into each read. So the raw
    response is checked for `contentBytes` too. Read-only: reads at most 20
    events, and skips when none has a file.
    """
    from msgraph.generated.users.item.events.item.attachments.attachments_request_builder import (
        AttachmentsRequestBuilder,
    )

    from outlook_mcp.pagination import build_request_config
    from outlook_mcp.tools.calendar_read import _ATTACHMENT_SELECT, get_event

    sdk = real_graph_client.sdk_client
    after, before = _wide_window()
    page = await list_events(sdk, after=after, before=before, count=20)
    for listed in page["events"]:
        detail = await get_event(sdk, listed["id"])
        if detail["attachments"]:
            break
    else:
        pytest.skip("No event with attachments among the first 20 in the window")

    for attachment in detail["attachments"]:
        assert attachment["id"] and attachment["name"]
        assert isinstance(attachment["size"], int) and attachment["size"] > 0
        assert isinstance(attachment["is_inline"], bool)

    raw = await sdk.me.events.by_event_id(detail["id"]).attachments.get(
        request_configuration=build_request_config(
            AttachmentsRequestBuilder.AttachmentsRequestBuilderGetQueryParameters,
            {"$select": _ATTACHMENT_SELECT},
        )
    )
    assert raw.value and all(getattr(a, "content_bytes", None) is None for a in raw.value)


async def test_the_event_listing_returns_the_show_as_it_selects(event_walk):
    """`showAs` is in `list_events`' $select, so it must come back populated.

    Not "the key is present" — `_format_event_summary` writes that key
    unconditionally, so asserting presence proves nothing about Graph and would
    pass against the very bug this guards. Only a non-empty value carries
    information.

    This was not a hypothetical shape when it was written: the same listing's
    `$select` omitted `type` while the formatter read it, so `type` came back
    `""` on every call and nothing in the mock suite could see it (issue #69).
    Its sibling below is that case, now guarded the same way.
    """
    if not event_walk["seen"]:
        pytest.skip("No events in a ±180-day window — nothing to assert against")
    if not event_walk["with_status"]:
        pytest.fail(
            f"walked {event_walk['seen']} events and every one returned an empty show_as. "
            f"Graph defaults showAs to 'busy' on every event, so this is a $select "
            f"that was not honoured, not a mailbox without the data."
        )

    valid = _sdk_enum_values(FreeBusyStatus)
    for event in event_walk["with_status"]:
        assert event["show_as"] in valid, (
            f"Graph returned an unknown freeBusyStatus {event['show_as']!r}; "
            f"the write path's accepted set needs widening to match."
        )


async def test_the_event_listing_returns_the_type_it_selects(event_walk):
    """Issue #69 itself: the field the `$select` never asked for.

    `_format_event_summary` has read `event.type` since 1.16.0 and the listing's
    `$select` never named it, so Graph honoured the `$select`, the SDK left the
    attribute `None`, and every listed event reported `type: ""` — present,
    empty, and indistinguishable from a real value. Seven releases (1.16.0
    through 1.22.0), a CHANGELOG entry claiming the opposite, and a green
    offline suite throughout, because a mock cannot tell an honoured `$select`
    from an ignored one.

    All-empty is a `fail`, not a `skip`: Graph assigns every event a type
    (`singleInstance` at minimum), so there is no such thing as a mailbox whose
    events genuinely have none. An empty walk is the only honest skip.

    That `fail` is the load-bearing half. The membership check after it cannot
    fail on this path and is kept for its error text rather than its coverage:
    kiota maps a value outside `EventType` to `None`, `_format_event_summary`
    renders that `""`, and `with_type` has already filtered it out — so a value
    Microsoft invents arrives here as an absence, caught by the `fail` above.
    The identical-looking check in the delta test *is* load-bearing: that path
    reads raw JSON, so a new string reaches the assertion verbatim.

    The values a listing can actually carry are the sibling test below.
    """
    if not event_walk["seen"]:
        pytest.skip("No events in a ±180-day window — nothing to assert against")
    if not event_walk["with_type"]:
        pytest.fail(
            f"walked {event_walk['seen']} events and every one returned an empty type. "
            f"Graph assigns every event a type, so this is a $select that was not "
            f"honoured, not a mailbox without the data — issue #69, returned."
        )

    valid = _sdk_enum_values(EventType)
    for event in event_walk["with_type"]:
        assert event["type"] in valid, (
            f"Graph returned an unknown event type {event['type']!r}; the value set "
            f"named in SKILL.md, the README and this guard needs widening to match."
        )


async def test_the_event_listing_never_returns_a_series_master(event_walk):
    """`SKILL.md`, the README and the CHANGELOG tell an agent this; pin it.

    `calendarView` returns expanded instances, so the discriminator on a
    listing is `occurrence`/`exception` against `singleInstance`. An agent that
    filters a listing for `seriesMaster` to find its recurring meetings matches
    nothing and reports there are none — #69's silent wrong answer, one value
    over, which is why the docs now say so and why this holds them to it.

    Unlike the concise-omission guard rejected at the foot of this file, this
    one *can* fail: `_format_event_summary` reports whatever `type` Graph
    sends, so a master appearing on `calendarView` would carry through to the
    key and redden here.

    It is the *listing* that has this property, not calendarView-in-general.
    `/me/calendarView/delta` is a different endpoint and does return masters,
    which is why the test below deliberately does not forbid them. Measured
    over one ±180-day window: the listing gave 276 `occurrence`, 206
    `singleInstance`, 18 `exception` and no masters; the delta gave 212
    `singleInstance`, 94 `seriesMaster` and 94 `occurrence`, three of which
    were read back by id and confirmed as masters carrying a real recurrence.

    Reuses the cached walk, so this costs no extra round trips.
    """
    if not event_walk["seen"]:
        pytest.skip("No events in a ±180-day window — nothing to assert against")

    # Premise: without an expanded instance in the walk this mailbox has no
    # series in the window, and an absence of masters would prove nothing
    # (#63's page-one lesson). Skip loudly with the tally rather than pass.
    instances = [e for e in event_walk["with_type"] if e["type"] in {"occurrence", "exception"}]
    if not instances:
        pytest.skip(
            f"walked {event_walk['seen']} events, none of them an occurrence or exception "
            f"— no expanded series in this window, so an absence of masters says nothing"
        )

    masters = [e for e in event_walk["with_type"] if e["type"] == "seriesMaster"]
    assert not masters, (
        f"calendarView returned {len(masters)} seriesMaster item(s) among "
        f"{len(event_walk['with_type'])} typed events, e.g. {masters[0]['subject']!r}, "
        f"alongside {len(instances)} expanded instance(s). Graph has started returning "
        f"series masters on the listing, so SKILL.md, the README and the CHANGELOG are "
        f"now wrong to tell an agent that only `outlook_get_event` shows one."
    )


async def test_the_event_delta_carries_type_without_a_select(real_graph_client):
    """The premise behind `_format_event_delta` reading `type` at all.

    `/me/calendarView/delta` takes no `$select`, so the claim is that Graph
    returns `type` regardless. That is a claim about an external service, and
    those expire: #68's consumer-vs-work lesson is that what Graph does for one
    account type is not what it does for another, and the same field can change
    under you. (#70 reports exactly that happening to `isOnlineMeeting` — but
    it is **open**, this repo's live pin still asserts the old behaviour, and
    nothing here should be read as its ruling.) **If this test fails, Graph has
    stopped sending `type` on the delta endpoint**, and
    `outlook_list_events_delta` is silently reporting
    `""` for a field `outlook_list_events` still populates — the key-set parity
    test cannot see that, because the key would still be there.

    Tombstones are excluded deliberately: `format_delta_item` short-circuits a
    removed item to `{id, is_deleted: True}` without ever calling the
    formatter, so counting them here would assert summary parity against a
    shape that is exempt from it by design.

    `seriesMaster` is a value this path really does carry, unlike its non-delta
    sibling above — so nothing here forbids it, and the docs say to expect
    master ids that a listing-seeded map never held.
    """
    after, before = _wide_window()
    page = await list_events_delta(real_graph_client, start=after, end=before, page_size=100)

    live = [e for e in page["events"] if not e.get("is_deleted")]
    if not live:
        pytest.skip(
            f"delta returned {len(page['events'])} items in a ±180-day window, none of them "
            f"live events — nothing to assert against"
        )

    with_type = [e for e in live if e.get("type")]
    if not with_type:
        pytest.fail(
            f"walked {len(live)} live delta events and every one returned an empty type. "
            f"The delta endpoint sends no $select, so this means Graph has stopped "
            f"including `type` in the raw JSON — not a mailbox without the data."
        )

    valid = _sdk_enum_values(EventType)
    for event in with_type:
        assert event["type"] in valid, (
            f"Graph returned an unknown event type {event['type']!r} on the delta path. "
            f"This path reads raw JSON, so the value arrives verbatim — unlike the "
            f"listing guard, where kiota would have turned it into an absence."
        )


# There is deliberately no live guard for concise mode's *omission* of showAs or
# type. The obvious one — assert the key is absent from a live concise listing —
# cannot fail: `_format_event_concise` builds a fixed key list that never
# contains either, so the assertion holds whatever the `$select` asked Graph for,
# and would pass against the very regression it appears to guard. The honest
# place for that contract is the `$select` itself, which is a string we send and
# can therefore be asserted offline — and since #69 it is asserted generally,
# for every field, by the `event concise` row in
# `tests/test_select_covers_the_formatter.py` rather than one field at a time.
# Nothing is left for the live tier to see here, because an omitted field has
# no wire behaviour for Graph to get wrong — unlike the positive cases above,
# where a `$select` Graph silently ignores is invisible to a mock.


# ── outlook_changes_since: the first sync's window ──
# The digest's bootstrap asks message delta for `receivedDateTime ge <bound>`
# and keeps contacts by `lastModifiedDateTime`. Mocks see the URL we build and
# the field we read; only Graph can say the filter is honoured and the field is
# sent.


async def test_the_mail_bootstrap_honours_its_window_or_reports_a_fallback(real_graph_client):
    """Prove the digest's bound on accounts that honour or ignore Graph's filter.

    Read-only. Every observed live message outside the bound must be absent
    from the digest, and ignored filters must be reported explicitly.
    """
    from unittest.mock import patch

    from outlook_mcp.tools import digest

    bound = datetime.now(dt_timezone.utc).replace(microsecond=0) - timedelta(days=7)
    older = await list_inbox(
        real_graph_client.sdk_client, before=bound.strftime("%Y-%m-%dT%H:%M:%SZ"), count=1
    )
    if not older["messages"]:
        pytest.skip("No inbox message older than the bound — nothing to exclude")

    observed = []

    async def read_page(*args, **kwargs):
        page = await list_inbox_delta(*args, **kwargs)
        observed.extend(page["messages"])
        return page

    class FrozenDatetime(datetime):
        @classmethod
        def now(cls, tz=None):
            return bound + timedelta(days=7)

    from unittest.mock import AsyncMock

    with (
        patch.object(digest, "datetime", FrozenDatetime),
        patch.object(digest, "list_inbox_delta", new=read_page),
        patch.object(
            digest,
            "list_events_delta",
            new=AsyncMock(return_value={"events": [], "delta_token": None, "has_more": False}),
        ),
        patch.object(
            digest,
            "list_contacts_delta",
            new=AsyncMock(return_value={"contacts": [], "delta_token": None, "has_more": False}),
        ),
    ):
        result = await digest.changes_since(real_graph_client, fallback_window_hours=7 * 24)
    token = result["delta_tokens"]["mail"]
    assert token, "Graph returned no cursor for a filtered first sync"
    assert require_graph_url(token, source="live-check", resource="mail") == token
    live = [item for item in observed if not item.get("is_deleted")]
    recent = [
        item
        for item in live
        if (
            parse_graph_datetime(item.get("received") or "")
            or datetime.min.replace(tzinfo=dt_timezone.utc)
        )
        >= bound
    ]
    assert result["mail"]["new_count"] == len(recent)
    assert all(
        parse_graph_datetime(item["received_at"]) >= bound
        for item in result["mail"]["urgent_flagged"]
    )
    ignored = any(
        (stamp := parse_graph_datetime(item.get("received") or "")) is not None and stamp < bound
        for item in live
    )
    print(
        f"mail delta received-time filter: {'ignored; client fallback' if ignored else 'honoured'}"
    )
    assert ("mail_window_fallback" in result.get("_meta", {})) is ignored
    assert result["delta_tokens"].get("mail_bootstrap_from") == (
        bound.strftime("%Y-%m-%dT%H:%M:%SZ") if result["has_more"]["mail"] else None
    )


async def test_the_contacts_delta_carries_last_modified_without_a_select(real_graph_client):
    """The premise behind `list_contacts_delta(modified_since=...)`.

    `/me/contacts/delta` takes neither `$filter` nor `$select`, so the digest's
    first sync keeps contacts by the `lastModifiedDateTime` Graph is assumed to
    send on every one. If Graph stops sending it, every contact is dropped and
    the first digest reports none changed, with no error. With a bound in the
    distant past, the filter must keep exactly the contacts an unfiltered call
    returns.
    """
    everything = await list_contacts_delta(real_graph_client, page_size=100)
    if everything["has_more"]:
        pytest.skip("More contacts than one call returns; two snapshots would not line up")
    live = {c["id"] for c in everything["contacts"] if not c["is_deleted"]}
    if not live:
        pytest.skip("No contacts to read lastModifiedDateTime from")

    kept = await list_contacts_delta(
        real_graph_client,
        page_size=100,
        modified_since=datetime(2000, 1, 1, tzinfo=dt_timezone.utc),
    )

    dropped = live - {c["id"] for c in kept["contacts"] if not c["is_deleted"]}
    assert not dropped, (
        f"{len(dropped)} of {len(live)} contacts carry no readable lastModifiedDateTime on the "
        f"delta path, so the digest's first sync cannot place them in its window."
    )


# ── todo: $expand=checklistItems on a single-task GET ──


async def test_get_task_checklist_expand_is_accepted(real_graph_client):
    """get_task's $expand=checklistItems must come back, not 400 and not silently empty.

    The mock suite can only assert we write "checklistItems" into the query
    configuration; only Graph decides whether the expansion is legal on this
    endpoint and actually populates task.checklistItems.

    This tier is read-only, so it cannot create sub-steps; it scans the
    default list for a task that already has some and skips (loudly) when
    there is none — an unconditional pass against an empty checklist was the
    condition that let the 1.22.0 `.value` crash ship.
    """
    page = await list_tasks(real_graph_client.sdk_client, count=10)
    detail = None
    for task in page["tasks"]:
        candidate = await get_task(real_graph_client.sdk_client, task["id"])
        if candidate["checklist_count"] > 0:
            detail = candidate
            break
    if detail is None:
        pytest.skip(
            "no task in the first 10 of the default list has checklist items — "
            "the write tier (TestChecklistRoundTrip) owns the guarantee instead"
        )

    assert detail["checklist_count"] > 0
    for item in detail["checklist_items"]:
        assert item["id"]
        assert isinstance(item["display_name"], str)
        assert isinstance(item["is_checked"], bool)
    # Unchecked-first ordering is part of the tool's contract: an agent reads
    # the first unchecked item as "the next step". Rather than compare the
    # emission to its own sort (which can only re-verify Python's sorted()),
    # assert the contract directly — pairwise, no checked item precedes an
    # unchecked one and `created` never runs backwards within a group.
    items = detail["checklist_items"]
    for earlier, later in zip(items, items[1:]):
        assert (earlier["is_checked"], earlier["created"]) <= (
            later["is_checked"],
            later["created"],
        ), f"checklist order broke the (is_checked, created) key: {items}"
    # And the load-bearing half for an agent loop: two identical calls on an
    # unchanged task emit the same order — a nondeterministic "next step" is
    # the bug this contract exists to prevent.
    again = await get_task(real_graph_client.sdk_client, detail["id"])
    assert [(i["is_checked"], i["created"]) for i in again["checklist_items"]] == [
        (i["is_checked"], i["created"]) for i in items
    ]


async def _a_contact_with_a_phone(client):
    """A listed contact whose summary phone has 7+ significant digits, or None."""
    from outlook_mcp.tools.contacts import _MIN_PHONE_DIGITS, _phone_digits

    cursor = None
    for _ in range(10):  # 1,000 contacts, bounded
        page = await list_contacts(client, count=100, cursor=cursor)
        for contact in page["contacts"]:
            if len(_phone_digits(contact["phone"])) >= _MIN_PHONE_DIGITS:
                return contact
        if not page["has_more"]:
            return None
        cursor = page["cursor"]
    return None


async def test_a_phone_search_finds_a_stored_number_however_it_is_written(real_graph_client):
    """`phone=` reads the listing in pages of 1,000 and compares digits.

    Two things only Graph can answer: that the listing accepts this `$top` with the
    summary `$select`, and that the pages are walked to the contact. The number is
    searched as its bare digits and as its last seven, neither of which is how it
    is stored. Read-only.
    """
    contact = await _a_contact_with_a_phone(real_graph_client.sdk_client)
    if contact is None:
        pytest.skip("No listed contact has a phone number of seven digits or more")
    digits = "".join(ch for ch in contact["phone"] if ch.isdigit())

    for written in (digits, digits[-7:]):
        found = await search_contacts(real_graph_client.sdk_client, phone=written, count=100)
        assert contact["id"] in {c["id"] for c in found["contacts"]}, (
            f"phone search missed a stored number written as {len(written)} digits; "
            f"{found['count']} other match(es)"
        )


async def test_contact_search_still_misses_a_part_of_a_number(real_graph_client):
    """The reason `phone=` exists, pinned: Graph's `$search` does not find a stored
    number from its last seven digits. If it starts to, the scan can go."""
    contact = await _a_contact_with_a_phone(real_graph_client.sdk_client)
    if contact is None:
        pytest.skip("No listed contact has a phone number of seven digits or more")
    digits = "".join(ch for ch in contact["phone"] if ch.isdigit())
    if len(digits) <= 7:
        pytest.skip("The harvested number has no more than seven digits")

    found = await search_contacts(real_graph_client.sdk_client, query=digits[-7:], count=25)
    assert contact["id"] not in {c["id"] for c in found["contacts"]}, (
        "Graph's contact $search now finds a number from its last digits; "
        "outlook_search_contacts(phone=...) may no longer need to read the listing"
    )


# ── #130: the opt-in read's $select ──
# Once `$select` names anything, Graph returns only what it names. A field the
# opt-in select forgot would come back empty only for a caller who asked for
# sender authentication, and a mock returns whatever it is handed.


async def test_the_sender_auth_read_returns_what_the_default_read_does(real_graph_client):
    page = await list_inbox(real_graph_client.sdk_client, count=5, concise=True)
    if not page["messages"]:
        pytest.skip("Mailbox inbox is empty — nothing to read")
    for summary in page["messages"]:
        default = await read_message(real_graph_client.sdk_client, summary["id"])
        opted_in = await read_message(
            real_graph_client.sdk_client, summary["id"], include_sender_auth=True
        )
        opted_in.pop("sender_authentication")
        assert opted_in == default


async def test_microsofts_own_header_parses_on_real_mail(real_graph_client):
    """Microsoft's stamp always parses; a reading of `unreadable` is a gap in the parser.

    Also checks the `$batch` path selects the same headers: its sub-request URLs
    carry the `$select` as text, where the single read goes through kiota.
    """
    page = await list_inbox(real_graph_client.sdk_client, count=10, concise=True)
    ids = [m["id"] for m in page["messages"]]
    if not ids:
        pytest.skip("Mailbox inbox is empty — nothing to read")
    single = [
        (await read_message(real_graph_client.sdk_client, i, include_sender_auth=True))[
            "sender_authentication"
        ]
        for i in ids
    ]
    batched = await read_messages(real_graph_client, ids, include_sender_auth=True)
    assert [m["sender_authentication"] for m in batched["messages"]] == single

    reasons = [v["reason"] for v in single]
    assert UNREADABLE not in reasons, reasons
    if all(r in (NO_HEADER, NO_FROM) for r in reasons):
        pytest.skip(f"none of the newest {len(ids)} messages carries Microsoft's header")
