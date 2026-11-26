"""One id family per process (contracts §6.3, decision D2), and the counted raw path.

A process that loads intent tools asks Graph for immutable ids on every request,
and learns from ``Preference-Applied`` which family it actually got. The header
is combined with any other preference, never assigned over it. ``IntentGraph``
is the one raw path the intent tools use: it adds the header, counts each
request against the call's budget before it leaves, holds the limiter, and turns
a refusal into a ``GraphAPIError``.
"""

import logging

import pytest

from outlook_mcp.errors import GraphAPIError, UntrustedURLError
from outlook_mcp.intents import ids, transport
from outlook_mcp.intents.budget import BudgetExhausted, RequestBudget, WeightedLimiter
from outlook_mcp.intents.ids import IdFamily, combine_prefer, immutable_applied
from outlook_mcp.intents.transport import IntentGraph
from tests.intent_fakes import FakeGraph, error, page

# ── The Prefer header ─────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("existing", "combined"),
    [
        (None, 'IdType="ImmutableId"'),
        ("", 'IdType="ImmutableId"'),
        ("odata.maxpagesize=50", 'odata.maxpagesize=50, IdType="ImmutableId"'),
        ('outlook.timezone="UTC"', 'outlook.timezone="UTC", IdType="ImmutableId"'),
        ('IdType="ImmutableId"', 'IdType="ImmutableId"'),
        ('odata.maxpagesize=5, idtype="immutableid"', 'odata.maxpagesize=5, idtype="immutableid"'),
    ],
)
def test_prefer_is_combined_never_replaced(existing, combined):
    assert combine_prefer(existing) == combined


@pytest.mark.parametrize(
    ("header", "applied"),
    [
        ('IdType="ImmutableId"', True),
        ('odata.maxpagesize=50, IdType="ImmutableId"', True),
        ("idtype=immutableid", True),
        ("odata.maxpagesize=50", False),
        ('IdType="RestId"', False),
        (None, False),
        ("", False),
    ],
)
def test_preference_applied_is_read_from_the_list(header, applied):
    assert immutable_applied(header) is applied


# ── The family ────────────────────────────────────────────────────────────


def test_id_family_applied():
    family = IdFamily()
    assert family.family is None
    family.observe('IdType="ImmutableId"')
    family.observe('IdType="ImmutableId"')
    assert family.family == "immutable"


def test_id_family_ignored():
    family = IdFamily()
    family.observe(None)
    assert family.family == "rest"


def test_id_family_mixed_is_logged_and_sticks(caplog):
    family = IdFamily()
    family.observe('IdType="ImmutableId"')
    with caplog.at_level(logging.WARNING, logger="outlook_mcp.intents.ids"):
        family.observe(None)
    assert family.family == "mixed"
    assert "mixed" in caplog.text
    family.observe('IdType="ImmutableId"')
    assert family.family == "mixed"


# ── IntentGraph ───────────────────────────────────────────────────────────


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    return graph


def _graph(fake, limit=30, family=None, limiter=None):
    return IntentGraph(
        fake.graph_client().credential,
        RequestBudget(limit),
        id_family=family or IdFamily(),
        limiter=limiter or WeightedLimiter(),
    )


async def test_a_get_carries_the_header_and_learns_the_family(fake):
    fake.on("/me/messages", lambda seen, m: page([{"id": "a"}]))
    family = IdFamily()
    async with _graph(fake, family=family) as graph:
        body = await graph.get("/me/messages", {"$top": 1})
    assert body["value"] == [{"id": "a"}]
    assert fake.seen[0].headers["prefer"] == 'IdType="ImmutableId"'
    assert fake.seen[0].headers["authorization"] == "Bearer not-a-real-token"
    assert family.family == "immutable"


async def test_a_reply_without_preference_applied_is_the_rest_family(fake):
    fake.preference_applied = None
    fake.on("/me/messages", lambda seen, m: page([]))
    family = IdFamily()
    async with _graph(fake, family=family) as graph:
        await graph.get("/me/messages")
    assert family.family == "rest"


async def test_a_resource_without_immutable_ids_does_not_vote(fake):
    fake.preference_applied = None
    fake.on("/me/todo/lists/L/tasks", lambda seen, m: page([]))
    family = IdFamily()
    async with _graph(fake, family=family) as graph:
        await graph.get("/me/todo/lists/L/tasks", ids=False)
    assert family.family is None
    assert fake.seen[0].headers["prefer"] == 'IdType="ImmutableId"'


async def test_a_batch_carries_the_header_on_the_envelope_and_every_sub_request(fake):
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)})
    family = IdFamily()
    async with _graph(fake, family=family) as graph:
        replies = await graph.batch(["/me/messages/a", "/me/messages/b", "/me/messages/zz"])
    assert [r["body"]["id"] for r in replies] == ["a", "b", "zz"]
    envelope = next(s for s in fake.seen if s.path == "/$batch")
    assert envelope.headers["prefer"] == 'IdType="ImmutableId"'
    subs = [s for s in fake.seen if s.in_batch]
    assert len(subs) == 3
    assert all(s.headers["prefer"] == 'IdType="ImmutableId"' for s in subs)
    assert family.family == "immutable"


async def test_a_batch_over_twenty_goes_in_more_than_one_envelope(fake):
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)})
    async with _graph(fake) as graph:
        replies = await graph.batch([f"/me/messages/m{n}" for n in range(25)])
    assert [r["body"]["id"] for r in replies] == [f"m{n}" for n in range(25)]
    assert sum(1 for s in fake.seen if s.path == "/$batch") == 2


async def test_each_request_is_counted_before_it_leaves(fake):
    fake.on("/me/messages", lambda seen, m: page([]))
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)})
    budget = RequestBudget(3)
    async with IntentGraph(fake.graph_client().credential, budget) as graph:
        await graph.get("/me/messages")
        with pytest.raises(BudgetExhausted):
            await graph.batch(["/me/messages/a", "/me/messages/b", "/me/messages/c"])
        await graph.batch(["/me/messages/a", "/me/messages/b"])
    assert budget.spent == 3
    assert fake.counted == 3


async def test_a_refusal_is_a_graph_api_error_with_graphs_code(fake):
    fake.on("/me/messages", lambda seen, m: error(400, "BadRequest", "Syntax error"))
    async with _graph(fake) as graph:
        with pytest.raises(GraphAPIError) as refused:
            await graph.get("/me/messages", {"$search": '"x"'})
    assert refused.value.status_code == 400
    assert refused.value.error_code == "BadRequest"


async def test_a_next_link_off_graph_is_refused_before_the_token_goes(fake):
    async with _graph(fake) as graph:
        with pytest.raises(UntrustedURLError):
            await graph.follow("https://evil.example.com/v1.0/me/messages?$skiptoken=x")
    assert fake.seen == []


async def test_a_next_link_is_followed_and_counted(fake):
    fake.on("/me/messages", lambda seen, m: page([{"id": seen.params.get("$skiptoken")}]))
    budget = RequestBudget(5)
    async with IntentGraph(fake.graph_client().credential, budget) as graph:
        body = await graph.follow("https://graph.microsoft.com/v1.0/me/messages?$skiptoken=p2")
    assert body["value"] == [{"id": "p2"}]
    assert budget.spent == 1


async def test_queries_sent_are_recorded_for_understood(fake):
    fake.on("/me/messages", lambda seen, m: page([]))
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)})
    async with _graph(fake) as graph:
        await graph.get("/me/messages", {"$search": '"budget"', "$top": 25})
        await graph.get("/me/messages", {"$filter": "isRead eq false"})
        await graph.batch(["/me/messages?$filter=receivedDateTime ge 1900-01-01T00:00:00Z"])
    assert graph.queries == [
        '"budget"',
        "isRead eq false",
        "receivedDateTime ge 1900-01-01T00:00:00Z",
    ]


async def test_requests_hold_the_limiter(fake):
    fake.on("/me/messages", lambda seen, m: page([]))
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)})
    limiter = WeightedLimiter()
    async with _graph(fake, limiter=limiter) as graph:
        await graph.get("/me/messages")
        assert limiter.peak == 1
        await graph.batch(["/me/messages/a"])
        assert limiter.peak == 4
    assert limiter.in_use == 0


async def test_sdk_requests_through_the_counted_client_spend_and_hold(fake):
    fake.on("/me/mailFolders", lambda seen, m: page([{"id": "f1", "displayName": "Travel"}]))
    budget = RequestBudget(1)
    limiter = WeightedLimiter()
    client = transport.counted(fake.graph_client().sdk_client, budget, limiter)
    folders = await client.me.mail_folders.get()
    assert folders.value[0].display_name == "Travel"
    assert budget.spent == 1 and limiter.peak == 1
    with pytest.raises(BudgetExhausted):
        await client.me.mail_folders.get()
    assert fake.counted == 1


def test_the_process_family_is_one_object():
    assert ids.ID_FAMILY is ids.ID_FAMILY
    assert isinstance(ids.ID_FAMILY, IdFamily)


# ── The process switch: every Outlook request in a process that loads intents ──


@pytest.fixture
def switched_on(monkeypatch):
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", True)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())


def test_the_switch_is_off_unless_an_intent_tool_loads():
    assert ids.wanted_for({"outlook_list_inbox", "outlook_whoami"}) is False
    assert ids.wanted_for({"outlook_list_inbox", "outlook_find"}) is True
    assert ids.wanted_for(set()) is False


def test_off_the_headers_are_left_exactly_as_they_were(monkeypatch):
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", False)
    headers = {"Prefer": "odata.maxpagesize=50", "Accept": "application/json"}
    assert ids.with_preference(headers) is headers
    assert ids.with_preference(None) == {}


def test_on_the_preference_is_combined_into_the_headers(switched_on):
    assert ids.with_preference({"Prefer": "odata.maxpagesize=50"}) == {
        "Prefer": 'odata.maxpagesize=50, IdType="ImmutableId"'
    }
    assert ids.with_preference({"Accept": "x"}) == {
        "Accept": "x",
        "Prefer": 'IdType="ImmutableId"',
    }
    # Any spelling of the header name is the same header.
    assert ids.with_preference({"prefer": 'outlook.body-content-type="text"'}) == {
        "prefer": 'outlook.body-content-type="text", IdType="ImmutableId"'
    }


@pytest.mark.parametrize(
    ("url", "carries"),
    [
        ("https://graph.microsoft.com/v1.0/me/messages?$top=1", True),
        ("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages", True),
        ("https://graph.microsoft.com/v1.0/me/messages/AAMk/attachments", True),
        ("https://graph.microsoft.com/v1.0/me/calendars/c1/calendarView", True),
        ("https://graph.microsoft.com/v1.0/me/events/e1", True),
        ("https://graph.microsoft.com/v1.0/me/contacts", True),
        ("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta", True),
        ("https://graph.microsoft.com/v1.0/me/mailFolders", False),
        ("https://graph.microsoft.com/v1.0/me/calendars", False),
        ("https://graph.microsoft.com/v1.0/me/todo/lists/l/tasks", False),
        ("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules", False),
        ("https://graph.microsoft.com/v1.0/me", False),
        ("https://graph.microsoft.com/v1.0/$batch", False),
    ],
)
def test_which_replies_speak_for_the_family(url, carries):
    assert ids.carries_outlook_ids(url) is carries


def _sdk_over(fake):
    from outlook_mcp.graph import GraphClient

    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(ids, "base_client", fake.http_client)
        return GraphClient(fake.graph_client().credential)


async def test_on_every_sdk_request_asks_and_the_reply_is_heard(switched_on):
    graph = FakeGraph()
    graph.on("/me/messages", lambda seen, m: page([{"id": "m1"}]))
    client = _sdk_over(graph)
    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    from outlook_mcp.pagination import build_request_config

    config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, {"$top": 1}
    )
    config.headers.add("Prefer", 'outlook.body-content-type="text"')
    await client.sdk_client.me.messages.get(request_configuration=config)
    assert (
        graph.seen[0].headers["prefer"] == 'outlook.body-content-type="text", IdType="ImmutableId"'
    )
    assert ids.ID_FAMILY.family == "immutable"


async def test_on_a_reply_without_ids_does_not_vote(switched_on):
    graph = FakeGraph()
    graph.preference_applied = None
    graph.on("/me/mailFolders", lambda seen, m: page([{"id": "f", "displayName": "F"}]))
    client = _sdk_over(graph)
    await client.sdk_client.me.mail_folders.get()
    assert graph.seen[0].headers["prefer"] == 'IdType="ImmutableId"'
    assert ids.ID_FAMILY.family is None


async def test_off_an_sdk_request_carries_nothing_new(monkeypatch):
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", False)
    graph = FakeGraph()
    graph.on("/me/messages", lambda seen, m: page([]))
    client = _sdk_over(graph)
    # Off, the server builds the SDK's own default client: the stand-in is never used.
    assert client.sdk_client.request_adapter._http_client is not None
    calls = []
    monkeypatch.setattr(ids, "base_client", lambda: calls.append(1))
    from outlook_mcp.graph import GraphClient

    GraphClient(graph.graph_client().credential)
    assert calls == []


# ── The classic tools' raw paths, in a process that loads intents ────────

_DELTA = "https://graph.microsoft.com/v1.0/me/calendarView/delta"


def _raw_over(fake, monkeypatch, module):
    import httpx

    real = httpx.AsyncClient
    monkeypatch.setattr(
        f"{module}.httpx.AsyncClient",
        lambda *a, **k: real(transport=httpx.MockTransport(fake.handle)),
    )


@pytest.mark.parametrize("on", [True, False])
async def test_a_delta_page_combines_the_preference_with_the_page_size(monkeypatch, on):
    from outlook_mcp.tools._delta import fetch_delta_pages

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", on)
    fake = FakeGraph()
    fake.on(
        "/me/calendarView/delta",
        lambda seen, m: {"value": [], "@odata.deltaLink": f"{_DELTA}?$deltatoken=t"},
    )
    _raw_over(fake, monkeypatch, "outlook_mcp.tools._delta")
    await fetch_delta_pages(
        fake.graph_client().credential,
        initial_url=f"{_DELTA}?startDateTime=2026-11-01T00:00:00Z&endDateTime=2026-12-01T00:00:00Z",
        delta_token=None,
        page_size=50,
        resource="calendar",
        headers={"Prefer": "odata.maxpagesize=50"},
    )
    expected = 'odata.maxpagesize=50, IdType="ImmutableId"' if on else "odata.maxpagesize=50"
    assert fake.seen[0].headers["prefer"] == expected


@pytest.mark.parametrize("on", [True, False])
async def test_a_bulk_read_asks_on_the_envelope_and_each_sub_request(monkeypatch, on):
    from outlook_mcp.tools.mail_read import read_messages

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", on)
    fake = FakeGraph()
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1), "subject": "s"})
    _raw_over(fake, monkeypatch, "outlook_mcp.tools.mail_read")
    await read_messages(fake.graph_client(), ["m1", "m2"], format="text")
    envelope = next(s for s in fake.seen if s.path == "/$batch")
    subs = [s for s in fake.seen if s.in_batch]
    text = 'outlook.body-content-type="text"'
    if on:
        assert envelope.headers["prefer"] == 'IdType="ImmutableId"'
        assert {s.headers["prefer"] for s in subs} == {f'{text}, IdType="ImmutableId"'}
    else:
        assert "prefer" not in envelope.headers
        assert {s.headers["prefer"] for s in subs} == {text}
        # Off, a sub-request is exactly what it was: no headers key added where none was.
        assert [s.body for s in fake.seen if s.path == "/$batch"][0]["requests"][0]["headers"] == {
            "Prefer": text
        }


async def test_a_bulk_read_without_its_own_headers_gains_none_when_off(monkeypatch):
    from outlook_mcp.tools.mail_read import read_messages

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", False)
    fake = FakeGraph()
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1), "subject": "s"})
    _raw_over(fake, monkeypatch, "outlook_mcp.tools.mail_read")
    await read_messages(fake.graph_client(), ["m1"], format="html")
    request = next(s for s in fake.seen if s.path == "/$batch").body["requests"][0]
    assert "headers" not in request


@pytest.mark.parametrize("on", [True, False])
async def test_batch_triage_asks_on_each_sub_request(monkeypatch, on):
    from outlook_mcp.config import Config
    from outlook_mcp.tools.batch import batch_triage

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", on)
    fake = FakeGraph()
    fake.on(r"/me/messages/(\w+)", lambda seen, m: {"id": m.group(1)}, method="PATCH")
    # On, the server's own client with its hooks; off, a plain SDK client over the fake.
    sdk = _sdk_over(fake).sdk_client if on else fake.graph_client().sdk_client
    await batch_triage(sdk, ["m1", "m2"], "mark_read", "true", config=Config())
    subs = [s for s in fake.seen if s.in_batch]
    assert len(subs) == 2
    for sub in subs:
        assert sub.headers["content-type"] == "application/json"
        assert ("prefer" in sub.headers) is on
    envelope = next(s for s in fake.seen if s.path == "/$batch")
    assert ("immutableid" in envelope.headers.get("prefer", "").lower()) is on
