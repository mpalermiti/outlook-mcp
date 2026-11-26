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
