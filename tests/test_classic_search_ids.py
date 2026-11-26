"""Classic `$search` tools hand back the process's id family (decision D2, live round 1).

Live, a `$search` answers with regular ids whatever the request asked for, while
lists and filters answer with immutable ones once a process has loaded an intent
tool. So in that process `outlook_search_mail` and `outlook_search_contacts`
translate their hits' ids with one `POST /me/translateExchangeIds`, as `find` does.
A process that doesn't load `intents` sends exactly what it did: no translation,
no extra request, the ids Graph gave.
"""

import ast
from pathlib import Path

import pytest

from outlook_mcp.intents import ids
from outlook_mcp.intents.ids import IdFamily
from outlook_mcp.tools.contacts import search_contacts
from outlook_mcp.tools.mail_read import search_mail
from tests.intent_fakes import REST, FakeGraph, Mailbox, contact, received

TOOLS_DIR = Path(__file__).resolve().parent.parent / "src" / "outlook_mcp" / "tools"


@pytest.fixture
def graph():
    fake = FakeGraph()
    Mailbox(
        received=[
            received("m1", "sam@example.com", "Budget review"),
            received("m2", "pat@example.org", "Budget offer"),
        ],
        contacts=[contact("c1", "Sammy Davis", "sammy@example.net")],
    ).install(fake)
    return fake


def translations(fake):
    return [s for s in fake.operations if s.path == "/me/translateExchangeIds"]


# ── A process that loads `intents` ─────────────────────────────────────────


@pytest.fixture
def immutable(monkeypatch):
    monkeypatch.setattr(ids, "IMMUTABLE_IDS", True)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())


async def test_search_mail_hands_back_the_process_ids(graph, immutable):
    result = await search_mail(graph.graph_client().sdk_client, "budget")
    assert sorted(m["id"] for m in result["messages"]) == ["m1", "m2"]
    (call,) = translations(graph)
    assert sorted(call.body["inputIds"]) == [f"{REST}m1", f"{REST}m2"]
    assert call.body["sourceIdType"] == "restId"
    assert call.body["targetIdType"] == "restImmutableEntryId"


async def test_search_contacts_hands_back_the_process_ids(graph, immutable):
    result = await search_contacts(graph.graph_client().sdk_client, query="sammy")
    assert [c["id"] for c in result["contacts"]] == ["c1"]
    assert len(translations(graph)) == 1


async def test_a_failed_translation_keeps_the_ids_and_the_family_says_mixed(immutable):
    fake = FakeGraph()
    box = Mailbox(received=[received("m1", "sam@example.com", "Budget review")])
    box.translate_supported = False
    box.install(fake)
    ids.ID_FAMILY.observe('IdType="ImmutableId"')  # what the lists said
    result = await search_mail(fake.graph_client().sdk_client, "budget")
    assert [m["id"] for m in result["messages"]] == [f"{REST}m1"]
    assert ids.ID_FAMILY.family == "mixed"


async def test_no_translation_when_the_family_is_regular(graph, immutable):
    ids.ID_FAMILY.observe(None)  # Graph gives this mailbox regular ids everywhere
    result = await search_mail(graph.graph_client().sdk_client, "budget")
    assert sorted(m["id"] for m in result["messages"]) == [f"{REST}m1", f"{REST}m2"]
    assert translations(graph) == []


async def test_no_translation_for_an_empty_search(graph, immutable):
    result = await search_mail(graph.graph_client().sdk_client, "nothing-matches-this")
    assert result["messages"] == []
    assert translations(graph) == []


# ── A process that doesn't (the default) ───────────────────────────────────


async def test_by_default_search_mail_is_unchanged(graph):
    assert ids.IMMUTABLE_IDS is False
    result = await search_mail(graph.graph_client().sdk_client, "budget")
    assert sorted(m["id"] for m in result["messages"]) == [f"{REST}m1", f"{REST}m2"]
    assert [s.path for s in graph.operations] == ["/me/messages"]


async def test_by_default_search_contacts_is_unchanged(graph):
    result = await search_contacts(graph.graph_client().sdk_client, query="sammy")
    assert [c["id"] for c in result["contacts"]] == [f"{REST}c1"]
    assert [s.path for s in graph.operations] == ["/me/contacts"]


# ── Every classic `$search` goes through the translation ──────────────────


def _sends_search(tree: ast.AST) -> bool:
    for node in ast.walk(tree):
        if isinstance(node, ast.Dict) and any(
            isinstance(k, ast.Constant) and k.value == "$search" for k in node.keys
        ):
            return True
        if isinstance(node, ast.Subscript) and isinstance(node.slice, ast.Constant):
            if node.slice.value == "$search":
                return True
    return False


def _translates(tree: ast.AST) -> bool:
    return any(
        isinstance(node, ast.Attribute) and node.attr == "translate_search_ids"
        for node in ast.walk(tree)
    )


def test_every_classic_tool_that_searches_translates_its_ids():
    """A new classic `$search` must route its ids through `ids.translate_search_ids`,
    or an intents process hands out two id families again."""
    searching = {}
    for path in sorted(TOOLS_DIR.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        if _sends_search(tree):
            searching[path.name] = _translates(tree)
    assert searching == {"contacts.py": True, "mail_read.py": True}
