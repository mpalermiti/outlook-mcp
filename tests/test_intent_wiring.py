"""``outlook_find`` wired into the server, in the opt-in ``intents`` group.

What loads by default does not change: ``intents`` is off by default, so a stock
install, and every selection that doesn't name it, registers what it did, sends
the same requests and reads the same instructions. A selection that names it
gets ``outlook_find``, read-only; the process asks Graph for immutable ids on
every request (decision D2); and the instructions gain one sentence about
``others_wrote``.
"""

from __future__ import annotations

import asyncio
import inspect
import json
import os
import subprocess
import sys
from types import SimpleNamespace

import pytest

from outlook_mcp import auth, server, toolsets
from outlook_mcp.config import Config
from outlook_mcp.intents import find as intent_find

OTHERS_WROTE_RULE = "anything under `others_wrote` was written by someone else"

_PROBE = """import asyncio, json
from mcp.client import Client
from outlook_mcp import server
from outlook_mcp.intents import ids
async def main():
    async with Client(server.mcp) as c:
        tools = {t.name: t.model_dump(exclude_none=True, by_alias=True)
                 for t in (await c.list_tools()).tools}
        instructions = c.instructions or ""
    print(json.dumps({"tools": tools, "instructions": instructions,
                      "immutable_ids": ids.IMMUTABLE_IDS}))
asyncio.run(main())
"""


def _process(selection: str | None) -> dict:
    env = {**os.environ}
    env.pop("PYTHONPATH", None)
    env.pop("OUTLOOK_MCP_TOOLSETS", None)
    if selection is not None:
        env["OUTLOOK_MCP_TOOLSETS"] = selection
    result = subprocess.run(
        [sys.executable, "-c", _PROBE],
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
        check=True,
    )
    return json.loads(result.stdout.strip().splitlines()[-1])


@pytest.fixture(scope="module")
def processes() -> dict[str | None, dict]:
    selections = [None, "classic,intents", "intents", "mail"]
    with __import__("concurrent.futures").futures.ThreadPoolExecutor(4) as pool:
        return dict(zip(selections, pool.map(_process, selections), strict=True))


def test_find_is_an_opt_in_read_only_intent_tool():
    assert toolsets.TOOL_GROUPS["outlook_find"] == toolsets.INTENTS
    assert toolsets.INTENTS in toolsets.OFF_BY_DEFAULT
    assert toolsets.INTENTS in toolsets.known_toolset_names()
    assert toolsets.annotation_for("outlook_find").read_only_hint is True


def test_a_default_process_is_what_it_was(processes):
    default = processes[None]
    assert "outlook_find" not in default["tools"]
    assert len(default["tools"]) == 71
    assert default["immutable_ids"] is False
    assert OTHERS_WROTE_RULE not in default["instructions"]


def test_classic_intents_is_the_default_plus_find(processes):
    default, both = processes[None], processes["classic,intents"]
    assert set(both["tools"]) == set(default["tools"]) | {"outlook_find"}
    assert {k: v for k, v in both["tools"].items() if k != "outlook_find"} == default["tools"]
    assert both["immutable_ids"] is True
    assert OTHERS_WROTE_RULE in both["instructions"]
    assert both["instructions"].startswith(default["instructions"])


def test_intents_alone_is_find_and_the_account_tools(processes):
    alone = processes["intents"]
    assert set(alone["tools"]) == {"outlook_auth_status", "outlook_whoami", "outlook_find"}
    assert alone["immutable_ids"] is True


def test_a_selection_without_intents_sends_nothing_new(processes):
    assert processes["mail"]["immutable_ids"] is False
    assert "outlook_find" not in processes["mail"]["tools"]


def test_the_advertised_tool(processes):
    find = processes["classic,intents"]["tools"]["outlook_find"]
    assert find["annotations"]["readOnlyHint"] is True
    properties = find["inputSchema"]["properties"]
    assert set(properties) == {
        "kind",
        "words",
        "sender",
        "after",
        "before",
        "window",
        "has_file",
        "folder",
        "task_list",
        "limit",
    }
    assert properties["kind"]["enum"] == list(intent_find.KINDS)
    assert set(find["inputSchema"]["required"]) == {"kind", "words"}
    assert len(find["description"]) <= 450


def test_the_tool_takes_the_implementations_arguments():
    tool = inspect.signature(server.outlook_find).parameters
    implementation = inspect.signature(intent_find.find_items).parameters
    assert set(tool) - {"ctx"} == {"kind", "words", "limit", *intent_find.APPLIES_TO}
    assert set(tool) - {"ctx"} <= set(implementation)
    assert tool["limit"].default == implementation["limit"].default


def test_the_tool_passes_the_zone_the_saved_username_and_the_loaded_tools(monkeypatch):
    seen: dict = {}

    async def fake_find(graph_client, **kwargs):
        seen.update(kwargs, graph_client=graph_client)
        return {"tool": "outlook_find"}

    graph = object()
    monkeypatch.setattr(intent_find, "find_items", fake_find)
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    monkeypatch.setattr(auth, "saved_username", lambda: "me@example.com")
    config = Config(timezone="America/Los_Angeles")
    ctx = SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )
    result = asyncio.run(
        server.outlook_find(ctx, kind="event", words="standup", window="today", limit=3)
    )
    assert result == {"tool": "outlook_find"}
    assert seen["graph_client"] is graph
    assert seen["zone"] == "America/Los_Angeles"
    assert seen["signed_in"] == "me@example.com"
    assert seen["loaded_tools"] == server.LOADED_TOOLS
    assert (seen["kind"], seen["words"], seen["window"], seen["limit"]) == (
        "event",
        "standup",
        "today",
        3,
    )
    assert seen["sender"] is None and seen["task_list"] is None


def test_the_saved_username_comes_from_the_sign_in_record(monkeypatch):
    monkeypatch.setattr(
        auth, "_load_auth_record", lambda: SimpleNamespace(username="Me@Example.com")
    )
    assert auth.saved_username() == "Me@Example.com"
    monkeypatch.setattr(auth, "_load_auth_record", lambda: None)
    assert auth.saved_username() is None
    monkeypatch.setattr(auth, "_load_auth_record", lambda: SimpleNamespace(username=""))
    assert auth.saved_username() is None
