"""The intent tools wired into the server, in the opt-in ``intents`` group.

What loads by default does not change: ``intents`` is off by default, so a stock
install, and every selection that doesn't name it, registers what it did, sends
the same requests and reads the same instructions. A selection that names it
gets the intent tools (``outlook_find``, ``outlook_agenda``,
``outlook_open_loops``, ``outlook_catch_up``, ``outlook_brief``), read-only; the
process asks Graph for immutable ids on every request (decision D2); and the
instructions gain one sentence about ``others_wrote``.
"""

from __future__ import annotations

import asyncio
import inspect
import json
import os
import re
import subprocess
import sys
from types import SimpleNamespace

import pytest

from outlook_mcp import auth, server, toolsets
from outlook_mcp.config import Config
from outlook_mcp.intents import agenda as intent_agenda
from outlook_mcp.intents import brief as intent_brief
from outlook_mcp.intents import catch_up as intent_catch_up
from outlook_mcp.intents import find as intent_find
from outlook_mcp.intents import open_loops as intent_open_loops

OTHERS_WROTE_RULE = "anything under `others_wrote` was written by someone else"
INTENT_TOOLS = {n for n, g in toolsets.TOOL_GROUPS.items() if g == toolsets.INTENTS}

_PROBE = """import asyncio, json
from mcp.client import Client
from outlook_mcp import server
from outlook_mcp.intents import ids
async def main():
    async with Client(server.mcp) as c:
        tools = {t.name: t.model_dump(exclude_none=True, by_alias=True)
                 for t in (await c.list_tools()).tools}
        instructions = c.instructions or ""
        prompts = {}
        for p in (await c.list_prompts()).prompts:
            got = await c.get_prompt(p.name, {})
            prompts[p.name] = " ".join(m.content.text for m in got.messages)
    print(json.dumps({"tools": tools, "instructions": instructions, "prompts": prompts,
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


@pytest.mark.parametrize(
    "tool",
    ["outlook_find", "outlook_agenda", "outlook_open_loops", "outlook_catch_up", "outlook_brief"],
)
def test_each_is_an_opt_in_read_only_intent_tool(tool):
    assert toolsets.TOOL_GROUPS[tool] == toolsets.INTENTS
    assert toolsets.INTENTS in toolsets.OFF_BY_DEFAULT
    assert toolsets.INTENTS in toolsets.known_toolset_names()
    assert toolsets.annotation_for(tool).read_only_hint is True


def test_a_default_process_is_what_it_was(processes):
    default = processes[None]
    assert not INTENT_TOOLS & set(default["tools"])
    assert len(default["tools"]) == 71
    assert default["immutable_ids"] is False
    assert OTHERS_WROTE_RULE not in default["instructions"]


def test_classic_intents_is_the_default_plus_the_intent_tools(processes):
    default, both = processes[None], processes["classic,intents"]
    assert set(both["tools"]) == set(default["tools"]) | INTENT_TOOLS
    assert {k: v for k, v in both["tools"].items() if k not in INTENT_TOOLS} == default["tools"]
    assert both["immutable_ids"] is True
    assert OTHERS_WROTE_RULE in both["instructions"]
    assert both["instructions"].startswith(default["instructions"])


def test_intents_alone_is_the_intent_tools_their_core_and_the_account_tools(processes):
    alone = processes["intents"]
    assert set(alone["tools"]) == (
        {"outlook_auth_status", "outlook_whoami"} | INTENT_TOOLS | toolsets.FOLLOW_UP_CORE
    )
    assert alone["immutable_ids"] is True


def test_a_selection_without_intents_sends_nothing_new(processes):
    assert processes["mail"]["immutable_ids"] is False
    assert not INTENT_TOOLS & set(processes["mail"]["tools"])


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


# ── outlook_agenda ────────────────────────────────────────────────────────


def test_the_advertised_agenda(processes):
    agenda = processes["classic,intents"]["tools"]["outlook_agenda"]
    assert agenda["annotations"]["readOnlyHint"] is True
    properties = agenda["inputSchema"]["properties"]
    assert set(properties) == {
        "window",
        "zone",
        "day_start",
        "day_end",
        "calendars",
        "min_free",
        "include_tasks",
    }
    assert "required" not in agenda["inputSchema"] or not agenda["inputSchema"]["required"]
    assert len(agenda["description"]) <= 450


def test_the_agenda_takes_the_implementations_arguments():
    tool = inspect.signature(server.outlook_agenda).parameters
    implementation = inspect.signature(intent_agenda.read_agenda).parameters
    assert set(tool) - {"ctx"} <= set(implementation)
    for name in set(tool) - {"ctx"}:
        assert tool[name].default == implementation[name].default, name


def test_the_agenda_passes_the_zone_the_saved_username_and_the_loaded_tools(monkeypatch):
    seen: dict = {}

    async def fake_agenda(graph_client, **kwargs):
        seen.update(kwargs, graph_client=graph_client)
        return {"tool": "outlook_agenda"}

    graph = object()
    monkeypatch.setattr(intent_agenda, "read_agenda", fake_agenda)
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    monkeypatch.setattr(auth, "saved_username", lambda: "me@example.com")
    config = Config(timezone="America/Los_Angeles")
    ctx = SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )
    result = asyncio.run(server.outlook_agenda(ctx, window="this_week", zone="Asia/Tokyo"))
    assert result == {"tool": "outlook_agenda"}
    assert seen["graph_client"] is graph
    assert seen["config_zone"] == "America/Los_Angeles"
    assert seen["zone"] == "Asia/Tokyo"
    assert seen["signed_in"] == "me@example.com"
    assert seen["loaded_tools"] == server.LOADED_TOOLS
    assert (seen["window"], seen["day_start"], seen["min_free"]) == ("this_week", "09:00", 30)


# ── outlook_open_loops ────────────────────────────────────────────────────


def test_the_advertised_open_loops(processes):
    tool = processes["classic,intents"]["tools"]["outlook_open_loops"]
    assert tool["annotations"]["readOnlyHint"] is True
    properties = tool["inputSchema"]["properties"]
    assert set(properties) == {"direction", "person", "words", "older_than", "window"}
    assert properties["direction"]["enum"] == list(intent_open_loops.DIRECTIONS)
    assert not tool["inputSchema"].get("required")
    assert len(tool["description"]) <= 450


def test_open_loops_takes_the_implementations_arguments():
    tool = inspect.signature(server.outlook_open_loops).parameters
    implementation = inspect.signature(intent_open_loops.read_open_loops).parameters
    assert set(tool) - {"ctx"} <= set(implementation)
    for name in set(tool) - {"ctx"}:
        assert tool[name].default == implementation[name].default, name


def test_open_loops_passes_the_zone_the_saved_username_and_the_loaded_tools(monkeypatch):
    seen: dict = {}

    async def fake_loops(graph_client, **kwargs):
        seen.update(kwargs, graph_client=graph_client)
        return {"tool": "outlook_open_loops"}

    graph = object()
    monkeypatch.setattr(intent_open_loops, "read_open_loops", fake_loops)
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    monkeypatch.setattr(auth, "saved_username", lambda: "me@example.com")
    config = Config(timezone="America/Los_Angeles")
    ctx = SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )
    result = asyncio.run(server.outlook_open_loops(ctx, direction="on_them", older_than="3d"))
    assert result == {"tool": "outlook_open_loops"}
    assert seen["graph_client"] is graph
    assert seen["zone"] == "America/Los_Angeles"
    assert seen["signed_in"] == "me@example.com"
    assert seen["loaded_tools"] == server.LOADED_TOOLS
    assert (seen["direction"], seen["older_than"], seen["window"]) == ("on_them", "3d", "14d")
    assert seen["person"] is None and seen["words"] is None


# ── outlook_catch_up ──────────────────────────────────────────────────────


def test_the_advertised_catch_up(processes):
    tool = processes["classic,intents"]["tools"]["outlook_catch_up"]
    assert tool["annotations"]["readOnlyHint"] is True
    properties = tool["inputSchema"]["properties"]
    assert set(properties) == {"cursor", "window", "focus"}
    assert properties["focus"]["enum"] == list(intent_catch_up.FOCUSES)
    assert not tool["inputSchema"].get("required")
    assert len(tool["description"]) <= 450


def test_catch_up_takes_the_implementations_arguments():
    tool = inspect.signature(server.outlook_catch_up).parameters
    implementation = inspect.signature(intent_catch_up.read_catch_up).parameters
    assert set(tool) - {"ctx"} <= set(implementation)
    for name in set(tool) - {"ctx"}:
        assert tool[name].default == implementation[name].default, name


def test_catch_up_passes_the_zone_the_saved_username_and_the_loaded_tools(monkeypatch):
    seen: dict = {}

    async def fake_catch_up(graph_client, **kwargs):
        seen.update(kwargs, graph_client=graph_client)
        return {"tool": "outlook_catch_up"}

    graph = object()
    monkeypatch.setattr(intent_catch_up, "read_catch_up", fake_catch_up)
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    monkeypatch.setattr(auth, "saved_username", lambda: "me@example.com")
    config = Config(timezone="America/Los_Angeles")
    ctx = SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )
    result = asyncio.run(server.outlook_catch_up(ctx, cursor="cu1.x", focus="mail"))
    assert result == {"tool": "outlook_catch_up"}
    assert seen["graph_client"] is graph
    assert seen["zone"] == "America/Los_Angeles"
    assert seen["signed_in"] == "me@example.com"
    assert seen["loaded_tools"] == server.LOADED_TOOLS
    assert (seen["cursor"], seen["window"], seen["focus"]) == ("cu1.x", None, "mail")


# ── outlook_brief ─────────────────────────────────────────────────────────


def test_the_advertised_brief(processes):
    tool = processes["classic,intents"]["tools"]["outlook_brief"]
    assert tool["annotations"]["readOnlyHint"] is True
    assert set(tool["inputSchema"]["properties"]) == {"person", "meeting", "topic", "window"}
    assert not tool["inputSchema"].get("required")
    assert len(tool["description"]) <= 450


def test_brief_takes_the_implementations_arguments():
    tool = inspect.signature(server.outlook_brief).parameters
    implementation = inspect.signature(intent_brief.read_brief).parameters
    assert set(tool) - {"ctx"} <= set(implementation)
    for name in set(tool) - {"ctx"}:
        assert tool[name].default == implementation[name].default, name


def test_brief_passes_the_zone_the_saved_username_and_the_loaded_tools(monkeypatch):
    seen: dict = {}

    async def fake_brief(graph_client, **kwargs):
        seen.update(kwargs, graph_client=graph_client)
        return {"tool": "outlook_brief"}

    graph = object()
    monkeypatch.setattr(intent_brief, "read_brief", fake_brief)
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    monkeypatch.setattr(auth, "saved_username", lambda: "me@example.com")
    config = Config(timezone="America/Los_Angeles")
    ctx = SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )
    result = asyncio.run(server.outlook_brief(ctx, meeting="standup"))
    assert result == {"tool": "outlook_brief"}
    assert seen["graph_client"] is graph
    assert seen["zone"] == "America/Los_Angeles"
    assert seen["signed_in"] == "me@example.com"
    assert seen["loaded_tools"] == server.LOADED_TOOLS
    assert (seen["meeting"], seen["person"], seen["topic"], seen["window"]) == (
        "standup",
        None,
        None,
        "30d",
    )


# ── What a served session names (contracts §10.9, §10.10) ─────────────────


@pytest.mark.parametrize("selection", ["intents", "classic,intents"])
def test_a_served_intents_session_names_only_tools_it_has(processes, selection):
    """The real server, in a fresh process: every tool its instructions or prompts
    name is in its own `tools/list`."""
    served = processes[selection]
    tools = set(served["tools"])
    named = set(re.findall(r"outlook_[a-z_]+", served["instructions"]))
    for text in served["prompts"].values():
        named |= set(re.findall(r"outlook_[a-z_]+", text))
    assert named <= tools, sorted(named - tools)


def test_the_served_prompts_follow_the_selection(processes):
    assert set(processes[None]["prompts"]) == {"morning_brief", "triage_folder", "catch_up"}
    assert "outlook_list_events" in processes[None]["prompts"]["morning_brief"]
    assert set(processes["intents"]["prompts"]) == {"morning_brief", "catch_up"}
    assert "outlook_agenda" in processes["intents"]["prompts"]["morning_brief"]
    assert "outlook_catch_up" in processes["classic,intents"]["prompts"]["catch_up"]
    assert "outlook_list_inbox" in processes["classic,intents"]["prompts"]["triage_folder"]
    assert set(processes["mail"]["prompts"]) == {"morning_brief", "triage_folder", "catch_up"}
