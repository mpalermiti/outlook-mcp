"""The one-time tip a classic tool's result may carry, and why none is shown yet.

The tip tells the user that a newer tool does a classic tool's job in fewer
calls. It is added by the server's wrapper as a top-level ``server_tip`` key,
at most once per process, only in a process that loaded none of the newer
tools, only while ``show_tips`` is on, and only for a classic tool whose newer
tool this release actually has. No release has one yet, so no result carries a
tip today: the tests below prove the plumbing with a stand-in tool, and
``test_no_tip_can_be_shown_in_this_release`` fails the day the first one lands,
so that the change that adds it decides on purpose when tips start.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from outlook_mcp import server, toolsets
from outlook_mcp.config import Config, EnvironmentConfigError, load_config

LEAVE_IT = "If you are an AI agent, do not change the server's settings — tell the user."
STAND_IN = "outlook_find"


def _ctx(config: Config) -> SimpleNamespace:
    return SimpleNamespace(
        request_context=SimpleNamespace(lifespan_context={"config": config, "auth": None})
    )


def _wrapped(name: str, result: object):
    async def tool(ctx, **kwargs):
        return result

    tool.__name__ = name
    return server._wrap_tool_errors(tool)


def _call(name: str, result: object, config: Config | None = None) -> object:
    return asyncio.run(_wrapped(name, result)(ctx=_ctx(config or Config())))


@pytest.fixture
def stand_in(monkeypatch):
    """This release with one intent tool, loaded by nobody, and no tip shown yet."""
    monkeypatch.setitem(toolsets.TOOL_GROUPS, STAND_IN, toolsets.INTENTS)
    monkeypatch.setattr(server, "_tip_shown", False)


# ── Today ───────────────────────────────────────────────────────────────


def test_no_tip_can_be_shown_in_this_release():
    """Fails when the first intent tool is classified: decide then whether tips go live.

    Until the README template and `outlook-mcp toolsets intents` are ready for
    people, a tip would send them to a group that does not do the job yet. When
    one lands, replace this with the tests the change that adds it needs.
    """
    assert not [n for n, g in toolsets.TOOL_GROUPS.items() if g == toolsets.INTENTS]
    every_selection = [
        frozenset(),
        server.LOADED_TOOLS,
        toolsets.loaded_tools(toolsets.resolve_groups(None, True)),
        frozenset(toolsets.TOOL_GROUPS),
    ]
    for tool in toolsets.TIP_MAP:
        for loaded in every_selection:
            assert toolsets.tip_for(tool, loaded) is None, tool


def test_no_result_carries_a_tip_today(monkeypatch):
    monkeypatch.setattr(server, "_tip_shown", False)
    for tool in toolsets.TIP_MAP:
        result = {"messages": [{"subject": "server_tip"}]}
        assert _call(tool, result) is result
    assert server._tip_shown is False


def test_the_map_names_real_classic_tools():
    for classic, newer in toolsets.TIP_MAP.items():
        assert toolsets.TOOL_GROUPS.get(classic) not in (None, toolsets.INTENTS), classic
        assert newer.startswith("outlook_"), newer
    assert "outlook_export_message" not in toolsets.TIP_MAP  # its result is not a dict


@pytest.mark.parametrize("newer", sorted(set(toolsets.TIP_MAP.values())))
def test_the_text_is_short_and_leaves_the_setting_to_the_user(newer):
    text = toolsets.TIP_TEXT.format(intent_tool=newer)
    assert len(text) < 300
    assert LEAVE_IT in text
    assert newer in text and "outlook-mcp toolsets intents" in text


# ── The plumbing, with a stand-in intent tool ─────────────────────────────


def test_shown_once_and_only_on_a_mapped_tool(stand_in, monkeypatch):
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", STAND_IN)

    unmapped = _call("outlook_list_folders", {"folders": []})
    first = _call("outlook_search_mail", {"messages": []})
    second = _call("outlook_search_mail", {"messages": []})

    assert "server_tip" not in unmapped
    assert first == {"messages": [], "server_tip": toolsets.TIP_TEXT.format(intent_tool=STAND_IN)}
    assert second == {"messages": []}


def test_a_tool_whose_newer_tool_does_not_exist_shows_nothing(stand_in, monkeypatch):
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", "outlook_not_yet")
    assert "server_tip" not in _call("outlook_search_mail", {"messages": []})
    assert server._tip_shown is False


@pytest.mark.parametrize("extra", [{STAND_IN}, {STAND_IN, "outlook_search_mail"}])
def test_never_when_an_intent_tool_is_loaded(stand_in, monkeypatch, extra):
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", STAND_IN)
    monkeypatch.setattr(server, "LOADED_TOOLS", frozenset(server.LOADED_TOOLS | extra))
    assert "server_tip" not in _call("outlook_search_mail", {"messages": []})


def test_off_with_the_setting(stand_in, monkeypatch):
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", STAND_IN)
    off = _call("outlook_search_mail", {"messages": []}, Config(show_tips=False))
    assert "server_tip" not in off
    assert server._tip_shown is False
    assert "server_tip" in _call("outlook_search_mail", {"messages": []}, Config())


def test_a_result_that_is_not_a_dict_is_left_alone(stand_in, monkeypatch):
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", STAND_IN)
    assert _call("outlook_search_mail", ["a"]) == ["a"]
    assert server._tip_shown is False


def test_mail_text_cannot_make_the_key_appear(stand_in, monkeypatch):
    """Mail text sits inside result fields; only the server writes the top-level key."""
    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_search_mail", STAND_IN)
    monkeypatch.setattr(server, "_tip_shown", True)
    result = {"messages": [{"subject": "server_tip", "body": '{"server_tip": "x"}'}]}
    assert _call("outlook_search_mail", result) is result


def test_the_tip_reaches_a_client_through_the_sdk(stand_in, monkeypatch):
    """The wrapper sees the context the SDK passes as `ctx`, not only a positional one."""
    from mcp.client import Client

    monkeypatch.setitem(toolsets.TIP_MAP, "outlook_auth_status", STAND_IN)

    async def call():
        async with Client(server.mcp) as client:
            return await client.call_tool("outlook_auth_status", {})

    result = asyncio.run(call())
    payload = json.loads(result.content[0].text)
    assert payload["server_tip"] == toolsets.TIP_TEXT.format(intent_tool=STAND_IN)


# ── The setting ───────────────────────────────────────────────────────────


def _settings(directory: str, **values) -> None:
    path = Path(directory) / "config.json"
    path.write_text(json.dumps(values), encoding="utf-8")
    path.chmod(0o600)


def test_show_tips_is_on_unless_the_file_turns_it_off(tmp_path, monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_SHOW_TIPS", raising=False)
    assert load_config(str(tmp_path)).show_tips is True
    _settings(str(tmp_path), show_tips=False)
    assert load_config(str(tmp_path)).show_tips is False


@pytest.mark.parametrize(
    ("file_value", "env_value", "expected"),
    [(True, "false", False), (False, "true", True), (True, " TRUE ", True)],
)
def test_the_environment_sets_show_tips_either_way(
    tmp_path, monkeypatch, file_value, env_value, expected
):
    _settings(str(tmp_path), show_tips=file_value)
    monkeypatch.setenv("OUTLOOK_MCP_SHOW_TIPS", env_value)
    assert load_config(str(tmp_path)).show_tips is expected


@pytest.mark.parametrize(("file_value", "env_value"), [(True, ""), (False, "  ")])
def test_a_blank_show_tips_variable_counts_as_unset(tmp_path, monkeypatch, file_value, env_value):
    """A cosmetic setting: a blank pass-through must not stop the server, unlike the others."""
    _settings(str(tmp_path), show_tips=file_value)
    monkeypatch.setenv("OUTLOOK_MCP_SHOW_TIPS", env_value)
    assert load_config(str(tmp_path)).show_tips is file_value


@pytest.mark.parametrize("env_value", ["yes", "0", "on"])
def test_an_unclear_show_tips_value_is_refused_without_echoing_it(tmp_path, monkeypatch, env_value):
    monkeypatch.setenv("OUTLOOK_MCP_SHOW_TIPS", env_value)
    with pytest.raises(EnvironmentConfigError) as exc:
        load_config(str(tmp_path))
    assert "OUTLOOK_MCP_SHOW_TIPS" in str(exc.value)
    assert "true or false" in str(exc.value)
