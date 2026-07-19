"""Tests for toolsets.py — tool annotations + config-gated tool selection."""

import asyncio

from mcp.server.fastmcp import FastMCP

from outlook_mcp import toolsets

# ── parse_toolsets ────────────────────────────────────────────────────


def test_parse_none_means_all():
    assert toolsets.parse_toolsets(None) is None


def test_parse_empty_means_all():
    assert toolsets.parse_toolsets("") is None
    assert toolsets.parse_toolsets("  ") is None


def test_parse_comma_list_normalizes():
    assert toolsets.parse_toolsets("mail, Calendar , digest") == {
        "mail",
        "calendar",
        "digest",
    }


# ── select_kept_tools ─────────────────────────────────────────────────


def test_select_none_keeps_everything():
    names = ["outlook_list_inbox", "outlook_create_event", "outlook_list_tasks"]
    assert toolsets.select_kept_tools(names, None) == set(names)


def test_select_keeps_enabled_groups_plus_always_on():
    names = [
        "outlook_list_inbox",  # mail
        "outlook_send_message",  # mail
        "outlook_create_event",  # calendar
        "outlook_list_tasks",  # todo
        "outlook_auth_status",  # account (always on)
        "outlook_whoami",  # account (always on)
    ]
    kept = toolsets.select_kept_tools(names, {"mail"})
    assert kept == {
        "outlook_list_inbox",
        "outlook_send_message",
        "outlook_auth_status",  # always-on survives even when not requested
        "outlook_whoami",
    }


def test_select_multiple_groups():
    names = ["outlook_list_inbox", "outlook_create_event", "outlook_changes_since"]
    kept = toolsets.select_kept_tools(names, {"calendar", "digest"})
    # mail tool dropped; calendar + digest kept
    assert kept == {"outlook_create_event", "outlook_changes_since"}


# ── annotations ───────────────────────────────────────────────────────


def test_read_tool_is_read_only():
    ann = toolsets.annotation_for("outlook_list_inbox")
    assert ann.readOnlyHint is True


def test_delete_tool_is_destructive():
    ann = toolsets.annotation_for("outlook_delete_message")
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is True


def test_send_tool_is_additive_write():
    ann = toolsets.annotation_for("outlook_send_message")
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


# ── configure applied to a real FastMCP instance ──────────────────────


def _mini_server():
    """A throwaway FastMCP with a few real-named tools from distinct groups."""
    m = FastMCP("mini")

    @m.tool(name="outlook_list_inbox")
    def _li() -> str:  # mail, read-only
        return "x"

    @m.tool(name="outlook_delete_message")
    def _dm() -> str:  # mail, destructive
        return "x"

    @m.tool(name="outlook_create_event")
    def _ce() -> str:  # calendar, additive
        return "x"

    @m.tool(name="outlook_auth_status")
    def _as() -> str:  # account, always-on
        return "x"

    return m


def _tool_names(m):
    return {t.name for t in asyncio.run(m.list_tools())}


def test_configure_gates_and_annotates():
    m = _mini_server()
    toolsets.configure(m, {"calendar"})
    names = _tool_names(m)
    # calendar + always-on account kept; mail tools dropped
    assert names == {"outlook_create_event", "outlook_auth_status"}
    # surviving read-only annotation applied
    ann = {t.name: t.annotations for t in asyncio.run(m.list_tools())}
    assert ann["outlook_auth_status"].readOnlyHint is True


def test_configure_none_keeps_all_but_still_annotates():
    m = _mini_server()
    toolsets.configure(m, None)
    tools = {t.name: t for t in asyncio.run(m.list_tools())}
    assert len(tools) == 4  # nothing gated
    assert tools["outlook_delete_message"].annotations.destructiveHint is True


# ── guardrail: every live tool is classified ──────────────────────────


def test_every_registered_tool_is_classified():
    """Catches drift: a new tool added to server.py without a group/annotation."""
    from outlook_mcp.server import mcp

    live = {t.name for t in asyncio.run(mcp.list_tools())}
    missing_group = live - set(toolsets.TOOL_GROUPS)
    assert not missing_group, f"tools missing from TOOL_GROUPS: {sorted(missing_group)}"
