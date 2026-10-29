"""Tests for toolsets.py — tool annotations + config-gated tool selection."""

import asyncio

from mcp.server.mcpserver import MCPServer

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
    assert ann.read_only_hint is True


def test_delete_tool_is_destructive():
    ann = toolsets.annotation_for("outlook_delete_message")
    assert ann.read_only_hint is False
    assert ann.destructive_hint is True


def test_send_tool_is_destructive():
    """Sending is irreversible and leaves the mailbox: not an "additive" write.

    Clients read `destructiveHint=False` as safe to auto-approve among writes,
    and an email cannot be unsent (see test_send_tools_are_marked_destructive).
    """
    ann = toolsets.annotation_for("outlook_send_message")
    assert ann.read_only_hint is False
    assert ann.destructive_hint is True


def test_an_additive_write_is_not_destructive():
    ann = toolsets.annotation_for("outlook_create_contact")
    assert ann.read_only_hint is False
    assert ann.destructive_hint is False


# ── configure applied to a real MCPServer instance ──────────────────────


def _mini_server():
    """A throwaway MCPServer with a few real-named tools from distinct groups."""
    m = MCPServer("mini")

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
    assert ann["outlook_auth_status"].read_only_hint is True


def test_configure_none_keeps_all_but_still_annotates():
    m = _mini_server()
    toolsets.configure(m, None)
    tools = {t.name: t for t in asyncio.run(m.list_tools())}
    assert len(tools) == 4  # nothing gated
    assert tools["outlook_delete_message"].annotations.destructive_hint is True


# ── resolve_groups: env selection + the config flag ───────────────────


def test_resolve_unset_env_leaves_off_by_default_groups_off():
    """No selection means every group except the off-by-default ones."""
    groups = toolsets.resolve_groups(None, mailbox_settings_write=False)
    assert groups == set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT
    assert "rules" not in groups


def test_resolve_unset_env_with_the_config_flag_adds_rules_and_only_rules():
    """Turning rule management on is itself the request for those tools.

    Not "everything": `None` here would load any opt-in group added later, and a
    group is opt-in precisely because it must not load by default. Measured
    against a second opt-in group, the flag loaded it too.
    """
    groups = toolsets.resolve_groups(None, mailbox_settings_write=True)

    assert "rules" in groups
    assert groups == (set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT) | {"rules"}


def test_resolve_explicit_selection_keeps_only_what_was_named():
    assert toolsets.resolve_groups({"mail"}, mailbox_settings_write=False) == {"mail"}


def test_resolve_explicit_selection_is_not_widened_by_the_flag():
    """A named selection wins: the flag must not hand back groups nobody asked for.

    A shared config with the flag on used to add `rules` to an agent gated down
    to `mail`, and there was no way to ask for the narrower surface.
    """
    assert toolsets.resolve_groups({"mail"}, mailbox_settings_write=True) == {"mail"}


def test_resolve_explicit_rules_selection_survives_the_flag_being_off():
    """Opting in by env still works on an install that never sets the flag."""
    assert toolsets.resolve_groups({"rules"}, mailbox_settings_write=False) == {"rules"}


def test_the_flag_does_not_load_the_other_off_by_default_group():
    """The flag is about rule writes, not about every off-by-default group."""
    groups = toolsets.resolve_groups(None, mailbox_settings_write=True)

    assert "rules" in groups
    assert "documents" not in groups


def test_no_selection_resolves_to_a_set_not_none():
    """`None` means "everything" to `select_kept_tools`, so it must not escape."""
    groups = toolsets.resolve_groups(None, mailbox_settings_write=False)

    assert isinstance(groups, set)
    assert groups == set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT


def test_select_kept_tools_none_matches_the_default_path():
    """The two "off by default" paths must name the same groups.

    With two sets, `select_kept_tools(None)` subtracted only `documents` and
    kept the four rule tools for a caller that meant the stock surface.
    """
    names = list(toolsets.TOOL_GROUPS)
    by_default = toolsets.select_kept_tools(names, None)
    by_resolver = toolsets.select_kept_tools(
        names, toolsets.resolve_groups(None, mailbox_settings_write=False)
    )

    assert by_default == by_resolver


def test_the_rule_tools_are_opt_in_and_off_by_default():
    """The most expensive group here must not be on the default surface."""
    names = list(toolsets.TOOL_GROUPS)
    kept = toolsets.select_kept_tools(names, toolsets.resolve_groups(None, False))
    assert not {n for n in kept if toolsets.TOOL_GROUPS[n] == "rules"}

    kept_with_flag = toolsets.select_kept_tools(names, toolsets.resolve_groups(None, True))
    assert {n for n in kept_with_flag if toolsets.TOOL_GROUPS[n] == "rules"}


# ── guardrail: every live tool is classified ──────────────────────────


def test_every_registered_tool_is_classified():
    """Catches drift: a new tool added to server.py without a group/annotation."""
    from outlook_mcp.server import mcp

    live = {t.name for t in asyncio.run(mcp.list_tools())}
    missing_group = live - set(toolsets.TOOL_GROUPS)
    assert not missing_group, f"tools missing from TOOL_GROUPS: {sorted(missing_group)}"
