"""Tests for MCP server tool registration."""

from outlook_mcp.server import mcp  # noqa: I001 - single import


EXPECTED_TOOLS = [
    # Auth (3)
    "outlook_login",
    "outlook_logout",
    "outlook_auth_status",
    # Mail read (4)
    "outlook_list_inbox",
    "outlook_read_message",
    "outlook_search_mail",
    "outlook_list_folders",
    # Mail write (3)
    "outlook_send_message",
    "outlook_reply",
    "outlook_forward",
    # Mail triage (5)
    "outlook_move_message",
    "outlook_delete_message",
    "outlook_flag_message",
    "outlook_categorize_message",
    "outlook_mark_read",
    # Calendar read (2)
    "outlook_list_events",
    "outlook_get_event",
    # Calendar write (4)
    "outlook_create_event",
    "outlook_update_event",
    "outlook_delete_event",
    "outlook_rsvp",
]


def test_tier1_tool_count():
    """Exactly 21 Tier 1 tools are registered."""
    registered = set(mcp._tool_manager._tools.keys())
    assert len(registered) == 21


def test_all_tier1_tools_registered():
    """Every expected tool name is registered on the server."""
    registered = set(mcp._tool_manager._tools.keys())
    for name in EXPECTED_TOOLS:
        assert name in registered, f"Missing tool: {name}"


def test_no_unexpected_tools():
    """No extra tools beyond the expected 21."""
    registered = set(mcp._tool_manager._tools.keys())
    expected = set(EXPECTED_TOOLS)
    extra = registered - expected
    assert not extra, f"Unexpected tools registered: {extra}"


def test_server_metadata():
    """Server has correct name."""
    assert mcp.name == "outlook-mcp"


def test_tools_have_descriptions():
    """Every registered tool has a non-empty description."""
    for name, tool in mcp._tool_manager._tools.items():
        assert tool.description, f"Tool {name} has no description"
