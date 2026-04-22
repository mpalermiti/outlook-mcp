"""Tests for MCP server tool registration."""

from outlook_mcp.server import mcp  # noqa: I001 - single import


EXPECTED_TOOLS = [
    # Auth (3)
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
    # Mail triage (6)
    "outlook_move_message",
    "outlook_delete_message",
    "outlook_flag_message",
    "outlook_categorize_message",
    "outlook_mark_read",
    "outlook_reclassify_message",
    # Calendar read (2)
    "outlook_list_events",
    "outlook_get_event",
    # Calendar write (4)
    "outlook_create_event",
    "outlook_update_event",
    "outlook_delete_event",
    "outlook_rsvp",
    # ── Tier 2 ──────────────────────────────────────────
    # Contacts (6)
    "outlook_list_contacts",
    "outlook_search_contacts",
    "outlook_get_contact",
    "outlook_create_contact",
    "outlook_update_contact",
    "outlook_delete_contact",
    # To Do (6)
    "outlook_list_task_lists",
    "outlook_list_tasks",
    "outlook_create_task",
    "outlook_update_task",
    "outlook_complete_task",
    "outlook_delete_task",
    # Mail drafts (5)
    "outlook_list_drafts",
    "outlook_create_draft",
    "outlook_update_draft",
    "outlook_send_draft",
    "outlook_delete_draft",
    # Mail attachments (5)
    "outlook_list_attachments",
    "outlook_download_attachment",
    "outlook_send_with_attachments",
    "outlook_attach_to_draft",
    "outlook_remove_draft_attachment",
    # Mail folders (3)
    "outlook_create_folder",
    "outlook_rename_folder",
    "outlook_delete_folder",
    # Mail thread (2)
    "outlook_list_thread",
    "outlook_copy_message",
    # Batch (1)
    "outlook_batch_triage",
    # User (2)
    "outlook_whoami",
    "outlook_list_calendars",
    # Admin (2)
    "outlook_list_categories",
    "outlook_get_mail_tips",
    # Multi-account (2)
    "outlook_list_accounts",
    "outlook_switch_account",
]


def test_tool_count():
    """All 54 tools are registered (auth is CLI-only now)."""
    registered = set(mcp._tool_manager._tools.keys())
    assert len(registered) == 54


def test_all_tools_registered():
    """Every expected tool name is registered on the server."""
    registered = set(mcp._tool_manager._tools.keys())
    for name in EXPECTED_TOOLS:
        assert name in registered, f"Missing tool: {name}"


def test_no_unexpected_tools():
    """No extra tools beyond the expected set."""
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
