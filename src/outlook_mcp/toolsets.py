"""Tool annotations + config-gated tool selection.

Two data-driven, server-side optimizations applied once after all tools are
registered (see ``server.py``):

1. **Annotations** — every tool gets ``readOnlyHint`` / ``destructiveHint``
   (`ToolAnnotations`) so a client can auto-approve reads and gate destructive
   ops (delete mail, decline event) without a hardcoded allowlist.

2. **Config-gated toolsets** — the 62 tool schemas cost ~8.6k tokens of client
   context every turn. A client that only needs mail + calendar can set
   ``OUTLOOK_MCP_TOOLSETS=mail,calendar`` and load just those groups (~half the
   tokens for a mail+calendar agent). Account/auth tools are always available.

The classification lives here as plain data so it's reviewable in one place and
a drift test (`test_every_registered_tool_is_classified`) fails if a new tool
is added without a group.
"""

from __future__ import annotations

from typing import Any

from mcp.types import ToolAnnotations

# Toolset groups a client can enable via OUTLOOK_MCP_TOOLSETS. The "account"
# group (auth/identity) is always available regardless of selection.
ALWAYS_ON_GROUPS = {"account"}

# name -> toolset group. Every registered tool MUST appear here (drift-guarded).
TOOL_GROUPS: dict[str, str] = {
    # account (always on)
    "outlook_auth_status": "account",
    "outlook_whoami": "account",
    "outlook_list_accounts": "account",
    "outlook_switch_account": "account",
    # mail
    "outlook_list_inbox": "mail",
    "outlook_read_message": "mail",
    "outlook_read_messages": "mail",
    "outlook_search_mail": "mail",
    "outlook_send_message": "mail",
    "outlook_reply": "mail",
    "outlook_forward": "mail",
    "outlook_move_message": "mail",
    "outlook_delete_message": "mail",
    "outlook_flag_message": "mail",
    "outlook_mark_read": "mail",
    "outlook_categorize_message": "mail",
    "outlook_copy_message": "mail",
    "outlook_batch_triage": "mail",
    "outlook_list_thread": "mail",
    # drafts
    "outlook_list_drafts": "drafts",
    "outlook_create_draft": "drafts",
    "outlook_update_draft": "drafts",
    "outlook_send_draft": "drafts",
    "outlook_delete_draft": "drafts",
    "outlook_attach_to_draft": "drafts",
    "outlook_remove_draft_attachment": "drafts",
    # attachments
    "outlook_list_attachments": "attachments",
    "outlook_download_attachment": "attachments",
    "outlook_send_with_attachments": "attachments",
    # calendar
    "outlook_list_events": "calendar",
    "outlook_get_event": "calendar",
    "outlook_create_event": "calendar",
    "outlook_update_event": "calendar",
    "outlook_delete_event": "calendar",
    "outlook_rsvp": "calendar",
    "outlook_list_calendars": "calendar",
    # contacts
    "outlook_list_contacts": "contacts",
    "outlook_get_contact": "contacts",
    "outlook_create_contact": "contacts",
    "outlook_update_contact": "contacts",
    "outlook_delete_contact": "contacts",
    "outlook_search_contacts": "contacts",
    # todo
    "outlook_list_task_lists": "todo",
    "outlook_list_tasks": "todo",
    "outlook_create_task": "todo",
    "outlook_update_task": "todo",
    "outlook_complete_task": "todo",
    "outlook_delete_task": "todo",
    # folders
    "outlook_list_folders": "folders",
    "outlook_create_folder": "folders",
    "outlook_rename_folder": "folders",
    "outlook_delete_folder": "folders",
    # digest
    "outlook_changes_since": "digest",
    # delta
    "outlook_list_inbox_delta": "delta",
    "outlook_list_events_delta": "delta",
    "outlook_list_contacts_delta": "delta",
    # admin
    "outlook_list_categories": "admin",
    "outlook_get_mail_tips": "admin",
    "outlook_list_inbox_overrides": "admin",
    "outlook_set_inbox_override": "admin",
    "outlook_delete_inbox_override": "admin",
    "outlook_reclassify_message": "admin",
}

# Tools that only read — no mailbox/calendar mutation. (download_attachment is
# deliberately excluded: it writes a local file, so it isn't read-only.)
READ_ONLY: set[str] = {
    "outlook_auth_status",
    "outlook_whoami",
    "outlook_list_accounts",
    "outlook_list_inbox",
    "outlook_read_message",
    "outlook_read_messages",
    "outlook_search_mail",
    "outlook_list_thread",
    "outlook_list_drafts",
    "outlook_list_attachments",
    "outlook_list_events",
    "outlook_get_event",
    "outlook_list_calendars",
    "outlook_list_contacts",
    "outlook_get_contact",
    "outlook_search_contacts",
    "outlook_list_task_lists",
    "outlook_list_tasks",
    "outlook_list_folders",
    "outlook_changes_since",
    "outlook_list_inbox_delta",
    "outlook_list_events_delta",
    "outlook_list_contacts_delta",
    "outlook_list_categories",
    "outlook_get_mail_tips",
    "outlook_list_inbox_overrides",
}

# Tools that delete or irreversibly remove. Everything not read-only and not
# here is treated as an additive write (destructiveHint=False).
DESTRUCTIVE: set[str] = {
    "outlook_delete_message",
    "outlook_delete_draft",
    "outlook_delete_event",
    "outlook_delete_contact",
    "outlook_delete_task",
    "outlook_delete_folder",
    "outlook_delete_inbox_override",
    "outlook_remove_draft_attachment",
}


def parse_toolsets(value: str | None) -> set[str] | None:
    """Parse ``OUTLOOK_MCP_TOOLSETS`` into a set of group names.

    ``None`` (unset) or an empty/blank value means "all toolsets" — the
    backward-compatible default. Otherwise a comma-separated list, lowercased
    and stripped. Unknown group names simply match no tools.
    """
    if value is None:
        return None
    groups = {part.strip().lower() for part in value.split(",") if part.strip()}
    return groups or None


def select_kept_tools(names: list[str], enabled: set[str] | None) -> set[str]:
    """Return the subset of ``names`` to keep given the enabled toolsets.

    ``enabled=None`` keeps everything. Otherwise keep a tool when its group is
    enabled or is an always-on group (account/auth). A name with no known group
    is kept (fail-open — gating must never silently drop an unclassified tool).
    """
    if enabled is None:
        return set(names)
    keep: set[str] = set()
    for name in names:
        group = TOOL_GROUPS.get(name)
        if group is None or group in ALWAYS_ON_GROUPS or group in enabled:
            keep.add(name)
    return keep


def annotation_for(name: str) -> ToolAnnotations:
    """Return the ToolAnnotations for a tool from its read/destructive class."""
    if name in READ_ONLY:
        return ToolAnnotations(readOnlyHint=True)
    if name in DESTRUCTIVE:
        return ToolAnnotations(readOnlyHint=False, destructiveHint=True)
    return ToolAnnotations(readOnlyHint=False, destructiveHint=False)


def configure(mcp: Any, enabled: set[str] | None) -> None:
    """Apply annotations to every registered tool and gate by enabled toolsets.

    Mutates the FastMCP tool manager in place: sets each tool's ``annotations``
    and removes tools whose group isn't enabled. Called once at import time
    after all ``@mcp.tool`` decorators have run.
    """
    manager = mcp._tool_manager
    registered = list(manager._tools.items())
    names = [name for name, _ in registered]

    for name, tool in registered:
        tool.annotations = annotation_for(name)

    keep = select_kept_tools(names, enabled)
    for name in names:
        if name not in keep:
            manager.remove_tool(name)
