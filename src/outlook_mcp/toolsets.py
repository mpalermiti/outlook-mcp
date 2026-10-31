"""Tool annotations, lean input schemas and config-gated tool selection.

Three server-side optimizations applied once after all tools are registered
(see ``server.py``):

1. **Annotations** — every tool gets ``readOnlyHint`` / ``destructiveHint``
   (`ToolAnnotations`) so a client can auto-approve reads and gate destructive
   ops (delete mail, send mail, an invitation or RSVP) without a hardcoded
   allowlist.

2. **Config-gated toolsets** — the full tool schemas cost chars/4 proxy tokens
   of client context every turn (the budget test's yardstick and its header
   hold the measured number). A client that only needs mail + calendar can set
   ``OUTLOOK_MCP_TOOLSETS=mail,calendar`` and load just those groups (~half the
   tokens for a mail+calendar agent). Account/auth tools are always available.

3. **Lean input schemas** — `lean_input_schema` drops what Pydantic writes into
   every schema that the model cannot use: a ``title`` restating each input's
   name and each tool's, and the ``null`` branch and ``"default": null`` of an
   optional input, which say nothing that omitting it does not. A quarter of
   the default surface, with no wording changed.

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

# Groups a stock install does not pay for: not loaded unless
# OUTLOOK_MCP_TOOLSETS names them. `documents` is main's opt-in reader; `rules`
# is the inbox-rule CRUD group, whose four schemas are the most expensive single
# feature here (~1.2k proxy tokens, ~8% of the default surface — mostly
# `outlook_create_rule`'s fourteen optional arguments), and three of the four
# cannot work at all without the write scope.
#
# One set, not two. With a separate "optional" and "opt-in" set, the default
# path subtracted both while `select_kept_tools(None)` subtracted only one, so
# the two disagreed about what a stock install gets and `rules` leaked through
# the second path.
# Original MIME export also stays off by default; it adds no write scope.
OFF_BY_DEFAULT = {"documents", "rules", "mime"}

# What `mailbox_settings_write: true` adds when nothing was named. Not every
# off-by-default group: that key asks for rule writes, not for the document
# reader.
SETTINGS_WRITE_ADDS = {"rules"}

# name -> toolset group. Every registered tool MUST appear here (drift-guarded).
TOOL_GROUPS: dict[str, str] = {
    # account (always on)
    "outlook_auth_status": "account",
    "outlook_whoami": "account",
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
    # MIME export is separate from ordinary mail reads and off by default.
    "outlook_export_message": "mime",
    # rules — their own group: the four schemas are the most expensive single
    # feature here (~1.2k proxy tokens, mostly `create_rule`'s argument list), and
    # a client that gated down to `mail` should not pay for them.
    "outlook_list_rules": "rules",
    "outlook_create_rule": "rules",
    "outlook_update_rule": "rules",
    "outlook_delete_rule": "rules",
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
    "outlook_read_attachment": "documents",
    "outlook_send_with_attachments": "attachments",
    # calendar
    "outlook_list_events": "calendar",
    "outlook_get_event": "calendar",
    "outlook_find_event_by_key": "calendar",
    "outlook_create_event": "calendar",
    "outlook_update_event": "calendar",
    "outlook_delete_event": "calendar",
    "outlook_rsvp": "calendar",
    "outlook_attach_to_event": "calendar",
    "outlook_remove_event_attachment": "calendar",
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
    "outlook_get_task": "todo",
    "outlook_create_task": "todo",
    "outlook_update_task": "todo",
    "outlook_complete_task": "todo",
    "outlook_delete_task": "todo",
    "outlook_add_checklist_item": "todo",
    "outlook_update_checklist_item": "todo",
    "outlook_delete_checklist_item": "todo",
    "outlook_list_task_attachments": "todo",
    "outlook_download_task_attachment": "todo",
    "outlook_upload_task_attachment": "todo",
    "outlook_delete_task_attachment": "todo",
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
    "outlook_export_message",
    "outlook_auth_status",
    "outlook_whoami",
    "outlook_list_inbox",
    "outlook_read_message",
    "outlook_read_messages",
    "outlook_search_mail",
    "outlook_list_thread",
    "outlook_list_rules",
    "outlook_list_drafts",
    "outlook_list_attachments",
    "outlook_read_attachment",
    "outlook_list_events",
    "outlook_get_event",
    "outlook_find_event_by_key",
    "outlook_list_calendars",
    "outlook_list_contacts",
    "outlook_get_contact",
    "outlook_search_contacts",
    "outlook_list_task_lists",
    "outlook_list_tasks",
    "outlook_get_task",
    "outlook_list_task_attachments",
    "outlook_list_folders",
    "outlook_changes_since",
    "outlook_list_inbox_delta",
    "outlook_list_events_delta",
    "outlook_list_contacts_delta",
    "outlook_list_categories",
    "outlook_get_mail_tips",
    "outlook_list_inbox_overrides",
}

# Tools whose effect cannot be taken back: they delete, or they send email to
# someone — an invitation and an RSVP are email too. Everything not read-only
# and not here is an additive write (destructiveHint=False), which a client may
# approve without asking, so a tool that can send must never fall through to
# it. `test_send_tools_are_marked_destructive.py` finds every tool gated on
# `mail_send` and fails if one is missing here.
DESTRUCTIVE: set[str] = {
    "outlook_send_message",
    "outlook_reply",
    "outlook_forward",
    "outlook_send_draft",
    "outlook_send_with_attachments",
    "outlook_create_event",
    "outlook_update_event",
    "outlook_rsvp",
    "outlook_delete_message",
    "outlook_delete_draft",
    "outlook_delete_event",
    "outlook_delete_contact",
    "outlook_delete_task",
    "outlook_delete_checklist_item",
    "outlook_delete_task_attachment",
    "outlook_delete_folder",
    "outlook_delete_inbox_override",
    # Creating a rule outlives the call: it can keep deleting or forwarding mail
    # indefinitely, so a client that asks before destructive calls must ask here.
    # `outlook_update_rule` joins them because re-enabling a forwarding rule
    # resumes the sending.
    "outlook_create_rule",
    "outlook_update_rule",
    "outlook_delete_rule",
    "outlook_remove_draft_attachment",
    "outlook_attach_to_event",
    "outlook_remove_event_attachment",
}


def parse_toolsets(value: str | None) -> set[str] | None:
    """Parse ``OUTLOOK_MCP_TOOLSETS`` into a set of group names.

    ``None`` (unset) or an empty/blank value means "all default toolsets" — the
    backward-compatible default. Otherwise a comma-separated list, lowercased
    and stripped. Unknown group names simply match no tools.
    """
    if value is None:
        return None
    groups = {part.strip().lower() for part in value.split(",") if part.strip()}
    return groups or None


def resolve_groups(enabled: set[str] | None, mailbox_settings_write: bool) -> set[str]:
    """Enabled groups, from the env selection and the config flag.

    ``enabled`` is what ``OUTLOOK_MCP_TOOLSETS`` said: ``None`` means "no
    selection", which resolves to every group **except** ``OFF_BY_DEFAULT``.
    ``mailbox_settings_write`` adds ``SETTINGS_WRITE_ADDS`` back, but only when
    nothing else was named: an agent gated down to ``mail`` must not have
    ``rules`` handed back by a shared config it did not write, with no way to
    ask for the narrower surface.

    Always a set. A ``None`` return would mean "everything" to
    ``select_kept_tools``, which would silently keep every off-by-default group
    for any caller that passed it straight through.
    """
    if enabled is None:
        groups = set(TOOL_GROUPS.values()) - OFF_BY_DEFAULT
        if mailbox_settings_write:
            groups |= SETTINGS_WRITE_ADDS
        return groups
    return set(enabled)


def select_kept_tools(names: list[str], enabled: set[str] | None) -> set[str]:
    """Return the subset of ``names`` to keep given the enabled toolsets.

    ``enabled=None`` keeps every group except the off-by-default ones. Otherwise
    keep a tool when its group is enabled or is an always-on group
    (account/auth). A name with no known group is kept (fail-open — gating must
    never silently drop an unclassified tool).
    """
    if enabled is None:
        # Same set the default path subtracts: with only `documents` here, this
        # kept the four rule tools for a caller that meant the stock surface.
        return {name for name in names if TOOL_GROUPS.get(name) not in OFF_BY_DEFAULT}
    keep: set[str] = set()
    for name in names:
        group = TOOL_GROUPS.get(name)
        if group is None or group in ALWAYS_ON_GROUPS or group in enabled:
            keep.add(name)
    return keep


# Tools a client may safely repeat with the same arguments — each one PATCHes an
# absolute value, so the second call lands on the state the first one produced.
# Opt-in, never inferred: a client that believes this may retry after a timeout,
# and a wrong entry here sends a second email or re-issues an invitation. The
# reasoning for every exclusion is in tests/test_idempotent_hints.py, which fails
# if this set changes. `openWorldHint` is deliberately absent — it is true by
# default in the schema and true for every tool, so stating it is pure token cost.
IDEMPOTENT: set[str] = {
    "outlook_flag_message",
    "outlook_mark_read",
    "outlook_categorize_message",
    "outlook_rename_folder",
    "outlook_set_inbox_override",
    "outlook_download_attachment",
    "outlook_download_task_attachment",
    "outlook_update_checklist_item",
}


def annotation_for(name: str) -> ToolAnnotations:
    """Return the ToolAnnotations for a tool from its read/destructive class."""
    if name in READ_ONLY:
        return ToolAnnotations(readOnlyHint=True)
    if name in DESTRUCTIVE:
        return ToolAnnotations(readOnlyHint=False, destructiveHint=True)
    if name in IDEMPOTENT:
        return ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=True)
    return ToolAnnotations(readOnlyHint=False, destructiveHint=False)


# Keywords whose value is a schema, or a list or map of schemas, in the subset of
# JSON Schema Pydantic emits. Anything under another keyword (`default`, `enum`,
# `const`, `examples`) is data and is copied untouched, so a `title` key inside a
# default value is never mistaken for the keyword.
_SCHEMA_KEYWORDS = ("items", "additionalProperties", "not")
_SCHEMA_LIST_KEYWORDS = ("anyOf", "oneOf", "allOf", "prefixItems")
_SCHEMA_MAP_KEYWORDS = ("properties", "$defs", "definitions")
_NULL = {"type": "null"}


def _without_titles(node: Any) -> Any:
    """``node`` with every ``title`` keyword removed, walked by keyword.

    Walking by keyword is what keeps an input *named* ``title`` (contacts and
    tasks have one): its name is a key of ``properties``, never a keyword.
    """
    if not isinstance(node, dict):
        return node
    lean = {key: value for key, value in node.items() if key != "title"}
    for key in _SCHEMA_KEYWORDS:
        if isinstance(lean.get(key), dict):
            lean[key] = _without_titles(lean[key])
    for key in _SCHEMA_LIST_KEYWORDS:
        if isinstance(lean.get(key), list):
            lean[key] = [_without_titles(child) for child in lean[key]]
    for key in _SCHEMA_MAP_KEYWORDS:
        if isinstance(lean.get(key), dict):
            lean[key] = {name: _without_titles(child) for name, child in lean[key].items()}
    return lean


def _without_null_default(prop: dict) -> dict:
    """An optional input with ``default: null`` minus its ``null`` branch and default.

    Only when the default *is* null: then sending null and leaving the input
    out mean the same thing, and the schema already says the input is optional
    by its absence from ``required``. A null allowed beside some other default
    may mean something of its own, so it stays.
    """
    branches = prop.get("anyOf")
    if (
        not isinstance(branches, list)
        or _NULL not in branches
        or prop.get("default", ...) is not None
    ):
        return prop
    rest = [branch for branch in branches if branch != _NULL]
    lean = {key: value for key, value in prop.items() if key not in ("anyOf", "default")}
    return {**rest[0], **lean} if len(rest) == 1 else {**lean, "anyOf": rest}


def lean_input_schema(schema: dict) -> dict:
    """The input schema a client is sent, without what the model cannot use.

    Advertising only: the SDK validates arguments against the function's own
    Pydantic model, never this schema, so an explicit null is still accepted.
    Returns a new schema; the one passed in is left as it was.
    """
    lean = _without_titles(schema)
    if isinstance(lean.get("properties"), dict):
        lean["properties"] = {
            name: _without_null_default(prop) if isinstance(prop, dict) else prop
            for name, prop in lean["properties"].items()
        }
    return lean


def configure(mcp: Any, enabled: set[str] | None) -> None:
    """Apply annotations and lean schemas to every registered tool, and gate by toolsets.

    Mutates the MCPServer tool manager in place: sets each tool's ``annotations``,
    replaces its advertised ``parameters`` with `lean_input_schema`'s, and removes
    tools whose group isn't enabled. Called once at import time after all
    ``@mcp.tool`` decorators have run.
    """
    manager = mcp._tool_manager
    registered = list(manager._tools.items())
    names = [name for name, _ in registered]

    for name, tool in registered:
        tool.annotations = annotation_for(name)
        tool.parameters = lean_input_schema(tool.parameters)

    keep = select_kept_tools(names, enabled)
    for name in names:
        if name not in keep:
            manager.remove_tool(name)
