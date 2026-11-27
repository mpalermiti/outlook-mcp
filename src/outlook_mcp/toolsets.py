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
   ``OUTLOOK_MCP_TOOLSETS=mail,calendar``, or ``"toolsets": "mail,calendar"`` in
   config.json, and load just those groups (~half the tokens for a
   mail+calendar agent). The variable wins when it is set and not blank
   (`resolve_selection`). Account/auth tools are always available.

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

from collections.abc import Collection
from typing import Any

from mcp.types import ToolAnnotations

from outlook_mcp.errors import LEAVE_SETTINGS_TO_THE_USER

# Toolset groups a client can enable via OUTLOOK_MCP_TOOLSETS or config.json's
# `toolsets`. The "account" group (auth/identity) is always available
# regardless of selection.
ALWAYS_ON_GROUPS = {"account"}

# Names a selection may use that are not groups. `classic` is what an unset
# selection loads, so a person can name the default and come back to it.
# `resolve_groups` expands it, which keeps `mailbox_settings_write` adding
# `rules` to it as it does to an unset selection, and lets `classic,documents`
# mean the default plus the document reader.
ALIASES = frozenset({"classic"})

# The variable that names a selection, read by server.py at import and by the CLI.
TOOLSETS_ENV = "OUTLOOK_MCP_TOOLSETS"

# Where a selection came from, as `resolve_selection` reports it, and how a
# warning names each source.
SOURCE_NAMES = {
    "env": TOOLSETS_ENV,
    "config": "config.json toolsets",
    "default": "the default",
}

# The home group of the request-level intent tools (`src/outlook_mcp/intents/`).
# Off by default and opt-in like the others below: `outlook_find` is its first
# tool, an early preview, so nothing loads it unless a selection names `intents`.
INTENTS = "intents"

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
# `contact_photos` too: accepted on #121 as an opt-in tool only, so the default
# surface and its ceiling do not move for a tool most clients never call.
# `intents`, the request-level tools, while they are a preview.
OFF_BY_DEFAULT = {"documents", "rules", "mime", "contact_photos", INTENTS}

# What `mailbox_settings_write: true` adds when nothing was named. Not every
# off-by-default group: that key asks for rule writes, not for the document
# reader.
SETTINGS_WRITE_ADDS = {"rules"}

# name -> toolset group. Every registered tool MUST appear here (drift-guarded).
TOOL_GROUPS: dict[str, str] = {
    # account (always on)
    "outlook_auth_status": "account",
    "outlook_whoami": "account",
    # intents (opt-in, preview)
    "outlook_find": INTENTS,
    "outlook_agenda": INTENTS,
    "outlook_open_loops": INTENTS,
    "outlook_catch_up": INTENTS,
    "outlook_brief": INTENTS,
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
    # A local file uploaded to Graph; opt-in (#121).
    "outlook_set_contact_photo": "contact_photos",
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

# The follow-up core (intent contracts §10.5): classic tools that the intent tools'
# `next` suggestions point to, or that an agent needs to act on what they find
# (open a message or a thread, read several at once, download a file, update or
# complete a task, list drafts). They keep their home group in `TOOL_GROUPS`, so
# every classification and drift guard keeps its meaning, and also belong to
# `intents`, so a selection that names `intents` alone never suggests a tool it
# lacks. A selection that doesn't name `intents` loads exactly what it did.
FOLLOW_UP_CORE = frozenset(
    {
        "outlook_read_message",
        "outlook_read_messages",
        "outlook_list_thread",
        "outlook_download_attachment",
        "outlook_update_task",
        "outlook_complete_task",
        "outlook_list_drafts",
    }
)
EXTRA_GROUPS: dict[str, frozenset[str]] = {name: frozenset({INTENTS}) for name in FOLLOW_UP_CORE}


def groups_of(name: str) -> frozenset[str]:
    """Every group ``name`` belongs to: its home group, and `intents` for the core.

    The only place a tool's groups are read when choosing what loads, on both the
    unset path and the named one, so the two cannot disagree (the "one set, not
    two" bug above).
    """
    home = TOOL_GROUPS.get(name)
    return (frozenset({home}) if home else frozenset()) | EXTRA_GROUPS.get(name, frozenset())


# Tools that only read — no mailbox/calendar mutation. (download_attachment is
# deliberately excluded: it writes a local file, so it isn't read-only.)
READ_ONLY: set[str] = {
    "outlook_find",
    "outlook_agenda",
    "outlook_open_loops",
    "outlook_catch_up",
    "outlook_brief",
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


def toolset_names(value: str | None) -> list[str]:
    """The names in a selection, in the order given, each once.

    The one grammar for ``OUTLOOK_MCP_TOOLSETS``, config.json's ``toolsets``
    and ``outlook-mcp toolsets``: comma-separated, lowercased and stripped;
    empty parts are skipped.
    """
    names: list[str] = []
    for part in (value or "").split(","):
        name = part.strip().lower()
        if name and name not in names:
            names.append(name)
    return names


def parse_toolsets(value: str | None) -> set[str] | None:
    """Parse a selection into a set of names.

    ``None`` (unset) or an empty/blank value means "all default toolsets" — the
    backward-compatible default. Otherwise a comma-separated list, lowercased
    and stripped. Unknown group names simply match no tools.
    """
    return set(toolset_names(value)) or None


def resolve_selection(env_value: str | None, file_value: str | None) -> tuple[set[str] | None, str]:
    """The selection in force and where it came from: ``"env"``, ``"config"`` or ``"default"``.

    ==========================  ==========================  ===============
    ``OUTLOOK_MCP_TOOLSETS``    config.json ``toolsets``    selection
    ==========================  ==========================  ===============
    set, not blank              anything                    the variable's
    unset or blank              set, not blank              the file's
    unset or blank              absent or blank             ``None``
    ==========================  ==========================  ===============

    ``None`` is the default set, as it always was. A blank variable falls
    through to the file rather than meaning "the default" outright: before the
    file key existed the two were the same, and a launch template that passes
    an empty value through should not hide the person's own setting.
    """
    enabled = parse_toolsets(env_value)
    if enabled is not None:
        return enabled, "env"
    enabled = parse_toolsets(file_value)
    if enabled is not None:
        return enabled, "config"
    return None, "default"


def known_toolset_names() -> set[str]:
    """Every name a selection can use: each group with tools, and the aliases."""
    return set(TOOL_GROUPS.values()) | ALIASES


def unknown_toolset_names(enabled: set[str] | None) -> set[str]:
    """The names in ``enabled`` that are neither a group nor an alias; they load nothing."""
    return set(enabled or ()) - known_toolset_names()


def resolve_groups(enabled: set[str] | None, mailbox_settings_write: bool) -> set[str]:
    """Enabled groups, from the selection and the config flag.

    ``enabled`` is the selection `resolve_selection` found: ``None`` means "no
    selection", which resolves to every group **except** ``OFF_BY_DEFAULT``.
    ``mailbox_settings_write`` adds ``SETTINGS_WRITE_ADDS`` back, but only when
    nothing else was named: an agent gated down to ``mail`` must not have
    ``rules`` handed back by a shared config it did not write, with no way to
    ask for the narrower surface. ``classic`` names that same unset default, so
    it is replaced by it here, flag included, and anything named beside it is
    added on top.

    Always a set. A ``None`` return would mean "everything" to
    ``select_kept_tools``, which would silently keep every off-by-default group
    for any caller that passed it straight through.
    """
    if enabled is None:
        groups = set(TOOL_GROUPS.values()) - OFF_BY_DEFAULT
        if mailbox_settings_write:
            groups |= SETTINGS_WRITE_ADDS
        return groups
    groups = set(enabled) - ALIASES
    if set(enabled) & ALIASES:
        groups |= resolve_groups(None, mailbox_settings_write)
    return groups


def select_kept_tools(names: list[str], enabled: set[str] | None) -> set[str]:
    """Return the subset of ``names`` to keep given the enabled toolsets.

    ``enabled=None`` keeps every group except the off-by-default ones. Otherwise
    keep a tool when one of its groups (`groups_of`) is enabled, or its group is
    an always-on one (account/auth). A name with no known group is kept (fail-open
    — gating must never silently drop an unclassified tool).
    """
    keep: set[str] = set()
    for name in names:
        group = TOOL_GROUPS.get(name)
        if group is None or group in ALWAYS_ON_GROUPS:
            keep.add(name)  # fail-open for an unclassified tool, as ever
        elif enabled is None:
            # Same set the default path subtracts: with only `documents` here, this
            # kept the four rule tools for a caller that meant the stock surface.
            if groups_of(name) - OFF_BY_DEFAULT:
                keep.add(name)
        elif groups_of(name) & enabled:
            keep.add(name)
    return keep


def loaded_tools(groups: set[str]) -> frozenset[str]:
    """The tools a process loads for ``groups``, from the classification alone.

    `configure` keeps the same tools on a registered server. This answers the
    question before any tool registers (``server.LOADED_TOOLS``) and without
    importing the server at all (``outlook-mcp status``).
    """
    return frozenset(select_kept_tools(list(TOOL_GROUPS), groups))


# ── The tip in classic results ────────────────────────────────────────────
#
# A classic tool's result may carry one short tip, once per process, naming the
# newer tool that does the same job in fewer calls. The server adds it (the
# `_wrap_tool_errors` wrapper in server.py) as a top-level `server_tip` key;
# tool modules never do, and mail text sits inside result fields, so it can
# neither create nor imitate that key. A tool whose result is not a dict
# (`outlook_export_message`) has no entry.
#
# `tip_for` names a tool only if this release has it. Addressed to the person:
# switching changes the server's settings, which an agent leaves to the user.
#
# Tips are on (`TIPS_LIVE`) from 1.35.0, the owner's go decision on the intent
# tools (step 6). Only a classic tool whose newer tool exists in this release
# carries one: today `outlook_find`, `outlook_agenda` and `outlook_catch_up`.
# `show_tips: false` (or `OUTLOOK_MCP_SHOW_TIPS=false`) turns them off; a process
# with `intents` loaded never shows one. `test_intent_tip.py` holds all of this.
TIPS_LIVE = True
TIP_MAP: dict[str, str] = {
    "outlook_list_inbox": "outlook_catch_up",
    "outlook_list_inbox_delta": "outlook_catch_up",
    "outlook_changes_since": "outlook_catch_up",
    "outlook_search_mail": "outlook_find",
    "outlook_search_contacts": "outlook_find",
    "outlook_find_event_by_key": "outlook_find",
    "outlook_list_events": "outlook_agenda",
    "outlook_list_events_delta": "outlook_agenda",
    "outlook_create_event": "outlook_schedule",
    "outlook_update_event": "outlook_change_meeting",
    "outlook_delete_event": "outlook_change_meeting",
    "outlook_rsvp": "outlook_answer_invites",
    "outlook_send_message": "outlook_compose",
    "outlook_reply": "outlook_compose",
    "outlook_forward": "outlook_compose",
    "outlook_create_draft": "outlook_compose",
    "outlook_batch_triage": "outlook_tidy",
    "outlook_move_message": "outlook_tidy",
    "outlook_delete_message": "outlook_tidy",
    "outlook_flag_message": "outlook_follow_up",
    "outlook_create_task": "outlook_capture_task",
    "outlook_create_contact": "outlook_save_contact",
    "outlook_update_contact": "outlook_save_contact",
}

TIP_TEXT = (
    "Tip for the user: outlook-mcp's opt-in intent tools can do this in fewer calls "
    "({intent_tool}). To turn them on, the user runs `outlook-mcp toolsets classic,intents` "
    "and restarts the client. " + LEAVE_SETTINGS_TO_THE_USER
)


def tip_for(tool: str, loaded: Collection[str]) -> str | None:
    """The tip ``tool``'s result may carry in a process that loaded ``loaded``, or None.

    None while `TIPS_LIVE` is off (it is on from 1.35.0). Otherwise None
    unless ``tool`` is in `TIP_MAP`, the tool it names exists in this release,
    and the process loaded no tool of that group: someone already using the
    newer tools has nothing to switch to. Whether the tip was already shown,
    and ``show_tips``, are the server's to check.
    """
    if not TIPS_LIVE:
        return None
    newer = TIP_MAP.get(tool)
    if newer is None or TOOL_GROUPS.get(newer) != INTENTS:
        return None
    if any(TOOL_GROUPS.get(name) == INTENTS for name in loaded):
        return None
    return TIP_TEXT.format(intent_tool=newer)


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
    # A PUT of the photo bytes: a repeat replaces the photo with the same image.
    "outlook_set_contact_photo",
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
