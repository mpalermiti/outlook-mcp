"""Account routing: which account serves a given tool call.

Multi-account configs route *capabilities* (mail / calendar / contacts / todo)
to accounts, so one Microsoft identity per concern — mail on the account that
receives notifications, To Do on the one that holds the task lists. The
classification reuses ``toolsets.TOOL_GROUPS`` (already drift-guarded) and maps
each group onto a capability:

- mail-centric groups (mail, drafts, attachments, folders, admin, digest) fold
  into ``mail`` — they all act on the mailbox.
- calendar / contacts / todo map one-to-one.
- ``account`` tools (whoami, auth_status, …) follow the *active* account, not
  a capability: ``capability_for`` returns None for them.
- ``delta`` tools span capabilities and split by name.
"""

from __future__ import annotations

from collections.abc import Mapping

from outlook_mcp.toolsets import TOOL_GROUPS

_GROUP_CAPABILITY: dict[str, str] = {
    "mail": "mail",
    "drafts": "mail",
    "attachments": "mail",
    "folders": "mail",
    # changes_since composes the mail/calendar/contacts deltas; routed with
    # mail, the dominant capability — split setups that route calendar away
    # from mail should call the per-capability delta tools instead.
    "digest": "mail",
    "admin": "mail",
    "calendar": "calendar",
    "contacts": "contacts",
    "todo": "todo",
}

_TOOL_CAPABILITY_OVERRIDES: dict[str, str] = {
    "outlook_list_inbox_delta": "mail",
    "outlook_list_events_delta": "calendar",
    "outlook_list_contacts_delta": "contacts",
}


def capability_for(tool_name: str | None) -> str | None:
    """Routing capability for a tool, or None when it follows the active account."""
    if tool_name in _TOOL_CAPABILITY_OVERRIDES:
        return _TOOL_CAPABILITY_OVERRIDES[tool_name]
    group = TOOL_GROUPS.get(tool_name or "")
    if group is None or group == "account":
        return None
    return _GROUP_CAPABILITY.get(group)


# outlook_changes_since runs its mail/events/contacts deltas against one
# client, so it is only correct when all three route to the same account.
COMPOSED_DIGEST_CAPABILITIES = ("mail", "calendar", "contacts")


def composed_digest_conflict(routing: Mapping[str, str | None]) -> str | None:
    """Error message when outlook_changes_since cannot serve its capabilities.

    In a split setup the composed digest would silently return one account's
    calendar under another's mail, so refuse and point at the per-capability
    delta tools, which each route to their own account.
    """
    accounts = {routing.get(cap) for cap in COMPOSED_DIGEST_CAPABILITIES}
    if len(accounts) > 1:
        detail = ", ".join(f"{cap}->{routing.get(cap)}" for cap in COMPOSED_DIGEST_CAPABILITIES)
        return (
            "outlook_changes_since spans mail, calendar and contacts, which are "
            f"routed to different accounts ({detail}); one call cannot serve "
            "them all. Use outlook_list_inbox_delta, outlook_list_events_delta "
            "and outlook_list_contacts_delta instead — each routes to its own "
            "account."
        )
    return None
