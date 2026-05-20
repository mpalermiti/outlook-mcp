"""Mailbox settings tools: get_timezone, set_timezone."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_MAILBOX_SETTINGS, check_permission
from outlook_mcp.validation import sanitize_output


async def get_timezone(graph_client: Any) -> dict:
    """Get the server-side mailbox timezone from /me/mailboxSettings.

    Returns the `timeZone` field of the user's MailboxSettings — this is the
    timezone Exchange uses for calendar items and OOF/auto-reply scheduling,
    and is distinct from the local `config.timezone` used by this server for
    relative-date math on user input.

    Returns an empty string when the mailbox has no timezone configured.
    """
    settings = await graph_client.me.mailbox_settings.get()
    time_zone = getattr(settings, "time_zone", None) if settings else None
    return {"timezone": time_zone or ""}


async def set_timezone(
    graph_client: Any,
    timezone: str,
    *,
    config: Config,
) -> dict:
    """Set the server-side mailbox timezone via PATCH /me/mailboxSettings.

    Accepts both IANA names (e.g. "America/Los_Angeles") and Windows display
    names (e.g. "Pacific Standard Time"); unknown values are rejected by
    Microsoft Graph, not by a local allowlist.
    """
    check_permission(config, CATEGORY_MAILBOX_SETTINGS, "outlook_set_timezone")

    if not timezone or not timezone.strip():
        raise ValueError("timezone must not be empty or whitespace")

    timezone = sanitize_output(timezone).strip()

    from msgraph.generated.models.mailbox_settings import MailboxSettings

    body = MailboxSettings()
    body.time_zone = timezone

    response = await graph_client.me.mailbox_settings.patch(body)
    echoed = getattr(response, "time_zone", None) if response else None
    return {"status": "updated", "timezone": echoed or timezone}
