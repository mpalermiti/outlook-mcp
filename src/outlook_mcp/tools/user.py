"""User tools: whoami, list_calendars."""

from __future__ import annotations

from typing import Any

from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.validation import sanitize_output


async def whoami(graph_client: Any) -> dict:
    """Get current user profile from /me.

    Returns display_name, email, and id.
    """
    user = await graph_client.me.get()

    email = user.mail or user.user_principal_name or ""
    display_name = user.display_name or ""

    return {
        "display_name": sanitize_output(display_name),
        "email": email,
        "id": user.id,
    }


async def list_calendars(graph_client: Any) -> dict:
    """List all calendars for the current user.

    Returns calendars with id, name, color, is_default, can_edit, owner,
    and can_view_private_items. Missing owner/permission metadata stays null.
    """
    calendars = []
    for cal in await fetch_all_calendars(graph_client):
        color = ""
        if cal.color and hasattr(cal.color, "value"):
            color = cal.color.value
        owner = getattr(cal, "owner", None)

        calendars.append(
            {
                "id": cal.id,
                "name": sanitize_output(cal.name or ""),
                "color": color,
                "is_default": bool(cal.is_default_calendar),
                "can_edit": bool(cal.can_edit),
                "owner": (
                    {
                        "name": sanitize_output(owner.name) if owner.name is not None else None,
                        "address": sanitize_output(owner.address)
                        if owner.address is not None
                        else None,
                    }
                    if owner is not None
                    else None
                ),
                "can_view_private_items": getattr(cal, "can_view_private_items", None),
            }
        )

    return {
        "calendars": calendars,
        "count": len(calendars),
    }
