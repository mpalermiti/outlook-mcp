"""User tools: whoami, list_calendars."""

from __future__ import annotations

from typing import Any

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

    Returns calendars with id, name, color, is_default, and can_edit.
    """
    response = await graph_client.me.calendars.get()

    calendars = []
    for cal in response.value or []:
        color = ""
        if cal.color and hasattr(cal.color, "value"):
            color = cal.color.value

        calendars.append(
            {
                "id": cal.id,
                "name": sanitize_output(cal.name or ""),
                "color": color,
                "is_default": bool(cal.is_default_calendar),
                "can_edit": bool(cal.can_edit),
            }
        )

    return {
        "calendars": calendars,
        "count": len(calendars),
    }
