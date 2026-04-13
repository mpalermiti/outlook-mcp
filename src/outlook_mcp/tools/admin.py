"""Admin tools: list_categories, get_mail_tips."""

from __future__ import annotations

from typing import Any

from outlook_mcp.validation import sanitize_output, validate_email


async def list_categories(graph_client: Any) -> dict:
    """List master categories from /me/outlook/masterCategories.

    Returns categories with id, display_name, and color.
    """
    response = await graph_client.me.outlook.master_categories.get()

    categories = []
    for cat in response.value or []:
        color = ""
        if cat.color and hasattr(cat.color, "value"):
            color = cat.color.value

        categories.append(
            {
                "id": cat.id,
                "display_name": sanitize_output(cat.display_name or ""),
                "color": color,
            }
        )

    return {
        "categories": categories,
        "count": len(categories),
    }


async def get_mail_tips(graph_client: Any, emails: list[str]) -> dict:
    """Get mail tips for a list of email addresses.

    Validates all emails, then POSTs to /me/getMailTips.
    Returns per-recipient delivery info: moderation, restrictions,
    max message size, and out-of-office messages.
    """
    # Validate all emails before making the API call
    for email in emails:
        validate_email(email)

    response = await graph_client.me.get_mail_tips.post(emails)

    tips = []
    for tip in response.value or []:
        email_addr = ""
        if tip.email_address and tip.email_address.address:
            email_addr = tip.email_address.address

        ooo_message = ""
        if tip.automatic_replies and tip.automatic_replies.message:
            ooo_message = sanitize_output(tip.automatic_replies.message)

        tips.append(
            {
                "email": email_addr,
                "is_delivery_restricted": bool(tip.delivery_restricted),
                "is_moderated": bool(tip.is_moderated),
                "max_message_size": tip.max_message_size or 0,
                "out_of_office_message": ooo_message,
            }
        )

    return {
        "tips": tips,
    }
