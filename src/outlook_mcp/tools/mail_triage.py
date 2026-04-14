"""Mail triage tools: move, delete, flag, categorize, mark_read."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.validation import validate_folder_name, validate_graph_id


async def move_message(
    graph_client: Any,
    message_id: str,
    folder: str,
    *,
    config: Config,
) -> dict:
    """Move a message to a folder."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_move_message")
    message_id = validate_graph_id(message_id)
    folder = validate_folder_name(folder)

    from msgraph.generated.users.item.messages.item.move.move_post_request_body import (
        MovePostRequestBody,
    )

    request_body = MovePostRequestBody()
    request_body.destination_id = folder

    await graph_client.me.messages.by_message_id(message_id).move.post(request_body)
    return {"status": "moved", "folder": folder}


async def delete_message(
    graph_client: Any,
    message_id: str,
    permanent: bool = False,
    *,
    config: Config,
) -> dict:
    """Delete a message. Soft delete (move to Deleted Items) by default."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_delete_message")
    message_id = validate_graph_id(message_id)

    if permanent:
        await graph_client.me.messages.by_message_id(message_id).delete()
        return {"status": "permanently_deleted"}
    else:
        return await move_message(graph_client, message_id, "deleteditems", config=config)


async def flag_message(
    graph_client: Any,
    message_id: str,
    status: str,
    *,
    config: Config,
) -> dict:
    """Set follow-up flag on a message."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_flag_message")
    message_id = validate_graph_id(message_id)

    valid_statuses = ("flagged", "complete", "notFlagged")
    if status not in valid_statuses:
        raise ValueError(f"flag status must be one of {valid_statuses}; got {status}")

    from msgraph.generated.models.followup_flag import FollowupFlag
    from msgraph.generated.models.followup_flag_status import FollowupFlagStatus
    from msgraph.generated.models.message import Message

    status_map = {
        "flagged": FollowupFlagStatus.Flagged,
        "complete": FollowupFlagStatus.Complete,
        "notFlagged": FollowupFlagStatus.NotFlagged,
    }

    msg = Message()
    msg.flag = FollowupFlag()
    msg.flag.flag_status = status_map[status]

    await graph_client.me.messages.by_message_id(message_id).patch(msg)
    return {"status": "flagged", "flag_status": status}


async def categorize_message(
    graph_client: Any,
    message_id: str,
    categories: list[str],
    *,
    config: Config,
) -> dict:
    """Set categories on a message."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_categorize_message")
    message_id = validate_graph_id(message_id)

    from msgraph.generated.models.message import Message

    msg = Message()
    msg.categories = categories

    await graph_client.me.messages.by_message_id(message_id).patch(msg)
    return {"status": "categorized", "categories": categories}


async def mark_read(
    graph_client: Any,
    message_id: str,
    is_read: bool,
    *,
    config: Config,
) -> dict:
    """Mark a message as read or unread."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_mark_read")
    message_id = validate_graph_id(message_id)

    from msgraph.generated.models.message import Message

    msg = Message()
    msg.is_read = is_read

    await graph_client.me.messages.by_message_id(message_id).patch(msg)
    return {"status": "updated", "is_read": is_read}
