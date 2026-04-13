"""Mail triage tools: move, delete, flag, categorize, mark_read."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.validation import validate_folder_name, validate_graph_id


def _check_read_only(read_only: bool, tool_name: str) -> None:
    if read_only:
        raise ReadOnlyError(tool_name)


async def move_message(
    graph_client: Any,
    message_id: str,
    folder: str,
    read_only: bool = False,
) -> dict:
    """Move a message to a folder."""
    _check_read_only(read_only, "outlook_move_message")
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
    read_only: bool = False,
) -> dict:
    """Delete a message. Soft delete (move to Deleted Items) by default."""
    _check_read_only(read_only, "outlook_delete_message")
    message_id = validate_graph_id(message_id)

    if permanent:
        await graph_client.me.messages.by_message_id(message_id).delete()
        return {"status": "permanently_deleted"}
    else:
        return await move_message(graph_client, message_id, "deleteditems")


async def flag_message(
    graph_client: Any,
    message_id: str,
    status: str,
    read_only: bool = False,
) -> dict:
    """Set follow-up flag on a message."""
    _check_read_only(read_only, "outlook_flag_message")
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
    read_only: bool = False,
) -> dict:
    """Set categories on a message."""
    _check_read_only(read_only, "outlook_categorize_message")
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
    read_only: bool = False,
) -> dict:
    """Mark a message as read or unread."""
    _check_read_only(read_only, "outlook_mark_read")
    message_id = validate_graph_id(message_id)

    from msgraph.generated.models.message import Message

    msg = Message()
    msg.is_read = is_read

    await graph_client.me.messages.by_message_id(message_id).patch(msg)
    return {"status": "updated", "is_read": is_read}
