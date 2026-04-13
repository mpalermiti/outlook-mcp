"""Mail thread tools: list_thread, copy_message."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import build_request_config
from outlook_mcp.tools.mail_read import _format_message_summary
from outlook_mcp.validation import validate_folder_name, validate_graph_id


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _check_read_only(read_only: bool, tool_name: str) -> None:
    if read_only:
        raise ReadOnlyError(tool_name)


async def list_thread(
    graph_client: Any,
    conversation_id: str,
    count: int = 50,
) -> dict:
    """List messages in a conversation thread, chronological order."""
    conversation_id = validate_graph_id(conversation_id)
    count = _clamp(count, 1, 100)

    query_params = {
        "$filter": f"conversationId eq '{conversation_id}'",
        "$orderby": "receivedDateTime asc",
        "$top": count,
        "$select": (
            "id,subject,from,receivedDateTime,isRead,importance,"
            "bodyPreview,hasAttachments,categories,flag,conversationId"
        ),
    }

    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    req_config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.messages.get(request_configuration=req_config)

    messages = [_format_message_summary(m) for m in (response.value or [])]

    return {
        "messages": messages,
        "count": len(messages),
    }


async def copy_message(
    graph_client: Any,
    message_id: str,
    folder: str,
    read_only: bool = False,
) -> dict:
    """Copy a message to a folder."""
    _check_read_only(read_only, "outlook_copy_message")
    message_id = validate_graph_id(message_id)
    folder = validate_folder_name(folder)

    from msgraph.generated.users.item.messages.item.copy.copy_post_request_body import (
        CopyPostRequestBody,
    )

    request_body = CopyPostRequestBody()
    request_body.destination_id = folder

    await graph_client.me.messages.by_message_id(message_id).copy.post(request_body)
    return {"status": "copied", "folder": folder}
