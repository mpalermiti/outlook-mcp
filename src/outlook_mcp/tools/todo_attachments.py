"""To Do task attachment tools: list, download, upload, delete.

To Do attachments (taskFileAttachment) are not mail FileAttachments: the
resource lives under ``/me/todo/lists/{id}/tasks/{taskId}/attachments``, the
content endpoint is ``.../attachments/{id}/$value`` (raw bytes, the SDK's
``.content`` builder), and creation only works through an upload session —
there is no inline base64 POST for this resource, so every upload size goes
through ``createUploadSession`` + chunked PUT. That is one code path for
0–25 MB instead of two, at the cost of an extra round trip on small files.

Graph caps a To Do attachment at 25 MB with no larger tier to fall back on
(unlike mail, where an oversize file could go in the message body instead).
"""

from __future__ import annotations

import mimetypes
import os
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_TODO_WRITE, check_permission
from outlook_mcp.tools.mail_attachments import _upload_large_file, resolve_attachment_path
from outlook_mcp.tools.todo import _resolve_list_id
from outlook_mcp.validation import sanitize_output, validate_graph_id

_MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024


async def list_task_attachments(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """List attachments on a To Do task.

    GET /me/todo/lists/{id}/tasks/{taskId}/attachments
    Returns {attachments: [{id, name, size, content_type, last_modified}], count}.
    """
    task_id = validate_graph_id(task_id)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    response = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.get()
    )
    attachments = response.value or []

    return {
        "attachments": [
            {
                "id": att.id,
                "name": sanitize_output(att.name or ""),
                "size": att.size,
                "content_type": att.content_type,
                "last_modified": str(att.last_modified_date_time or ""),
            }
            for att in attachments
        ],
        "count": len(attachments),
    }


async def download_task_attachment(
    graph_client: Any,
    task_id: str,
    attachment_id: str,
    save_path: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Download a To Do task attachment's content to a local file.

    GET /me/todo/lists/{id}/tasks/{taskId}/attachments/{attId}/$value
    Writes the raw bytes to save_path (confined to attachments_dir) and
    returns the path.
    """
    task_id = validate_graph_id(task_id)
    attachment_id = validate_graph_id(attachment_id)
    save_path = resolve_attachment_path(save_path, config.attachments_dir)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    content = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.by_attachment_base_id(attachment_id)
        .content.get()
    )

    with open(save_path, "wb") as f:
        f.write(content)
    return {
        "saved_to": save_path,
        "size": len(content),
    }


async def upload_task_attachment(
    graph_client: Any,
    task_id: str,
    file_path: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Attach a local file to a To Do task via an upload session.

    POST .../attachments/createUploadSession, then chunked PUT to the
    uploadUrl (0–25 MB; Graph has no larger tier for task attachments).
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_upload_task_attachment")
    task_id = validate_graph_id(task_id)

    file_path = resolve_attachment_path(file_path, config.attachments_dir)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Attachment file not found: {file_path}")

    file_size = os.path.getsize(file_path)
    if file_size == 0:
        # An empty range set makes the session un-PUTtable: there is no first
        # chunk to send, so the attachment can never complete.
        raise ValueError("Cannot attach an empty file (0 bytes).")
    if file_size > _MAX_ATTACHMENT_SIZE:
        raise ValueError(
            f"Attachment is {file_size} bytes; To Do caps task attachments at "
            f"25 MB ({_MAX_ATTACHMENT_SIZE} bytes)."
        )

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.attachment_info import AttachmentInfo
    from msgraph.generated.models.attachment_type import AttachmentType
    from msgraph.generated.users.item.todo.lists.item.tasks.item.attachments.create_upload_session.create_upload_session_post_request_body import (  # noqa: E501
        CreateUploadSessionPostRequestBody,
    )

    content_type, _ = mimetypes.guess_type(file_path)
    info = AttachmentInfo()
    info.attachment_type = AttachmentType.File
    info.name = os.path.basename(file_path)
    info.size = file_size
    info.content_type = content_type or "application/octet-stream"

    body = CreateUploadSessionPostRequestBody()
    body.attachment_info = info

    session = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.create_upload_session.post(body)
    )
    await _upload_large_file(session.upload_url, file_path, file_size)

    return {
        "status": "uploaded",
        "task_id": task_id,
        "name": info.name,
        "size": file_size,
    }


async def delete_task_attachment(
    graph_client: Any,
    task_id: str,
    attachment_id: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Remove an attachment from a To Do task.

    DELETE /me/todo/lists/{id}/tasks/{taskId}/attachments/{attId}
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_delete_task_attachment")
    task_id = validate_graph_id(task_id)
    attachment_id = validate_graph_id(attachment_id)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.by_attachment_base_id(attachment_id)
        .delete()
    )

    return {
        "status": "deleted",
        "task_id": task_id,
        "attachment_id": attachment_id,
    }
