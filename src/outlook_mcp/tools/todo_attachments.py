"""To Do task attachment tools: list, download, upload, delete.

To Do attachments (taskFileAttachment) are not mail FileAttachments: the
resource lives under ``/me/todo/lists/{id}/tasks/{taskId}/attachments`` and
the content endpoint is ``.../attachments/{id}/$value``.

Creation is an inline base64 POST — a ``taskFileAttachment`` with
``contentBytes`` sent to the attachments collection — and on the accounts this
server targets that is not one option among two: the upload-session endpoint
(``POST .../tasks/{id}/attachmentSessions``) answers **404** on consumer
outlook.com mailboxes (verified live), so the session+chunked-PUT path cannot
work there at all. Inline POST was verified live on the same mailbox from 64
bytes to 4 MB.

Size ceiling: the inline POST is one JSON document, Graph refuses bodies over
30 MB, and base64 inflates content 4/3 — the hard limit is ~22.5 MB of raw
file. The client-side ceiling is 20 MiB, leaving room for the JSON envelope
and the filename.

Downloads read the attachment entity and take ``contentBytes`` from it (the
same style as ``mail_attachments.download_attachment``; kiota base64-decodes
into raw bytes) rather than minting a raw bearer token against ``$value``.
"""

from __future__ import annotations

import mimetypes
import os
import tempfile
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_TODO_WRITE, check_permission
from outlook_mcp.tools.mail_attachments import resolve_attachment_path
from outlook_mcp.tools.todo import _iso_datetime, _resolve_list_id
from outlook_mcp.validation import sanitize_output, validate_graph_id

# 20 MiB, not Graph's nominal 25 MB per taskFileAttachment: the inline POST is
# one JSON body, Graph rejects those at 30 MB, and base64 inflates the payload
# 4/3 — 25 MiB of file is ~33.4 MB of JSON and a guaranteed 400.
_MAX_ATTACHMENT_SIZE = 20 * 1024 * 1024


async def list_task_attachments(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List attachments on a To Do task.

    GET /me/todo/lists/{id}/tasks/{taskId}/attachments
    Returns {attachments: [{id, name, size, content_type, last_modified}],
    count, has_more, next_cursor} — same pagination shape as outlook_list_tasks.
    """
    task_id = validate_graph_id(task_id)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    query_params = apply_pagination({}, count, cursor)

    from msgraph.generated.users.item.todo.lists.item.tasks.item.attachments.attachments_request_builder import (  # noqa: E501
        AttachmentsRequestBuilder,
    )

    req_config = build_request_config(
        AttachmentsRequestBuilder.AttachmentsRequestBuilderGetQueryParameters,
        query_params,
    )
    response = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.get(request_configuration=req_config)
    )
    attachments = response.value or []
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "attachments": [
            {
                "id": att.id,
                "name": sanitize_output(att.name or ""),
                "size": att.size,
                "content_type": sanitize_output(att.content_type or ""),
                "last_modified": _iso_datetime(att.last_modified_date_time),
            }
            for att in attachments
        ],
        "count": len(attachments),
        "has_more": next_cursor is not None,
        "next_cursor": next_cursor,
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

    GET /me/todo/lists/{id}/tasks/{taskId}/attachments/{attId}, bytes from
    contentBytes (kiota decodes the base64). All bytes are fetched *before*
    the destination is touched, and the write lands via a temp file + atomic
    replace — an empty or failed download never truncates a staged file under
    its trusted name. A 0-byte attachment writes an honest 0-byte file.
    """
    task_id = validate_graph_id(task_id)
    attachment_id = validate_graph_id(attachment_id)
    save_path = resolve_attachment_path(save_path, config.attachments_dir)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    attachment = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.by_attachment_base_id(attachment_id)
        .get()
    )
    if attachment is None:
        raise ValueError(
            f"Graph returned no attachment for id {attachment_id} — it may be "
            "gone; re-list with outlook_list_task_attachments"
        )
    # contentBytes is Optional[bytes]; a 0-byte attachment legitimately comes
    # back empty, and a missing property is indistinguishable from empty here.
    content = attachment.content_bytes or b""

    fd, tmp_path = tempfile.mkstemp(
        dir=os.path.dirname(save_path), prefix=".download-", suffix=".tmp"
    )
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(content)
        os.replace(tmp_path, save_path)
    except BaseException:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise

    return {
        "saved_to": save_path,
        "name": sanitize_output(attachment.name or ""),
        "size": len(content),
        "content_type": sanitize_output(attachment.content_type or ""),
    }


async def upload_task_attachment(
    graph_client: Any,
    task_id: str,
    file_path: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Attach a local file to a To Do task via inline base64 POST.

    POST .../attachments with a taskFileAttachment (contentBytes base64).
    Accepted size: 1 byte – 20 MiB (the JSON body cap, see module docstring).
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_upload_task_attachment")
    task_id = validate_graph_id(task_id)

    file_path = resolve_attachment_path(file_path, config.attachments_dir)
    if not os.path.isfile(file_path):
        # ValueError, not FileNotFoundError: _wrap_tool_errors only forwards
        # ValueError text to the model, and the path is the thing it needs.
        raise ValueError(f"Attachment file not found: {file_path}")

    file_size = os.stat(file_path).st_size
    if file_size < 1:
        raise ValueError(
            f"Attachment file is empty ({file_size} bytes) — accepted size is "
            "1 byte – 20 MiB"
        )
    if file_size > _MAX_ATTACHMENT_SIZE:
        raise ValueError(
            f"Attachment is {file_size} bytes; task attachments are limited to "
            f"20 MiB ({_MAX_ATTACHMENT_SIZE} bytes). Graph caps the request "
            "body at 30 MB and base64 inflates the file 4/3, so anything "
            "larger is rejected before it leaves."
        )

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.task_file_attachment import TaskFileAttachment

    content_type, _ = mimetypes.guess_type(file_path)
    att = TaskFileAttachment()
    att.odata_type = "#microsoft.graph.taskFileAttachment"
    att.name = os.path.basename(file_path)
    att.content_type = content_type or "application/octet-stream"
    att.size = file_size
    with open(file_path, "rb") as f:
        # One read into memory — bounded by the 20 MiB gate above.
        att.content_bytes = f.read()

    response = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .attachments.post(att)
    )

    if response is None or response.id is None:
        # Optional[AttachmentBase] on the SDK side; an empty 201/204 must not
        # crash — and the attachment may exist server-side, so say that
        # instead of letting an agent's retry create a duplicate.
        raise ValueError(
            "Attachment upload got no id back from Graph — do not retry "
            "blindly (that may attach a duplicate); verify with "
            "outlook_list_task_attachments"
        )

    return {
        "status": "attached",
        "task_id": task_id,
        "attachment_id": response.id,
        "name": sanitize_output(response.name or att.name),
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
