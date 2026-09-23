"""To Do task attachment tools: list, download, upload, delete.

To Do attachments (taskFileAttachment) are not mail FileAttachments: the
resource lives under ``/me/todo/lists/{id}/tasks/{taskId}/attachments`` and
the content endpoint is ``.../attachments/{id}/$value``.

Creation is an inline base64 POST — a ``taskFileAttachment`` with
``contentBytes`` sent to the attachments collection. The upload-session
route also exists on consumer outlook.com mailboxes
(``POST .../attachments/createUploadSession``, verified live), but its
upload URL is a ``graph.microsoft.com`` route rather than a pre-authenticated
``outlook.office.com`` one like mail's: every chunk PUT needs an
``Authorization`` header (401 "Access token is empty" without it, verified)
and a ``Content-Type`` header (400 without it, verified), plus response
checking and offset-following of ``nextExpectedRanges``. At the sizes this
tool accepts, inline is one round-trip through kiota's throttling and retry;
it was verified live on a consumer mailbox from 64 bytes to the full 20 MiB
ceiling. The session route is the documented future path above that.

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
from pathlib import Path
from typing import Any

import httpx

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

# Sentinel for "the response entity has no contentBytes property at all" —
# distinct from None, which a 0-byte attachment legitimately carries.
_NO_CONTENT_BYTES = object()


def _verified_content_bytes(attachment: Any, attachment_id: str) -> bytes:
    """contentBytes off the entity, cross-checked against its ``size``.

    A missing ``contentBytes`` used to be smoothed over with ``or b""``: the
    download wrote a 0-byte file under the trusted name and reported success,
    which is exactly what the atomic-write docstring promises can never
    happen. So the entity's own ``size`` — already in the response — is the
    witness: no bytes where size says there are some (or the reverse) is a
    broken response, not an empty attachment, and errors out before any file
    is touched.

    The property check exists because kiota picks the model class from
    ``@odata.type``: a response without it deserializes as ``AttachmentBase``,
    which has no ``content_bytes`` attribute at all, and a bare attribute read
    would raise AttributeError whose text never reaches the model.
    """
    content = getattr(attachment, "content_bytes", _NO_CONTENT_BYTES)
    if content is _NO_CONTENT_BYTES:
        raise ValueError(
            f"Attachment {attachment_id} came back with no contentBytes "
            "property (the response may lack @odata.type, so the SDK built "
            "the wrong entity type) — nothing was written; re-list with "
            "outlook_list_task_attachments and report the response"
        )
    declared = getattr(attachment, "size", None)
    # Only an int witnesses anything; anything else (mocks, absent property)
    # cannot contradict the payload and must not block an honest 0-byte write.
    if isinstance(declared, int):
        if not content and declared > 0:
            raise ValueError(
                f"Attachment {attachment_id} declares {declared} bytes but "
                "Graph returned no contentBytes — refusing to write an empty "
                "file and call it the attachment; nothing was written, "
                "re-try the download or re-list with "
                "outlook_list_task_attachments"
            )
        if content and declared == 0:
            raise ValueError(
                f"Attachment {attachment_id} returned {len(content)} bytes "
                "but declares size 0 — the response contradicts itself; "
                "nothing was written"
            )
    # contentBytes is Optional[bytes]; a 0-byte attachment legitimately
    # carries None or b"" past the witness check above.
    return content or b""


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


def _validated_download_target(save_path: str, attachments_dir: str) -> str:
    """Refuse a download target that cannot land, before any Graph call.

    Two shapes used to slip past the confinement check and blow up late:
    ``save_path`` resolving to the attachments directory itself (``"."`` is
    relative to it, and ``is_relative_to`` is satisfied) sent ``dirname()`` one
    level *above* the fence, so the temp file landed next to ``config.json``;
    and a path whose parent does not exist made it all the way through the
    fetch before ``mkstemp`` raised FileNotFoundError — an OS error whose text
    never reaches the model. Both are caller-input problems, so both are
    ValueError, checked before a single byte is requested.
    """
    resolved = Path(save_path)
    if resolved.is_dir():
        raise ValueError(
            f"Attachment download target is a directory, not a file: "
            f"{resolved}. Pass a file path inside {attachments_dir}."
        )
    if not resolved.parent.is_dir():
        raise ValueError(
            f"Attachment download directory does not exist: {resolved.parent}. "
            "Create it first — downloads do not create directories on demand."
        )
    return save_path


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
    contentBytes (kiota decodes the base64). The payload is cross-checked
    against the entity's own `size` — a response whose bytes contradict its
    declared size errors out rather than writing half an attachment. All
    bytes are fetched *before* the destination is touched, and the write lands
    via a temp file + atomic replace — an empty or failed download never
    truncates a staged file under its trusted name. A 0-byte attachment
    (size 0, no contentBytes) writes an honest 0-byte file.
    """
    task_id = validate_graph_id(task_id)
    attachment_id = validate_graph_id(attachment_id)
    save_path = resolve_attachment_path(save_path, config.attachments_dir)
    save_path = _validated_download_target(save_path, config.attachments_dir)
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
    content = _verified_content_bytes(attachment, attachment_id)

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

    # Race-free size gate. `os.stat` then an unbounded `f.read()` let a file
    # that grows in the window between the two bypass the cap entirely (the
    # gate can see 1 KiB while the wire gets 26 MB) and reported a stale
    # size. The read itself is the gate: capped at MAX+1 bytes, so an
    # over-cap file is caught without ever being read whole, and the sent
    # size is always the length actually read.
    with open(file_path, "rb") as f:
        content = f.read(_MAX_ATTACHMENT_SIZE + 1)
    file_size = len(content)
    if file_size < 1:
        raise ValueError(
            f"Attachment file is empty ({file_size} bytes) — accepted size is "
            "1 byte – 20 MiB"
        )
    if file_size > _MAX_ATTACHMENT_SIZE:
        raise ValueError(
            f"Attachment is over the 20 MiB ceiling (read {file_size} bytes "
            f"of it; the cap is {_MAX_ATTACHMENT_SIZE} bytes). Graph caps "
            "request bodies at 30 MB and base64 inflates the file 4/3, so "
            "anything larger is rejected before it leaves."
        )

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.task_file_attachment import TaskFileAttachment

    content_type, _ = mimetypes.guess_type(file_path)
    att = TaskFileAttachment()
    att.odata_type = "#microsoft.graph.taskFileAttachment"
    att.name = os.path.basename(file_path)
    att.content_type = content_type or "application/octet-stream"
    att.size = file_size
    att.content_bytes = content

    from kiota_abstractions.base_request_configuration import RequestConfiguration
    from kiota_http.middleware.options.retry_handler_option import RetryHandlerOption

    # One POST, no retries. kiota's RetryHandler re-sends on 429/503/504, and
    # for this call that means up to three more copies of a ~28 MB JSON body
    # (112 MB of upload for one attachment) — and a 504 that arrives *after*
    # the server committed creates duplicate attachments that no read-back
    # guard can see, because the failure looks transport-level. A per-request
    # option disables just this POST; the handler stays on for everything
    # else the client sends.
    one_shot = RequestConfiguration(options=[RetryHandlerOption(should_retry=False)])

    try:
        response = await (
            graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
            .tasks.by_todo_task_id(task_id)
            .attachments.post(att, request_configuration=one_shot)
        )
    except httpx.TimeoutException as exc:
        # With retries off, a transport timeout is a single unanswered
        # attempt — but the server may still have committed it. Say the
        # verify-first thing, same as the no-id branch below.
        raise ValueError(
            "Attachment upload timed out — the POST is not auto-retried "
            "(re-sending a ~28 MB body may attach a duplicate). Verify with "
            "outlook_list_task_attachments whether it landed before retrying."
        ) from exc

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
