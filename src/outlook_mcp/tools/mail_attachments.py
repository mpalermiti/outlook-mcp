"""Mail attachment tools: list, download, send / attach to drafts."""

from __future__ import annotations

import mimetypes
import os
from pathlib import Path
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import (
    CATEGORY_MAIL_DRAFTS,
    CATEGORY_MAIL_SEND,
    check_permission,
)
from outlook_mcp.validation import validate_email, validate_graph_id

# 3MB threshold — files above this use upload sessions
_LARGE_FILE_THRESHOLD = 3 * 1024 * 1024
# Chunk size for upload sessions (320 KiB aligned, as required by Graph)
_UPLOAD_CHUNK_SIZE = 320 * 1024 * 10  # 3.2 MB chunks


def resolve_attachment_path(path: str, attachments_dir: str) -> str:
    """Resolve ``path`` inside ``attachments_dir``, or refuse.

    These tools take a filesystem path from the model, and the model takes
    instructions from email. So the path is untrusted input: "attach the file at
    <path> and reply" is a working exfiltration primitive unless the reachable
    set is bounded. Every read and write goes through here.

    Confinement is resolved, not textual. A substring test for ``..`` — which is
    all this did before 1.20.0 — passes an absolute path to anywhere, and passes
    a symlink sitting innocently inside the directory. Resolving first collapses
    both, and comparing against the resolved base means a sibling directory
    sharing a name prefix cannot slip through either.

    A relative path is taken as relative to ``attachments_dir``, so an agent that
    passes a bare filename lands somewhere predictable instead of the process's
    working directory.
    """
    if not path or not path.strip() or "\x00" in path:
        raise ValueError("Attachment path must be a non-empty path containing no null bytes.")

    base = Path(os.path.expanduser(attachments_dir))
    created = not base.exists()
    base.mkdir(parents=True, exist_ok=True)
    if created:
        # Ours to lock down. Never chmod a directory the user pointed us at.
        base.chmod(0o700)
    base = base.resolve()

    candidate = Path(os.path.expanduser(path))
    if not candidate.is_absolute():
        candidate = base / candidate
    # strict=False by default: download writes a file that does not exist yet.
    resolved = candidate.resolve()

    if not resolved.is_relative_to(base):
        raise ValueError(
            f"Attachment path is outside the permitted directory: {path}. "
            f"Attachments may only be read from or written to {attachments_dir} "
            f"(set `attachments_dir` in ~/.outlook-mcp/config.json to change it)."
        )
    return str(resolved)


def _make_inline_attachment(file_path: str) -> Any:
    """Build a FileAttachment SDK object from a local file (<=3MB path)."""
    from msgraph.generated.models.file_attachment import FileAttachment

    att = FileAttachment()
    att.name = os.path.basename(file_path)
    with open(file_path, "rb") as f:
        att.content_bytes = f.read()
    content_type, _ = mimetypes.guess_type(file_path)
    att.content_type = content_type or "application/octet-stream"
    att.odata_type = "#microsoft.graph.fileAttachment"
    return att


async def list_attachments(
    graph_client: Any,
    message_id: str,
) -> dict:
    """List attachments on a message.

    GET /me/messages/{id}/attachments
    Returns {attachments: [{id, name, size, content_type}], count}.
    """
    message_id = validate_graph_id(message_id)

    response = await graph_client.me.messages.by_message_id(message_id).attachments.get()
    attachments = response.value or []

    return {
        "attachments": [
            {
                "id": att.id,
                "name": att.name,
                "size": att.size,
                "content_type": att.content_type,
            }
            for att in attachments
        ],
        "count": len(attachments),
    }


async def download_attachment(
    graph_client: Any,
    message_id: str,
    attachment_id: str,
    save_path: str,
    *,
    config: Config,
) -> dict:
    """Download an attachment.

    GET /me/messages/{id}/attachments/{att_id}
    Writes the file bytes to save_path and returns the path.
    """
    message_id = validate_graph_id(message_id)
    attachment_id = validate_graph_id(attachment_id)
    save_path = resolve_attachment_path(save_path, config.attachments_dir)

    attachment = (
        await graph_client.me.messages.by_message_id(message_id)
        .attachments.by_attachment_id(attachment_id)
        .get()
    )
    # The msgraph SDK (Kiota) already base64-decodes contentBytes into raw bytes
    # during deserialization, so content_bytes is the raw file content. Decoding
    # again corrupts binary files / raises UnicodeDecodeError (issue #25).
    content = attachment.content_bytes

    with open(save_path, "wb") as f:
        f.write(content)
    return {
        "saved_to": save_path,
        "name": attachment.name,
        "size": attachment.size,
        "content_type": attachment.content_type,
    }


async def _upload_large_file(
    upload_url: str,
    file_path: str,
    file_size: int,
) -> None:
    """Upload a large file in chunks via an upload session.

    Uses httpx to PUT chunks to the upload URL provided by Graph.
    """
    import httpx

    async with httpx.AsyncClient() as client:
        offset = 0
        with open(file_path, "rb") as f:
            while offset < file_size:
                chunk = f.read(_UPLOAD_CHUNK_SIZE)
                chunk_size = len(chunk)
                end = offset + chunk_size - 1
                headers = {
                    "Content-Range": f"bytes {offset}-{end}/{file_size}",
                    "Content-Length": str(chunk_size),
                }
                await client.put(upload_url, content=chunk, headers=headers)
                offset += chunk_size


async def send_with_attachments(
    graph_client: Any,
    to: list[str],
    subject: str,
    body: str,
    attachment_paths: list[str],
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    reply_to: list[str] | None = None,
    *,
    config: Config,
) -> dict:
    """Send a message with file attachments.

    For files under 3MB: inline as base64 FileAttachment.
    For files over 3MB: create draft, use createUploadSession + chunked upload, then send.
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_with_attachments")

    # Validate emails
    validated_to = [validate_email(e) for e in to]
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []
    validated_reply_to = [validate_email(e) for e in reply_to] if reply_to else []

    # Confine every path before touching the filesystem, then validate existence.
    attachment_paths = [
        resolve_attachment_path(path, config.attachments_dir) for path in attachment_paths
    ]
    for path in attachment_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")

    # Partition files into small (inline) and large (upload session)
    small_files = []
    large_files = []
    for path in attachment_paths:
        file_size = os.path.getsize(path)
        if file_size > _LARGE_FILE_THRESHOLD:
            large_files.append((path, file_size))
        else:
            small_files.append(path)

    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.message import Message
    from msgraph.generated.models.recipient import Recipient

    def _make_recipient(email: str) -> Recipient:
        r = Recipient()
        r.email_address = EmailAddress()
        r.email_address.address = email
        return r

    def _build_message() -> Message:
        msg = Message()
        msg.subject = subject
        msg.body = ItemBody()
        msg.body.content = body
        msg.body.content_type = BodyType.Html if is_html else BodyType.Text
        msg.to_recipients = [_make_recipient(e) for e in validated_to]
        if validated_cc:
            msg.cc_recipients = [_make_recipient(e) for e in validated_cc]
        if validated_bcc:
            msg.bcc_recipients = [_make_recipient(e) for e in validated_bcc]
        if validated_reply_to:
            msg.reply_to = [_make_recipient(e) for e in validated_reply_to]
        importance_map = {
            "low": Importance.Low,
            "normal": Importance.Normal,
            "high": Importance.High,
        }
        msg.importance = importance_map.get(importance, Importance.Normal)
        return msg

    if not large_files:
        # All small — send inline via sendMail
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        msg = _build_message()
        msg.attachments = [_make_inline_attachment(p) for p in small_files]

        request_body = SendMailPostRequestBody()
        request_body.message = msg
        request_body.save_to_sent_items = True

        await graph_client.me.send_mail.post(request_body)
    else:
        # Has large files — create draft, attach via upload sessions, then send
        from msgraph.generated.models.attachment_item import AttachmentItem
        from msgraph.generated.models.attachment_type import AttachmentType
        from msgraph.generated.users.item.messages.item.attachments.create_upload_session.create_upload_session_post_request_body import (  # noqa: E501
            CreateUploadSessionPostRequestBody,
        )

        msg = _build_message()
        # Attach small files inline on the draft
        msg.attachments = [_make_inline_attachment(p) for p in small_files]

        # Create draft message
        draft = await graph_client.me.messages.post(msg)

        # Upload each large file
        for file_path, file_size in large_files:
            content_type, _ = mimetypes.guess_type(file_path)
            att_item = AttachmentItem()
            att_item.attachment_type = AttachmentType.File
            att_item.name = os.path.basename(file_path)
            att_item.size = file_size
            att_item.content_type = content_type or "application/octet-stream"

            upload_body = CreateUploadSessionPostRequestBody()
            upload_body.attachment_item = att_item

            session = (
                await graph_client.me.messages.by_message_id(draft.id)
                .attachments.create_upload_session.post(upload_body)
            )

            await _upload_large_file(session.upload_url, file_path, file_size)

        # Send the draft
        await graph_client.me.messages.by_message_id(draft.id).send.post()

    return {
        "status": "sent",
        "attachment_count": len(attachment_paths),
    }


async def attach_to_draft(
    graph_client: Any,
    draft_id: str,
    attachment_paths: list[str],
    *,
    config: Config,
) -> dict:
    """Add one or more attachments to an existing draft message.

    For files under 3MB: POST a FileAttachment directly.
    For files over 3MB: createUploadSession + chunked upload.

    Returns the new attachment IDs so callers can reference or
    remove individual attachments later.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_attach_to_draft")
    draft_id = validate_graph_id(draft_id)

    # Confine every path before touching the filesystem, then validate existence.
    attachment_paths = [
        resolve_attachment_path(path, config.attachments_dir) for path in attachment_paths
    ]
    for path in attachment_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")

    # Partition files into small (inline) and large (upload session)
    small_files: list[str] = []
    large_files: list[tuple[str, int]] = []
    for path in attachment_paths:
        file_size = os.path.getsize(path)
        if file_size > _LARGE_FILE_THRESHOLD:
            large_files.append((path, file_size))
        else:
            small_files.append(path)

    attachment_ids: list[str] = []
    msg_builder = graph_client.me.messages.by_message_id(draft_id)

    # Small files — POST each as an inline FileAttachment
    for file_path in small_files:
        att = _make_inline_attachment(file_path)
        created = await msg_builder.attachments.post(att)
        if created is not None and getattr(created, "id", None):
            attachment_ids.append(created.id)

    # Large files — upload session
    if large_files:
        from msgraph.generated.models.attachment_item import AttachmentItem
        from msgraph.generated.models.attachment_type import AttachmentType
        from msgraph.generated.users.item.messages.item.attachments.create_upload_session.create_upload_session_post_request_body import (  # noqa: E501
            CreateUploadSessionPostRequestBody,
        )

        for file_path, file_size in large_files:
            content_type, _ = mimetypes.guess_type(file_path)
            att_item = AttachmentItem()
            att_item.attachment_type = AttachmentType.File
            att_item.name = os.path.basename(file_path)
            att_item.size = file_size
            att_item.content_type = content_type or "application/octet-stream"

            upload_body = CreateUploadSessionPostRequestBody()
            upload_body.attachment_item = att_item

            session = await msg_builder.attachments.create_upload_session.post(upload_body)
            await _upload_large_file(session.upload_url, file_path, file_size)

    return {
        "status": "attached",
        "draft_id": draft_id,
        "attachment_count": len(attachment_paths),
        "attachment_ids": attachment_ids,
    }


async def remove_draft_attachment(
    graph_client: Any,
    draft_id: str,
    attachment_id: str,
    *,
    config: Config,
) -> dict:
    """Remove a single attachment from a draft message.

    DELETE /me/messages/{draft_id}/attachments/{attachment_id}.
    Only useful on drafts — sent messages are immutable.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_remove_draft_attachment")
    draft_id = validate_graph_id(draft_id)
    attachment_id = validate_graph_id(attachment_id)

    await (
        graph_client.me.messages.by_message_id(draft_id)
        .attachments.by_attachment_id(attachment_id)
        .delete()
    )

    return {
        "status": "removed",
        "draft_id": draft_id,
        "attachment_id": attachment_id,
    }
