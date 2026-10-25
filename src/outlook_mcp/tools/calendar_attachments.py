"""Calendar event attachments: attach files from attachments_dir, remove one.

The mail attachment code does the file handling, as it is: the paths go through
``resolve_attachment_path`` and its ``attachments_dir`` fence, a file up to 3 MB
is built by ``_make_inline_attachment``, and a larger one goes through Graph's
upload session for events and ``_upload_large_file``. Only the request builders
differ, because an event's attachments hang off ``/me/events/{id}``.

Listing an event's attachments is ``outlook_get_event``'s, in ``calendar_read``.
"""

from __future__ import annotations

import mimetypes
import os
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import (
    CATEGORY_CALENDAR_WRITE,
    check_permission,
    check_sends_mail,
    mail_send_withheld,
)
from outlook_mcp.tools.calendar_write import _has_attendees
from outlook_mcp.tools.mail_attachments import (
    _LARGE_FILE_THRESHOLD,
    _make_inline_attachment,
    _upload_large_file,
    resolve_attachment_path,
)
from outlook_mcp.validation import validate_graph_id


def _confined_files(attachment_paths: list[str], attachments_dir: str) -> list[tuple[str, int]]:
    """Each path resolved inside ``attachments_dir``, with its size; refused otherwise.

    Every path is checked before anything is sent, so one bad entry attaches
    nothing. A missing file is a ``ValueError`` rather than the
    ``FileNotFoundError`` the mail tools raise: an ``OSError`` reaches the model
    as a bare "Error executing tool", and this is the caller's mistake to fix.
    """
    if not attachment_paths:
        raise ValueError("Pass at least one file in attachment_paths.")
    files = []
    for path in attachment_paths:
        resolved = resolve_attachment_path(path, attachments_dir)
        if not os.path.isfile(resolved):
            raise ValueError(
                f"Not a file: {resolved}. Pass files inside {attachments_dir}; nothing "
                "was attached."
            )
        files.append((resolved, os.path.getsize(resolved)))
    return files


async def attach_to_event(
    graph_client: Any,
    event_id: str,
    attachment_paths: list[str],
    *,
    config: Config,
) -> dict:
    """Attach files to a calendar event.

    ``POST /me/events/{id}/attachments`` for a file up to 3 MB;
    ``/me/events/{id}/attachments/createUploadSession`` and a chunked upload for
    a larger one, as ``outlook_attach_to_draft`` does for a draft.

    Attendees are not sent the files when they are attached. Observed on an
    Exchange Online mailbox: they arrive with the organizer's next update of the
    event that changes something. So a file on an event with other people on it
    leaves the mailbox later, which is the same reason rewording such an event
    needs ``mail_send`` under a restricted policy, and attaching does too. Only
    that policy pays for the read that finds out.

    Returns the IDs of the attachments Graph reports, for
    ``outlook_remove_event_attachment``. A file over 3 MB has none yet:
    ``_upload_large_file`` checks every chunk's reply (#125) but does not take
    the attachment's ID from the last one.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_attach_to_event")
    event_id = validate_graph_id(event_id)
    files = _confined_files(attachment_paths, config.attachments_dir)

    builder = graph_client.me.events.by_event_id(event_id)
    if mail_send_withheld(config) and _has_attendees(await builder.get()):
        check_sends_mail(
            config,
            "outlook_attach_to_event",
            "attach a file to an event that has attendees",
        )

    attachment_ids: list[str] = []
    for path, size in files:
        if size <= _LARGE_FILE_THRESHOLD:
            created = await builder.attachments.post(_make_inline_attachment(path))
            if created is not None and getattr(created, "id", None):
                attachment_ids.append(created.id)
            continue

        from msgraph.generated.models.attachment_item import AttachmentItem
        from msgraph.generated.models.attachment_type import AttachmentType
        from msgraph.generated.users.item.events.item.attachments.create_upload_session.create_upload_session_post_request_body import (  # noqa: E501
            CreateUploadSessionPostRequestBody,
        )

        item = AttachmentItem()
        item.attachment_type = AttachmentType.File
        item.name = os.path.basename(path)
        item.size = size
        item.content_type = mimetypes.guess_type(path)[0] or "application/octet-stream"
        body = CreateUploadSessionPostRequestBody()
        body.attachment_item = item

        session = await builder.attachments.create_upload_session.post(body)
        await _upload_large_file(session.upload_url, path, size)

    return {
        "status": "attached",
        "event_id": event_id,
        "attachment_count": len(files),
        "attachment_ids": attachment_ids,
    }


async def remove_event_attachment(
    graph_client: Any,
    event_id: str,
    attachment_id: str,
    *,
    config: Config,
) -> dict:
    """Remove one attachment from a calendar event.

    ``DELETE /me/events/{id}/attachments/{attachment_id}``. The IDs come from
    ``outlook_attach_to_event`` or from ``outlook_get_event``'s ``attachments``.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_remove_event_attachment")
    event_id = validate_graph_id(event_id)
    attachment_id = validate_graph_id(attachment_id)

    await (
        graph_client.me.events.by_event_id(event_id)
        .attachments.by_attachment_id(attachment_id)
        .delete()
    )

    return {"status": "removed", "event_id": event_id, "attachment_id": attachment_id}
