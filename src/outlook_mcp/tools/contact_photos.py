"""Contact photos: set or replace one from a file in ``attachments_dir`` (#121).

An opt-in group (``contact_photos``), off by default so the default tool list
does not grow. Set only: there is no read tool.

The file is read through ``resolve_attachment_path`` like every other host path,
because the path comes from the model and the model reads email. The bytes, not
the file name, decide the type: JPEG, PNG or GIF, detected by
``mail_attachments._image_type`` (signature plus a truncation check, not a full
decode). Graph decodes the image itself and refuses one it cannot read with
``400 InputNotAnImage``, which reaches the agent through the tool error wrapper.

The PUT goes through the SDK's own ``photo.content.put``, which labels the body
``application/octet-stream``. Graph accepts that: on a personal account on
2026-11-14 a PNG, a JPEG and a GIF each stored a photo sent that way, and came
back as ``image/jpeg``. A repeat of the same PUT lands on the same state, so the
SDK's normal retry stays on.
"""

from __future__ import annotations

import os
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.errors import ToolInputError
from outlook_mcp.permissions import CATEGORY_CONTACTS_WRITE, check_permission
from outlook_mcp.tools.mail_attachments import _image_type, resolve_attachment_path
from outlook_mcp.validation import validate_graph_id

# This server's limit, not Graph's. Graph's own limit is the mailbox store's item
# size, checked after it re-encodes the image: on one personal account it took a
# 26.2 MB PNG and refused a 47.2 MB one with 400 ErrorMessageSizeExceeded. 20 MiB
# sits under the largest upload seen accepted, covers any phone photo, and is the
# To Do attachment ceiling, read the same way (the read is the gate).
_MAX_PHOTO_BYTES = 20 * 1024 * 1024

_PHOTO_TYPES = {"image/jpeg", "image/png", "image/gif"}


def _read_photo(file_path: str, config: Config) -> bytes:
    """The file's bytes, confined to ``attachments_dir`` and checked, or a refusal.

    ``OSError`` from the resolver (which may create the directory) or from the
    read becomes text the model can see: a Windows lock, a permission, a file
    that went away after ``isfile``. Confinement refusals are ``ValueError`` and
    pass through unchanged.
    """
    try:
        path = resolve_attachment_path(
            file_path, config.attachments_dir, file_root=config._attachments_file_root
        )
        if not os.path.isfile(path):
            raise ToolInputError(
                f"Photo file not found: {path}. Pass a file inside {config.attachments_dir}; "
                "no photo was set."
            )
        with open(path, "rb") as f:
            content = f.read(_MAX_PHOTO_BYTES + 1)
    except OSError as exc:
        raise ToolInputError(
            f"Could not open {file_path} in attachments_dir: {exc.strerror or exc}; "
            "no photo was set."
        ) from exc

    name = os.path.basename(path)
    if not content:
        raise ToolInputError(f"{name} is empty, so no photo was set.")
    if len(content) > _MAX_PHOTO_BYTES:
        raise ToolInputError(
            f"{name} is over this server's 20 MiB limit for a contact photo (read "
            f"{len(content)} bytes of it), so nothing was sent. Ask the user for a smaller copy."
        )
    if _image_type(content) not in _PHOTO_TYPES:
        raise ToolInputError(
            f"{name} is not a JPEG, PNG or GIF file, so no photo was set. "
            "Pass a JPEG, PNG or GIF file."
        )
    return content


async def set_contact_photo(
    graph_client: Any,
    contact_id: str,
    file_path: str,
    *,
    config: Config,
) -> dict:
    """Set or replace a contact's photo from a JPEG, PNG or GIF file in ``attachments_dir``."""
    check_permission(config, CATEGORY_CONTACTS_WRITE, "outlook_set_contact_photo")
    contact_id = validate_graph_id(contact_id)
    content = _read_photo(file_path, config)

    await graph_client.me.contacts.by_contact_id(contact_id).photo.content.put(content)

    return {"status": "updated", "contact_id": contact_id, "uploaded_bytes": len(content)}
