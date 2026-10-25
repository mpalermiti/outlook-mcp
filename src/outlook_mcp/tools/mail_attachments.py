"""Mail attachment tools: list, download, send / attach to drafts."""

from __future__ import annotations

import base64
import json
import mimetypes
import ntpath
import os
import re
import tempfile
from pathlib import Path
from typing import Any
from urllib.parse import quote

from mcp.types import (
    BlobResourceContents,
    ContentBlock,
    EmbeddedResource,
    ImageContent,
    TextContent,
)

from outlook_mcp.config import Config
from outlook_mcp.errors import LEAVE_SETTINGS_TO_THE_USER, MESSAGE_MAY_HAVE_BEEN_SENT
from outlook_mcp.pagination import build_request_config
from outlook_mcp.permissions import (
    CATEGORY_MAIL_DRAFTS,
    CATEGORY_MAIL_SEND,
    check_permission,
)
from outlook_mcp.throttle import call_without_replay
from outlook_mcp.tools.mail_drafts import require_draft
from outlook_mcp.validation import (
    has_control_characters,
    sanitize_output,
    validate_email,
    validate_graph_id,
    validate_inline_attachment,
)

# 3MB threshold — files above this use upload sessions. Also the cap on an inline
# attachment (`attachments`), which never takes the upload-session path.
_LARGE_FILE_THRESHOLD = 3 * 1024 * 1024
# Bound normalization too; allow twice the encoded cap for MIME whitespace.
_INLINE_BASE64_MAX = 2 * 4 * ((_LARGE_FILE_THRESHOLD + 2) // 3)
# Chunk size for upload sessions (320 KiB aligned, as required by Graph)
_UPLOAD_CHUNK_SIZE = 320 * 1024 * 10  # 3.2 MB chunks
_UPLOAD_NOT_ATTACHED = (
    "This file was not attached and nothing was sent. Try again; if it keeps failing, send the "
    "file another way."
)

# Where resolving a path can reach the network — see _network_path_outside.
_WINDOWS = os.name == "nt"

# `as_content=True` puts the attachment into the model's context rather than onto the disk,
# so what it may return is capped. Over 10 MiB it is refused and save_path is the way.
_CONTENT_MAX_BYTES = 10 * 1024 * 1024
# Some model APIs refuse a larger image outright, so above this an image is a resource.
_IMAGE_CONTENT_MAX_BYTES = 5 * 1024 * 1024
# About 25,000 tokens, the default cap some clients put on one tool result. Longer text
# is a resource too: whole, never truncated.
_TEXT_CONTENT_MAX_CHARS = 100_000
# The image types the major model APIs read as images.
_IMAGE_TYPES = frozenset({"image/png", "image/jpeg", "image/gif", "image/webp"})


def _network_path_outside(path: str, bases: tuple[str, ...]) -> bool:
    r"""True if ``path`` names a network or device location outside every base.

    Windows only, and lexical on purpose — ``ntpath`` reads the string and never
    the filesystem. ``Path.resolve()`` opens a path to canonicalise it, and for
    ``\\host\share\x`` that means connecting to ``host`` and signing in as the
    logged-in user, before the confinement check has had its chance to refuse.
    So this one class is judged by its text, and resolving stays the authority
    for everything else.

    A path that begins with two separators covers UNC (``\\host\share``,
    ``//host/share``), the long-path form (``\\?\UNC\host\share``) and the
    device namespace (``\\.\…``). The two leading characters are read
    straight off the string rather than from ``ntpath.splitdrive``: before
    Python 3.12 that function finds no drive in some of these spellings that
    pathlib then treats as one, and the check has to agree with what gets
    resolved, on every version.

    An ``attachments_dir`` that itself sits on a share is the operator's
    choice, so a path lexically inside one of ``bases`` goes through to the
    resolver. Only a base that is a network path can vouch for one — a local
    root normalises to an empty prefix, which every path starts with.
    """
    if not _starts_with_two_separators(path):
        return False
    target = ntpath.normcase(ntpath.normpath(path))
    for base in bases:
        if not _starts_with_two_separators(base):
            continue
        root = ntpath.normcase(ntpath.normpath(base)).rstrip("\\")
        if root and (target == root or target.startswith(root + "\\")):
            return False
    return True


def _starts_with_two_separators(path: str) -> bool:
    return len(path) >= 2 and path[0] in "\\/" and path[1] in "\\/"


def resolve_attachment_path(
    path: str, attachments_dir: str, *, file_root: str | None = None
) -> str:
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

    One class is refused by its text first: on Windows a network path is turned
    away before it is resolved, because there resolving is itself the harm (see
    ``_network_path_outside``). That check only ever adds a refusal — whatever
    it lets through still has to resolve inside the directory.

    An environment-narrowed directory also carries the file's frozen resolved
    ``file_root``. Recheck the base on every call, before creating anything:
    a child directory can become an escaping symlink after settings load.
    """
    if not path or not path.strip() or "\x00" in path:
        raise ValueError("Attachment path must be a non-empty path containing no null bytes.")

    configured = Path(os.path.expanduser(attachments_dir))
    if file_root is not None and not configured.resolve().is_relative_to(Path(file_root)):
        raise ValueError(
            "The resolved attachments_dir is outside the file's configured root. "
            + LEAVE_SETTINGS_TO_THE_USER
        )
    base = configured
    created = not base.exists()
    base.mkdir(parents=True, exist_ok=True)
    if created:
        # Ours to lock down. Never chmod a directory the user pointed us at.
        # POSIX-only: on Windows os.chmod sets nothing but the read-only attribute, so this
        # cannot enforce owner-only access and the path's Windows ACL governs instead — see
        # the note above `_ensure_dir` in config.py (#85).
        base.chmod(0o700)
    base = base.resolve()
    if file_root is not None and not base.is_relative_to(Path(file_root)):
        raise ValueError(
            "The resolved attachments_dir is outside the file's configured root. "
            + LEAVE_SETTINGS_TO_THE_USER
        )

    expanded = os.path.expanduser(path)
    # Before resolve(), which on Windows is what makes the connection.
    if _WINDOWS and _network_path_outside(expanded, (str(configured), str(base))):
        raise ValueError(
            f"Attachment path is a network or device location outside the permitted "
            f"directory: {path}. Attachments may only be read from or written to "
            f"{attachments_dir}."
        )
    candidate = Path(expanded)
    if not candidate.is_absolute():
        candidate = base / candidate
    # strict=False by default: download writes a file that does not exist yet.
    resolved = candidate.resolve()

    if not resolved.is_relative_to(base):
        raise ValueError(
            f"Attachment path is outside the permitted directory: {path}. "
            f"Attachments may only be read from or written to {attachments_dir}, "
            "the server's attachments_dir setting. To send a file, ask the user to "
            "put it there; to save one, pass a path inside it (a bare filename lands "
            f"there). {LEAVE_SETTINGS_TO_THE_USER}"
        )
    return str(resolved)


def _file_attachment(name: str, content: bytes, content_type: str | None) -> Any:
    """Build a FileAttachment SDK object; the type defaults to a guess from the name."""
    from msgraph.generated.models.file_attachment import FileAttachment

    att = FileAttachment()
    att.name = name
    att.content_bytes = content
    att.content_type = content_type or mimetypes.guess_type(name)[0] or "application/octet-stream"
    att.odata_type = "#microsoft.graph.fileAttachment"
    return att


def _make_inline_attachment(file_path: str) -> Any:
    """Build a FileAttachment SDK object from a local file (<=3MB path)."""
    with open(file_path, "rb") as f:
        content = f.read()
    return _file_attachment(os.path.basename(file_path), content, None)


def _require_something_to_attach(
    attachment_paths: list[str] | None, attachments: list[dict] | None, tool: str, remedy: str
) -> None:
    """Refuse a call that would attach nothing.

    Both arguments are optional, and the server ignores an unknown top-level
    argument, so a misspelt one arrives as neither. Without this the mail went
    out, or the draft reported "attached", with no file: #41's class, a call that
    succeeds and does nothing.
    """
    if not attachment_paths and not attachments:
        raise ValueError(
            f"{tool} needs attachment_paths or attachments (inline files); neither was "
            f"given, so nothing would be attached. {remedy}"
        )


def _over_the_inline_cap(where: str, name: str, size: str) -> ValueError:
    return ValueError(
        f"{where} ({name}) is {size} bytes, over the 3 MiB limit for an inline attachment "
        "(Graph's simple-attachment size). For a larger file, ask the user to put it in "
        "attachments_dir and pass its path in attachment_paths."
    )


def _inline_attachments(items: list[dict] | None) -> list[Any]:
    """FileAttachments for ``attachments`` (inline content), validated before anything is sent.

    Inline content never touches the disk, so the attachments_dir fence has nothing to
    confine here: it still governs every path. What is capped is size, at Graph's
    simple-attachment limit, because an inline file never takes the upload-session path.
    """
    built = []
    for index, item in enumerate(items or []):
        where = f"attachments[{index}]"
        encoded = item.get("content_base64") if isinstance(item, dict) else None
        if isinstance(encoded, str):
            if len(encoded) > _INLINE_BASE64_MAX:
                name = item.get("name")
                label = sanitize_output(name.strip()) if isinstance(name, str) else "no name"
                raise ValueError(
                    f"{where} ({label}) has too much base64 input, including whitespace, "
                    "for the 3 MiB inline limit. For a larger file, use attachments_dir "
                    "and pass its path in attachment_paths."
                )
            # Use decoded length, including padding, without allocating the bytes.
            # Ignore MIME whitespace as the decoder does; invalid excess padding
            # must not reduce a large input below the cap.
            compact = "".join(encoded.split())
            padding = min(2, len(compact) - len(compact.rstrip("=")))
            size = len(compact) // 4 * 3 - padding
            if size > _LARGE_FILE_THRESHOLD:
                if len(compact) % 4 or re.fullmatch(r"[A-Za-z0-9+/]*={0,2}", compact) is None:
                    raise ValueError(f"{where}: content_base64 is not valid base64.")
                name = item.get("name")
                label = sanitize_output(name.strip()) if isinstance(name, str) else "no name"
                raise _over_the_inline_cap(where, label, str(size))
        name, content_type, content = validate_inline_attachment(item, where)
        if len(content) > _LARGE_FILE_THRESHOLD:
            raise _over_the_inline_cap(where, name, str(len(content)))
        built.append(_file_attachment(name, content, content_type))
    return built


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


def download_destination(
    save_path: str | None, as_content: bool, attachments_dir: str, *, file_root: str | None = None
) -> str | None:
    """Where a download goes: a confined file, or None when it is returned as content.

    Exactly one of the two, checked before any Graph call. With ``as_content`` nothing is
    written, so no path is resolved at all and the attachments directory is not created.
    """
    if as_content:
        if save_path is not None:
            raise ValueError(
                "Pass save_path or as_content=True, not both: as_content returns the "
                "attachment and writes nothing."
            )
        return None
    if save_path is None:
        raise ValueError(
            "Pass save_path to save the attachment inside attachments_dir, or "
            "as_content=True to receive it in the result."
        )
    save_path = resolve_attachment_path(save_path, attachments_dir, file_root=file_root)
    return _validated_download_target(save_path, attachments_dir)


def _media_type(content_type: str, name: str) -> tuple[str, str | None]:
    """(lower-case media type, charset), guessed from the name when the type is generic."""
    media, _, params = content_type.partition(";")
    media = media.strip().lower()
    charset = None
    for param in params.split(";"):
        key, _, value = param.partition("=")
        if key.strip().lower() == "charset":
            charset = value.strip().strip('"') or None
    if media in ("", "application/octet-stream"):
        media = (mimetypes.guess_type(name)[0] or "application/octet-stream").lower()
    return media, charset


def _readable_text(
    media: str, charset: str | None, content: bytes, *, for_extraction: bool = False
) -> str | None:
    """Decode supported text, returning None when its declared codec cannot decode it.

    Ordinary byte-content delivery leaves decoded text unaltered; control-bearing text
    goes out whole as a resource instead. Extraction normalizes lines, strips controls
    and replaces unpaired surrogates before paging.
    """
    if not (
        media.startswith("text/")
        or media in ("application/json", "application/xml")
        or media.endswith(("+json", "+xml"))
    ):
        return None
    try:
        text = content.decode(charset or "utf-8-sig")
    except (LookupError, ValueError):
        return None
    if for_extraction:
        # A declared codec can decode otherwise valid bytes to lone surrogates
        # (UTF-7 and unicode_escape both can). Match the parser worker's
        # replacement before the MCP reply is encoded as UTF-8.
        text = text.encode("utf-8", errors="replace").decode("utf-8")
        return sanitize_output(text.replace("\r\n", "\n").replace("\r", "\n"), multiline=True)
    if len(text) > _TEXT_CONTENT_MAX_CHARS or has_control_characters(text):
        return None
    return text


# The terminal PNG chunk as a file carries it: zero-length payload plus its CRC.
# Searching for the four bytes "IEND" accepted a file with nothing in it — no IHDR,
# no pixels — and sent it to the client as an image, which is the failure this
# check exists to prevent. Allowing it anywhere in the last 16 bytes still leaves
# room for a little trailing padding.
_PNG_IEND_CHUNK = b"\x00\x00\x00\x00IEND\xaeB`\x82"


def _jpeg_complete(content: bytes) -> bool:
    """True if ``content`` reaches a scan and ends with end-of-image after it.

    Not ``endswith`` on ``FF D9``: phones write bytes after EOI on purpose.
    A Samsung camera appends a ``SEFH…SEFT`` trailer, and Google's Motion Photo —
    the default on Pixel — appends a whole MP4, so the file does not end at the
    image. Decoders stop at EOI either way, so those files were never the problem,
    and refusing them would make a phone photo unviewable.

    So walk the segments to the first start-of-scan and look for EOI after it.
    Scan data byte-stuffs ``0xFF``, so ``FF D9`` inside the scan can only be the end
    of the image, and an EXIF thumbnail's own EOI is skipped along with its APP1
    segment.
    """
    i = 2
    while i + 4 <= len(content) and content[i] == 0xFF:
        marker = content[i + 1]
        if marker == 0xFF:  # fill byte
            i += 1
            continue
        length = int.from_bytes(content[i + 2 : i + 4], "big")
        if marker == 0xDA:  # start of scan
            return b"\xff\xd9" in content[i + 2 + length :]
        i += 2 + length
    return False


def _image_type(content: bytes) -> str | None:
    """The image type ``content`` opens as, from its start and its end.

    A signature is a handful of bytes, and a sender can stop right after it: a PNG
    truncated before its IEND cannot be opened by the client, and a client that
    rejects one unreadable image can reject the whole tool result. So each type has
    to finish the way it starts — the terminal chunk for PNG and GIF, end-of-image
    for JPEG, and a RIFF length that fits the bytes actually present for WEBP.
    Anything else is returned as a resource, which is always safe.
    """
    if content.startswith(b"\x89PNG\r\n\x1a\n") and _PNG_IEND_CHUNK in content[-16:]:
        return "image/png"
    if content.startswith(b"\xff\xd8\xff") and _jpeg_complete(content):
        return "image/jpeg"
    if content[:6] in (b"GIF87a", b"GIF89a") and content.endswith(b";"):
        return "image/gif"
    if content[:4] == b"RIFF" and content[8:12] == b"WEBP":
        declared = int.from_bytes(content[4:8], "little")
        if declared + 8 <= len(content):
            return "image/webp"
    return None


def attachment_content(name: str, content_type: str, content: bytes) -> list[ContentBlock]:
    """An attachment as MCP content (``as_content=True``), never written to disk.

    A JSON block with ``name``, ``size`` (the bytes returned) and ``content_type`` comes
    first, then the attachment: an image a model can view as ``ImageContent``, text as text,
    and anything else, PDFs included, as an ``EmbeddedResource`` holding a base64 blob and
    its media type. Over ``_CONTENT_MAX_BYTES`` it is refused. Shared by the mail and To Do
    downloads; ``name`` and ``content_type`` arrive sanitised.
    """
    if len(content) > _CONTENT_MAX_BYTES:
        raise ValueError(
            f"Attachment {name!r} is {len(content)} bytes, over the 10 MiB limit for "
            "as_content. Save it with save_path instead (it lands in attachments_dir)."
        )
    meta = {"name": name, "size": len(content), "content_type": content_type}
    blocks: list[ContentBlock] = [
        TextContent(type="text", text=json.dumps(meta, ensure_ascii=False))
    ]
    media, charset = _media_type(content_type, name)
    encoded = base64.b64encode(content).decode()

    # The name and type are the sender's; the bytes decide. Some model APIs reject a
    # whole result whose image does not decode, so non-image bytes fall through.
    image = _image_type(content) if media in _IMAGE_TYPES else None
    if image and len(content) <= _IMAGE_CONTENT_MAX_BYTES:
        blocks.append(ImageContent(type="image", data=encoded, mime_type=image))
        return blocks
    text = _readable_text(media, charset, content)
    if text is not None:
        blocks.append(TextContent(type="text", text=text))
        return blocks
    # The URI names the attachment and is never a path: the last segment of the name, quoted.
    file_name = name.replace("\\", "/").rsplit("/", 1)[-1] or "attachment"
    resource = BlobResourceContents(
        uri=f"attachment:///{quote(file_name, safe='')}", mime_type=media, blob=encoded
    )
    blocks.append(EmbeddedResource(type="resource", resource=resource))
    return blocks


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


def write_download(save_path: str, content: bytes) -> None:
    """Land downloaded bytes at ``save_path`` atomically, readable by the owner only.

    All bytes are in hand before this is called, and they go to a temp file in
    the target's own directory (``mkstemp``: created 0600, exclusively) that
    ``os.replace`` then moves into place. So a failed write never truncates a
    file already under that name, and a symlink planted at the target after the
    path check is replaced rather than written through — the bytes cannot be
    redirected to whatever it pointed at. Shared by the mail and To Do downloads.
    """
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


async def download_attachment(
    graph_client: Any,
    message_id: str,
    attachment_id: str,
    save_path: str | None = None,
    *,
    config: Config,
    as_content: bool = False,
) -> dict | list[ContentBlock]:
    """Download an attachment.

    GET /me/messages/{id}/attachments/{att_id}
    Writes the file bytes to save_path and returns the path, or with ``as_content``
    returns them as MCP content and writes nothing.
    """
    message_id = validate_graph_id(message_id)
    attachment_id = validate_graph_id(attachment_id)
    save_path = download_destination(
        save_path, as_content, config.attachments_dir, file_root=config._attachments_file_root
    )

    attachment, content = await fetch_attachment(graph_client, message_id, attachment_id)
    name = sanitize_output(attachment.name or "")
    content_type = sanitize_output(attachment.content_type or "")
    if save_path is None:
        return attachment_content(name, content_type, content)

    write_download(save_path, content)
    return {
        "saved_to": save_path,
        "name": name,
        "size": attachment.size,
        "content_type": content_type,
    }


async def fetch_attachment(
    graph_client: Any, message_id: str, attachment_id: str, *, max_bytes: int | None = None
) -> tuple[Any, bytes]:
    """The SDK attachment and its already-decoded bytes, shared by download and reading."""
    message_id = validate_graph_id(message_id)
    attachment_id = validate_graph_id(attachment_id)
    request = graph_client.me.messages.by_message_id(message_id).attachments.by_attachment_id(
        attachment_id
    )
    if max_bytes is not None:
        from msgraph.generated.users.item.messages.item.attachments.item import (
            attachment_item_request_builder as attachment_builder,
        )

        builder = attachment_builder.AttachmentItemRequestBuilder

        metadata = await request.get(
            request_configuration=build_request_config(
                builder.AttachmentItemRequestBuilderGetQueryParameters,
                {"$select": ["id", "name", "contentType", "size"]},
            )
        )
        size = getattr(metadata, "size", None)
        if type(size) is not int or size < 0:
            raise ValueError(
                "The attachment size is unavailable; download it instead of reading it."
            )
        if size > max_bytes:
            raise ValueError(
                "This attachment exceeds the 10 MiB reading limit. Download it instead."
            )
    attachment = await request.get()
    # The msgraph SDK (Kiota) already base64-decodes contentBytes into raw bytes
    # during deserialization, so content_bytes is the raw file content. Decoding
    # again corrupts binary files / raises UnicodeDecodeError (issue #25).
    # An attached email or a link to a cloud file is a different model with no
    # content_bytes at all; a file attachment can come back with None. Either
    # way there is nothing to save, and nothing is touched.
    content = getattr(attachment, "content_bytes", None)
    if content is None:
        kind = getattr(attachment, "odata_type", None)
        kind = f" ({kind})" if isinstance(kind, str) else ""
        raise ValueError(
            f"Attachment {attachment_id} has no file content to read or download{kind} — it may "
            "be an attached email or a link to a cloud file. Nothing was written."
        )

    return attachment, content


async def _upload_large_file(
    upload_url: str,
    file_path: str,
    file_size: int,
) -> None:
    """Upload a large file in chunks via an upload session.

    Uses httpx to PUT chunks to the upload URL provided by Graph. Every chunk's reply is
    checked (#125): a refused chunk raises ``httpx.HTTPStatusError``, which
    ``_wrap_tool_errors`` reports as Graph's own error (the upload URL, which carries its own
    authorization, stays out of it), and an upload whose last reply is not a completed one
    (200 or 201), or a file that comes up short while it is read, raises ``GraphAPIError``.
    Either way the caller stops before it sends the draft or reports the attachment as added.
    """
    import httpx

    from outlook_mcp.errors import GraphAPIError

    status = None
    async with httpx.AsyncClient() as client:
        offset = 0
        with open(file_path, "rb") as f:
            while offset < file_size:
                chunk = f.read(_UPLOAD_CHUNK_SIZE)
                if not chunk:
                    # The file shrank after its size was read. An empty PUT would carry a
                    # backwards Content-Range, and a 2xx reply to it would never move ``offset``.
                    raise GraphAPIError(
                        0,
                        "IncompleteUpload",
                        f"The file changed while it was being uploaded: it ended after "
                        f"{offset:,} of the {file_size:,} bytes Graph was told to expect.",
                        action=_UPLOAD_NOT_ATTACHED,
                    )
                chunk_size = len(chunk)
                end = offset + chunk_size - 1
                headers = {
                    "Content-Range": f"bytes {offset}-{end}/{file_size}",
                    "Content-Length": str(chunk_size),
                }
                response = await client.put(upload_url, content=chunk, headers=headers)
                response.raise_for_status()
                status = response.status_code
                offset += chunk_size
    if status not in (200, 201):
        raise GraphAPIError(
            status or 0,
            "IncompleteUpload",
            f"Graph did not report the attachment upload as complete (last reply: {status}).",
            action=_UPLOAD_NOT_ATTACHED,
        )


def _draft_left_behind(exc: Exception, draft_id: str) -> Exception:
    """``exc`` from a failed upload in ``send_with_attachments``, naming the draft it left.

    The draft (recipients, body, any files that did upload) stays in Drafts, and without its
    id an agent's natural retry makes a second one. Graph's refusals (SDK or raw httpx),
    ``IncompleteUpload`` and an interrupted connection come back as ``GraphAPIError`` with the
    draft in the action; anything else is a crash and is returned unchanged.
    """
    import httpx

    from outlook_mcp.errors import GraphAPIError, wrap_graph_error, wrap_http_status_error

    if isinstance(exc, GraphAPIError):
        error = exc
    elif isinstance(exc, httpx.HTTPStatusError):
        error = wrap_http_status_error(exc)
    elif isinstance(exc, httpx.TransportError):
        # Type only: the exception's text can name the upload URL.
        error = GraphAPIError(
            0, "UploadInterrupted", f"The upload did not finish ({type(exc).__name__})."
        )
    else:
        try:
            error = wrap_graph_error(exc)
        except TypeError:
            return exc
    action = (
        f"Nothing was sent. The message is saved in Drafts (id {draft_id}); delete it, or "
        "attach the file there with outlook_attach_to_draft."
    )
    if error.action and error.error_code != "IncompleteUpload":
        action = f"{action} {error.action}"
    return GraphAPIError(error.status_code, error.error_code, error.message, action=action)


async def send_with_attachments(
    graph_client: Any,
    to: list[str],
    subject: str,
    body: str,
    attachment_paths: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    reply_to: list[str] | None = None,
    *,
    config: Config,
    attachments: list[dict] | None = None,
) -> dict:
    """Send a message with file attachments.

    For files under 3MB: inline as base64 FileAttachment.
    For files over 3MB: create draft, use createUploadSession + chunked upload, then send.
    ``attachments`` are inline files (``{name, content_base64, content_type?}``, ≤3MB each),
    which go out like a small file and are never read from or written to disk.
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_with_attachments")
    _require_something_to_attach(
        attachment_paths,
        attachments,
        "outlook_send_with_attachments",
        "For a message with no attachment, use outlook_send_message.",
    )

    # Validate emails
    validated_to = [validate_email(e) for e in to]
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []
    validated_reply_to = [validate_email(e) for e in reply_to] if reply_to else []

    # Confine every path before touching the filesystem, then validate existence.
    attachment_paths = [
        resolve_attachment_path(
            path, config.attachments_dir, file_root=config._attachments_file_root
        )
        for path in attachment_paths or []
    ]
    for path in attachment_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")
    inline = _inline_attachments(attachments)

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
        msg.attachments = [_make_inline_attachment(p) for p in small_files] + inline

        request_body = SendMailPostRequestBody()
        request_body.message = msg
        request_body.save_to_sent_items = True

        await call_without_replay(
            lambda once: graph_client.me.send_mail.post(request_body, request_configuration=once),
            outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT,
        )
    else:
        # Has large files — create draft, attach via upload sessions, then send
        from msgraph.generated.models.attachment_item import AttachmentItem
        from msgraph.generated.models.attachment_type import AttachmentType
        from msgraph.generated.users.item.messages.item.attachments.create_upload_session.create_upload_session_post_request_body import (  # noqa: E501
            CreateUploadSessionPostRequestBody,
        )

        msg = _build_message()
        # Attach small files inline on the draft
        msg.attachments = [_make_inline_attachment(p) for p in small_files] + inline

        # Create draft message
        draft = await graph_client.me.messages.post(msg)

        # Upload each large file. A failure from here on leaves the draft in Drafts, so the
        # error names it (see _draft_left_behind).
        try:
            for file_path, file_size in large_files:
                content_type, _ = mimetypes.guess_type(file_path)
                att_item = AttachmentItem()
                att_item.attachment_type = AttachmentType.File
                att_item.name = os.path.basename(file_path)
                att_item.size = file_size
                att_item.content_type = content_type or "application/octet-stream"

                upload_body = CreateUploadSessionPostRequestBody()
                upload_body.attachment_item = att_item

                session = await graph_client.me.messages.by_message_id(
                    draft.id
                ).attachments.create_upload_session.post(upload_body)

                await _upload_large_file(session.upload_url, file_path, file_size)
        except Exception as exc:
            named = _draft_left_behind(exc, draft.id)
            if named is exc:
                raise
            raise named from exc

        # Send the draft
        send_builder = graph_client.me.messages.by_message_id(draft.id).send
        await call_without_replay(
            lambda once: send_builder.post(request_configuration=once),
            outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT,
        )

    return {
        "status": "sent",
        "attachment_count": len(attachment_paths) + len(inline),
    }


async def attach_to_draft(
    graph_client: Any,
    draft_id: str,
    attachment_paths: list[str] | None = None,
    *,
    config: Config,
    attachments: list[dict] | None = None,
) -> dict:
    """Add one or more attachments to an existing draft message.

    For files under 3MB: POST a FileAttachment directly.
    For files over 3MB: createUploadSession + chunked upload.
    ``attachments`` are inline files (``{name, content_base64, content_type?}``, ≤3MB
    each), POSTed like a small file and never read from or written to disk.

    Returns the new attachment IDs so callers can reference or
    remove individual attachments later.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_attach_to_draft")
    draft_id = validate_graph_id(draft_id)
    _require_something_to_attach(
        attachment_paths,
        attachments,
        "outlook_attach_to_draft",
        "Check the argument names: attachment_paths takes file paths, attachments takes "
        "{name, content_base64} items.",
    )

    # Confine every path before touching the filesystem, then validate existence.
    attachment_paths = [
        resolve_attachment_path(
            path, config.attachments_dir, file_root=config._attachments_file_root
        )
        for path in attachment_paths or []
    ]
    for path in attachment_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")
    inline = _inline_attachments(attachments)

    # Partition files into small (inline) and large (upload session)
    small_files: list[str] = []
    large_files: list[tuple[str, int]] = []
    for path in attachment_paths:
        file_size = os.path.getsize(path)
        if file_size > _LARGE_FILE_THRESHOLD:
            large_files.append((path, file_size))
        else:
            small_files.append(path)

    if attachment_paths or inline:
        # After the paths are confined and found, before anything is uploaded.
        await require_draft(graph_client, draft_id, "outlook_attach_to_draft")

    attachment_ids: list[str] = []
    msg_builder = graph_client.me.messages.by_message_id(draft_id)

    # Small files and inline content — POST each as an inline FileAttachment
    for att in [_make_inline_attachment(p) for p in small_files] + inline:
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
        "attachment_count": len(attachment_paths) + len(inline),
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

    await require_draft(graph_client, draft_id, "outlook_remove_draft_attachment")
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
