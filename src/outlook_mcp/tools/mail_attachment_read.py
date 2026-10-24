"""Read document attachments without requiring access to the server's filesystem."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any

from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.tools import _document_worker
from outlook_mcp.tools.mail_attachments import fetch_attachment
from outlook_mcp.validation import sanitize_output

MAX_BYTES = 10 * 2**20
WORKER = Path(_document_worker.__file__)


async def extract(data: bytes, kind: str) -> dict:
    # Windows needs SystemRoot to load system DLLs. No mailbox credentials or
    # Python path overrides are passed to the parser; -I ignores user site code.
    env = {key: value for key, value in os.environ.items() if key.lower() == "systemroot"}
    process = await asyncio.create_subprocess_exec(
        sys.executable,
        "-I",
        "-X",
        "utf8",
        str(WORKER),
        kind,
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.DEVNULL,
        env=env,
    )
    try:
        stdout, _ = await asyncio.wait_for(process.communicate(data), timeout=30)
    except BaseException:
        if process.returncode is None:
            process.kill()
        await process.communicate()
        raise
    if process.returncode:
        raise OutlookMCPError(
            "attachment_read_failed",
            "The document parser exceeded its resource limit or could not finish.",
            "Download the original for local extraction.",
        )
    try:
        parsed = json.loads(stdout)
        if not isinstance(parsed, dict) or not isinstance(
            parsed.get("text", parsed.get("error")), str
        ):
            raise ValueError("Invalid parser result")
    except (ValueError, UnicodeError) as exc:
        raise OutlookMCPError(
            "attachment_read_failed",
            "The document parser returned no readable result.",
            "Download the original for local extraction.",
        ) from exc
    if "error" in parsed:
        raise OutlookMCPError("attachment_read_failed", parsed["error"])
    return parsed


async def read_attachment(
    graph_client: Any,
    message_id: str,
    attachment_id: str,
    offset: int = 0,
    max_chars: int = 20_000,
) -> dict:
    if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
        raise ValueError("offset must be a non-negative integer.")
    if (
        isinstance(max_chars, bool)
        or not isinstance(max_chars, int)
        or not 1 <= max_chars <= 50_000
    ):
        raise ValueError("max_chars must be an integer from 1 to 50000.")
    attachment, data = await fetch_attachment(graph_client, message_id, attachment_id)
    if len(data) > MAX_BYTES:
        raise ValueError("This attachment exceeds the 10 MiB reading limit. Download it instead.")
    name = sanitize_output(attachment.name or "attachment")
    mime = sanitize_output(attachment.content_type or "application/octet-stream")
    media = mime.partition(";")[0].strip().lower()
    suffix = Path(name).suffix.lower()
    if suffix == ".pdf" or media == "application/pdf":
        kind = "pdf"
    elif (
        suffix == ".xlsx"
        or media == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        kind = "xlsx"
    elif media.startswith("text/") or suffix in (".txt", ".csv", ".json", ".xml", ".md"):
        kind = "text"
    else:
        raise ValueError(
            "Readable extraction supports PDF, XLSX and UTF-8 text. Download other formats."
        )
    try:
        parsed = await extract(data, kind)
    except asyncio.TimeoutError as exc:
        raise OutlookMCPError(
            "attachment_read_failed",
            "The document took over 30 seconds to read.",
            "Download the original for local extraction.",
        ) from exc
    text = parsed.pop("text")
    if offset > len(text):
        raise ValueError(f"offset exceeds the extracted text length ({len(text)} characters).")
    end = min(offset + max_chars, len(text))
    return {
        "name": name,
        "content_type": mime,
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "trust": "untrusted_external_content",
        **parsed,
        "text": text[offset:end],
        "offset": offset,
        "next_offset": end if end < len(text) else None,
        "total_chars": len(text),
        "complete": end == len(text),
    }
