"""Read document attachments without requiring access to the server's filesystem."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import signal
import sys
from pathlib import Path
from typing import Any

from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.tools import _document_worker
from outlook_mcp.tools.mail_attachments import _media_type, _readable_text, fetch_attachment
from outlook_mcp.validation import sanitize_output

MAX_BYTES = 10 * 2**20
MAX_SHEETS = 100
WORKER = Path(_document_worker.__file__)


MEMORY_LIMIT = 256 * 2**20
WALL_SECONDS = 30
MAX_OUTPUT_BYTES = 8 * 2**20
_WORKERS = asyncio.Semaphore(1)


def _failure(message: str) -> OutlookMCPError:
    return OutlookMCPError(
        "attachment_read_failed", message, "Download the original for local extraction."
    )


def _parser_result(stdout: bytes) -> dict:
    try:
        parsed = json.loads(stdout)
        if not isinstance(parsed, dict):
            raise ValueError
        if "error" in parsed:
            if not isinstance(parsed["error"], str):
                raise ValueError
            raise _failure(sanitize_output(parsed["error"], multiline=True)[:300])
        text = parsed.get("text")
        if not isinstance(text, str) or len(text) > 1_000_000:
            raise ValueError
        result = {
            "text": sanitize_output(text.replace("\r\n", "\n").replace("\r", "\n"), multiline=True)
        }
        if "pages" in parsed:
            pages = parsed["pages"]
            if type(pages) is not int or not 0 <= pages <= 200:
                raise ValueError
            result["pages"] = pages
        if "sheets" in parsed:
            sheets = parsed["sheets"]
            if not isinstance(sheets, list) or len(sheets) > MAX_SHEETS:
                raise ValueError
            result["sheets"] = []
            for sheet in sheets:
                if not isinstance(sheet, dict) or not all(
                    isinstance(sheet.get(k), str) for k in ("name", "state")
                ):
                    raise ValueError
                if len(sheet["name"]) > 31 or sheet["state"] not in (
                    "visible",
                    "hidden",
                    "veryHidden",
                ):
                    raise ValueError
                result["sheets"].append({k: sanitize_output(sheet[k]) for k in ("name", "state")})
        warnings = parsed.get("warnings", [])
        if (
            not isinstance(warnings, list)
            or len(warnings) > 200
            or not all(isinstance(w, str) for w in warnings)
        ):
            raise ValueError
        result["warnings"] = [sanitize_output(w, multiline=True)[:300] for w in warnings]
        return result
    except (ValueError, UnicodeError, RecursionError) as exc:
        raise _failure("The document parser returned no readable result.") from exc


async def extract(data: bytes, kind: str) -> dict:
    try:
        import psutil
    except ImportError as exc:
        raise _failure(
            "PDF/XLSX reading needs outlook-graph-mcp[documents]. "
            "Ask the operator to install it and restart the server."
        ) from exc
    async with _WORKERS:
        # Clean environment for crash and resource containment, not a sandbox:
        # the worker retains this user's filesystem and network privileges.
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
            **({"start_new_session": True} if os.name == "posix" else {}),
        )
        child = psutil.Process(process.pid)

        async def communicate() -> bytes:
            async def feed() -> None:
                try:
                    process.stdin.write(data)
                    await process.stdin.drain()
                except (BrokenPipeError, ConnectionResetError):
                    pass
                finally:
                    process.stdin.close()

            feeding = asyncio.create_task(feed())
            try:
                chunks, size = [], 0
                while chunk := await process.stdout.read(65536):
                    size += len(chunk)
                    if size > MAX_OUTPUT_BYTES:
                        raise _failure("The document parser exceeded its output limit.")
                    chunks.append(chunk)
                await feeding
                await process.wait()
                if process.returncode:
                    raise _failure(
                        "The document parser exceeded its resource limit or could not finish."
                    )
                return b"".join(chunks)
            finally:
                feeding.cancel()
                await asyncio.gather(feeding, return_exceptions=True)

        async def watch_memory() -> None:
            while True:
                try:
                    rss = child.memory_info().rss
                    for descendant in child.children(recursive=True):
                        try:
                            rss += descendant.memory_info().rss
                        except psutil.NoSuchProcess:
                            pass
                    if rss > MEMORY_LIMIT:
                        raise _failure("The document parser exceeded its memory limit.")
                except psutil.NoSuchProcess:
                    pass
                except psutil.AccessDenied as exc:
                    raise _failure(
                        "The document parser's memory usage could not be monitored."
                    ) from exc
                await asyncio.sleep(0.01)

        exchange = asyncio.create_task(communicate())
        watchdog = asyncio.create_task(watch_memory())
        try:

            async def run() -> bytes:
                done, _ = await asyncio.wait(
                    (exchange, watchdog), return_when=asyncio.FIRST_COMPLETED
                )
                if watchdog in done:
                    await watchdog
                return await exchange

            return _parser_result(await asyncio.wait_for(run(), timeout=WALL_SECONDS))
        finally:
            # Kill the group even if the immediate child already exited: a
            # descendant may still hold stdout. Never wait indefinitely on it.
            if os.name == "posix":
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except (ProcessLookupError, PermissionError):
                    # Nothing left to kill. macOS says EPERM, not ESRCH, while
                    # the exited worker awaits reaping; a live member is killed.
                    pass
            else:
                try:
                    for descendant in child.children(recursive=True):
                        try:
                            descendant.kill()
                        except psutil.NoSuchProcess:
                            pass
                except psutil.NoSuchProcess:
                    pass
                if process.returncode is None:
                    process.kill()
            exchange.cancel()
            watchdog.cancel()
            await asyncio.gather(exchange, watchdog, return_exceptions=True)
            try:
                await asyncio.wait_for(process.wait(), timeout=2)
            except asyncio.TimeoutError:
                pass


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
    attachment, data = await fetch_attachment(
        graph_client, message_id, attachment_id, max_bytes=MAX_BYTES
    )
    if len(data) > MAX_BYTES:
        raise ValueError("This attachment exceeds the 10 MiB reading limit. Download it instead.")
    name = sanitize_output(attachment.name or "attachment")
    mime = sanitize_output(attachment.content_type or "application/octet-stream")
    media, charset = _media_type(attachment.content_type or "", name)
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
            "Readable extraction supports PDF, XLSX and text with its declared charset. "
            "Download other formats."
        )
    try:
        if kind == "text":
            text = _readable_text(
                "text/plain" if suffix in (".txt", ".csv", ".json", ".xml", ".md") else media,
                charset,
                data,
                for_extraction=True,
            )
            if text is None:
                raise _failure(
                    "This text could not be decoded with its declared charset (UTF-8 by default)."
                )
            if len(text) > 1_000_000:
                raise _failure("The document exceeds one million extracted characters.")
            parsed = {"text": text, "warnings": []}
        else:
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
    # Sheet names are sender-controlled content too. Include their bounded
    # summary only on the first page when it fits alongside progress in text.
    sheets = parsed.pop("sheets", None)
    budget = max_chars
    if offset == 0 and sheets is not None:
        size = len(json.dumps(sheets, ensure_ascii=False))
        if size <= budget - bool(text):
            parsed["sheets"] = sheets
            budget -= size
    # A caller can resume in the middle of a quoted PDF line, even immediately
    # after '| '. Re-quote the continuation inside the same output budget.
    # Tiny chunks cannot contain a complete '[Page N]' marker and still advance.
    prefix = ""
    if kind == "pdf" and offset and offset < len(text) and text[offset - 1] != "\n" and budget > 2:
        prefix = "| "
    end = min(offset + budget - len(prefix), len(text))
    return {
        **parsed,
        "name": name,
        "content_type": mime,
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "trust": "untrusted_external_content",
        "text": prefix + text[offset:end],
        "continuation_prefix": prefix,
        "offset": offset,
        "next_offset": end if end < len(text) else None,
        "total_chars": len(text),
        "complete": end == len(text),
    }
