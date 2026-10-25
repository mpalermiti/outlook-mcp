"""Independent guards for each document-worker trust and cleanup boundary.

Fixtures are synthetic; resource tests have their own finite time/RAM ceilings.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import subprocess
import sys
import time
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest

from outlook_mcp.tools import mail_attachment_read as reader
from tests.test_readable_attachments import ARGS, graph, pdf, xlsx


def test_worker_known_keys_filter_is_independent_of_parent_field_precedence():
    result = reader._parser_result(
        json.dumps(
            {"text": "ok", "trust": "forged", "private_extra": "forged", "offset": 3}
        ).encode()
    )
    assert result == {"text": "ok", "warnings": []}


async def test_parent_fields_are_authoritative_even_if_worker_filter_regresses(monkeypatch):
    data = b"synthetic content"
    forged = {
        "text": "ok",
        "warnings": [],
        "name": "forged",
        "content_type": "forged",
        "bytes": 99,
        "sha256": "forged",
        "trust": "trusted",
        "offset": 99,
        "next_offset": 99,
        "continuation_prefix": "forged",
        "total_chars": 99,
        "complete": False,
    }
    # Inject after validation to test the independent last-writer boundary.
    # Keeping _parser_result's whitelist must not mask a regression here.
    monkeypatch.setattr(reader, "extract", AsyncMock(return_value=forged))
    sdk, _ = graph(data, "original.pdf", "application/pdf")
    result = await reader.read_attachment(sdk, **ARGS)
    assert {key: result[key] for key in forged if key not in ("text", "warnings")} == {
        "name": "original.pdf",
        "content_type": "application/pdf",
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "trust": "untrusted_external_content",
        "offset": 0,
        "next_offset": None,
        "continuation_prefix": "",
        "total_chars": 2,
        "complete": True,
    }


async def test_worker_exit_wait_is_bounded_even_after_group_kill(monkeypatch):
    import psutil

    waits = 0
    never = asyncio.Event()

    async def wait():
        nonlocal waits
        waits += 1
        if waits == 1:  # communicate completed, but final process reaping stalls
            return 0
        await never.wait()

    process = SimpleNamespace(
        pid=424242,
        returncode=0,
        wait=wait,
        stdin=SimpleNamespace(write=lambda _data: None, drain=AsyncMock(), close=lambda: None),
        stdout=SimpleNamespace(read=AsyncMock(side_effect=[b'{"text":"ok"}', b""])),
    )
    child = SimpleNamespace(memory_info=lambda: SimpleNamespace(rss=0), children=lambda **_: [])
    monkeypatch.setattr(asyncio, "create_subprocess_exec", AsyncMock(return_value=process))
    monkeypatch.setattr(psutil, "Process", lambda _pid: child)
    if os.name == "posix":
        monkeypatch.setattr(os, "killpg", lambda *_: None)
    started = time.monotonic()
    result = await asyncio.wait_for(reader.extract(b"abc", "pdf"), timeout=3.5)
    assert result["text"] == "ok" and waits == 2
    assert 1.5 <= time.monotonic() - started < 3.5


@pytest.mark.parametrize("kind", ["pdf", "xlsx"])
async def test_document_worker_text_is_cleaned_for_each_binary_format(monkeypatch, tmp_path, kind):
    worker = tmp_path / "controls.py"
    hostile = "before\x1b[2Jafter\x00\x7f\r\nsecond\rthird\tend"
    worker.write_text(
        "import json,sys\nsys.stdin.buffer.read()\n"
        f"assert sys.argv[1] == {kind!r}\n"
        f"print(json.dumps({{'text':{hostile!r}}}))\n"
    )
    monkeypatch.setattr(reader, "WORKER", worker)
    sdk, _ = graph(b"synthetic", f"controls.{kind}", "application/octet-stream")
    result = await reader.read_attachment(sdk, **ARGS)
    assert result["text"] == "beforeafter\nsecond\nthird end"


@pytest.mark.parametrize("kind", ["pdf", "xlsx"])
async def test_real_pdf_and_xlsx_control_glyphs_are_sanitized(kind):
    if kind == "pdf":
        data = pdf("before\x7fafter")
    else:
        from io import BytesIO
        from zipfile import ZipFile

        source = xlsx()
        output = BytesIO()
        with ZipFile(BytesIO(source)) as original, ZipFile(output, "w") as changed:
            for item in original.infolist():
                content = original.read(item)
                if item.filename == "xl/worksheets/sheet1.xml":
                    content = content.replace(b"Service", b"before\x7fafter")
                changed.writestr(item, content)
        data = output.getvalue()
    sdk, _ = graph(data, f"controls.{kind}", "application/octet-stream")
    result = await reader.read_attachment(sdk, **ARGS)
    assert "beforeafter" in result["text"] and "\x7f" not in result["text"]


def test_documents_opt_in_registration_preserves_reader_error_wrapper(tmp_path):
    code = """
import asyncio
from types import SimpleNamespace
from unittest.mock import AsyncMock
from mcp.client import Client
from outlook_mcp import server

assert "outlook_read_attachment" in server.mcp._tool_manager._tools
server._get_graph_client = lambda ctx: SimpleNamespace(sdk_client=object())
server.mail_attachment_read.read_attachment = AsyncMock(
    side_effect=ValueError("reader range sentinel")
)
async def main():
    async with Client(server.mcp) as client:
        reply = await client.call_tool(
            "outlook_read_attachment",
            {"message_id":"synthetic", "attachment_id":"synthetic"},
        )
        assert reply.is_error
        assert "reader range sentinel" in reply.content[0].text, reply.content
asyncio.run(main())
"""
    environment = {
        key: value for key, value in os.environ.items() if not key.startswith("OUTLOOK_MCP_")
    }
    environment.update(
        {
            "HOME": str(tmp_path),
            "OUTLOOK_MCP_CONFIG_DIR": str(tmp_path / "config"),
            "OUTLOOK_MCP_TOOLSETS": "documents",
        }
    )
    done = subprocess.run(
        [sys.executable, "-c", code], env=environment, capture_output=True, text=True, timeout=20
    )
    assert done.returncode == 0, done.stdout + done.stderr
