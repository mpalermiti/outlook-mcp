"""Document contents reach an MCP client without opening a server-side saved path.

All PDFs and spreadsheets are synthetic. Run with the optional documents extra;
CI's --all-extras includes it. The ordinary byte-download tests remain unchanged.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import sys
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from mcp.client import Client

from outlook_mcp import server
from outlook_mcp.config import Config
from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.tools import mail_attachment_read
from tests.test_attachments_as_content import _graph_attachment, _mail_client

ARGS = {"message_id": "AAMkAG123=", "attachment_id": "att1"}


def pdf(*pages: str | None, password: str | None = None) -> bytes:
    pypdf = pytest.importorskip("pypdf")
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

    writer = pypdf.PdfWriter()
    for text in pages:
        page = writer.add_blank_page(width=300, height=300)
        if text is not None:
            font = DictionaryObject(
                {
                    NameObject("/Type"): NameObject("/Font"),
                    NameObject("/Subtype"): NameObject("/Type1"),
                    NameObject("/BaseFont"): NameObject("/Helvetica"),
                }
            )
            page[NameObject("/Resources")] = DictionaryObject(
                {NameObject("/Font"): DictionaryObject({NameObject("/F1"): font})}
            )
            content = DecodedStreamObject()
            content.set_data(f"BT /F1 12 Tf 10 200 Td ({text}) Tj ET".encode())
            page[NameObject("/Contents")] = content
    if password is not None:
        writer.encrypt(password)
    file = BytesIO()
    writer.write(file)
    return file.getvalue()


def xlsx() -> bytes:
    openpyxl = pytest.importorskip("openpyxl")
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Charges"
    sheet.append(["Item", "Amount", "Total"])
    sheet.append(["Service", 125.5, "=SUM(B2:B3)"])
    sheet.append(["Fee", 24.5, "=B2+B3"])
    sheet["B2"].number_format = '"EUR" 0.00'
    hidden = book.create_sheet("History")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "Deposit"
    hidden["B1"] = 50
    file = BytesIO()
    book.save(file)
    output = BytesIO()
    with ZipFile(file) as original, ZipFile(output, "w") as rewritten:
        for item in original.infolist():
            data = original.read(item)
            if item.filename == "xl/worksheets/sheet1.xml":
                for empty in (b"<v></v>", b"<v />", b"<v/>"):
                    data = data.replace(
                        b"<f>SUM(B2:B3)</f>" + empty, b"<f>SUM(B2:B3)</f><v>150</v>"
                    )
                # The declared range must not hide populated rows.
                data = data.replace(b'dimension ref="A1:C3"', b'dimension ref="A1:A1"')
            rewritten.writestr(item, data)
    return output.getvalue()


def graph(data: bytes, name: str = "note.txt", mime: str = "text/plain"):
    return _mail_client(_graph_attachment(data, name, mime))


async def test_pdf_contents_reach_mcp_with_lossless_paging(monkeypatch, tmp_path):
    data = pdf("Invoice 123", "Due 2026-11-01")
    sdk, get = graph(data, "invoice.pdf", "application/pdf")
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: SimpleNamespace(sdk_client=sdk))
    config = Config(read_only=True, attachments_dir=str(tmp_path / "no-files"))
    monkeypatch.setattr(server, "_get_config", lambda ctx: config)
    pieces, offset = [], 0
    async with Client(server.mcp) as client:
        while offset is not None:
            reply = await client.call_tool(
                "outlook_read_attachment", {**ARGS, "offset": offset, "max_chars": 17}
            )
            assert not reply.is_error
            assert all(block.type == "text" for block in reply.content)
            # Dict results are also JSON text for clients without structuredContent.
            part = json.loads(reply.content[0].text)
            pieces.append(part["text"])
            assert part["offset"] == offset
            offset = part["next_offset"]
    text = "".join(pieces)
    assert "[Page 1]" in text and "Invoice 123" in text and "Due 2026-11-01" in text
    assert part["pages"] == 2 and part["complete"] and part["warnings"] == []
    assert len(text) == part["total_chars"]
    assert part["sha256"] == hashlib.sha256(data).hexdigest() and part["bytes"] == len(data)
    assert part["trust"] == "untrusted_external_content"
    assert not (tmp_path / "no-files").exists()
    assert get.await_count > 1


async def test_xlsx_preserves_sheets_formulas_saved_results_and_number_formats():
    sdk, _ = graph(xlsx(), "charges.xlsx", "application/octet-stream")
    read = await mail_attachment_read.read_attachment(sdk, **ARGS)
    text = read["text"]
    assert 'C2: "=SUM(B2:B3)" | saved_result=150' in text
    assert 'C3: "=B2+B3" | saved_result=null' in text
    assert 'B2: 125.5 | number_format="\\"EUR\\" 0.00"' in text
    assert '[Sheet "History"; hidden]' in text and 'A1: "Deposit"' in text
    assert read["sheets"] == [
        {"name": "Charges", "state": "visible"},
        {"name": "History", "state": "hidden"},
    ]
    assert "not recalculated" in read["warnings"][0] and read["complete"]


@pytest.mark.parametrize(
    "case,expected", [("empty", "OCR"), ("locked", "password"), ("corrupt", "Could not extract")]
)
async def test_unreadable_pdf_errors_reach_mcp(monkeypatch, case, expected):
    pytest.importorskip("pypdf")
    data = {
        "empty": lambda: pdf(None),
        "locked": lambda: pdf("Secret", password="locked"),
        "corrupt": lambda: b"not a PDF",
    }[case]()
    sdk, _ = graph(data, "scan.pdf", "application/pdf")
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: SimpleNamespace(sdk_client=sdk))
    async with Client(server.mcp) as client:
        reply = await client.call_tool("outlook_read_attachment", ARGS)
    assert reply.is_error and expected in reply.content[0].text


async def test_mixed_pdf_reports_pages_without_text():
    sdk, _ = graph(pdf("Annex", None), "mixed.pdf", "application/pdf")
    read = await mail_attachment_read.read_attachment(sdk, **ARGS)
    assert "[2]" in read["warnings"][0] and "Annex" in read["text"]


@pytest.mark.parametrize(
    "arguments", [{"offset": -1}, {"offset": True}, {"max_chars": 0}, {"max_chars": 50001}]
)
async def test_ranges_are_validated_before_graph(arguments):
    sdk, get = graph(b"abc")
    with pytest.raises(ValueError):
        await mail_attachment_read.read_attachment(sdk, **ARGS, **arguments)
    get.assert_not_awaited()


async def test_missing_file_content_and_unsupported_types():
    sdk, get = graph(b"abc", "document.docx", "application/msword")
    with pytest.raises(ValueError, match="other formats"):
        await mail_attachment_read.read_attachment(sdk, **ARGS)
    get.return_value = SimpleNamespace(
        content_bytes=None, odata_type="#microsoft.graph.itemAttachment"
    )
    with pytest.raises(ValueError, match="no file content"):
        await mail_attachment_read.read_attachment(sdk, **ARGS)


async def test_archive_expansion_is_bounded():
    pytest.importorskip("openpyxl")
    file = BytesIO()
    with ZipFile(file, "w", ZIP_DEFLATED) as archive:
        archive.writestr("big.xml", b"x" * (50 * 2**20 + 1))
    sdk, _ = graph(file.getvalue(), "large.xlsx", "application/octet-stream")
    with pytest.raises(OutlookMCPError, match="expands beyond"):
        await mail_attachment_read.read_attachment(sdk, **ARGS)


async def test_text_is_preserved_and_offsets_are_checked():
    data = "€125.50\n".encode()
    sdk, _ = graph(data)
    read = await mail_attachment_read.read_attachment(sdk, **ARGS)
    assert read["text"] == "€125.50\n" and read["bytes"] == len(data)
    with pytest.raises(ValueError, match="exceeds the extracted text"):
        await mail_attachment_read.read_attachment(sdk, **ARGS, offset=999)
    end = await mail_attachment_read.read_attachment(sdk, **ARGS, offset=read["total_chars"])
    assert end["text"] == "" and end["next_offset"] is None


async def test_too_large_file_and_extracted_text_are_explicit_errors():
    sdk, _ = graph(b"x" * (mail_attachment_read.MAX_BYTES + 1))
    with pytest.raises(ValueError, match="10 MiB"):
        await mail_attachment_read.read_attachment(sdk, **ARGS)
    sdk, _ = graph(b"x" * 1_000_001)
    with pytest.raises(OutlookMCPError, match="one million"):
        await mail_attachment_read.read_attachment(sdk, **ARGS)


async def test_worker_gets_no_credentials_and_timeout_kills_it(monkeypatch):
    child = SimpleNamespace(
        returncode=None, communicate=AsyncMock(return_value=(b"", b"")), kill=MagicMock()
    )
    spawn = AsyncMock(return_value=child)

    async def timed_out(task, timeout):
        assert timeout == 30
        await task
        raise asyncio.TimeoutError

    monkeypatch.setattr(asyncio, "create_subprocess_exec", spawn)
    monkeypatch.setattr(asyncio, "wait_for", timed_out)
    monkeypatch.setenv("PRIVATE_TEST_CREDENTIAL", "must-not-reach-parser")
    with pytest.raises(asyncio.TimeoutError):
        await mail_attachment_read.extract(b"abc", "text")
    child.kill.assert_called_once()
    assert all(key.lower() == "systemroot" for key in spawn.call_args.kwargs["env"])


def test_missing_optional_parsers_has_an_install_hint():
    # -S omits site-packages, proving a default install gets a useful error
    # without changing the test environment or relying on installed extras.
    import subprocess

    result = subprocess.run(
        [sys.executable, "-I", "-S", str(mail_attachment_read.WORKER), "pdf"],
        input=b"%PDF-1.4",
        capture_output=True,
        env={key: value for key, value in os.environ.items() if key.lower() == "systemroot"},
    )
    assert result.returncode == 0
    assert "outlook-graph-mcp[documents]" in json.loads(result.stdout)["error"]
