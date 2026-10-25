"""Attachments as content (#129): the bytes travel in the tool result and the tool arguments.

The attachment tools move files through `attachments_dir`, which only a client that
shares the server's disk can reach. `as_content=True` on the two download tools returns
the attachment in the result instead, and `attachments` on the two mail upload tools takes
files inline. Neither reads or writes the server's disk, so the fence
(`test_attachment_paths.py`) is not loosened: it still governs every path argument, and
the tests below check that the new arguments leave it in force.

What is new is size in the model's context, so each form has a cap, tested at its real
value: 10 MiB for a download, 3 MiB (Graph's simple-attachment size, the inline threshold
paths already use) for an inline upload.

Fidelity is checked the way #25's regression test checks it, end to end: the bytes start
as Graph JSON, go through kiota's real deserializer, the real tool and a real MCP client,
and must come back byte for byte. The upload side goes through kiota's real serializer in
`test_write_payloads_reach_the_wire.py`. Every fixture is synthetic, built here.
"""

from __future__ import annotations

import base64
import json
import struct
import zlib
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest
from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory
from mcp.client import Client
from msgraph.generated.models.attachment import Attachment
from msgraph.generated.models.attachment_base import AttachmentBase

from outlook_mcp import server
from outlook_mcp.config import Config
from outlook_mcp.tools import mail_attachments, todo_attachments
from outlook_mcp.tools.mail_attachments import (
    attach_to_draft,
    download_attachment,
    send_with_attachments,
)

MESSAGE_ID = "AAMkAG123="
CAP = 10 * 1024 * 1024
INLINE_CAP = 3 * 1024 * 1024


def _png() -> bytes:
    """A real 1x1 PNG. Its signature alone is not valid UTF-8 (0x89, CR LF, 0x1a)."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        crc = struct.pack(">I", zlib.crc32(kind + data))
        return struct.pack(">I", len(data)) + kind + data + crc

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    pixels = zlib.compress(b"\x00\xff\x00\x00")
    return (
        b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", pixels) + chunk(b"IEND", b"")
    )


# A real 1x1 JPEG (macOS `sips`, Exif segment kept so the segment walk has one to
# skip). Written from base64 rather than as a literal: a JPEG is not readable as bytes.
JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAAAaADAAQAAAABAAAAAQAAAAD/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9sAQwACAgICAgIDAgIDBQMDAwUGBQUFBQYIBgYGBgYICggICAgICAoKCgoKCgoKDAwMDAwMDg4ODg4PDw8PDw8PDw8P/9sAQwECAgIEBAQHBAQHEAsJCxAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ/90ABAAB/9oADAMBAAIRAxEAPwD4vooor+Uz/fw//9k="
)

PNG = _png()

# Real 1x1 GIF and WebP (GIF written by `sips`, WebP lossless by Pillow), not
# headers with filler after them: a fixture that cannot be opened tests that the
# server forwards unreadable files, which is the opposite of the point.
GIF = base64.b64decode("R0lGODdhAQABAJEAAAAAAP8AAP///wAAACH5BAQAAAAALAAAAAABAAEAAAICTAEAOw==")
WEBP = base64.b64decode("UklGRhwAAABXRUJQVlA4TA8AAAAvAAAAAAcQ/Y/+ByKi/wEA")

# A minimal PDF with the binary comment line real PDFs carry, which is not valid UTF-8.
PDF = (
    b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
    b"2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
    b"trailer\n<< /Root 1 0 R >>\n%%EOF\n"
)

DOCX = b"PK\x03\x04\x14\x00\x00\x00\x08\x00\xff\xfe\x00\x01binary"


def _graph_attachment(content: bytes, name: str, content_type: str, *, todo: bool = False):
    """The attachment as kiota builds it from Graph's JSON, base64 contentBytes and all."""
    odata_type = "taskFileAttachment" if todo else "fileAttachment"
    doc = {
        "@odata.type": f"#microsoft.graph.{odata_type}",
        "id": "att1",
        "name": name,
        "contentType": content_type,
        # Graph's size for a mail attachment includes metadata, so it is larger than the file.
        "size": len(content) + (0 if todo else 120),
        "contentBytes": base64.b64encode(content).decode(),
    }
    node = JsonParseNodeFactory().get_root_parse_node("application/json", json.dumps(doc).encode())
    return node.get_object_value(AttachmentBase if todo else Attachment)


def _mail_client(attachment):
    client = MagicMock()
    builder = client.me.messages.by_message_id.return_value.attachments.by_attachment_id
    builder.return_value.get = AsyncMock(return_value=attachment)
    return client, builder.return_value.get


def _file(content: bytes | None, name: str, content_type: str):
    att = MagicMock()
    att.name = name
    att.size = len(content or b"")
    att.content_type = content_type
    att.content_bytes = content
    return att


async def _blocks(content: bytes, name: str, content_type: str, tmp_path):
    client, _ = _mail_client(_file(content, name, content_type))
    return await download_attachment(
        client,
        message_id=MESSAGE_ID,
        attachment_id="att1",
        as_content=True,
        config=Config(attachments_dir=str(tmp_path / "att")),
    )


def _meta(blocks) -> dict:
    assert blocks[0].type == "text"
    return json.loads(blocks[0].text)


@pytest.fixture
def serve(monkeypatch, tmp_path):
    """The real server, with Graph replaced by ``sdk_client`` and a fresh attachments_dir."""

    def _serve(sdk_client):
        config = Config(client_id="test", attachments_dir=str(tmp_path / "att"))
        monkeypatch.setattr(
            server, "_get_graph_client", lambda ctx: SimpleNamespace(sdk_client=sdk_client)
        )
        monkeypatch.setattr(server, "_get_config", lambda ctx: config)
        return tmp_path / "att"

    return _serve


async def _call(tool: str, arguments: dict):
    async with Client(server.mcp) as client:
        return await client.call_tool(tool, arguments)


# ── Downloads reach the client byte for byte ─────────────────────────────────


class TestDownloadReachesTheClient:
    async def test_a_png_arrives_as_an_image_byte_for_byte(self, serve):
        client, _ = _mail_client(_graph_attachment(PNG, "pixel.png", "image/png"))
        attachments_dir = serve(client)

        result = await _call(
            "outlook_download_attachment",
            {"message_id": MESSAGE_ID, "attachment_id": "att1", "as_content": True},
        )

        assert result.is_error is False
        # Not echoed into structuredContent too, which would put the bytes in context twice.
        assert result.structured_content is None
        assert [b.type for b in result.content] == ["text", "image"]
        assert base64.b64decode(result.content[1].data) == PNG
        assert result.content[1].mime_type == "image/png"
        assert _meta(result.content) == {
            "name": "pixel.png",
            "size": len(PNG),
            "content_type": "image/png",
        }
        assert not attachments_dir.exists(), "as_content must not touch the server's disk"

    async def test_a_pdf_arrives_as_an_embedded_resource_byte_for_byte(self, serve):
        client, _ = _mail_client(_graph_attachment(PDF, "statement.pdf", "application/pdf"))
        attachments_dir = serve(client)

        result = await _call(
            "outlook_download_attachment",
            {"message_id": MESSAGE_ID, "attachment_id": "att1", "as_content": True},
        )

        assert result.is_error is False
        assert [b.type for b in result.content] == ["text", "resource"]
        resource = result.content[1].resource
        assert base64.b64decode(resource.blob) == PDF
        assert resource.mime_type == "application/pdf"
        assert not attachments_dir.exists()

    async def test_a_task_attachment_arrives_byte_for_byte(self, serve):
        from tests.test_todo_attachments import _build_mock_client

        entity = _graph_attachment(PNG, "pixel.png", "image/png", todo=True)
        attachments_dir = serve(_build_mock_client(attachment_entity=entity))

        result = await _call(
            "outlook_download_task_attachment",
            {"task_id": "task1", "attachment_id": "att1", "as_content": True},
        )

        assert result.is_error is False
        assert [b.type for b in result.content] == ["text", "image"]
        assert base64.b64decode(result.content[1].data) == PNG
        assert not attachments_dir.exists()

    async def test_saving_to_a_path_is_unchanged_and_still_the_default(self, serve):
        client, _ = _mail_client(_graph_attachment(PDF, "statement.pdf", "application/pdf"))
        attachments_dir = serve(client)
        attachments_dir.mkdir()

        result = await _call(
            "outlook_download_attachment",
            {"message_id": MESSAGE_ID, "attachment_id": "att1", "save_path": "statement.pdf"},
        )

        assert result.is_error is False
        assert json.loads(result.content[0].text)["saved_to"].endswith("statement.pdf")
        assert (attachments_dir / "statement.pdf").read_bytes() == PDF

    async def test_an_attachment_over_the_cap_reaches_the_model_as_a_refusal(self, serve):
        client, _ = _mail_client(_file(b"\x00" * (CAP + 1), "dump.bin", "application/zip"))
        serve(client)

        result = await _call(
            "outlook_download_attachment",
            {"message_id": MESSAGE_ID, "attachment_id": "att1", "as_content": True},
        )

        assert result.is_error is True
        assert "10 MiB limit for as_content" in result.content[0].text
        assert "save_path" in result.content[0].text


# ── What each kind of file comes back as ─────────────────────────────────────


class TestContentByType:
    async def test_text_is_returned_as_text(self, tmp_path):
        text = "name\tamount\r\nTea\t3.50\r\nCafé\t4.20\r\n"

        blocks = await _blocks(text.encode(), "receipt.tsv", "text/tab-separated-values", tmp_path)

        assert [b.type for b in blocks] == ["text", "text"]
        assert blocks[1].text == text

    async def test_json_and_xml_are_text(self, tmp_path):
        for name, content_type in (("a.json", "application/json"), ("a.xml", "text/xml")):
            blocks = await _blocks(b'{"a": 1}', name, content_type, tmp_path)
            assert blocks[1].type == "text", content_type

    async def test_a_declared_charset_is_honoured(self, tmp_path):
        blocks = await _blocks(
            "Café".encode("latin-1"), "a.txt", "text/plain; charset=latin-1", tmp_path
        )

        assert blocks[1].text == "Café"

    async def test_text_that_does_not_decode_comes_back_whole_as_a_resource(self, tmp_path):
        """Never a lossy decode: replacement characters would hand over different bytes."""
        content = b"caf\xe9 au lait"

        blocks = await _blocks(content, "note.txt", "text/plain", tmp_path)

        assert blocks[1].type == "resource"
        assert base64.b64decode(blocks[1].resource.blob) == content
        assert blocks[1].resource.mime_type == "text/plain"

    async def test_text_with_control_characters_comes_back_as_a_resource(self, tmp_path):
        """Text is never shown altered, and never with a sender's terminal escapes in it."""
        content = b"ok \x1b[2J\x1b[31mcleared\n"

        blocks = await _blocks(content, "note.txt", "text/plain", tmp_path)

        assert blocks[1].type == "resource"
        assert base64.b64decode(blocks[1].resource.blob) == content

    async def test_text_over_the_limit_comes_back_whole_as_a_resource(self, tmp_path):
        limit = mail_attachments._TEXT_CONTENT_MAX_CHARS
        assert limit == 100_000

        at_limit = await _blocks(b"a" * limit, "a.txt", "text/plain", tmp_path)
        over = await _blocks(b"a" * (limit + 1), "a.txt", "text/plain", tmp_path)

        assert at_limit[1].type == "text"
        assert over[1].type == "resource"
        assert base64.b64decode(over[1].resource.blob) == b"a" * (limit + 1)

    async def test_an_image_over_the_limit_comes_back_as_a_resource(self, tmp_path):
        """Some model APIs refuse an image over 5 MB; the bytes still arrive whole."""
        limit = mail_attachments._IMAGE_CONTENT_MAX_BYTES
        assert limit == 5 * 1024 * 1024
        # Pad *before* the terminal chunk, so the file at the limit is still a PNG
        # that ends the way it starts; `over` gains a byte after IEND and is not.
        tail, head = PNG[-12:], PNG[:-12]
        at_limit = head + b"\x00" * (limit - len(PNG)) + tail
        over = at_limit + b"\x00"

        assert (await _blocks(at_limit, "a.png", "image/png", tmp_path))[1].type == "image"
        blocks = await _blocks(over, "a.png", "image/png", tmp_path)
        assert blocks[1].type == "resource"
        assert blocks[1].resource.mime_type == "image/png"
        assert base64.b64decode(blocks[1].resource.blob) == over

    @pytest.mark.parametrize(
        ("content_type", "content"),
        [
            # Complete files: end-of-image, the GIF trailer, a RIFF length that
            # matches. A signature on its own is not an image a client can open.
            ("image/jpeg", JPEG),
            ("image/gif", GIF),
            ("image/webp", WEBP),
        ],
    )
    async def test_every_image_type_a_model_reads_is_an_image(
        self, tmp_path, content_type, content
    ):
        blocks = await _blocks(content, "photo", content_type, tmp_path)

        assert blocks[1].type == "image"
        assert blocks[1].mime_type == content_type

    @pytest.mark.parametrize(
        ("content_type", "content"),
        [
            # Each carries a valid signature and nothing after it — the shape a
            # truncated download or a sender's hand-made file takes.
            ("image/png", b"\x89PNG\r\n\x1a\n"),
            ("image/jpeg", b"\xff\xd8\xff\xe0\x00\x10JFIF"),
            ("image/gif", b"GIF89a\x01\x00\x01\x00"),
        ],
    )
    async def test_a_truncated_image_is_a_resource_not_an_image(
        self, tmp_path, content_type, content
    ):
        """A client that cannot open the image can reject the whole tool result."""
        blocks = await _blocks(content, "photo", content_type, tmp_path)

        assert blocks[1].type == "resource"
        assert base64.b64decode(blocks[1].resource.blob) == content

    async def test_a_png_cut_before_its_terminal_chunk_is_a_resource(self, tmp_path):
        """Signature, IHDR and IDAT, no IEND: half a file is not a file."""
        blocks = await _blocks(PNG[:-12], "photo.png", "image/png", tmp_path)

        assert blocks[1].type == "resource"

    async def test_an_image_name_over_bytes_that_are_no_image_is_a_resource(self, tmp_path):
        """A sender can name anything `chart.png`; some model APIs reject a whole result
        whose image does not decode, so only bytes that open as an image are an image."""
        content = b"%PDF-1.7 not a picture"
        blocks = await _blocks(content, "chart.png", "image/png", tmp_path)

        assert blocks[1].type == "resource"
        assert blocks[1].resource.mime_type == "image/png"
        assert base64.b64decode(blocks[1].resource.blob) == content

    async def test_an_image_labelled_as_another_image_type_is_sent_as_what_it_is(self, tmp_path):
        blocks = await _blocks(PNG, "photo.jpg", "image/jpeg", tmp_path)

        assert blocks[1].type == "image"
        assert blocks[1].mime_type == "image/png"
        assert _meta(blocks)["content_type"] == "image/jpeg"

    async def test_an_image_type_a_model_cannot_read_is_a_resource(self, tmp_path):
        blocks = await _blocks(b"II*\x00", "scan.tiff", "image/tiff", tmp_path)

        assert blocks[1].type == "resource"
        assert blocks[1].resource.mime_type == "image/tiff"

    async def test_anything_else_is_a_base64_resource_with_its_type(self, tmp_path):
        docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        blocks = await _blocks(DOCX, "report.docx", docx, tmp_path)

        assert blocks[1].type == "resource"
        assert blocks[1].resource.mime_type == docx
        assert base64.b64decode(blocks[1].resource.blob) == DOCX
        assert _meta(blocks) == {"name": "report.docx", "size": len(DOCX), "content_type": docx}

    async def test_a_generic_type_is_read_from_the_name(self, tmp_path):
        blocks = await _blocks(PNG, "pixel.png", "application/octet-stream", tmp_path)

        assert blocks[1].type == "image"
        assert blocks[1].mime_type == "image/png"

    async def test_no_type_and_no_extension_is_octet_stream(self, tmp_path):
        blocks = await _blocks(DOCX, "blob", "", tmp_path)

        assert blocks[1].resource.mime_type == "application/octet-stream"

    async def test_the_resource_uri_carries_a_file_name_and_never_a_path(self, tmp_path):
        blocks = await _blocks(PDF, "..\\../etc/My Report.pdf", "application/pdf", tmp_path)

        uri = str(blocks[1].resource.uri)
        assert uri == "attachment:///My%20Report.pdf"

    async def test_the_name_and_type_it_reports_are_sanitised(self, tmp_path):
        blocks = await _blocks(
            PDF, "report\x1b[31m\x07.pdf", "application/pdf\r\nX-Injected: 1", tmp_path
        )

        meta = _meta(blocks)
        assert meta["name"] == "report.pdf"
        assert "\r" not in meta["content_type"] and "\n" not in meta["content_type"]


# ── The cap, and what as_content never does ──────────────────────────────────


class TestDownloadLimits:
    async def test_the_cap_is_10_mib_and_inclusive(self, tmp_path):
        blocks = await _blocks(b"\x00" * CAP, "a.zip", "application/zip", tmp_path)

        assert len(base64.b64decode(blocks[1].resource.blob)) == CAP

    async def test_over_the_cap_is_refused_naming_size_cap_and_save_path(self, tmp_path):
        with pytest.raises(ValueError) as exc:
            await _blocks(b"\x00" * (CAP + 1), "a.zip", "application/zip", tmp_path)

        message = str(exc.value)
        assert str(CAP + 1) in message
        assert "10 MiB" in message
        assert "save_path" in message

    async def test_the_cap_applies_to_task_attachments_too(self, tmp_path):
        from tests.test_todo_attachments import _build_mock_client

        entity = _graph_attachment(b"\x00" * (CAP + 1), "a.zip", "application/zip", todo=True)

        with pytest.raises(ValueError, match="save_path"):
            await todo_attachments.download_task_attachment(
                _build_mock_client(attachment_entity=entity),
                task_id="task1",
                attachment_id="att1",
                as_content=True,
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

    async def test_a_task_attachment_that_contradicts_its_size_is_still_refused(self, tmp_path):
        """The To Do download's size witness holds in content mode too."""
        from tests.test_todo_attachments import _build_mock_client

        entity = MagicMock()
        entity.name = "a.pdf"
        entity.content_type = "application/pdf"
        entity.size = 4096
        entity.content_bytes = None

        with pytest.raises(ValueError, match="declares 4096 bytes"):
            await todo_attachments.download_task_attachment(
                _build_mock_client(attachment_entity=entity),
                task_id="task1",
                attachment_id="att1",
                as_content=True,
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

    async def test_an_attached_email_is_refused_in_content_mode_too(self, tmp_path):
        item = MagicMock(spec=["name", "size", "content_type", "odata_type"])
        item.name = "Fwd: invoice"
        item.odata_type = "#microsoft.graph.itemAttachment"
        client, _ = _mail_client(item)

        with pytest.raises(ValueError, match="no file content"):
            await download_attachment(
                client,
                message_id=MESSAGE_ID,
                attachment_id="att1",
                as_content=True,
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

    @pytest.mark.parametrize(
        ("save_path", "as_content", "match"),
        [
            (None, False, "as_content=True"),
            ("report.pdf", True, "not both"),
        ],
        ids=["neither", "both"],
    )
    async def test_exactly_one_destination_before_graph_is_asked(
        self, tmp_path, save_path, as_content, match
    ):
        from tests.test_todo_attachments import _build_mock_client

        client, get = _mail_client(_file(PDF, "report.pdf", "application/pdf"))
        config = Config(attachments_dir=str(tmp_path / "att"))

        with pytest.raises(ValueError, match=match):
            await download_attachment(
                client,
                message_id=MESSAGE_ID,
                attachment_id="att1",
                save_path=save_path,
                as_content=as_content,
                config=config,
            )
        get.assert_not_called()

        todo_client = _build_mock_client()
        with pytest.raises(ValueError, match=match):
            await todo_attachments.download_task_attachment(
                todo_client,
                task_id="task1",
                attachment_id="att1",
                save_path=save_path,
                as_content=as_content,
                config=config,
            )
        assert not (tmp_path / "att").exists()


# ── Inline uploads ───────────────────────────────────────────────────────────


def _inline(content: bytes, name: str = "pixel.png", **extra) -> dict:
    return {"name": name, "content_base64": base64.b64encode(content).decode(), **extra}


def _draft_client():
    client = MagicMock()
    builder = client.me.messages.by_message_id.return_value
    builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
    created = iter(["att_a", "att_b", "att_c"])
    builder.attachments.post = AsyncMock(side_effect=lambda att: MagicMock(id=next(created)))
    return client, builder


def _posted(builder) -> list:
    return [call.args[0] for call in builder.attachments.post.call_args_list]


class TestInlineUploads:
    async def test_inline_files_are_attached_with_their_bytes_and_ids(self, tmp_path):
        client, builder = _draft_client()

        result = await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[_inline(PNG), _inline(PDF, "statement.pdf")],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert result["attachment_count"] == 2
        assert result["attachment_ids"] == ["att_a", "att_b"]
        png, pdf = _posted(builder)
        assert (png.name, png.content_type, png.content_bytes) == ("pixel.png", "image/png", PNG)
        assert (pdf.name, pdf.content_type, pdf.content_bytes) == (
            "statement.pdf",
            "application/pdf",
            PDF,
        )

    async def test_a_given_content_type_wins_over_the_name(self, tmp_path):
        client, builder = _draft_client()

        await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[_inline(b"a,b\n", "data.txt", content_type="text/csv")],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert _posted(builder)[0].content_type == "text/csv"

    async def test_no_type_and_no_extension_is_octet_stream(self, tmp_path):
        client, builder = _draft_client()

        await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[_inline(DOCX, "blob")],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert _posted(builder)[0].content_type == "application/octet-stream"

    async def test_whitespace_in_the_base64_is_tolerated(self, tmp_path):
        """Line-wrapped base64 (MIME style) decodes to the same bytes."""
        client, builder = _draft_client()
        encoded = base64.encodebytes(PDF).decode()
        assert "\n" in encoded

        await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[{"name": "s.pdf", "content_base64": encoded}],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert _posted(builder)[0].content_bytes == PDF

    async def test_paths_and_inline_files_mix(self, tmp_path):
        attachments_dir = tmp_path / "att"
        attachments_dir.mkdir()
        (attachments_dir / "on-disk.pdf").write_bytes(PDF)
        client, builder = _draft_client()

        result = await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachment_paths=["on-disk.pdf"],
            attachments=[_inline(PNG)],
            config=Config(attachments_dir=str(attachments_dir)),
        )

        assert result["attachment_count"] == 2
        assert result["attachment_ids"] == ["att_a", "att_b"]
        assert [a.name for a in _posted(builder)] == ["on-disk.pdf", "pixel.png"]

    async def test_inline_files_never_touch_the_servers_disk(self, tmp_path):
        attachments_dir = tmp_path / "att"
        client, _ = _draft_client()
        client.me.send_mail.post = AsyncMock()
        config = Config(attachments_dir=str(attachments_dir))

        await attach_to_draft(
            client, draft_id="AAMkAG123=", attachments=[_inline(PNG)], config=config
        )
        await send_with_attachments(
            client,
            to=["someone@example.com"],
            subject="s",
            body="b",
            attachments=[_inline(PNG)],
            config=config,
        )

        assert not attachments_dir.exists()

    async def test_send_puts_inline_files_on_the_message(self, tmp_path):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        result = await send_with_attachments(
            client,
            to=["someone@example.com"],
            subject="s",
            body="b",
            attachments=[_inline(PNG), _inline(PDF, "statement.pdf")],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert result == {"status": "sent", "attachment_count": 2}
        message = client.me.send_mail.post.call_args.args[0].message
        assert [(a.name, a.content_bytes) for a in message.attachments] == [
            ("pixel.png", PNG),
            ("statement.pdf", PDF),
        ]

    async def test_send_with_a_large_path_puts_inline_files_on_the_draft(self, tmp_path):
        attachments_dir = tmp_path / "att"
        attachments_dir.mkdir()
        (attachments_dir / "big.bin").write_bytes(b"\x00" * (INLINE_CAP + 1))
        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=MagicMock(id="draft1"))
        draft = client.me.messages.by_message_id.return_value
        draft.attachments.create_upload_session.post = AsyncMock(
            return_value=MagicMock(upload_url="https://upload.example/session")
        )
        draft.send.post = AsyncMock()
        upload = AsyncMock()

        with pytest.MonkeyPatch.context() as mp:
            mp.setattr(mail_attachments, "_upload_large_file", upload)
            result = await send_with_attachments(
                client,
                to=["someone@example.com"],
                subject="s",
                body="b",
                attachment_paths=["big.bin"],
                attachments=[_inline(PNG)],
                config=Config(attachments_dir=str(attachments_dir)),
            )

        assert result["attachment_count"] == 2
        message = client.me.messages.post.call_args.args[0]
        assert [(a.name, a.content_bytes) for a in message.attachments] == [("pixel.png", PNG)]
        upload.assert_awaited_once()
        draft.send.post.assert_awaited_once()


class TestInlineUploadsAreRefusedBeforeAnythingIsSent:
    async def test_a_message_that_is_not_a_draft_is_refused_for_inline_files_too(self, tmp_path):
        """The draft check ran only when there were paths; inline files alone must not skip it."""
        client, builder = _draft_client()
        builder.get = AsyncMock(return_value=MagicMock(is_draft=False))

        with pytest.raises(ValueError, match="not a draft"):
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[_inline(PNG)],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        builder.attachments.post.assert_not_called()

    @pytest.mark.parametrize(
        ("item", "match"),
        [
            pytest.param("pixel.png", "must be an object", id="not-an-object"),
            pytest.param({"content_base64": "AAAA"}, "needs a name", id="no-name"),
            pytest.param({"name": "  ", "content_base64": "AAAA"}, "needs a name", id="blank-name"),
            pytest.param(
                {"name": "a\r\nb.pdf", "content_base64": "AAAA"}, "needs a name", id="name-crlf"
            ),
            pytest.param({"name": "..", "content_base64": "AAAA"}, "not a file name", id="dot-dot"),
            pytest.param({"name": ".", "content_base64": "AAAA"}, "not a file name", id="dot"),
            pytest.param(
                {"name": "a" * 252 + ".pdf", "content_base64": "AAAA"}, "255", id="name-too-long"
            ),
            pytest.param({"name": "a.pdf"}, "needs content_base64", id="no-content"),
            pytest.param(
                {"name": "a.pdf", "content_base64": "not base64!"}, "not valid base64", id="bad-b64"
            ),
            pytest.param(
                {"name": "a.pdf", "content_base64": "AAAA", "content_type": ""},
                "content_type",
                id="blank-type",
            ),
            pytest.param(
                {"name": "a.pdf", "content": "AAAA"}, r"unknown key\(s\) \['content'\]", id="typo"
            ),
        ],
    )
    async def test_a_malformed_item_is_refused(self, tmp_path, item, match):
        client, builder = _draft_client()

        with pytest.raises(ValueError, match=match):
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[_inline(PNG), item],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        builder.get.assert_not_called()
        builder.attachments.post.assert_not_called()

    async def test_a_name_is_stripped_and_255_characters_still_fit(self, tmp_path):
        client, builder = _draft_client()
        longest = "a" * 251 + ".png"

        await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[_inline(PNG, "  pixel.png "), _inline(PNG, longest)],
            config=Config(attachments_dir=str(tmp_path / "att")),
        )

        assert [att.name for att in _posted(builder)] == ["pixel.png", longest]

    async def test_far_over_the_cap_is_refused_before_it_is_decoded(self, tmp_path, monkeypatch):
        """Decoding 30 MiB of base64 only to refuse it is waste; its length says enough."""
        from outlook_mcp import validation

        def no_decoding(*args, **kwargs):
            raise AssertionError("decoded")

        monkeypatch.setattr(validation.base64, "b64decode", no_decoding)
        client, builder = _draft_client()

        with pytest.raises(ValueError, match="3 MiB") as exc:
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[{"name": "dump.bin", "content_base64": "A" * (40 * 1024 * 1024)}],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )
        assert "dump.bin" in str(exc.value) and "attachments_dir" in str(exc.value)
        builder.attachments.post.assert_not_called()

    @pytest.mark.parametrize("wrapped", [False, True])
    async def test_one_byte_over_the_cap_is_refused_before_decoding(
        self, tmp_path, monkeypatch, wrapped
    ):
        from outlook_mcp import validation

        encoded = base64.b64encode(b"x" * (INLINE_CAP + 1)).decode()
        if wrapped:
            encoded = "\r\n".join(encoded[i : i + 76] for i in range(0, len(encoded), 76))

        def no_decoding(*args, **kwargs):
            raise AssertionError("decoded an oversized file")

        monkeypatch.setattr(validation.base64, "b64decode", no_decoding)
        client, builder = _draft_client()
        with pytest.raises(ValueError, match=f"{INLINE_CAP + 1} bytes"):
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[{"name": "over.bin", "content_base64": encoded}],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )
        builder.get.assert_not_called()
        builder.attachments.post.assert_not_called()

    async def test_the_inline_cap_is_graphs_3_mib_and_inclusive(self, tmp_path):
        client, builder = _draft_client()
        config = Config(attachments_dir=str(tmp_path / "att"))

        await attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[_inline(b"\x00" * INLINE_CAP)],
            config=config,
        )
        assert len(_posted(builder)[0].content_bytes) == INLINE_CAP

        with pytest.raises(ValueError) as exc:
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[_inline(b"\x00" * (INLINE_CAP + 1), "big.bin")],
                config=config,
            )
        message = str(exc.value)
        assert "big.bin" in message and str(INLINE_CAP + 1) in message
        assert "3 MiB" in message and "attachments_dir" in message
        assert builder.attachments.post.call_count == 1

    async def test_send_refuses_before_sending(self, tmp_path):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        with pytest.raises(ValueError, match="not valid base64"):
            await send_with_attachments(
                client,
                to=["someone@example.com"],
                subject="s",
                body="b",
                attachments=[{"name": "a.pdf", "content_base64": "%%%"}],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        client.me.send_mail.post.assert_not_called()

    async def test_read_only_still_refuses_inline_uploads(self, tmp_path):
        from outlook_mcp.errors import ReadOnlyError

        client, builder = _draft_client()
        client.me.send_mail.post = AsyncMock()
        config = Config(attachments_dir=str(tmp_path / "att"), read_only=True)

        with pytest.raises(ReadOnlyError):
            await attach_to_draft(
                client, draft_id="AAMkAG123=", attachments=[_inline(PNG)], config=config
            )
        with pytest.raises(ReadOnlyError):
            await send_with_attachments(
                client,
                to=["someone@example.com"],
                subject="s",
                body="b",
                attachments=[_inline(PNG)],
                config=config,
            )
        builder.attachments.post.assert_not_called()
        client.me.send_mail.post.assert_not_called()


class TestAnUploadWithNothingToAttachIsRefused:
    """Both upload tools take `attachment_paths` and `attachments`, and both are optional,
    so a call with neither, or with a misspelt one (the server ignores unknown top-level
    arguments), would send the mail or report "attached" with no file: the #41 class, a
    call that succeeds and does nothing. Each is refused before Graph is called, through
    the real server, so the error is what the model sees."""

    @staticmethod
    def _graph():
        client, builder = _draft_client()
        client.me.send_mail.post = AsyncMock()
        client.me.messages.post = AsyncMock()
        return client, builder

    @staticmethod
    def _untouched(client, builder):
        client.me.send_mail.post.assert_not_called()
        client.me.messages.post.assert_not_called()
        builder.get.assert_not_called()
        builder.attachments.post.assert_not_called()

    @pytest.mark.parametrize(
        "extra",
        [{"attachment": ["report.pdf"]}, {}],
        ids=["misspelt-argument", "no-attachment-argument"],
    )
    async def test_send_with_attachments_refuses_and_sends_nothing(self, serve, extra):
        client, builder = self._graph()
        serve(client)

        result = await _call(
            "outlook_send_with_attachments",
            {"to": ["someone@example.com"], "subject": "s", "body": "b", **extra},
        )

        assert result.is_error is True
        text = result.content[0].text
        assert "attachment_paths" in text and "attachments" in text
        assert "outlook_send_message" in text
        self._untouched(client, builder)

    @pytest.mark.parametrize(
        "extra",
        [{"attachment": ["report.pdf"]}, {}],
        ids=["misspelt-argument", "no-attachment-argument"],
    )
    async def test_attach_to_draft_refuses_and_touches_nothing(self, serve, extra):
        client, builder = self._graph()
        serve(client)

        result = await _call("outlook_attach_to_draft", {"draft_id": "AAMkAG123=", **extra})

        assert result.is_error is True
        text = result.content[0].text
        assert "attachment_paths" in text and "attachments" in text
        self._untouched(client, builder)

    async def test_empty_lists_are_refused_too(self, tmp_path):
        client, builder = self._graph()
        config = Config(attachments_dir=str(tmp_path / "att"))

        with pytest.raises(ValueError, match="attachment_paths"):
            await attach_to_draft(
                client, draft_id="AAMkAG123=", attachment_paths=[], attachments=[], config=config
            )
        with pytest.raises(ValueError, match="attachment_paths"):
            await send_with_attachments(
                client,
                to=["someone@example.com"],
                subject="s",
                body="b",
                attachment_paths=[],
                attachments=[],
                config=config,
            )
        self._untouched(client, builder)


class TestTheFenceStillGovernsEveryPath:
    """Inline content adds a way in that has no path; it opens nothing for the paths."""

    @pytest.fixture
    def secret(self, tmp_path):
        path = tmp_path / "id_ed25519"
        path.write_text("PRIVATE KEY")
        return path

    async def test_attach_refuses_a_path_outside_beside_inline_files(self, tmp_path, secret):
        client, builder = _draft_client()

        with pytest.raises(ValueError, match="attachments_dir"):
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachment_paths=[str(secret)],
                attachments=[_inline(PNG)],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        builder.attachments.post.assert_not_called()

    async def test_send_refuses_a_path_outside_beside_inline_files(self, tmp_path, secret):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        with pytest.raises(ValueError, match="attachments_dir"):
            await send_with_attachments(
                client,
                to=["someone@example.com"],
                subject="s",
                body="b",
                attachment_paths=[str(secret)],
                attachments=[_inline(PNG)],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        client.me.send_mail.post.assert_not_called()

    @pytest.mark.parametrize("separator", ["/", "\\"])
    async def test_an_inline_name_that_is_a_path_is_refused(self, tmp_path, secret, separator):
        """`name` is the attachment's file name. A model that passes a path there may think
        it attaches that file; it would not, and the refusal says where paths go."""
        client, builder = _draft_client()
        name = str(secret).replace("/", separator)

        with pytest.raises(ValueError, match="attachment_paths"):
            await attach_to_draft(
                client,
                draft_id="AAMkAG123=",
                attachments=[_inline(b"inline bytes", name)],
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        builder.attachments.post.assert_not_called()

    async def test_a_download_with_both_destinations_writes_nowhere(self, tmp_path, secret):
        client, get = _mail_client(_file(b"attacker bytes", "x", "application/pdf"))

        with pytest.raises(ValueError):
            await download_attachment(
                client,
                message_id=MESSAGE_ID,
                attachment_id="att1",
                save_path=str(secret),
                as_content=True,
                config=Config(attachments_dir=str(tmp_path / "att")),
            )

        assert secret.read_text() == "PRIVATE KEY"
        get.assert_not_called()


async def test_send_with_attachments_takes_inline_files_through_the_server(serve):
    client = MagicMock()
    client.me.send_mail.post = AsyncMock()
    attachments_dir = serve(client)

    result = await _call(
        "outlook_send_with_attachments",
        {
            "to": ["someone@example.com"],
            "subject": "s",
            "body": "b",
            "attachments": [_inline(PDF, "statement.pdf")],
        },
    )

    assert result.is_error is False, result.content[0].text
    assert json.loads(result.content[0].text) == {"status": "sent", "attachment_count": 1}
    message = client.me.send_mail.post.call_args.args[0].message
    assert [(a.name, a.content_bytes) for a in message.attachments] == [("statement.pdf", PDF)]
    assert not attachments_dir.exists()


# ── The whole loop: download as content, attach inline ───────────────────────


async def test_a_downloaded_pdf_attaches_back_byte_for_byte(serve):
    """What a remote client does: read an attachment, then send it on, never via disk."""
    mail, _ = _mail_client(_graph_attachment(PDF, "statement.pdf", "application/pdf"))
    _, builder = _draft_client()
    mail.me.messages.by_message_id.return_value.get = builder.get
    mail.me.messages.by_message_id.return_value.attachments.post = builder.attachments.post
    attachments_dir = serve(mail)

    downloaded = await _call(
        "outlook_download_attachment",
        {"message_id": MESSAGE_ID, "attachment_id": "att1", "as_content": True},
    )
    resource = downloaded.content[1].resource
    attached = await _call(
        "outlook_attach_to_draft",
        {
            "draft_id": "AAMkAG456=",
            "attachments": [
                {
                    "name": "statement.pdf",
                    "content_base64": resource.blob,
                    "content_type": resource.mime_type,
                }
            ],
        },
    )

    assert attached.is_error is False, attached.content[0].text
    assert json.loads(attached.content[0].text)["attachment_ids"] == ["att_a"]
    assert builder.attachments.post.call_args.args[0].content_bytes == PDF
    assert not attachments_dir.exists()


@pytest.mark.asyncio
@pytest.mark.parametrize("send", [False, True])
async def test_whitespace_heavy_raw_input_is_refused_before_normalization(tmp_path, send):
    class NoSplit(str):
        def split(self, *args, **kwargs):
            raise AssertionError("normalized unbounded whitespace-heavy input")

    encoded = NoSplit("AA " * (40 * 1024 * 1024 // 3))
    client, builder = _draft_client()
    kwargs = {
        "attachments": [{"name": "dump.bin", "content_base64": encoded}],
        "config": Config(attachments_dir=str(tmp_path / "att")),
    }
    with pytest.raises(ValueError, match="base64 input.*3 MiB"):
        if send:
            await send_with_attachments(client, ["a@example.com"], "s", "b", **kwargs)
        else:
            await attach_to_draft(client, "AAMkAG123=", **kwargs)
    builder.get.assert_not_called()
    builder.attachments.post.assert_not_called()
    client.me.send_mail.post.assert_not_called()


@pytest.mark.asyncio
@pytest.mark.parametrize("send", [False, True])
@pytest.mark.parametrize("encoded", ["", " \t\r\n"])
async def test_empty_inline_content_is_refused_without_graph_calls(tmp_path, send, encoded):
    client, builder = _draft_client()
    kwargs = {
        "attachments": [{"name": "empty.bin", "content_base64": encoded}],
        "config": Config(attachments_dir=str(tmp_path / "att")),
    }
    with pytest.raises(ValueError, match="content_base64 must not be empty"):
        if send:
            await send_with_attachments(client, ["a@example.com"], "s", "b", **kwargs)
        else:
            await attach_to_draft(client, "AAMkAG123=", **kwargs)
    builder.get.assert_not_called()
    builder.attachments.post.assert_not_called()
    client.me.send_mail.post.assert_not_called()


@pytest.mark.asyncio
@pytest.mark.parametrize("send", [False, True])
@pytest.mark.parametrize("suffix", ["!!!!", "====", "AAAAA"])
async def test_invalid_near_cap_content_reports_base64_without_decoding(
    tmp_path, monkeypatch, send, suffix
):
    from outlook_mcp import validation

    def no_decoding(*args, **kwargs):
        raise AssertionError("decoded an oversized input")

    monkeypatch.setattr(validation.base64, "b64decode", no_decoding)
    client, builder = _draft_client()
    kwargs = {
        "attachments": [
            {"name": "bad.bin", "content_base64": "A" * (4 * INLINE_CAP // 3) + suffix}
        ],
        "config": Config(attachments_dir=str(tmp_path / "att")),
    }
    with pytest.raises(ValueError, match="content_base64 is not valid base64"):
        if send:
            await send_with_attachments(client, ["a@example.com"], "s", "b", **kwargs)
        else:
            await attach_to_draft(client, "AAMkAG123=", **kwargs)
    builder.get.assert_not_called()
    builder.attachments.post.assert_not_called()
    client.me.send_mail.post.assert_not_called()


class TestJpegTrailers:
    """A phone photo is not required to end where its image ends.

    Samsung's camera app writes a `SEFH…SEFT` trailer after end-of-image, and
    Google's Motion Photo — the default on Pixel — appends a whole MP4. Both open
    fine, because decoders stop at EOI. Turning them away would make a phone photo
    unviewable, which is worse than the truncated file this check exists for.
    """

    SAMSUNG = b"SEFH" + b"\x00" * 12 + b"SEFT" + b"\x00" * 8
    MOTION_PHOTO = b"\x00\x00\x00\x18ftypmp42\x00\x00\x00\x00mp42isom" + b"\x00" * 2048

    @pytest.mark.parametrize(
        "trailer",
        [SAMSUNG, MOTION_PHOTO, b"\x00\x00"],
        ids=["samsung-trailer", "motion-photo-mp4", "two-zero-bytes"],
    )
    async def test_bytes_after_end_of_image_are_still_an_image(self, tmp_path, trailer):
        blocks = await _blocks(JPEG + trailer, "photo.jpg", "image/jpeg", tmp_path)

        assert blocks[1].type == "image"

    async def test_a_jpeg_cut_in_half_is_a_resource(self, tmp_path):
        blocks = await _blocks(JPEG[: len(JPEG) // 2], "photo.jpg", "image/jpeg", tmp_path)

        assert blocks[1].type == "resource"

    async def test_a_jpeg_cut_just_after_the_scan_header_is_a_resource(self, tmp_path):
        """The header arrives, the scan data does not: nothing to open."""
        scan = JPEG.index(b"\xff\xda")
        blocks = await _blocks(JPEG[: scan + 6], "photo.jpg", "image/jpeg", tmp_path)

        assert blocks[1].type == "resource"


class TestIncompleteImages:
    """A signature, or a stray marker, is not a file.

    The terminal bytes have to be the terminal *chunk*: the four ASCII bytes
    "IEND" appear in a file only as part of one, so searching for them accepted
    `signature + b"IEND"` — no dimensions, no pixels — and sent it as an image.
    """

    SIGNATURE = b"\x89PNG\r\n\x1a\n"

    async def test_signature_and_the_letters_iend_are_not_a_png(self, tmp_path):
        blocks = await _blocks(self.SIGNATURE + b"IEND", "photo.png", "image/png", tmp_path)

        assert blocks[1].type == "resource"

    async def test_an_iend_chunk_with_the_wrong_crc_is_not_a_png(self, tmp_path):
        broken = self.SIGNATURE + b"\x00\x00\x00\x00IEND\x00\x00\x00\x00"
        blocks = await _blocks(broken, "photo.png", "image/png", tmp_path)

        assert blocks[1].type == "resource"

    async def test_a_png_with_trailing_padding_is_still_a_png(self, tmp_path):
        """Padding after the terminal chunk is tolerated: readers stop at IEND."""
        blocks = await _blocks(PNG + b"\x00\x00", "photo.png", "image/png", tmp_path)

        assert blocks[1].type == "image"

    async def test_a_gif_cut_before_its_trailer_is_a_resource(self, tmp_path):
        blocks = await _blocks(GIF[:-1], "anim.gif", "image/gif", tmp_path)

        assert blocks[1].type == "resource"

    async def test_a_webp_shorter_than_its_declared_length_is_a_resource(self, tmp_path):
        blocks = await _blocks(WEBP[:-4], "photo.webp", "image/webp", tmp_path)

        assert blocks[1].type == "resource"
