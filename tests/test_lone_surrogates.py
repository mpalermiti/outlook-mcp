"""A lone surrogate in a reply must not end the server (#186).

A lone UTF-16 surrogate (U+D800-U+DFFF) is a legal code point in a Python ``str`` and has
no UTF-8 encoding. Where it lands decides what breaks. In a content block or in error text
it reaches the SDK's stdio writer, which raises: the server process exits, the call gets no
answer and the session is over. A text attachment downloaded as content, its name, and
error text naming sender data all went that way. In a dict result the SDK's own conversion
fails first, so that one call fails with a bare ``Error executing tool``. The sender has
more than one way in: a declared charset decodes ``+2AA-`` (utf-7) or ``\\ud800``
(unicode_escape) to one, and a JSON ``\\ud800`` escape in anything Graph returns decodes to
one too.

The in-memory client the other tests use never encodes the reply, so it passes either
way. ``test_the_stdio_server_answers_and_stays_up`` runs the real server over real stdio.
"""

from __future__ import annotations

import asyncio
import base64
import json
import os
import sys

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import CallToolResult, EmbeddedResource, TextContent, TextResourceContents

from outlook_mcp.config import Config
from outlook_mcp.errors import OutlookMCPError, ToolInputError
from outlook_mcp.server import _wrap_tool_errors
from outlook_mcp.tools.mail_attachments import download_attachment
from outlook_mcp.validation import has_control_characters, sanitize_output
from tests.test_attachments_as_content import _graph_attachment, _mail_client

MESSAGE_ID = "AAMkAG123="


def _encodes(value: object) -> None:
    """Raise if ``value`` cannot travel as UTF-8 JSON, as a stdio reply must."""
    json.dumps(value, ensure_ascii=False).encode("utf-8")


# ── The cleaner every echo point uses ────────────────────────────────────────


@pytest.mark.parametrize(
    "raw",
    ["a\ud800b", "a\udfffb", "\ud800", "end\udc00"],
    ids=["high", "low", "alone", "at-the-end"],
)
def test_sanitize_output_replaces_every_lone_surrogate(raw):
    for multiline in (False, True):
        cleaned = sanitize_output(raw, multiline=multiline)
        cleaned.encode("utf-8")
        assert chr(0xFFFD) in cleaned


def test_sanitize_output_joins_a_split_pair():
    """Two halves a decoder left unjoined are the character they encode, not two errors."""
    assert sanitize_output(chr(0xD83D) + chr(0xDE00)) == chr(0x1F600)


def test_sanitize_output_keeps_real_characters():
    assert sanitize_output("café 😀 日本") == "café 😀 日本"


def test_a_lone_surrogate_counts_as_a_character_the_cleaner_changes():
    """So text that carries one is never sent on as unaltered text."""
    assert has_control_characters("a\ud800b")
    assert not has_control_characters("café 😀")


# ── Download as content: the bytes, unaltered ────────────────────────────────


@pytest.mark.parametrize(
    ("content", "charset"),
    [(b"hello +2AA- world", "utf-7"), (b"hello \\ud800 world", "unicode_escape")],
)
async def test_a_text_attachment_that_decodes_to_a_lone_surrogate_arrives_as_its_bytes(
    tmp_path, content, charset
):
    client, _ = _mail_client(
        _graph_attachment(content, "note.txt", f"text/plain; charset={charset}")
    )

    blocks = await download_attachment(
        client,
        message_id=MESSAGE_ID,
        attachment_id="att1",
        as_content=True,
        config=Config(attachments_dir=str(tmp_path / "att")),
    )

    assert [b.type for b in blocks] == ["text", "resource"]
    assert base64.b64decode(blocks[1].resource.blob) == content
    CallToolResult(content=blocks).model_dump_json()


# ── The backstop every tool passes through ───────────────────────────────────


async def test_the_wrapper_replaces_a_lone_surrogate_anywhere_in_a_result():
    @_wrap_tool_errors
    async def tool():
        return {
            "messages": [{"subject": "hi \ud800", "tags": ("ok", "\udc00")}],
            "count": 1,
            "\ud800key": None,
        }

    result = await tool()

    _encodes(result)
    assert result["messages"][0]["subject"] == "hi �"
    assert result["messages"][0]["tags"] == ("ok", "�")
    assert "�key" in result


async def test_the_wrapper_reaches_into_content_blocks():
    @_wrap_tool_errors
    async def tool():
        resource = TextResourceContents(
            uri="attachment:///x", mime_type="text/plain", text="b\udfff"
        )
        return [
            TextContent(type="text", text="a\ud800"),
            EmbeddedResource(type="resource", resource=resource),
        ]

    blocks = await tool()

    assert blocks[0].text == "a�"
    assert blocks[1].resource.text == "b�"
    CallToolResult(content=blocks).model_dump_json()


async def test_a_clean_result_comes_back_as_the_same_object():
    payload = {"messages": [{"subject": "café 😀"}], "blocks": [TextContent(type="text", text="x")]}

    @_wrap_tool_errors
    async def tool():
        return payload

    assert await tool() is payload


def test_error_text_never_carries_a_lone_surrogate():
    error = OutlookMCPError("bad_input", "Not a file: \ud800", "Pass \udc00 instead.")

    assert str(error) == "Not a file: � Pass � instead."


async def test_a_value_error_with_a_lone_surrogate_reaches_the_model_cleaned():
    @_wrap_tool_errors
    async def tool():
        raise ValueError("bad name \ud800")

    with pytest.raises(ToolInputError) as caught:
        await tool()

    assert str(caught.value) == "bad name �"


# ── End to end, over real stdio ──────────────────────────────────────────────

# Written to a file and run as the server. Raw, so the `\ud800` escapes stay text here
# and only become surrogates inside the server process.
_STUB_SERVER = r"""
import base64, json, sys
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory
from msgraph.generated.models.attachment import Attachment

from outlook_mcp import server
from outlook_mcp.config import Config

content = b"hello +2AA- world"
doc = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    "id": "att1",
    "name": "note\ud800.txt",
    "contentType": "text/plain; charset=utf-7",
    "size": len(content) + 120,
    "contentBytes": base64.b64encode(content).decode(),
}
node = JsonParseNodeFactory().get_root_parse_node("application/json", json.dumps(doc).encode())
client = MagicMock()
attachments = client.me.messages.by_message_id.return_value.attachments
found = MagicMock()
found.get = AsyncMock(return_value=node.get_object_value(Attachment))
gone = MagicMock()
gone.get = AsyncMock(side_effect=ValueError("Attachment \ud800 is gone."))
attachments.by_attachment_id.side_effect = lambda attachment_id: (
    found if attachment_id == "att1" else gone
)


class Refusal:
    def __str__(self):
        return "Graph refused \ud800 the refresh."


server._get_graph_client = lambda ctx: SimpleNamespace(sdk_client=client)
server._get_config = lambda ctx: Config(client_id="test", attachments_dir=sys.argv[1])
server._get_auth = lambda ctx: SimpleNamespace(
    is_authenticated=lambda: False,
    config=SimpleNamespace(read_only=False),
    startup_error=Refusal(),
)
server.mcp.run("stdio")
"""


async def test_the_stdio_server_answers_and_stays_up(tmp_path):
    """Before #186, the first call got no answer and the second found the server gone."""
    stub = tmp_path / "stub_server.py"
    stub.write_text(_STUB_SERVER, encoding="utf-8")
    params = StdioServerParameters(
        command=sys.executable,
        args=[str(stub), str(tmp_path / "att")],
        env={
            **os.environ,
            "HOME": str(tmp_path),
            "OUTLOOK_MCP_CONFIG_DIR": str(tmp_path / "settings"),
        },
        cwd=os.getcwd(),
    )

    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        await asyncio.wait_for(session.initialize(), 30)

        download = await asyncio.wait_for(
            session.call_tool(
                "outlook_download_attachment",
                {"message_id": MESSAGE_ID, "attachment_id": "att1", "as_content": True},
            ),
            30,
        )
        assert [b.type for b in download.content] == ["text", "resource"]
        assert "note" + chr(0xFFFD) + ".txt" in download.content[0].text
        assert base64.b64decode(download.content[1].resource.blob) == b"hello +2AA- world"

        refused = await asyncio.wait_for(
            session.call_tool(
                "outlook_download_attachment",
                {"message_id": MESSAGE_ID, "attachment_id": "att2", "as_content": True},
            ),
            30,
        )
        assert refused.is_error
        assert "Attachment " + chr(0xFFFD) + " is gone." in refused.content[0].text

        status = await asyncio.wait_for(session.call_tool("outlook_auth_status", {}), 30)
        assert "Graph refused " + chr(0xFFFD) + " the refresh." in status.content[0].text

        tools = await asyncio.wait_for(session.list_tools(), 30)
        assert tools.tools
