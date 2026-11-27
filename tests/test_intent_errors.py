"""The intent tools' new refusal (contracts §11): ``NoMatchError``.

An ambiguous name is a result (``candidates``), not an error. A reference that
matches nothing is a refusal for every caller except ``find``, and its text has
to reach the model whole: what was searched for, where, and what to pass instead.
"""

import pytest
from mcp.client import Client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mcp.errors import NoMatchError, OutlookMCPError
from outlook_mcp.server import _wrap_tool_errors


def test_it_is_an_anticipated_failure_with_its_code():
    error = NoMatchError("person", "Sam Jones", ["sent mail", "contacts"])
    assert isinstance(error, OutlookMCPError)
    assert isinstance(error, ToolError)
    assert error.code == "no_match"


def test_the_text_names_the_reference_the_sources_and_the_way_out():
    text = str(NoMatchError("person", "Sam Jones", ["sent mail", "contacts"]))
    assert "Nothing matched person 'Sam Jones' (searched: sent mail, contacts)." in text
    assert text.endswith("Pass an exact address or id, or different words.")


def test_the_words_are_sanitized_and_cut():
    text = str(NoMatchError("message", "a\x1b[31mred\nline" + "x" * 300, ["mail"]))
    assert "\x1b" not in text
    assert "\n" not in text
    assert len(text) < 260


@pytest.mark.asyncio
async def test_the_text_reaches_the_client():
    server = MCPServer("no-match-guard", version="0")

    @server.tool()
    @_wrap_tool_errors
    async def failing_tool() -> dict:
        raise NoMatchError("event", "standup", ["your calendars"])

    async with Client(server) as client:
        result = await client.call_tool("failing_tool", {})

    assert result.is_error is True
    text = result.content[0].text
    assert "Nothing matched event 'standup'" in text
    assert "Pass an exact address or id, or different words." in text
