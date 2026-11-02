"""Contact field meanings survive MCP discovery and agree with partial writes."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest
from mcp.client import Client

from outlook_mcp import server
from outlook_mcp.config import Config
from tests.test_write_payloads_reach_the_wire import wire


async def test_contact_update_meanings_are_in_the_tool_description():
    async with Client(server.mcp) as client:
        tools = {tool.name: tool for tool in (await client.list_tools()).tools}

    tool = tools["outlook_update_contact"]
    assert "job_title is the job" in tool.description
    assert "company the employer" in tool.description
    assert "name prefix" in tool.description
    assert tool.annotations.read_only_hint is False
    assert tool.annotations.destructive_hint is False
    for field in ("job_title", "company", "title"):
        assert tool.input_schema["properties"][field] == {"type": "string"}
        assert field not in tool.input_schema["required"]


async def test_update_description_explains_omission_and_clearing():
    async with Client(server.mcp) as client:
        tools = {tool.name: tool for tool in (await client.list_tools()).tools}

    description = tools["outlook_update_contact"].description
    assert "only provided fields change" in description
    assert 'A text field passed as "" is cleared' in description


@pytest.mark.parametrize("clear", [False, True], ids=["set", "clear"])
@pytest.mark.parametrize(
    ("field", "graph_field", "value"),
    [
        ("first_name", "givenName", "Alex"),
        ("last_name", "surname", "Example"),
        ("middle_name", "middleName", "Morgan"),
        ("personal_notes", "personalNotes", "Synthetic note"),
        ("job_title", "jobTitle", "Engineer"),
        ("company", "companyName", "Example Company"),
        ("title", "title", "Dr."),
    ],
)
async def test_a_single_contact_field_reaches_graph_without_other_changes(
    monkeypatch, field, graph_field, value, clear
):
    value = "" if clear else value
    sdk = MagicMock()
    builder = sdk.me.contacts.by_contact_id.return_value
    builder.patch = AsyncMock()
    monkeypatch.setattr(server, "_get_graph_client", lambda _ctx: MagicMock(sdk_client=sdk))
    monkeypatch.setattr(server, "_get_config", lambda _ctx: Config(client_id="test"))

    async with Client(server.mcp) as client:
        result = await client.call_tool(
            "outlook_update_contact", {"contact_id": "contact123", field: value}
        )

    assert not result.is_error
    sdk.me.contacts.by_contact_id.assert_called_once_with("contact123")
    builder.patch.assert_awaited_once()
    payload = json.loads(wire(builder.patch.call_args.args[0]))
    payload.pop("@odata.type", None)
    assert payload == {graph_field: value}
