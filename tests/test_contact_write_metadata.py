"""Contact field meanings survive MCP discovery and agree with partial writes."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest
from mcp.client import Client

from outlook_mcp import server
from outlook_mcp.config import Config
from tests.test_write_payloads_reach_the_wire import wire


@pytest.mark.parametrize("name", ["outlook_create_contact", "outlook_update_contact"])
async def test_contact_field_meanings_are_in_the_advertised_schema(name):
    async with Client(server.mcp) as client:
        tools = {tool.name: tool for tool in (await client.list_tools()).tools}

    tool = tools[name]
    properties = tool.input_schema["properties"]
    for field, meaning in (
        ("job_title", "employment title"),
        ("company", "employer"),
        ("title", "name prefix"),
    ):
        assert properties[field]["type"] == "string"
        assert meaning in properties[field].get("description", "").lower()
        assert field not in tool.input_schema["required"]
    assert "job_title" in tool.description
    assert "company" in tool.description
    assert "name prefix" in tool.description
    assert tool.annotations.read_only_hint is False
    assert tool.annotations.destructive_hint is False


async def test_update_discovery_explains_omission_and_clearing():
    async with Client(server.mcp) as client:
        tools = {tool.name: tool for tool in (await client.list_tools()).tools}

    tool = tools["outlook_update_contact"]
    for field in ("job_title", "company", "title"):
        description = tool.input_schema["properties"][field].get("description", "")
        assert "omit" in description.lower()
        assert '""' in description and "clear" in description.lower()
    assert "only provided fields change" in tool.description


@pytest.mark.parametrize(
    ("field", "graph_field", "value"),
    [
        ("job_title", "jobTitle", "Engineer"),
        ("company", "companyName", "Example Company"),
        ("title", "title", "Dr."),
        ("job_title", "jobTitle", ""),
        ("company", "companyName", ""),
        ("title", "title", ""),
    ],
)
async def test_a_single_contact_field_reaches_graph_without_other_changes(
    monkeypatch, field, graph_field, value
):
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
