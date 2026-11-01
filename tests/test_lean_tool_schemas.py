"""The advertised tool schemas carry nothing the model cannot use.

Every client pays for every tool schema on every turn. Pydantic writes two kinds
of filler into them: a ``title`` on every input that restates its name
(``"title": "Unread Only"`` on ``unread_only``) plus one per tool
(``"outlook_list_inboxArguments"``), and, for each optional input, a ``null``
branch and ``"default": null`` that say nothing omitting the input does not.
Together they were a quarter of the default surface.

What the server *accepts* is not decided by this schema: the SDK validates
arguments against the function's own Pydantic model, so an explicit ``null``
still works after the schema stops advertising it. These tests pin both halves.
"""

import asyncio

import pytest
from mcp.client import Client

from outlook_mcp import toolsets
from outlook_mcp.server import mcp

# Keywords whose value is a schema, or a map / list of schemas, in the subset of
# JSON Schema Pydantic emits. Values under any other keyword (`default`, `enum`,
# `const`, `examples`) are data, so a `title` key inside one is not filler.
_SCHEMA = ("items", "additionalProperties", "not")
_SCHEMA_LISTS = ("anyOf", "oneOf", "allOf", "prefixItems")
_SCHEMA_MAPS = ("properties", "$defs", "definitions")


def _schema_nodes(node):
    """Every schema node in ``node``, walked by keyword rather than by key name.

    A key name alone cannot tell the `title` keyword from an input called
    `title` (contacts and tasks have one); the keyword that holds it can.
    """
    if not isinstance(node, dict):
        return
    yield node
    for key in _SCHEMA:
        yield from _schema_nodes(node.get(key))
    for key in _SCHEMA_LISTS:
        for child in node.get(key) or []:
            yield from _schema_nodes(child)
    for key in _SCHEMA_MAPS:
        for child in (node.get(key) or {}).values():
            yield from _schema_nodes(child)


def _advertised():
    async def main():
        async with Client(mcp) as client:
            return {t.name: t.input_schema for t in (await client.list_tools()).tools}

    return asyncio.run(main())


def test_no_advertised_schema_carries_a_generated_title():
    titled = sorted(
        name
        for name, schema in _advertised().items()
        if any("title" in node for node in _schema_nodes(schema))
    )
    assert not titled, f"generated `title` labels still advertised by: {titled}"


def test_no_optional_input_advertises_a_null_branch_or_a_null_default():
    nullable = sorted(
        f"{name}.{prop}"
        for name, schema in _advertised().items()
        for prop, node in schema.get("properties", {}).items()
        if {"type": "null"} in (node.get("anyOf") or [])
        or ("default" in node and node["default"] is None)
    )
    assert not nullable, f"optional inputs still advertising null: {nullable}"


def test_inputs_called_title_are_still_advertised():
    schemas = _advertised()
    task, contact = schemas["outlook_create_task"], schemas["outlook_create_contact"]
    assert task["properties"]["title"] == {"type": "string"}
    assert "title" in task["required"]
    assert contact["properties"]["title"] == {"type": "string"}
    assert "title" not in contact.get("required", [])


def test_an_optional_input_keeps_every_type_except_null():
    props = _advertised()["outlook_create_event"]["properties"]
    assert props["recurrence"] == {
        "anyOf": [{"additionalProperties": True, "type": "object"}, {"type": "string"}]
    }
    assert props["attendees"] == {"items": {"type": "string"}, "type": "array"}


def test_a_non_null_default_survives():
    props = _advertised()["outlook_list_inbox"]["properties"]
    assert props["count"] == {"default": 25, "type": "integer"}
    assert props["folder"] == {"default": "inbox", "type": "string"}


def test_an_explicit_null_is_still_accepted():
    """The trim is advertising only: the SDK validates with the function's model."""
    arg_model = mcp._tool_manager._tools["outlook_list_inbox"].fn_metadata.arg_model
    parsed = arg_model.model_validate({"from_address": None, "after": None})
    assert parsed.from_address is None and parsed.after is None


# ── lean_input_schema, on schemas the server does not happen to have ──────


def test_a_null_with_a_meaning_is_kept():
    """`None` as a default is what makes null redundant. Elsewhere it may mean something."""
    schema = {
        "type": "object",
        "properties": {
            "reminder": {"anyOf": [{"type": "integer"}, {"type": "null"}], "default": 15},
        },
    }
    assert toolsets.lean_input_schema(schema)["properties"]["reminder"] == {
        "anyOf": [{"type": "integer"}, {"type": "null"}],
        "default": 15,
    }


def test_nested_titles_go_and_data_stays():
    schema = {
        "type": "object",
        "title": "fArguments",
        "properties": {
            "title": {"title": "Title", "type": "string"},
            "slots": {
                "title": "Slots",
                "type": "array",
                "items": {"title": "Slot", "anyOf": [{"title": "A", "type": "string"}]},
                "default": [{"title": "kept, this is data"}],
            },
        },
        "required": ["title"],
    }
    assert toolsets.lean_input_schema(schema) == {
        "type": "object",
        "properties": {
            "title": {"type": "string"},
            "slots": {
                "type": "array",
                "items": {"anyOf": [{"type": "string"}]},
                "default": [{"title": "kept, this is data"}],
            },
        },
        "required": ["title"],
    }


def test_the_input_schema_is_not_mutated():
    schema = {"type": "object", "title": "fArguments", "properties": {}}
    toolsets.lean_input_schema(schema)
    assert schema["title"] == "fArguments"


@pytest.mark.parametrize("keyword", ["description", "format"])
def test_other_keywords_on_a_collapsed_input_are_kept(keyword):
    schema = {
        "type": "object",
        "properties": {
            "when": {
                "anyOf": [{"type": "string"}, {"type": "null"}],
                "default": None,
                keyword: "x",
            },
        },
    }
    assert toolsets.lean_input_schema(schema)["properties"]["when"] == {
        "type": "string",
        keyword: "x",
    }
