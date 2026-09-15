"""Prompts carry the workflows; docstrings carry the tools.

Sequencing guidance — which tools to call, in what order, carrying what forward —
is not per-tool knowledge, so paying for it in 70 docstrings on every turn is the
wrong shape. A prompt costs a name and one line in `prompts/list` until someone
invokes it, and unlike a SKILL.md it reaches every MCP client rather than only
the ones that load skills.

The three here are the workflows the trajectory archive actually shows being
assembled by hand, call by call.
"""

import re

import pytest
from mcp.client import Client

from outlook_mcp.server import mcp

EXPECTED_PROMPTS = {"morning_brief", "triage_folder", "catch_up"}


@pytest.mark.asyncio
async def test_the_workflow_prompts_are_registered():
    async with Client(mcp) as client:
        names = {p.name for p in (await client.list_prompts()).prompts}
    assert EXPECTED_PROMPTS <= names


@pytest.mark.asyncio
async def test_every_prompt_has_a_description():
    """`prompts/list` is a menu; an entry with no description is unusable."""
    async with Client(mcp) as client:
        prompts = (await client.list_prompts()).prompts
    for prompt in prompts:
        assert prompt.description, f"{prompt.name} has no description"


@pytest.mark.asyncio
async def test_morning_brief_names_the_tools_it_wants_used():
    """The point of the prompt is the sequence, so the sequence has to be in it."""
    async with Client(mcp) as client:
        result = await client.get_prompt("morning_brief", {})
    text = " ".join(m.content.text for m in result.messages)

    assert "outlook_list_events" in text
    assert "outlook_list_inbox" in text
    assert "concise" in text


@pytest.mark.asyncio
async def test_triage_folder_passes_the_folder_through():
    async with Client(mcp) as client:
        result = await client.get_prompt("triage_folder", {"folder": "Junk Email"})
    text = " ".join(m.content.text for m in result.messages)

    assert "Junk Email" in text


@pytest.mark.asyncio
async def test_triage_folder_defaults_to_the_inbox():
    async with Client(mcp) as client:
        result = await client.get_prompt("triage_folder", {})
    text = " ".join(m.content.text for m in result.messages)

    assert "inbox" in text.lower()


@pytest.mark.asyncio
async def test_catch_up_steers_to_the_delta_path():
    """The composed tools exist and go unused; this is where they get named."""
    async with Client(mcp) as client:
        result = await client.get_prompt("catch_up", {"since": "24h"})
    text = " ".join(m.content.text for m in result.messages)

    assert "outlook_changes_since" in text
    assert "24h" in text


@pytest.mark.asyncio
async def test_prompts_do_not_invent_tool_names():
    """A prompt that names a tool we don't have sends the agent hunting."""
    async with Client(mcp) as client:
        tool_names = {t.name for t in (await client.list_tools()).tools}
        prompts = (await client.list_prompts()).prompts

        for prompt in prompts:
            args = {a.name: "inbox" for a in (prompt.arguments or []) if a.required}
            result = await client.get_prompt(prompt.name, args)
            text = " ".join(m.content.text for m in result.messages)

            for name in re.findall(r"outlook_[a-z_]+", text):
                assert name in tool_names, f"{prompt.name} names unknown tool {name}"
