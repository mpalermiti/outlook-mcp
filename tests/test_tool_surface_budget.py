"""The tool surface is a per-turn cost, so it gets a ceiling.

Every client pays for all 62 tool schemas on every turn that includes them.
ROADMAP records a measured baseline of ~8,644 tokens (o200k proxy, 2026-07), and
that number is what makes the `OUTLOOK_MCP_TOOLSETS` gating worth having — but
nothing stopped it drifting upward one well-meant docstring at a time.

This is a budget, not a golden file. A golden snapshot of the schemas would fail
on every wording change and get regenerated without being read, which teaches
nobody anything. A ceiling only fires when the surface gets materially more
expensive, and that is the moment worth a conversation: adding a tool, or adding
a field to all 62.

Raising the ceiling is a legitimate outcome — a new tool has to go somewhere. It
just has to be deliberate, with the new number written down next to the reason.
"""

import json

import pytest
from mcp.client import Client

from outlook_mcp.server import mcp

# A chars/4 proxy. This is NOT the same measure as the ~8,644 o200k figure in
# ROADMAP.md — that one ran a real tokenizer, this one divides. Do not compare
# the two numbers or "reconcile" them; this file only ever compares itself to
# itself, which is all a drift guard needs.
CHARS_PER_TOKEN = 4

# Ceiling for the full 62-tool surface. Measured at 11,206 on 1.20.0, with ~7%
# headroom: enough for a docstring fix or another annotation, not enough for a
# field added to all 62 or a batch of new tools.
TOOL_SURFACE_CEILING = 12_000

# `prompts/list` is the cheap half of the bargain struck in 1.20.0: workflow
# guidance costs a name and one line here until someone invokes it. If that ever
# stops being true, the guidance has moved back into a place that is paid for
# on every turn.
PROMPT_SURFACE_CEILING = 200


def _proxy_tokens(payload: object) -> int:
    return len(json.dumps(payload, default=str)) // CHARS_PER_TOKEN


@pytest.mark.asyncio
async def test_the_tool_surface_stays_within_budget():
    async with Client(mcp) as client:
        tools = (await client.list_tools()).tools

    cost = _proxy_tokens([t.model_dump(exclude_none=True, by_alias=True) for t in tools])

    assert cost <= TOOL_SURFACE_CEILING, (
        f"The 62-tool surface now costs ~{cost} proxy tokens per turn, over the "
        f"{TOOL_SURFACE_CEILING} ceiling. Every client pays this on every turn. "
        f"Either trim it, or raise the ceiling deliberately and say why."
    )


@pytest.mark.asyncio
async def test_prompt_listing_stays_cheap():
    """Prompts earn their place by being nearly free until invoked."""
    async with Client(mcp) as client:
        prompts = (await client.list_prompts()).prompts

    cost = _proxy_tokens([p.model_dump(exclude_none=True, by_alias=True) for p in prompts])

    assert cost <= PROMPT_SURFACE_CEILING, (
        f"`prompts/list` now costs ~{cost} proxy tokens. A prompt's body is supposed "
        f"to live behind `prompts/get`, not in the listing."
    )


@pytest.mark.asyncio
async def test_gating_actually_reduces_the_surface():
    """The lever ROADMAP's baseline justifies has to still work."""
    from outlook_mcp import toolsets

    async with Client(mcp) as client:
        full = (await client.list_tools()).tools

    mail_only = [t for t in full if toolsets.TOOL_GROUPS.get(t.name) in {"mail", "account"}]

    assert _proxy_tokens(
        [t.model_dump(exclude_none=True, by_alias=True) for t in mail_only]
    ) < _proxy_tokens([t.model_dump(exclude_none=True, by_alias=True) for t in full]) // 2
