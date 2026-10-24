"""The tool surface is a per-turn cost, so it gets a ceiling.

Every client pays for all 72 tool schemas on every turn that includes them.
ROADMAP records a measured baseline of ~8,644 tokens (o200k proxy, 2026-07), and
that number is what makes the `OUTLOOK_MCP_TOOLSETS` gating worth having — but
nothing stopped it drifting upward one well-meant docstring at a time.

This is a budget, not a golden file. A golden snapshot of the schemas would fail
on every wording change and be regenerated without being read, which teaches
nobody anything. A ceiling only fires when the surface gets materially more
expensive, and that is the moment worth a conversation: adding a tool, or adding
a field to all 68.

Raising the ceiling is a legitimate outcome — a new tool has to go somewhere. It
just has to be deliberate, with the new number written down next to the reason.
"""

import inspect
import json

import pytest
from mcp.client import Client

from outlook_mcp.server import mcp

# A chars/4 proxy. This is NOT the same measure as the ~8,644 o200k figure in
# ROADMAP.md — that one ran a real tokenizer, this one divides. Do not compare
# the two numbers or "reconcile" them; this file only ever compares itself to
# itself, which is all a drift guard needs.
CHARS_PER_TOKEN = 4

# Ceiling for the full tool surface. Measured at 11,206 on 1.20.0 with a
# 12,000 ceiling; raised to 13,500 for the eight To Do detail tools
# (get_task, checklist-item CRUD, task attachments), then to 13,650 after
# the attachment rewrite (measured 12,855 at the time) — each a deliberate
# "a new tool has to go somewhere" raise, not drift. Re-measured 13,154
# after the rebase onto main picked up #76's time-zone anchoring (which
# grew create/get/update_event, not just `outlook_create_event`'s new
# parameter) and the upload-session docstring note.
#
# 13,156 on the 68-tool surface: the two in-process account-switching tools
# came off (barely ~20 tokens between them — two of the smallest schemas on
# the surface), and the attachment-path docstrings stopped naming a default
# an override moves anyway (+22). Ceiling 13,300 holds ~1% headroom: room
# for a wording fix, not for another batch.
#
# 13,245 with `outlook_update_event`'s `timezone` argument and the sentence that
# says it re-anchors (+55), and one sentence saying a series with edited or
# deleted occurrences will not be reshaped (+34) — silent data loss, so it goes
# where the model reads it. The other refusals carry their remedies in the error
# text rather than in the schema every turn pays for.
#
# 13,004 on every Python once the descriptions are dedented (-239). Every number
# above was taken on an interpreter that keeps docstring indentation (3.10-3.12);
# the same 1.25.1 surface measured 13,243 there and 13,024 on 3.13, which strips
# it. Compare numbers from here on only with each other. Ceiling unchanged.
#
# 13,203 with the contact fixes, on the dedented measure: `job_title` on create
# and update (#119), then `phone_type` on both and `title` on update (#120,
# +101), then update's docstring saying `title` is the name prefix, as create's
# does (+23) — undescribed, models fill it with the job (#119). No raise.
#
# 13,261 with `clear_phones` on `outlook_update_contact` and its one-line example
# (+58): without it, moving a number from mobile to business left a copy in
# mobile (#120 follow-up). No raise; ~39 tokens of headroom left.
#
# 13,298 with `calendar` on `outlook_create_event` (#141, +37): the argument and
# one short line; that it takes a display name is in INSTRUCTIONS already, which
# every session gets once. No raise; 2 tokens of headroom left.
#
# 13,308 once #146 (`outlook_read_message` takes a Message-ID, +10) and #147
# (`calendar` on create, +37) landed together on 2026-10-23: each fit on its own,
# the pair did not. Raised to 13,500 at merge time — a deliberate step for the
# contributor batch reviewed that day, several of which carry their own raise to
# the same number. The ~1% convention resumes from whatever the batch measures.
#
# 13,326 with `birthday` on `outlook_create_contact` (+31) and
# `outlook_update_contact` (+34): the argument on each, and the words saying it is a
# date, YYYY-MM-DD. `outlook_get_contact` already returned it; nothing could write
# it. Two arguments do not fit the ~39 tokens left, so this is a deliberate raise to
# 13,500: ~1% headroom, as before.
#
# 13,397 with `company` on `outlook_update_contact` (+25 on today's main, 13,372):
# create set it and update could not, so a contact who changed employer kept the
# old one (#173). The argument alone, no new words. No raise.
#
# 13,586 with `middle_name`, `personal_notes` and `categories` on
# `outlook_create_contact` and `outlook_update_contact` (+189, on 13,397 with #174)
# and update's sentence that `categories` replaces the whole list ([] clears it):
# without it, adding one category with a one-item list drops the rest, silently.
# `outlook_get_contact` returned categories and notes that nothing could write
# (#175). Six arguments do not fit the ~100 tokens left, so this is a deliberate
# raise to 13,700: ~0.8% headroom.
#
# 14,805 on the 72-tool surface, measured after the rebase onto that 13,586 main
# and two review rounds: `outlook_list_rules` returns conditions and actions (+25 in
# the tool docstring, which is what the model reads), and update/delete document the
# ordering gate that starts a hidden forwarding rule (+71) — a gate the model cannot
# respect without being told it exists:
# the four inbox-rule schemas (+1,123, ~281 each), of which `outlook_create_rule`
# is most of it — 14 optional arguments, the price of exposing Graph's whole
# predicate/action set in one call. Review asked for a sentence in both rule
# docstrings saying a rule outlives the session, which is worth paying: a rule is
# the write most likely to be induced by the mail it acts on. Deliberate raise to
# 15,000: ~1.3% headroom, room for a wording fix, not for another batch. The rule
# tools are their own `rules` toolset group, so a client that gates down to `mail`
# does not pay this.
TOOL_SURFACE_CEILING = 15_000

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
        f"The 72-tool surface now costs ~{cost} proxy tokens per turn, over the "
        f"{TOOL_SURFACE_CEILING} ceiling. Every client pays this on every turn. "
        f"Either trim it, or raise the ceiling deliberately and say why."
    )


@pytest.mark.asyncio
async def test_descriptions_carry_no_source_indentation():
    """The surface above is the same surface on every Python.

    Python 3.13 strips a docstring's indentation when it compiles it; 3.10-3.12
    keep it, and the SDK publishes ``__doc__`` as it finds it. So on the older
    versions every client paid for this file's indentation on every turn (~240
    proxy tokens on 1.25.1), and the budget measured a different surface
    depending on which Python ran it: #122 measured 13,224 on 3.13 and 13,448 on
    3.12 against a 13,300 ceiling.
    """
    async with Client(mcp) as client:
        tools = (await client.list_tools()).tools

    indented = [t.name for t in tools if t.description != inspect.cleandoc(t.description or "")]

    assert indented == []


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

    assert (
        _proxy_tokens([t.model_dump(exclude_none=True, by_alias=True) for t in mail_only])
        < _proxy_tokens([t.model_dump(exclude_none=True, by_alias=True) for t in full]) // 2
    )

    # The rule tools are their own group: the most expensive feature here must be
    # opt-in, so a `mail`-only client does not silently pay for it.
    assert not [t for t in mail_only if t.name.endswith("_rule") or t.name == "outlook_list_rules"]
    assert {n for n, g in toolsets.TOOL_GROUPS.items() if g == "rules"} == {
        "outlook_list_rules",
        "outlook_create_rule",
        "outlook_update_rule",
        "outlook_delete_rule",
    }
