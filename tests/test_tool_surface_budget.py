"""The tool surface is a per-turn cost, so it gets a ceiling.

Every client pays for all 71 tool schemas on every turn that includes them.
ROADMAP records a measured baseline of ~8,644 tokens (o200k proxy, 2026-07), and
that number is what makes the `OUTLOOK_MCP_TOOLSETS` gating worth having — but
nothing stopped it drifting upward one well-meant docstring at a time.

This is a budget, not a golden file. A golden snapshot of the schemas would fail
on every wording change and be regenerated without being read, which teaches
nobody anything. A ceiling only fires when the surface gets materially more
expensive, and that is the moment worth a conversation: adding a tool, or adding
a field to all 71.

Raising the ceiling is a legitimate outcome — a new tool has to go somewhere. It
just has to be deliberate, with the new number written down next to the reason.
"""

import inspect
import json
import os
import subprocess
import sys

import pytest
from mcp.client import Client

from outlook_mcp import toolsets
from outlook_mcp.server import mcp

# A chars/4 proxy. This is NOT the same measure as the ~8,644 o200k figure in
# ROADMAP.md — that one ran a real tokenizer, this one divides. Do not compare
# the two numbers or "reconcile" them; this file only ever compares itself to
# itself, which is all a drift guard needs.
CHARS_PER_TOKEN = 4

# Ceiling for the full tool surface. Measured at 11,206 on 1.20.0 with a
#
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
# 13,724 with reply and forward drafts on `outlook_create_draft` (#143, +138 on
# 13,586): `reply_to_message_id`, `reply_all` and `forward_message_id` (+80), `to`
# and `subject` made optional because Graph fills a reply's (+20), and one docstring
# line saying so (+38). A new tool for it would have cost a whole second schema.
# Raised to 13,900 as one ceiling for this and #168 (`is_markdown`, stacked on
# this), which together measure ~13,820: ~0.6% headroom over the pair.
#
# 13,821 with `is_markdown` on the five tools that write a mail body (#168, +97 on
# 13,724 with #143): `outlook_send_message`, `outlook_reply`, `outlook_forward`,
# `outlook_create_draft` and `outlook_update_draft`, about 20 each for the boolean's
# schema. No docstring grew: the supported subset is one INSTRUCTIONS bullet, sent
# once per session. One ceiling for the pair, the 13,900 #143 set: reply and forward
# drafts (+138) and Markdown bodies (+97) together, ~0.6% headroom.
#
# 13,676 with add_emails and remove_emails on outlook_update_contact (#161):
# two list arguments and the line distinguishing them from email, which replaces
# the whole list (+90 on main at 13,586). The 13,700 ceiling still holds.
#
# 13,911 once #161 (`add_emails` and `remove_emails` on `outlook_update_contact`, +90)
# merged onto the mail stack (13,821, #131 #167 #143 #168): each fit its own base,
# together they do not. Raised at merge time to 14,100 as one ceiling for the contacts
# pair (#161, then #165 stacked on it, ~13,964 together): ~1% headroom.
#
# 13,729 with phone on outlook_search_contacts (#165, +53 on #162 at 13,676):
# the optional query argument and phone matching description. The contacts stack
# includes birthday, company, categories, notes and middle name now on main;
# its combined surface needs a 13,900 ceiling (~1% headroom), rather than the
# 13,600 discussed before those fields landed.
#
# 13,605 with attachments as content (#129, +233 on today's main, 13,372):
# `as_content` and an optional `save_path` on the two downloads, inline
# `attachments` and an optional `attachment_paths` on the two mail uploads (+159 in
# the schemas), and one short sentence each (+74). Four new arguments do not fit the
# ~128 tokens left, so this is a deliberate raise to 13,750: ~1% headroom.
#
# 13,819 after rebasing #129 onto main with #174 and #176 (+233 on 13,586).
# The added contact fields lifted main's ceiling to 13,700; this independent
# attachment feature needs a deliberate 14,000 ceiling (~1.3% headroom).
#
# 14,197 once #129 (downloads returned as content, inline uploads, +233) merged onto
# the contacts pair (13,964): over both the 14,100 set for the pair and the 14,000
# #129 set on its own base. Raised at merge time to 14,350: ~1% headroom.
#
# 13,594 with folder paths (#144, +8 on 13,586): `outlook_list_folders` says each
# folder carries the `path` folder arguments accept. The path rule itself went into
# INSTRUCTIONS, which every tool that takes a folder shares and which is sent once
# per session, not into five docstrings. No raise.
#
# 14,204 with folder paths after rebasing onto main with the mail stack,
# contacts pair and attachments as content; the 14,350 ceiling still holds.
#
# 13,802 on the 69-tool surface with event keys (#148, +208 on 13,594 with #144),
# raised to 13,900: `outlook_find_event_by_key` is +137 on its own, nearly all of
# it the schema every tool pays for; `key` on `outlook_create_event` (+39, the
# argument and one line), `outlook_update_event` (+30) and the word on
# `outlook_get_event` (+1). A `key` filter on `outlook_list_events` would have cost
# ~30, but that listing reads a seven-day window by default and returns
# occurrences, so the lookup before a create would miss a booking three weeks out
# and hand back occurrence IDs for a series. The guidance on using keys is in
# INSTRUCTIONS, paid once per session. ~0.7% headroom.
#
# 14,412 with event keys on the rebased folder-path stack. This exceeds the
# 14,350 base ceiling; 14,600 gives about 1.3% headroom.
#
# 14,091 on the 71-tool surface with event attachments (#153, +289 on 13,802 with
# #144 and #148), raised to 14,200: `outlook_attach_to_event` (+163) and
# `outlook_remove_event_attachment` (+117), most of each the schema every tool pays
# for, and `outlook_get_event` naming the attachment list it now returns (+6).
# Removal is its own tool, as for drafts and To Do; the listing rides on
# `outlook_get_event` rather than a third tool, at the cost of a second request
# only for an event that has files. The last of the three, so this is their
# combined number. ~0.8% headroom.
#
# 14,701 with event attachments on the rebased folder-path/event-key stack.
# The 14,600 parent ceiling no longer covers it; 14,850 gives about 1% headroom.
#
# 13,662 with include_sender_auth on the two read tools (#130, +76 on main
# at 13,586). The existing 13,700 ceiling covers the optional authentication
# field; guidance is in INSTRUCTIONS, paid once per session.
#
# 14,272 with sender authentication on current main (14,197); the existing
# 14,350 ceiling covers the opt-in read without a raise.
# 13,594 after the digest describes has_more and preserving the whole token
# dictionary (#156, +8 on main at 13,586). The ignored-filter fallback and
# bootstrap context add response metadata, not schema fields. No ceiling raise.
#
# 14,205 with the digest bootstrap fix on current main (14,197); the
# existing 14,350 ceiling covers it without a raise.
#
# 14,210 after clarifying that stored token dictionaries must be replaced,
# not merged (+5); still covered by the 14,350 ceiling.
# 14,790 measured after rebasing the digest onto current main.
# The existing 14,850 ceiling covers it; no raise.
# 14,813 on current main with the HTTP follow-ups: transport checks and
# operator errors add no tools or schema fields. The same measure holds on
# Python 3.10 and 3.13; no ceiling raise.
# 14,846 with `outlook_get_event`'s optional `message_id` (#196, +33); #194's
# calendar owner fields are response data, not schema. 72 / 15,024 with every
# group. No raise, but the default ceiling now has 4 tokens of headroom.
#
# 11,130 once `toolsets.lean_input_schema` drops Pydantic's filler from every
# advertised schema (-3,716, a quarter): the `title` restating each input's name
# and each tool's (-2,199), and an optional input's `null` branch and
# `"default": null` (-1,517). No wording changed and nothing the server accepts
# changed. Same on Python 3.10 and 3.13. The ceiling comes DOWN to 11,250 (~1%
# headroom), so the freed quarter cannot be spent a sentence at a time.
#
# Re-measured on 1.29.0 with this branch's rules group in the tree: the default surface
# is 71 / 11,130 — upstream's own head measures 11,130 as well, so the four opt-in tools
# change nothing a stock install pays. With the `rules` group (every group except
# `documents`) the surface is 75 / 11,951, documents 72 / 11,274 and all groups 76 / 12,095.
# Measured with the same `by_alias=True` dump the tests use, on Python 3.10.
# The rules ceiling is 12,150 (~1%), because a surface nobody loads by default should
# not get to raise everyone's ceiling.
# 11,142 once `outlook_read_message` says `full` is the stored body, not "both" (+12).
# No raise.
TOOL_SURFACE_CEILING = 12_150
DEFAULT_SURFACE_CEILING = 11_250

# 14,776 on the 71-tool default surface before the latest rebase;
# 14,954 with the optional documents group enabled (+178 for its one reader).
# Re-measured 14,813 / 14,991 after rebasing onto main c832373 (+37 on both
# surfaces from the digest and reply-draft descriptions); the reader still adds 178.
# Re-measured 2026-10-25 after the third document review on Python 3.10 and 3.13:
# default 71 / 14,813; every group 72 / 14,991; documents-only 3 / 302. Both
# INSTRUCTIONS are separate, sent once per session: unchanged 3,176 characters by
# default; 3,379 with PDF-only chunk guidance enabled.
# 72 / 11,274 with the lean schemas (documents-only 3 / 249); the documents ceiling
# comes down to 11,400 with it, again ~1% headroom.
DOCUMENT_SURFACE_CEILING = 11_400

# `prompts/list` is the cheap half of the bargain struck in 1.20.0: workflow
# guidance costs a name and one line here until someone invokes it. If that ever
# stops being true, the guidance has moved back into a place that is paid for
# on every turn.
PROMPT_SURFACE_CEILING = 200


def _proxy_tokens(payload: object) -> int:
    return len(json.dumps(payload, default=str)) // CHARS_PER_TOKEN


@pytest.mark.asyncio
async def test_the_default_surface_stays_within_budget():
    """What a stock install pays on every turn — 71 tools, rules not loaded.

    Measured by dropping the opt-in group from whatever this process registered,
    so the number is the same whether or not the machine running the suite has
    `mailbox_settings_write` on.
    """
    async with Client(mcp) as client:
        tools = (await client.list_tools()).tools

    hidden = toolsets.OFF_BY_DEFAULT
    default = [t for t in tools if toolsets.TOOL_GROUPS.get(t.name) not in hidden]
    cost = _proxy_tokens([t.model_dump(exclude_none=True, by_alias=True) for t in default])

    assert len(default) == 71
    assert cost <= DEFAULT_SURFACE_CEILING, (
        f"The 71-tool default surface now costs ~{cost} proxy tokens per turn, over the "
        f"{DEFAULT_SURFACE_CEILING} ceiling. Every user pays this on every turn. "
        f"Either trim it, or raise the ceiling deliberately and say why."
    )


def test_the_rules_group_fits_its_own_ceiling():
    """The opt-in surface is bounded too, in a process that actually loads it.

    A separate interpreter because tool selection happens at import: this is the
    only way to measure both surfaces without the suite depending on its own
    environment. `mailbox_settings_write` is unset there, so the group is loaded
    the way an env opt-in loads it.
    """
    code = (
        "import asyncio, json\n"
        "from mcp.client import Client\n"
        "from outlook_mcp import toolsets\n"
        "from outlook_mcp.server import mcp\n"
        "async def main():\n"
        "    async with Client(mcp) as c:\n"
        "        tools = (await c.list_tools()).tools\n"
        "    loaded = {t.name for t in tools if toolsets.TOOL_GROUPS.get(t.name) == 'rules'}\n"
        "    payload = [t.model_dump(exclude_none=True, by_alias=True) for t in tools]\n"
        "    cost = len(json.dumps(payload, default=str)) // 4\n"
        "    print(json.dumps({'cost': cost, 'rules': len(loaded), 'count': len(tools)}))\n"
        "asyncio.run(main())\n"
    )
    # Default groups plus `rules`: this measures the rules surface an install
    # that asks for rules actually loads, so this ceiling and the documents
    # ceiling bound different surfaces instead of the same one twice.
    env = {
        **os.environ,
        "OUTLOOK_MCP_TOOLSETS": ",".join(
            sorted(set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT | {"rules"})
        ),
    }
    env.pop("PYTHONPATH", None)
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        env=env,
        timeout=120,
        check=True,
    )
    measured = json.loads(result.stdout.strip().splitlines()[-1])

    assert measured["rules"] == 4, result.stdout
    assert measured["count"] == 75, result.stdout
    assert measured["cost"] <= TOOL_SURFACE_CEILING, (
        f"The rules-enabled surface costs ~{measured['cost']} proxy tokens, over the "
        f"{TOOL_SURFACE_CEILING} ceiling. It is opt-in, so this is the budget for the "
        f"install that asked for it — trim it or raise the ceiling deliberately."
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


def test_documents_enabled_surface_has_an_explicit_budget():
    import os
    import subprocess
    import sys

    from outlook_mcp.toolsets import TOOL_GROUPS

    code = """import asyncio,json
from mcp.client import Client
from outlook_mcp.server import mcp
async def main():
    async with Client(mcp) as c:
        tools=(await c.list_tools()).tools
    payload=[t.model_dump(exclude_none=True,by_alias=True) for t in tools]
    print(json.dumps({"count":len(tools),"cost":len(json.dumps(payload,default=str))//4}))
asyncio.run(main())
"""
    result = subprocess.run(
        [sys.executable, "-c", code],
        env={
            **os.environ,
            # Default groups plus `documents`: the documents surface, 72 tools.
            # Loading rules too made this measure the same 76-tool surface as the
            # rules budget, so the documents reader could grow without this test
            # noticing.
            "OUTLOOK_MCP_TOOLSETS": ",".join(
                sorted(set(TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT | {"documents"})
            ),
        },
        capture_output=True,
        text=True,
        check=True,
    )
    measured = json.loads(result.stdout)
    assert measured["count"] == 72
    assert measured["cost"] <= DOCUMENT_SURFACE_CEILING, measured


# MIME export: default stays 71 / 11,130; default + mime is 72 / 11,209
# (+79 proxy tokens). All groups: 77 / 12,175. Measured on Python 3.13 with
# the by_alias=True dump above. Its own ceiling gives about 1% headroom,
# without raising the default, rules or document ceilings.
MIME_SURFACE_CEILING = 11_350


def test_mime_surface_has_its_own_budget_and_preserves_default_schemas():
    code = """import asyncio,json
from mcp.client import Client
from outlook_mcp.server import mcp
async def main():
    async with Client(mcp) as c:
        tools=(await c.list_tools()).tools
    print(json.dumps([t.model_dump(exclude_none=True,by_alias=True) for t in tools]))
asyncio.run(main())
"""
    default_groups = set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT

    def measure(groups):
        env = {**os.environ, "OUTLOOK_MCP_TOOLSETS": ",".join(sorted(groups))}
        env.pop("PYTHONPATH", None)
        result = subprocess.run(
            [sys.executable, "-c", code],
            env=env,
            capture_output=True,
            text=True,
            timeout=120,
            check=True,
        )
        return json.loads(result.stdout.strip().splitlines()[-1])

    default = measure(default_groups)
    enabled = measure(default_groups | {"mime"})
    mime = [tool for tool in enabled if tool["name"] == "outlook_export_message"]
    assert len(default) == 71
    assert len(enabled) == 72
    assert len(mime) == 1
    assert mime[0]["annotations"]["readOnlyHint"] is True
    assert [tool for tool in enabled if tool["name"] != "outlook_export_message"] == default
    assert _proxy_tokens(default) <= DEFAULT_SURFACE_CEILING
    assert _proxy_tokens(enabled) <= MIME_SURFACE_CEILING
