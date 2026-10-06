"""Every tool that can send email to someone carries `destructiveHint=True`.

MCP clients use the annotations to decide what to ask the user about: a client
that auto-approves writes marked non-destructive treats `destructiveHint=False`
as "only adds to your own data". An email cannot be unsent, and an invitation
or RSVP is an email. These tools were all marked additive, so such a client
would send without asking.

The set is found, not listed: any tool whose implementation gates on
`mail_send` — directly, or through `check_sends_mail` for the part of a call
that emails someone — is a tool that can send. A new one is covered the day
it is written.
"""

import ast
from pathlib import Path

import pytest
from mcp.client import Client

from outlook_mcp import toolsets
from outlook_mcp.server import mcp

TOOLS_DIR = Path(__file__).resolve().parent.parent / "src" / "outlook_mcp" / "tools"


def _tools_that_can_send() -> set[str]:
    """Tool names passed to a `mail_send` gate anywhere under tools/."""
    found: set[str] = set()
    for path in TOOLS_DIR.glob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name):
                continue
            args = node.args
            if (
                node.func.id == "check_permission"
                and len(args) >= 3
                and isinstance(args[1], ast.Name)
                and args[1].id == "CATEGORY_MAIL_SEND"
            ):
                name = args[2]
            elif node.func.id == "check_sends_mail" and len(args) >= 2:
                name = args[1]
            else:
                continue
            if isinstance(name, ast.Constant) and isinstance(name.value, str):
                found.add(name.value)
    return found


def test_the_scan_finds_every_tool_that_can_send():
    """Guards the guard: an empty or shrunken scan would pass anything."""
    assert _tools_that_can_send() == {
        "outlook_send_message",
        "outlook_reply",
        "outlook_forward",
        "outlook_send_draft",
        "outlook_send_with_attachments",
        "outlook_create_event",
        "outlook_update_event",
        "outlook_rsvp",
    }


def test_every_tool_that_can_send_is_destructive():
    missing = _tools_that_can_send() - toolsets.DESTRUCTIVE
    assert not missing, f"can send email but marked additive: {sorted(missing)}"


@pytest.mark.asyncio
async def test_a_client_sees_them_as_destructive():
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    for name in _tools_that_can_send():
        annotations = tools[name].annotations
        assert annotations.read_only_hint is False, name
        assert annotations.destructive_hint is True, name
