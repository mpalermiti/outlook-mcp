"""No live tier sends mail, except the one built to.

``live_write`` creates and deletes real items, drafts among them, and the read-only
tiers read a real mailbox; a sent message cannot be taken back. So no file in those
tiers may import a sending tool or call one of the SDK's sending actions; a draft
there is read back and deleted, never sent. (The Graph client those tiers share also
refuses a send on the wire: ``tests/live_guard.NeverSends``.)

``live_send`` is the exception, and is kept apart so this check can stay a check on
files: a file that sends carries that marker and no other live one.
"""

from __future__ import annotations

import ast
from pathlib import Path

_TESTS = Path(__file__).parent
_NEVER_SEND = {"integration", "live", "live_write"}
_SENDING_TOOLS = {"send_message", "send_draft", "send_with_attachments", "reply", "forward"}
_SENDING_ACTIONS = {"send", "send_mail", "reply", "reply_all", "forward"}


def _markers(tree: ast.AST) -> set[str]:
    """Every ``pytest.mark.<name>`` the file applies."""
    return {
        node.attr
        for node in ast.walk(tree)
        if isinstance(node, ast.Attribute)
        and isinstance(node.value, ast.Attribute)
        and node.value.attr == "mark"
        and isinstance(node.value.value, ast.Name)
        and node.value.value.id == "pytest"
    }


def _files() -> dict[str, ast.AST]:
    return {p.name: ast.parse(p.read_text("utf-8")) for p in sorted(_TESTS.glob("test_*.py"))}


def _never_send_files() -> dict[str, ast.AST]:
    return {name: tree for name, tree in _files().items() if _markers(tree) & _NEVER_SEND}


def _sends(tree: ast.AST) -> list[str]:
    found = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and (node.module or "").startswith("outlook_mcp"):
            if node.module.endswith("mail_write"):
                found.append(f"imports {node.module}")
            found += [f"imports {a.name}" for a in node.names if a.name in _SENDING_TOOLS]
        elif isinstance(node, ast.Attribute) and node.attr in _SENDING_ACTIONS:
            found.append(f".{node.attr}")
    return found


def test_the_check_finds_the_tiers():
    assert set(_never_send_files()) >= {
        "test_live_calendar_write.py",
        "test_live_drafts_write.py",
        "test_live_query_shape.py",
        "test_integration.py",
    }


def test_no_file_outside_the_send_tier_sends():
    offenders = {}
    for name, tree in _never_send_files().items():
        found = _sends(tree)
        if found:
            offenders[name] = found
    assert not offenders, f"only the live_send tier may send mail: {offenders}"


def test_the_send_tier_keeps_to_its_own_files():
    mixed = {
        name: sorted(_markers(tree) & _NEVER_SEND)
        for name, tree in _files().items()
        if "live_send" in _markers(tree) and _markers(tree) & _NEVER_SEND
    }
    assert not mixed, f"a live_send file may carry no other live marker: {mixed}"
    assert "test_live_send.py" in {n for n, t in _files().items() if "live_send" in _markers(t)}


def test_the_check_sees_a_send():
    """Control: each shape is caught, or the test above passes on anything."""
    for source in (
        "from outlook_mcp.tools.mail_write import reply",
        "from outlook_mcp.tools.mail_drafts import send_draft",
        "async def f(sdk):\n    await sdk.me.messages.by_message_id('x').send.post()",
        "async def f(sdk, b):\n    await sdk.me.send_mail.post(b)",
    ):
        assert _sends(ast.parse(source)), source


def test_the_check_sees_a_marker():
    """Control: every way a marker is applied here is found."""
    tree = ast.parse(
        "import pytest\n"
        "pytestmark = pytest.mark.live_write\n"
        "pytestmark = [pytest.mark.live, pytest.mark.asyncio]\n"
        "@pytest.mark.integration\n"
        "def test_x(): pass\n"
    )
    assert _markers(tree) >= {"live_write", "live", "integration"}
