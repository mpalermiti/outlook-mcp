"""The write tier never sends mail.

``live_write`` creates and deletes real items, drafts among them, and a sent message
cannot be taken back. So no file in that tier may import a sending tool or call one
of the SDK's sending actions; a draft there is read back and deleted, never sent.
"""

from __future__ import annotations

import ast
from pathlib import Path

_TESTS = Path(__file__).parent
_SENDING_TOOLS = {"send_message", "send_draft", "send_with_attachments", "reply", "forward"}
_SENDING_ACTIONS = {"send", "send_mail", "reply", "reply_all", "forward"}


def _write_tier_files() -> list[Path]:
    return sorted(
        p for p in _TESTS.glob("test_*.py") if "pytest.mark.live_write" in p.read_text("utf-8")
    )


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


def test_the_write_tier_has_files():
    assert {p.name for p in _write_tier_files()} >= {
        "test_live_calendar_write.py",
        "test_live_drafts_write.py",
    }


def test_no_write_tier_file_sends():
    offenders = {}
    for path in _write_tier_files():
        found = _sends(ast.parse(path.read_text("utf-8")))
        if found:
            offenders[path.name] = found
    assert not offenders, f"the live_write tier must not send mail: {offenders}"


def test_the_check_sees_a_send():
    """Control: each shape is caught, or the test above passes on anything."""
    for source in (
        "from outlook_mcp.tools.mail_write import reply",
        "from outlook_mcp.tools.mail_drafts import send_draft",
        "async def f(sdk):\n    await sdk.me.messages.by_message_id('x').send.post()",
        "async def f(sdk, b):\n    await sdk.me.send_mail.post(b)",
    ):
        assert _sends(ast.parse(source)), source
