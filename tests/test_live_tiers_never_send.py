"""No live tier sends mail, except the one built to.

``live_write`` creates and deletes real items, drafts among them, and the read-only
tiers read a real mailbox; a sent message cannot be taken back. So no file in those
tiers may import a sending tool or call one of the SDK's sending actions; a draft
there is read back and deleted, never sent. (The Graph client those tiers share also
refuses a send on the wire: ``tests/live_guard.NeverSends``.)

``live_send`` is the exception, and is kept apart so this check can stay a check on
files: a file that sends carries that marker and no other live one, and makes its
requests only through the send tier's fixtures (no raw ``httpx``, no Graph client of
its own).

This reads source, so it sees the obvious shapes, through the SDK or raw: a sending
tool, an SDK sending action, a URL naming a sending action, an ``attendees`` key or
argument. The guarantee is at run time, where every request a live test makes is
held to its tier's guard (``tests/live_guard.guarding``).
"""

from __future__ import annotations

import ast
import re
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


# A URL fragment naming an action that sends, as a raw request would spell it.
_SENDING_URL = re.compile(
    r"sendmail|/(send|reply|replyall|forward|cancel|accept|decline|tentativelyaccept)\b"
    r"|messagerules|automaticrepliessetting|calendarpermissions",
    re.IGNORECASE,
)


def _docstrings(tree: ast.AST) -> set[int]:
    """ids of the docstring nodes: prose, not requests."""
    found = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = getattr(node, "body", [])
            if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant):
                found.add(id(body[0].value))
    return found


def _empty(node: ast.AST) -> bool:
    return (isinstance(node, ast.Constant) and node.value is None) or (
        isinstance(node, (ast.List, ast.Tuple)) and not node.elts
    )


def _sends(tree: ast.AST) -> list[str]:
    found = []
    prose = _docstrings(tree)
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Constant)
            and isinstance(node.value, str)
            and id(node) not in prose
            and _SENDING_URL.search(node.value)
        ):
            found.append(f"string {node.value[:40]!r}")
        elif isinstance(node, ast.Dict):
            for key, value in zip(node.keys, node.values):
                if isinstance(key, ast.Constant) and key.value == "attendees" and not _empty(value):
                    found.append("an 'attendees' key")
        elif isinstance(node, ast.keyword) and node.arg == "attendees" and not _empty(node.value):
            found.append("attendees=")
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


def test_the_send_tier_goes_through_its_fixtures():
    """No raw client and no Graph client of its own: the send guard is installed
    for the whole test either way, but a raw call is a shape this tier never needs."""
    tree = _files()["test_live_send.py"]
    imported = {
        (node.module or "") if isinstance(node, ast.ImportFrom) else alias.name
        for node in ast.walk(tree)
        if isinstance(node, (ast.Import, ast.ImportFrom))
        for alias in node.names
    }
    assert not {m for m in imported if m.split(".")[0] in ("httpx", "requests")}, imported
    built = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id in ("GraphClient", "GraphServiceClient")
    ]
    assert not built, "test_live_send.py builds a Graph client; use send_graph_client"


def test_the_check_sees_a_send():
    """Control: each shape is caught, or the test above passes on anything."""
    for source in (
        "from outlook_mcp.tools.mail_write import reply",
        "from outlook_mcp.tools.mail_drafts import send_draft",
        "async def f(sdk):\n    await sdk.me.messages.by_message_id('x').send.post()",
        "async def f(sdk, b):\n    await sdk.me.send_mail.post(b)",
        "def f(c):\n    c.post('https://graph.microsoft.com/v1.0/me/sendMail', json={})",
        "def f(c, i):\n    c.post(f'https://graph.microsoft.com/v1.0/me/messages/{i}/reply')",
        "def f(c):\n    c.post(url, json={'subject': 's', 'attendees': [{'emailAddress': {}}]})",
        "async def f(s):\n    await create_event(s, attendees=['a@example.com'])",
        "def f(c):\n    c.post('/me/mailFolders/inbox/messageRules', json={})",
    ):
        assert _sends(ast.parse(source)), source


def test_the_check_leaves_prose_and_empty_attendees_alone():
    source = (
        '"""Never calls /reply or sendMail."""\n'
        "def f(c):\n"
        '    """A docstring may say /forward."""\n'
        "    c.post(url, json={'attendees': []})\n"
        "    create_event(s, attendees=None)\n"
    )
    assert _sends(ast.parse(source)) == []


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
