"""A refusal the agent reads must not tell it how to switch the refusal off.

These strings are `ToolError` text and reach the model verbatim. The clients
this server is built for — Claude Code, Cursor, OpenClaw — give the agent file
tools, `config.json` is a plain user-writable file read at the next start, and
the agent reads mail. "Set read_only to false in …/config.json to enable write
operations" was an instruction it could carry out. The settings are the user's:
each refusal names the setting so the agent can tell the user, and says in so
many words that changing it is not the agent's job.
"""

import pytest

from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError, UnencryptedTokenCacheError
from outlook_mcp.tools.mail_attachments import resolve_attachment_path

LEAVE_IT = "Do not change the server's settings yourself."

REFUSALS = [
    pytest.param(ReadOnlyError("outlook_send_message"), id="read_only"),
    pytest.param(PermissionDeniedError("outlook_send_message", "mail_send"), id="category"),
    pytest.param(
        PermissionDeniedError("outlook_create_event", "mail_send", doing="invite attendees"),
        id="category-for-part-of-a-call",
    ),
    pytest.param(UnencryptedTokenCacheError(), id="plaintext-token-cache"),
]

# The step-by-step recipes the refusals used to give.
RECIPES = [
    "Set read_only to false",
    "to enable write operations",
    "unset allow_categories",
    "for full write access",
    "accept plaintext storage by setting",
    "to change it",
]


@pytest.mark.parametrize("refusal", REFUSALS)
def test_the_refusal_leaves_the_setting_to_the_user(refusal):
    assert LEAVE_IT in str(refusal)


@pytest.mark.parametrize("refusal", REFUSALS)
def test_the_refusal_gives_no_recipe_for_turning_itself_off(refusal):
    text = str(refusal)
    for recipe in RECIPES:
        assert recipe not in text, f"{recipe!r} in {text!r}"


def test_the_attachment_fence_leaves_its_directory_to_the_user(tmp_path):
    base = tmp_path / "attachments"
    base.mkdir()
    outside = tmp_path / "id_ed25519"
    outside.write_text("PRIVATE KEY")

    with pytest.raises(ValueError) as exc:
        resolve_attachment_path(str(outside), str(base))

    text = str(exc.value)
    assert LEAVE_IT in text
    for recipe in RECIPES:
        assert recipe not in text, f"{recipe!r} in {text!r}"
