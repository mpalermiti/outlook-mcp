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

from outlook_mcp.config import DEFAULT_CONFIG_DIR
from outlook_mcp.errors import (
    ConfigLoadError,
    PermissionDeniedError,
    ReadOnlyError,
    UnencryptedTokenCacheError,
)
from outlook_mcp.tools.mail_attachments import resolve_attachment_path

# Addressed to the agent by name, because the same text is printed to the
# operator by `outlook-mcp auth` and logged to stderr — and the operator is the
# one who should change the setting.
LEAVE_IT = "If you are an AI agent, do not change the server's settings — tell the user."

REFUSALS = [
    pytest.param(ReadOnlyError("outlook_send_message"), id="read_only"),
    pytest.param(PermissionDeniedError("outlook_send_message", "mail_send"), id="category"),
    pytest.param(
        PermissionDeniedError("outlook_create_event", "mail_send", doing="invite attendees"),
        id="category-for-part-of-a-call",
    ),
    pytest.param(UnencryptedTokenCacheError(), id="plaintext-token-cache"),
    pytest.param(
        ConfigLoadError(ValueError("read_only_consent: true needs read_only: true"), "/cfg"),
        id="config-load",
    ),
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


def test_the_plaintext_refusal_does_not_hand_the_agent_the_config_path():
    """Naming the setting is for the user; where the file lives is not the agent's business."""
    assert DEFAULT_CONFIG_DIR not in str(UnencryptedTokenCacheError())


def test_a_refused_download_target_says_how_to_retry(tmp_path):
    """The fence also guards downloads, where there is no file to move yet.

    The recovery there is a path inside the directory — a bare filename lands
    in it — not asking the user to move something that does not exist.
    """
    base = tmp_path / "attachments"
    base.mkdir()

    with pytest.raises(ValueError) as exc:
        resolve_attachment_path(str(tmp_path / "Desktop" / "a.pdf"), str(base))

    text = str(exc.value)
    assert "bare filename" in text
    assert "To send a file" in text
