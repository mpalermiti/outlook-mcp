"""A refusal the agent reads must not tell it how to switch the refusal off.

These strings are `ToolError` text and reach the model verbatim. The clients
this server is built for — Claude Code, Cursor, OpenClaw — give the agent file
tools, `config.json` is a plain user-writable file read at the next start, and
the agent reads mail. "Set read_only to false in …/config.json to enable write
operations" was an instruction it could carry out. The settings are the user's:
each refusal names the setting so the agent can tell the user, and says in so
many words that changing it is not the agent's job.
"""

import re

import pytest

from outlook_mcp.config import DEFAULT_CONFIG_DIR
from outlook_mcp.errors import (
    RULE_WRITE_NEEDS_SETTINGS_WRITE,
    AccountMismatchError,
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


# Each text names its setting so the agent can tell the user, and never spells
# out the value or the edit that would widen what the server may do.
SETTING_REFUSALS = [
    pytest.param(
        str(RULE_WRITE_NEEDS_SETTINGS_WRITE), "mailbox_settings_write", id="rule-write-403"
    ),
    pytest.param(
        str(UnencryptedTokenCacheError()), "allow_unencrypted_token_cache", id="plaintext-cache"
    ),
]


@pytest.mark.parametrize(("text", "setting"), SETTING_REFUSALS)
def test_a_text_that_names_a_setting_does_not_give_its_value_or_the_edit(text, setting):
    assert setting in text
    assert LEAVE_IT in text
    assert ": true" not in text
    assert "= true" not in text
    assert not re.search(rf"\bset\s+`?{setting}", text), text


def test_the_token_cache_refusal_points_at_the_readme_section():
    assert "Token cache stored unencrypted" in str(UnencryptedTokenCacheError())


def test_the_account_mismatch_does_not_claim_where_the_setting_came_from():
    """`expected_username` can come from config.json or the environment."""
    text = str(AccountMismatchError())
    assert "expected_username" in text
    assert "config.json" not in text
    assert LEAVE_IT in text


def test_a_real_send_refusal_keeps_its_wording():
    text = str(PermissionDeniedError("outlook_create_event", "mail_send", doing="invite attendees"))
    assert text.startswith(
        "Cannot use outlook_create_event to invite attendees — that sends email, and category "
        "'mail_send' is not in allow_categories."
    )
