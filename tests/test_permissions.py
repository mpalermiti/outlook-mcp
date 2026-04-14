"""Tests for the granular permission system."""

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.permissions import (
    CATEGORY_CALENDAR_WRITE,
    CATEGORY_CONTACTS_WRITE,
    CATEGORY_MAIL_DRAFTS,
    CATEGORY_MAIL_FOLDERS,
    CATEGORY_MAIL_SEND,
    CATEGORY_MAIL_TRIAGE,
    CATEGORY_TODO_WRITE,
    VALID_CATEGORIES,
    check_permission,
)

# ---------------------------------------------------------------------------
# Category constants
# ---------------------------------------------------------------------------


def test_category_constants_are_snake_case_strings():
    """All 7 category constants are non-empty lowercase snake_case strings."""
    categories = [
        CATEGORY_MAIL_SEND,
        CATEGORY_MAIL_DRAFTS,
        CATEGORY_MAIL_TRIAGE,
        CATEGORY_MAIL_FOLDERS,
        CATEGORY_CALENDAR_WRITE,
        CATEGORY_CONTACTS_WRITE,
        CATEGORY_TODO_WRITE,
    ]
    for cat in categories:
        assert isinstance(cat, str)
        assert cat == cat.lower()
        assert " " not in cat
        assert cat


def test_valid_categories_contains_all_7():
    """VALID_CATEGORIES includes every declared category constant."""
    assert VALID_CATEGORIES == {
        CATEGORY_MAIL_SEND,
        CATEGORY_MAIL_DRAFTS,
        CATEGORY_MAIL_TRIAGE,
        CATEGORY_MAIL_FOLDERS,
        CATEGORY_CALENDAR_WRITE,
        CATEGORY_CONTACTS_WRITE,
        CATEGORY_TODO_WRITE,
    }
    assert len(VALID_CATEGORIES) == 7


def test_category_values_are_expected_names():
    """Category string values match the published contract."""
    assert CATEGORY_MAIL_SEND == "mail_send"
    assert CATEGORY_MAIL_DRAFTS == "mail_drafts"
    assert CATEGORY_MAIL_TRIAGE == "mail_triage"
    assert CATEGORY_MAIL_FOLDERS == "mail_folders"
    assert CATEGORY_CALENDAR_WRITE == "calendar_write"
    assert CATEGORY_CONTACTS_WRITE == "contacts_write"
    assert CATEGORY_TODO_WRITE == "todo_write"


# ---------------------------------------------------------------------------
# check_permission: read_only mode
# ---------------------------------------------------------------------------


def test_read_only_true_raises_regardless_of_allow_categories():
    """read_only=True always raises ReadOnlyError, even if category is whitelisted."""
    config = Config(read_only=True, allow_categories=[CATEGORY_MAIL_SEND])
    with pytest.raises(ReadOnlyError) as exc_info:
        check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_message")
    assert "outlook_send_message" in exc_info.value.message


def test_read_only_true_with_empty_allow_categories_raises():
    """read_only=True with no whitelist raises ReadOnlyError."""
    config = Config(read_only=True)
    with pytest.raises(ReadOnlyError):
        check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")


def test_read_only_true_preempts_permission_denied():
    """ReadOnlyError is raised (not PermissionDeniedError) when both would apply."""
    # Category not in allow_categories AND read_only=True — ReadOnlyError wins.
    config = Config(read_only=True, allow_categories=[CATEGORY_CALENDAR_WRITE])
    with pytest.raises(ReadOnlyError):
        check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_message")


# ---------------------------------------------------------------------------
# check_permission: empty allow_categories (fully open mode)
# ---------------------------------------------------------------------------


def test_empty_allow_categories_allows_all_when_not_read_only():
    """Empty allow_categories + read_only=False means fully open: no error."""
    config = Config(read_only=False)
    assert config.allow_categories == []
    for cat in VALID_CATEGORIES:
        # Should not raise
        result = check_permission(config, cat, f"tool_for_{cat}")
        assert result is None


# ---------------------------------------------------------------------------
# check_permission: non-empty allow_categories (whitelist mode)
# ---------------------------------------------------------------------------


def test_single_category_allows_that_category():
    """A single allowed category permits matching tools."""
    config = Config(read_only=False, allow_categories=[CATEGORY_CALENDAR_WRITE])
    result = check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")
    assert result is None


def test_single_category_denies_other_categories():
    """A single allowed category rejects all others with PermissionDeniedError."""
    config = Config(read_only=False, allow_categories=[CATEGORY_CALENDAR_WRITE])
    denied = VALID_CATEGORIES - {CATEGORY_CALENDAR_WRITE}
    for cat in denied:
        with pytest.raises(PermissionDeniedError) as exc_info:
            check_permission(config, cat, f"tool_for_{cat}")
        assert exc_info.value.code == "permission_denied"
        assert f"tool_for_{cat}" in exc_info.value.message
        assert cat in exc_info.value.message


def test_permission_denied_error_action_guides_user():
    """PermissionDeniedError action tells user how to fix it."""
    config = Config(read_only=False, allow_categories=[CATEGORY_MAIL_SEND])
    with pytest.raises(PermissionDeniedError) as exc_info:
        check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")
    err = exc_info.value
    assert err.action is not None
    assert CATEGORY_CALENDAR_WRITE in err.action
    assert "allow_categories" in err.action


def test_multiple_categories_whitelist():
    """Multiple allowed categories each permit matching tools, deny non-matching."""
    allowed = [CATEGORY_MAIL_SEND, CATEGORY_MAIL_DRAFTS, CATEGORY_CALENDAR_WRITE]
    config = Config(read_only=False, allow_categories=allowed)

    # Allowed ones pass
    for cat in allowed:
        assert check_permission(config, cat, "some_tool") is None

    # Everything else denied
    for cat in VALID_CATEGORIES - set(allowed):
        with pytest.raises(PermissionDeniedError):
            check_permission(config, cat, "some_tool")


def test_all_categories_whitelisted_allows_everything():
    """Whitelisting every category is equivalent to fully open."""
    config = Config(read_only=False, allow_categories=list(VALID_CATEGORIES))
    for cat in VALID_CATEGORIES:
        assert check_permission(config, cat, "some_tool") is None


# ---------------------------------------------------------------------------
# Config-level validation of allow_categories
# ---------------------------------------------------------------------------


def test_config_accepts_valid_category():
    """Config accepts any single valid category."""
    for cat in VALID_CATEGORIES:
        config = Config(allow_categories=[cat])
        assert config.allow_categories == [cat]


def test_config_accepts_all_valid_categories():
    """Config accepts the full list of valid categories."""
    config = Config(allow_categories=list(VALID_CATEGORIES))
    assert set(config.allow_categories) == VALID_CATEGORIES


def test_config_accepts_empty_allow_categories():
    """Config accepts an empty allow_categories list (the default)."""
    config = Config()
    assert config.allow_categories == []

    config2 = Config(allow_categories=[])
    assert config2.allow_categories == []


def test_config_rejects_unknown_category():
    """Config raises a validation error for an unknown category name."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        Config(allow_categories=["not_a_real_category"])


def test_config_rejects_mixed_valid_and_invalid_category():
    """Even one bad category poisons the whole list."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        Config(allow_categories=[CATEGORY_MAIL_SEND, "bogus_category"])


def test_config_rejects_empty_string_category():
    """Empty string is not a valid category."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        Config(allow_categories=[""])
