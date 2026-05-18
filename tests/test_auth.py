"""Tests for auth module."""

import logging
from unittest.mock import patch

import pytest

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager, _unencrypted_fallback_will_be_used
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError


@pytest.fixture(autouse=True)
def _reset_unencrypted_warning_latch():
    """Reset the once-per-process warning latch between tests."""
    auth_module._warned_unencrypted_fallback = False
    yield
    auth_module._warned_unencrypted_fallback = False


def test_auth_manager_init():
    """AuthManager initializes with config."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.config is config
    assert auth.credential is None


def test_auth_scopes_default():
    """Default scopes include read-write."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.ReadWrite" in scopes
    assert "Mail.Send" in scopes
    assert "Calendars.ReadWrite" in scopes
    # offline_access is reserved — MSAL adds it automatically
    assert "offline_access" not in scopes


def test_auth_scopes_read_only():
    """Read-only mode uses read scopes."""
    config = Config(client_id="test-id", read_only=True)
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.Read" in scopes
    assert "Mail.ReadWrite" not in scopes
    assert "Mail.Send" not in scopes
    assert "Calendars.Read" in scopes
    assert "Calendars.ReadWrite" not in scopes


def test_auth_not_authenticated():
    """is_authenticated returns False before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False


def test_auth_get_credential_raises_when_not_authenticated():
    """get_credential raises AuthRequiredError before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    with pytest.raises(AuthRequiredError):
        auth.get_credential()


def test_login_interactive_requires_client_id():
    """login_interactive raises if client_id is not configured."""
    config = Config()  # No client_id
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="client_id"):
        auth.login_interactive(auth.get_scopes())


def test_try_cached_token_returns_false_without_client_id():
    """try_cached_token returns False if client_id is not set."""
    config = Config()
    auth = AuthManager(config)
    assert auth.try_cached_token(auth.get_scopes()) is False


class TestUnencryptedFallbackDetection:
    """_unencrypted_fallback_will_be_used mirrors msal_extensions' check."""

    def test_macos_is_never_fallback(self):
        """macOS uses Keychain — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "darwin"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_windows_is_never_fallback(self):
        """Windows uses DPAPI — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "win32"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_with_gi_available(self):
        """Linux with PyGObject importable uses libsecret — no fallback."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=object()),
        ):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_without_gi_uses_fallback(self):
        """Linux without PyGObject (issue #7) triggers the fallback path."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=None),
        ):
            assert _unencrypted_fallback_will_be_used() is True


class TestUnencryptedFallbackWarning:
    """_make_credential emits a warning at most once when fallback is in use."""

    def test_warning_fires_once_when_fallback_active(self, caplog):
        """A single warning is logged on the first credential build."""
        config = Config(client_id="test-id")
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=True,
            ),
        ):
            auth._make_credential()
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert len(fallback_warnings) == 1
        assert fallback_warnings[0].levelno == logging.WARNING

    def test_no_warning_when_fallback_inactive(self, caplog):
        """No fallback warning is logged when encrypted storage is available."""
        config = Config(client_id="test-id")
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
        ):
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert fallback_warnings == []
