"""Tests for auth module."""

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError


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
    assert "offline_access" in scopes


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


def test_auth_requires_client_id():
    """Login raises if client_id is not configured."""
    config = Config()  # No client_id
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="client_id"):
        auth.login()
