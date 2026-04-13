"""Tests for multi-account support."""

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import AccountConfig, Config


def test_account_config_model():
    """AccountConfig has expected fields and defaults."""
    acc = AccountConfig(name="personal", client_id="abc-123")
    assert acc.name == "personal"
    assert acc.client_id == "abc-123"
    assert acc.tenant_id == "consumers"


def test_list_accounts_empty():
    """Single client_id config shows as 'default' account."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["name"] == "default"
    assert accounts[0]["active"] is True


def test_list_accounts_multiple():
    """Multiple accounts listed from config."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
            AccountConfig(name="family", client_id="id2-efgh"),
        ]
    )
    auth = AuthManager(config)
    accounts = auth.list_accounts()
    assert len(accounts) == 2
    assert accounts[0]["name"] == "personal"
    assert accounts[1]["name"] == "family"


def test_switch_account_valid():
    """Switching to a valid account succeeds."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
            AccountConfig(name="family", client_id="id2-efgh"),
        ]
    )
    auth = AuthManager(config)
    result = auth.switch_account("family")
    assert result["status"] == "switched"
    assert result["account"] == "family"


def test_switch_account_not_found():
    """Switching to nonexistent account raises ValueError."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
        ]
    )
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="not found"):
        auth.switch_account("nonexistent")


def test_default_account_from_config():
    """default_account in config sets the active account."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
            AccountConfig(name="family", client_id="id2-efgh"),
        ],
        default_account="family",
    )
    auth = AuthManager(config)
    accounts = auth.list_accounts()
    active = [a for a in accounts if a["active"]]
    assert len(active) == 1
    assert active[0]["name"] == "family"


def test_backward_compatible_single_account():
    """Single client_id config still works (no accounts array)."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["active"] is True


def test_config_accounts_default_empty():
    """Config defaults to empty accounts list."""
    config = Config()
    assert config.accounts == []
    assert config.default_account is None


def test_switch_account_clears_credential():
    """Switching to account without cached credential sets credential to None."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
            AccountConfig(name="family", client_id="id2-efgh"),
        ]
    )
    auth = AuthManager(config)
    # Manually set a credential to simulate authenticated state
    auth.credential = "fake-credential"
    auth.switch_account("family")
    assert auth.credential is None


def test_list_accounts_masks_client_id():
    """Client ID is masked in list_accounts output."""
    config = Config(
        accounts=[
            AccountConfig(name="personal", client_id="abcdefgh-1234"),
        ]
    )
    auth = AuthManager(config)
    accounts = auth.list_accounts()
    assert accounts[0]["client_id"] == "abcdefgh..."


def test_config_roundtrip_with_accounts(tmp_path):
    """Config with accounts saves and loads correctly."""
    from outlook_mcp.config import load_config, save_config

    config_dir = str(tmp_path / ".outlook-mcp")
    original = Config(
        accounts=[
            AccountConfig(name="personal", client_id="id1-abcd"),
            AccountConfig(name="family", client_id="id2-efgh"),
        ],
        default_account="personal",
    )
    save_config(original, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert len(loaded.accounts) == 2
    assert loaded.accounts[0].name == "personal"
    assert loaded.accounts[1].name == "family"
    assert loaded.default_account == "personal"
