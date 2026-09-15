"""Tests for multi-account support: config, auth state, routing, and the gate."""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import AccountConfig, Config


def _two_account_config(**overrides) -> Config:
    base = dict(
        accounts=[
            AccountConfig(name="net", client_id="id1-abcd"),
            AccountConfig(name="neko", client_id="id2-efgh"),
        ],
        capability_accounts={"mail": "net", "calendar": "net", "todo": "neko"},
    )
    base.update(overrides)
    return Config(**base)


# ── Config model ─────────────────────────────────────────


def test_account_config_model():
    """AccountConfig has expected fields and defaults."""
    acc = AccountConfig(name="personal", client_id="abc-123")
    assert acc.name == "personal"
    assert acc.client_id == "abc-123"
    assert acc.tenant_id == "consumers"


def test_account_name_must_be_filesystem_safe():
    """Names become cache filenames and CLI args; refuse the exotic ones."""
    with pytest.raises(ValueError, match="letters, digits"):
        AccountConfig(name="../evil name", client_id="abc-123")


def test_default_account_defaults_to_first():
    """Without an explicit default_account, the first account is it."""
    config = _two_account_config(default_account=None)
    assert config.default_account == "net"


def test_capability_routing_unknown_account_rejected():
    """A routing that names a missing account is a config typo, caught at load."""
    with pytest.raises(ValueError, match="unknown accounts"):
        _two_account_config(capability_accounts={"mail": "ghost"})


def test_capability_routing_unknown_capability_rejected():
    with pytest.raises(ValueError, match="Unknown capabilities"):
        _two_account_config(capability_accounts={"mailbox": "net"})


def test_capability_routing_requires_accounts():
    with pytest.raises(ValueError, match="'accounts' is empty"):
        Config(client_id="abc-123", capability_accounts={"mail": "net"})


def test_duplicate_account_names_rejected():
    with pytest.raises(ValueError, match="Duplicate account names"):
        Config(
            accounts=[
                AccountConfig(name="net", client_id="id1-abcd"),
                AccountConfig(name="net", client_id="id2-efgh"),
            ]
        )


def test_default_account_must_exist():
    with pytest.raises(ValueError, match="default_account 'ghost' not in"):
        _two_account_config(default_account="ghost")


def test_config_roundtrip_with_accounts(tmp_path):
    """Config with accounts saves and loads correctly."""
    from outlook_mcp.config import load_config, save_config

    config_dir = str(tmp_path / ".outlook-mcp")
    original = _two_account_config(allow_cross_account=True)
    save_config(original, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert len(loaded.accounts) == 2
    assert loaded.capability_accounts == {"mail": "net", "calendar": "net", "todo": "neko"}
    assert loaded.allow_cross_account is True


# ── Backward compatibility ───────────────────────────────


def test_list_accounts_empty():
    """Single client_id config shows as 'default' account."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["name"] == "default"
    assert accounts[0]["active"] is True


def test_backward_compatible_single_account():
    """Single client_id config still works (no accounts array)."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False
    assert auth.resolve_capability_account("mail") is None
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["active"] is True


def test_config_accounts_default_empty():
    """Config defaults to empty accounts list."""
    config = Config()
    assert config.accounts == []
    assert config.default_account is None
    assert config.capability_accounts == {}
    assert config.allow_cross_account is False


# ── The gate: allow_cross_account ────────────────────────


def test_gate_closed_hides_other_accounts():
    """With the gate closed the agent sees one merged account, not a menu."""
    auth = AuthManager(_two_account_config())
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["name"] == "net"  # active identity only
    assert "client_id" not in accounts[0]  # and no per-account detail


def test_gate_closed_refuses_switch():
    """switch_account with the gate closed refuses — there is nothing to switch to."""
    auth = AuthManager(_two_account_config())
    with pytest.raises(ValueError, match="allow_cross_account"):
        auth.switch_account("neko")


def test_gate_open_lists_all_accounts():
    auth = AuthManager(_two_account_config(allow_cross_account=True))
    accounts = auth.list_accounts()
    assert [a["name"] for a in accounts] == ["net", "neko"]
    assert accounts[0]["client_id"] == "id1-abcd"[:8] + "..."


def test_gate_open_switch_active_account():
    config = _two_account_config(allow_cross_account=True)
    auth = AuthManager(config)
    result = auth.switch_account("neko")
    assert result == {"status": "switched", "account": "neko"}
    assert auth.resolve_capability_account("contacts") == "neko"  # unrouted follows active
    assert auth.resolve_capability_account("mail") == "net"  # routed stays config-driven


def test_gate_open_switch_one_capability():
    """A capability override redirects just that capability."""
    config = _two_account_config(allow_cross_account=True)
    auth = AuthManager(config)
    result = auth.switch_account("neko", capability="mail")
    assert result == {"status": "switched", "capability": "mail", "account": "neko"}
    assert auth.resolve_capability_account("mail") == "neko"
    assert auth.resolve_capability_account("todo") == "neko"  # unchanged routing
    assert auth.resolve_capability_account("calendar") == "net"


def test_switch_rejects_unknown_account_and_capability():
    auth = AuthManager(_two_account_config(allow_cross_account=True))
    with pytest.raises(ValueError, match="not found in config"):
        auth.switch_account("ghost")
    with pytest.raises(ValueError, match="Unknown capability"):
        auth.switch_account("neko", capability="mailbox")


# ── Capability routing resolution ────────────────────────


def test_routing_resolution_config_only():
    """Gate closed: routing is a pure function of the config."""
    auth = AuthManager(_two_account_config())
    assert auth.resolve_capability_account("mail") == "net"
    assert auth.resolve_capability_account("calendar") == "net"
    assert auth.resolve_capability_account("todo") == "neko"
    assert auth.resolve_capability_account("contacts") == "net"  # unrouted -> default
    assert auth.resolve_capability_account(None) == "net"  # identity tools -> active


def test_default_account_from_config():
    """default_account in config sets the active account."""
    config = _two_account_config(allow_cross_account=True, default_account="neko")
    auth = AuthManager(config)
    active = [a for a in auth.list_accounts() if a["active"]]
    assert len(active) == 1
    assert active[0]["name"] == "neko"


# ── Per-account credentials and storage names ────────────


def test_get_account_credential_unknown_and_unauthenticated():
    auth = AuthManager(_two_account_config())
    with pytest.raises(ValueError, match="not configured"):
        auth.get_account_credential("ghost")
    from outlook_mcp.errors import AuthRequiredError

    with pytest.raises(AuthRequiredError, match="auth neko"):
        auth.get_account_credential("neko")


def test_per_account_cache_and_record_names():
    """Each account gets its own DPAPI/keyring cache and auth record file."""
    from outlook_mcp.auth import _auth_record_path, _cache_name

    assert _cache_name(None) == "outlook-mcp"
    assert _cache_name("net") == "outlook-mcp-net"
    assert _auth_record_path(None).name == "auth_record.json"
    assert _auth_record_path("neko").name == "auth_record-neko.json"


def test_try_cached_token_loads_each_account(monkeypatch):
    """Startup authenticates every account it can, quietly skipping failures."""
    config = _two_account_config()
    auth = AuthManager(config)

    def fake_make(_self, account=None, prompt_callback=None, auth_record=None, *, silent=False):
        cred = MagicMock()
        cred.get_token = MagicMock(side_effect=Exception("boom") if account == "net" else None)
        # azure hands back an _auth_record only on interactive login; silent
        # reuse doesn't reach for it, so the mock doesn't need one.
        return cred

    records = {"net": None, "neko": object()}
    monkeypatch.setattr(
        "outlook_mcp.auth._load_auth_record", lambda account=None: records.get(account)
    )
    monkeypatch.setattr(AuthManager, "_make_credential", fake_make)

    assert auth.try_cached_token() is True
    assert "neko" in auth._credentials
    assert "net" not in auth._credentials
    # Active account (net) failed; identity falls back to the authenticated one.
    assert auth.credential is auth._credentials["neko"]
    assert auth._active_account == "neko"


def test_try_cached_token_single_account_untouched(monkeypatch):
    """No accounts configured: the legacy single-record path still applies."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    monkeypatch.setattr("outlook_mcp.auth._load_auth_record", lambda account=None: object())
    cred = MagicMock()
    monkeypatch.setattr(AuthManager, "_make_credential", lambda *a, **k: cred)
    assert auth.try_cached_token() is True
    assert auth.credential is cred


def test_logout_scopes_to_one_account(tmp_path, monkeypatch):
    """Logout must unlink only the named account's record — and only in a
    sandboxed config dir: the real paths live in the operator's home, and a
    test account named like a configured one once deleted a live record."""
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: record_dir / f"record-{account or 'default'}.json",
    )
    (record_dir / "record-neko.json").write_text("{}")

    config = _two_account_config()
    auth = AuthManager(config)
    auth._credentials["net"] = MagicMock()
    auth._credentials["neko"] = MagicMock()
    auth.credential = auth._credentials["net"]

    result = auth.logout("neko")

    assert "neko" not in auth._credentials
    assert "net" in auth._credentials
    assert result["status"] == "logged_out"
    assert not (record_dir / "record-neko.json").exists()
