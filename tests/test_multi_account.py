"""Tests for multi-account support: auth state, fail-closed routing, the gate."""

from unittest.mock import MagicMock

import pytest
from azure.identity import CredentialUnavailableError

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import AccountConfig, Config
from outlook_mcp.errors import AuthRequiredError, UnencryptedTokenCacheError


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


def _authenticated_manager(config) -> AuthManager:
    """Manager with a fake credential per account, as startup would leave it."""
    auth = AuthManager(config)
    for acc in config.accounts:
        auth._credentials[acc.name] = MagicMock()
    active = auth._active_account
    if active in auth._credentials:
        auth.credential = auth._credentials[active]
    return auth


# ── Config model ─────────────────────────────────────────


def test_account_config_model():
    """AccountConfig has expected fields and defaults."""
    acc = AccountConfig(name="personal", client_id="abc-123")
    assert acc.name == "personal"
    assert acc.client_id == "abc-123"
    assert acc.tenant_id == "consumers"


def test_default_account_defaults_to_first():
    """Without an explicit default_account, the first account is it."""
    config = _two_account_config(default_account=None)
    assert config.default_account == "net"


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
    """Single client_id config still works (no accounts array). Routing answers
    None so the legacy single credential serves everything."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False
    assert auth.resolve_capability_account("mail") is None
    assert auth.resolve_capability_account(None) is None
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


# ── The gate: allow_cross_account ────────────────────────


def test_gate_closed_hides_other_accounts():
    """With the gate closed the agent sees one merged account, not a menu."""
    auth = AuthManager(_two_account_config())
    accounts = auth.list_accounts()
    assert len(accounts) == 1
    assert accounts[0]["name"] == "net"  # active identity only
    assert "client_id" not in accounts[0]  # and no per-account detail


def test_gate_closed_refuses_switch():
    """switch_account with the gate closed refuses — there is nothing to
    switch to. Refusal comes before any name validation, so it leaks
    nothing about which accounts exist."""
    auth = _authenticated_manager(_two_account_config())
    with pytest.raises(ValueError, match="allow_cross_account"):
        auth.switch_account("neko")
    with pytest.raises(ValueError, match="allow_cross_account"):
        auth.switch_account("not-even-an-account")


def test_gate_open_lists_all_accounts():
    auth = _authenticated_manager(_two_account_config(allow_cross_account=True))
    accounts = auth.list_accounts()
    assert [a["name"] for a in accounts] == ["net", "neko"]
    assert accounts[0]["client_id"] == "id1-abcd"[:8] + "..."


def test_gate_open_switch_active_account():
    config = _two_account_config(allow_cross_account=True)
    auth = _authenticated_manager(config)
    result = auth.switch_account("neko")
    assert result == {"status": "switched", "account": "neko"}
    assert auth.resolve_capability_account("contacts") == "neko"  # unrouted follows active
    assert auth.resolve_capability_account("mail") == "net"  # routed stays config-driven


def test_gate_open_switch_one_capability():
    """A capability override redirects just that capability."""
    config = _two_account_config(allow_cross_account=True)
    auth = _authenticated_manager(config)
    result = auth.switch_account("neko", capability="mail")
    assert result == {"status": "switched", "capability": "mail", "account": "neko"}
    assert auth.resolve_capability_account("mail") == "neko"
    assert auth.resolve_capability_account("calendar") == "net"  # unchanged routing
    assert auth.resolve_capability_account("todo") == "neko"


def test_switch_rejects_unknown_account_and_capability():
    auth = _authenticated_manager(_two_account_config(allow_cross_account=True))
    with pytest.raises(ValueError, match="not found in config"):
        auth.switch_account("ghost")
    with pytest.raises(ValueError, match="Unknown capability"):
        auth.switch_account("neko", capability="mailbox")


def test_switch_to_unauthenticated_account_fails_closed_for_identity():
    """Explicit switch gets honest results, not a fallback: identity tools on
    an unauthenticated active account raise naming it."""
    config = _two_account_config(allow_cross_account=True)
    auth = _authenticated_manager(config)
    auth._credentials.pop("neko", None)

    auth.switch_account("neko")
    assert auth.resolve_capability_account(None) == "neko"
    with pytest.raises(AuthRequiredError, match="auth neko"):
        auth.get_account_credential(auth.resolve_capability_account(None))


# ── Capability routing resolution ────────────────────────


def test_routing_resolution_config_only():
    """Gate closed: routing is a pure function of the config."""
    auth = _authenticated_manager(_two_account_config())
    assert auth.resolve_capability_account("mail") == "net"
    assert auth.resolve_capability_account("calendar") == "net"
    assert auth.resolve_capability_account("todo") == "neko"
    assert auth.resolve_capability_account("contacts") == "net"  # unrouted -> default
    assert auth.resolve_capability_account(None) == "net"  # identity tools -> active


def test_default_account_from_config():
    """default_account in config sets the active account."""
    config = _two_account_config(allow_cross_account=True, default_account="neko")
    auth = _authenticated_manager(config)
    active = [a for a in auth.list_accounts() if a["active"]]
    assert len(active) == 1
    assert active[0]["name"] == "neko"


# ── Per-account credentials and storage names ────────────


def test_get_account_credential_unknown_and_unauthenticated():
    auth = AuthManager(_two_account_config())
    with pytest.raises(ValueError, match="not configured"):
        auth.get_account_credential("ghost")
    with pytest.raises(AuthRequiredError, match="auth neko"):
        auth.get_account_credential("neko")


def test_one_cache_name_for_every_account():
    """All accounts share ONE token cache name; the per-account
    AuthenticationRecord (one auth_record-<name>.json each) pins the identity.

    Per-account cache names look right on Windows (file per name) and Linux
    (libsecret keyed by name) but on macOS the Keychain item is fixed
    (Microsoft.Developer.IdentityService / MSALCache) — the name only picks
    the signal file, so per-account names there overwrite each other
    (review of #61). MSAL caches are multi-account by design."""
    import inspect

    from outlook_mcp import auth as auth_module
    from outlook_mcp.auth import _auth_record_path

    assert not hasattr(auth_module, "_cache_name")  # the per-name helper is gone
    src = inspect.getsource(auth_module.AuthManager._make_credential)
    assert "name=CACHE_NAME" in src  # and every credential uses the shared name
    assert _auth_record_path(None).name == "auth_record.json"
    assert _auth_record_path("neko").name == "auth_record-neko.json"


def test_try_cached_token_is_fail_closed_when_default_fails(monkeypatch, caplog):
    """The review's central case, pinned as the EXPECTED behaviour (it was
    previously pinned as a leak: the active account silently became whichever
    account authenticated first, and unrouted writes ran on it).

    net (the default) has no valid token; neko authenticates. After startup:
    - the active account is still net — never re-pointed;
    - a data capability routed to net raises AuthRequiredError('net');
    - a capability routed to neko still works;
    - identity tools fall back to neko (identity-only), with a warning.
    """
    config = _two_account_config()
    auth = AuthManager(config)

    def fake_make(_self, account=None, prompt_callback=None, auth_record=None, *, silent=False):
        cred = MagicMock()
        # Silent-mode cache miss on net: azure-identity's own "cannot refresh"
        # signal, i.e. the realistic way a default account's token is gone.
        cred.get_token = MagicMock(
            side_effect=CredentialUnavailableError("no token for net") if account == "net" else None
        )
        return cred

    records = {"net": object(), "neko": object()}
    monkeypatch.setattr(
        "outlook_mcp.auth._load_auth_record", lambda account=None: records.get(account)
    )
    monkeypatch.setattr(AuthManager, "_make_credential", fake_make)

    with caplog.at_level("WARNING", logger="outlook_mcp.auth"):
        assert auth.try_cached_token() is True

    assert "neko" in auth._credentials
    assert "net" not in auth._credentials
    # The configured default stays the active account — no substitution.
    assert auth._active_account == "net"
    assert auth.credential is None  # so is_authenticated() is honest
    # Data capability on the failed default: fail closed, naming the fix.
    assert auth.resolve_capability_account("mail") == "net"
    with pytest.raises(AuthRequiredError, match="outlook-mcp auth net"):
        auth.get_account_credential(auth.resolve_capability_account("mail"))
    # A capability routed to the healthy account keeps working.
    assert (
        auth.get_account_credential(auth.resolve_capability_account("todo"))
        is auth._credentials["neko"]
    )
    # Identity tools fall back — identity-only, and loudly.
    assert auth.resolve_capability_account(None) == "neko"
    assert "fail closed" in caplog.text


def test_transient_token_error_does_not_flip_anything(monkeypatch, caplog):
    """A transient 5xx/timeout on the token endpoint is not a logout and not
    an identity switch: the account is simply unauthenticated this run."""
    from azure.core.exceptions import ServiceRequestError

    config = _two_account_config()
    auth = AuthManager(config)

    def fake_make(_self, account=None, prompt_callback=None, auth_record=None, *, silent=False):
        cred = MagicMock()
        cred.get_token = MagicMock(
            side_effect=ServiceRequestError("503") if account == "net" else None
        )
        return cred

    monkeypatch.setattr("outlook_mcp.auth._load_auth_record", lambda account=None: object())
    monkeypatch.setattr(AuthManager, "_make_credential", fake_make)

    with caplog.at_level("WARNING", logger="outlook_mcp.auth"):
        assert auth.try_cached_token() is True

    assert auth._active_account == "net"
    assert "neko" in auth._credentials
    assert "Transient error refreshing the token" in caplog.text


def test_unexpected_errors_are_not_swallowed(monkeypatch):
    """No bare `except Exception` on the refresh path: a bug in our own code
    must surface, not read as 're-run auth' (review of #61)."""
    config = _two_account_config()
    auth = AuthManager(config)

    def fake_make(_self, account=None, prompt_callback=None, auth_record=None, *, silent=False):
        cred = MagicMock()
        cred.get_token = MagicMock(side_effect=RuntimeError("our bug"))
        return cred

    monkeypatch.setattr("outlook_mcp.auth._load_auth_record", lambda account=None: object())
    monkeypatch.setattr(AuthManager, "_make_credential", fake_make)

    with pytest.raises(RuntimeError, match="our bug"):
        auth.try_cached_token()


def test_try_cached_token_single_account_untouched(monkeypatch):
    """No accounts configured: the legacy single-record path still applies."""
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    monkeypatch.setattr("outlook_mcp.auth._load_auth_record", lambda account=None: object())
    cred = MagicMock()
    monkeypatch.setattr(AuthManager, "_make_credential", lambda *a, **k: cred)
    assert auth.try_cached_token() is True
    assert auth.credential is cred


def test_get_account_credential_surfaces_startup_error():
    """A host that cannot store tokens safely (no libsecret) sets
    startup_error with empty _credentials — every routed tool must report the
    config fix, not advise an `outlook-mcp auth <name>` that fails the same
    way (review of #61)."""
    config = _two_account_config()
    auth = AuthManager(config)
    auth.startup_error = UnencryptedTokenCacheError()

    for account in ("net", "neko"):
        with pytest.raises(UnencryptedTokenCacheError):
            auth.get_account_credential(account)


# ── Logout ───────────────────────────────────────────────


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
    (record_dir / "record-net.json").write_text("{}")

    config = _two_account_config()
    auth = _authenticated_manager(config)

    result = auth.logout("neko")

    assert "neko" not in auth._credentials
    assert "net" in auth._credentials
    assert result["status"] == "logged_out"
    assert not (record_dir / "record-neko.json").exists()
    assert (record_dir / "record-net.json").exists()  # other records untouched


def test_logout_of_non_active_account_keeps_is_authenticated(tmp_path, monkeypatch):
    """logout('neko') while net is active and credentialed: net stays
    authenticated. The old code cleared self.credential unconditionally,
    leaving every account working but is_authenticated() False."""
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: record_dir / f"record-{account or 'default'}.json",
    )
    config = _two_account_config()
    auth = _authenticated_manager(config)

    auth.logout("neko")

    assert auth.is_authenticated() is True
    assert auth.credential is auth._credentials["net"]


def test_logout_of_active_account_clears_it(tmp_path, monkeypatch):
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: record_dir / f"record-{account or 'default'}.json",
    )
    config = _two_account_config()
    auth = _authenticated_manager(config)

    auth.logout()  # no argument: the default account (net)

    assert auth.is_authenticated() is False
    assert "net" not in auth._credentials
    assert "neko" in auth._credentials  # the other account is untouched
    assert auth._active_account == "net"  # and still the configured default


def test_logout_single_account_removes_legacy_record(tmp_path, monkeypatch):
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: record_dir / f"record-{account or 'default'}.json",
    )
    (record_dir / "record-default.json").write_text("{}")
    config = Config(client_id="test-id-1234")
    auth = AuthManager(config)
    auth.credential = MagicMock()

    result = auth.logout()

    assert result["status"] == "logged_out"
    assert auth.credential is None
    assert not (record_dir / "record-default.json").exists()


def test_logout_clears_identity_fallback_when_it_is_the_logout_target(tmp_path, monkeypatch):
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: record_dir / f"record-{account or 'default'}.json",
    )
    config = _two_account_config()
    auth = AuthManager(config)
    # Startup state: default (net) unauthenticated, identity fell back to neko.
    auth._credentials["neko"] = MagicMock()
    auth._identity_fallback_account = "neko"

    auth.logout("neko")

    assert auth._identity_fallback_account is None
    assert auth.resolve_capability_account(None) == "net"  # identity follows config again


# ── Auth surfaces reflect the active account ─────────────


@pytest.mark.asyncio
async def test_auth_status_names_the_active_account_when_unauthenticated():
    """auth_status reflects the ACTIVE account only, and its remedy names it
    (review of #61: the auth surfaces must not describe a different account
    than the one whose credentials will actually be used)."""
    from outlook_mcp import server as server_mod

    config = _two_account_config()
    auth = AuthManager(config)  # startup found nothing authenticatable

    ctx = MagicMock()
    ctx.request_context.lifespan_context = {"auth": auth, "config": config}

    result = await server_mod.outlook_auth_status(ctx)

    assert result["authenticated"] is False
    assert "outlook-mcp auth net" in result["action_required"]


# ── 1.21 → multi-account upgrade keeps the login ────────────────────


def test_legacy_login_is_adopted_for_default_account(tmp_path, monkeypatch, caplog):
    """A 1.21 install with 'accounts' populated had ONE login (auth_record.json)
    that served everything — the list did nothing. The upgrade must not
    silently drop it: the legacy record becomes the default account's."""
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    (record_dir / "legacy.json").write_text("{}")
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: (
            record_dir / ("legacy.json" if account is None else f"record-{account}.json")
        ),
    )
    records: dict = {None: object()}
    monkeypatch.setattr(
        "outlook_mcp.auth._load_auth_record", lambda account=None: records.get(account)
    )
    monkeypatch.setattr(
        "outlook_mcp.auth._save_auth_record",
        lambda record, account=None: records.__setitem__(account, record),
    )

    config = _two_account_config()
    auth = AuthManager(config)

    def fake_make(_self, account=None, prompt_callback=None, auth_record=None, *, silent=False):
        cred = MagicMock()
        cred.get_token = MagicMock(return_value=None)
        return cred

    monkeypatch.setattr(AuthManager, "_make_credential", fake_make)

    with caplog.at_level("WARNING", logger="outlook_mcp.auth"):
        assert auth.try_cached_token() is True

    # The default account runs on yesterday's login, not on nothing. The other
    # account legitimately starts unauthenticated — on 1.21 it had no login —
    # and its tools fail closed with their own remedy.
    assert "net" in auth._credentials
    assert "neko" not in auth._credentials
    assert auth.credential is auth._credentials["net"]
    assert "Adopting it as 'net'" in caplog.text


def test_no_adoption_when_the_default_account_has_its_own_record(tmp_path, monkeypatch, caplog):
    """Per-account record present: the legacy file is left alone (and a fresh
    multi-account install, with no legacy file at all, never adopts)."""
    record_dir = tmp_path / "records"
    record_dir.mkdir()
    (record_dir / "legacy.json").write_text("{}")
    (record_dir / "record-net.json").write_text("{}")
    monkeypatch.setattr(
        "outlook_mcp.auth._auth_record_path",
        lambda account=None: (
            record_dir / ("legacy.json" if account is None else f"record-{account}.json")
        ),
    )
    records: dict = {"net": object(), "neko": object()}
    monkeypatch.setattr(
        "outlook_mcp.auth._load_auth_record", lambda account=None: records.get(account)
    )

    config = _two_account_config()
    auth = AuthManager(config)
    cred = MagicMock()
    cred.get_token = MagicMock(return_value=None)
    monkeypatch.setattr(AuthManager, "_make_credential", lambda *a, **k: cred)

    with caplog.at_level("WARNING", logger="outlook_mcp.auth"):
        assert auth.try_cached_token() is True

    assert "Adopting" not in caplog.text
