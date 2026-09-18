"""Tests for config management."""

import json

import pytest

from outlook_mcp.config import Config, load_config, save_config


def test_default_config():
    """Default config has sensible values."""
    config = Config()
    assert config.client_id is None
    assert config.tenant_id == "consumers"
    assert config.read_only is False
    assert config.timezone == "UTC"


def test_config_dir_created(tmp_path, monkeypatch):
    """Config directory is created with 0700 permissions."""
    config_dir = tmp_path / ".outlook-mcp"
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(config_dir))
    save_config(Config(), config_dir=str(config_dir))
    assert config_dir.exists()
    assert oct(config_dir.stat().st_mode & 0o777) == "0o700"


def test_config_file_permissions(tmp_path, monkeypatch):
    """Config file is written with 0600 permissions."""
    config_dir = tmp_path / ".outlook-mcp"
    config_dir.mkdir(mode=0o700)
    save_config(Config(), config_dir=str(config_dir))
    config_file = config_dir / "config.json"
    assert config_file.exists()
    assert oct(config_file.stat().st_mode & 0o777) == "0o600"


def test_config_roundtrip(tmp_path):
    """Config saves and loads correctly."""
    config_dir = str(tmp_path / ".outlook-mcp")
    original = Config(
        client_id="my-app-uuid",
        timezone="America/Los_Angeles",
        read_only=True,
    )
    save_config(original, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert loaded.client_id == "my-app-uuid"
    assert loaded.timezone == "America/Los_Angeles"
    assert loaded.read_only is True


def test_config_rejects_symlink(tmp_path):
    """Config refuses to load from a symlinked file."""
    config_dir = tmp_path / ".outlook-mcp"
    config_dir.mkdir(mode=0o700)
    real_file = tmp_path / "evil_config.json"
    real_file.write_text(json.dumps({"timezone": "Evil/Zone"}))
    symlink = config_dir / "config.json"
    symlink.symlink_to(real_file)
    with pytest.raises(PermissionError, match="symlink"):
        load_config(config_dir=str(config_dir))


def test_config_override_client_id(tmp_path):
    """Client ID set via config."""
    config_dir = str(tmp_path / ".outlook-mcp")
    config = Config(client_id="custom-client-id-uuid")
    save_config(config, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert loaded.client_id == "custom-client-id-uuid"


def test_load_missing_config_returns_defaults(tmp_path):
    """Loading from nonexistent dir returns default config."""
    config_dir = str(tmp_path / "nonexistent")
    loaded = load_config(config_dir=config_dir)
    assert loaded.client_id is None
    assert loaded.tenant_id == "consumers"


def test_unencrypted_token_cache_is_off_unless_asked_for():
    """The secure default has to survive a config file that never mentions it."""
    assert Config().allow_unencrypted_token_cache is False


# ── Multi-account routing fields ─────────────────────────────────────


def _account(name, client_id):
    from outlook_mcp.config import AccountConfig

    return AccountConfig(name=name, client_id=client_id)


def test_capability_routing_roundtrip(tmp_path):
    """accounts + routing + gate survive a save/load cycle."""
    config_dir = str(tmp_path / ".outlook-mcp")
    original = Config(
        accounts=[_account("net", "id1-abcd"), _account("neko", "id2-efgh")],
        capability_accounts={"mail": "net", "todo": "neko"},
        allow_cross_account=True,
    )
    save_config(original, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert loaded.default_account == "net"
    assert loaded.capability_accounts == {"mail": "net", "todo": "neko"}
    assert loaded.allow_cross_account is True


def test_default_account_defaults_to_first_account():
    config = Config(accounts=[_account("net", "id1"), _account("neko", "id2")])
    assert config.default_account == "net"


# ── Legacy-shape compatibility: accept and warn, never kill the install ──
# 1.21 accepted `accounts` and `default_account` without cross-field
# validation, and load_config() runs uncaught in the lifespan and every CLI
# command — a validator that rejects what main accepted bricks the install
# at startup (review of #61). Each shape below was loadable on main, so it
# must stay loadable here.


def test_legacy_default_account_without_accounts_is_ignored(caplog):
    """default_account set, accounts empty: main accepted it; we warn + drop."""
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(client_id="abc", default_account="net")
    assert config.accounts == []
    assert config.default_account is None
    assert "ignoring it" in caplog.text


def test_legacy_capability_accounts_without_accounts_is_dropped(caplog):
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(client_id="abc", capability_accounts={"mail": "net"})
    assert config.capability_accounts == {}
    assert "ignoring the routing" in caplog.text


def test_legacy_unusual_account_name_is_kept_with_warning(caplog):
    """A name 1.21 accepted (spaces, unicode) stays loadable — warn, don't kill."""
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(accounts=[_account("my account", "id1")])
    assert config.accounts[0].name == "my account"
    assert "recommended shape" in caplog.text


def test_unusable_account_names_are_rejected():
    """Names that cannot be a filename fail loudly — they could never work."""
    for bad in ("../evil", "a/b", "a\\b", "", "net\n", ".", "..", "x" * 65):
        with pytest.raises(ValueError, match="not usable as a filename"):
            Config(accounts=[_account(bad, "id1")])


def test_windows_reserved_names_are_rejected():
    """CON/NUL/COM1 match the canonical regex but cannot be filenames on
    Windows — the reserved check has to run ahead of it."""
    for bad in ("CON", "con", "NUL.txt", "COM1"):
        with pytest.raises(ValueError, match="reserved filename"):
            Config(accounts=[_account(bad, "id1")])


def test_account_name_newline_is_not_canonical():
    """re.fullmatch, not match-with-$: 'net\\n' must not slip through the regex.

    `re.match(r'^...$', 'net\\n')` accepts the trailing newline; a name like
    that must not reach a filename (review of #61).
    """
    from outlook_mcp.config import _ACCOUNT_NAME_RE

    assert _ACCOUNT_NAME_RE.match("net\n")  # the trap the review flagged
    assert not _ACCOUNT_NAME_RE.fullmatch("net\n")
    assert _ACCOUNT_NAME_RE.fullmatch("net")
    assert _ACCOUNT_NAME_RE.fullmatch("a" + "b" * 31)
    assert not _ACCOUNT_NAME_RE.fullmatch("a" + "b" * 32)


def test_legacy_duplicate_account_names_keep_first(caplog):
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(
            accounts=[_account("net", "id1"), _account("net", "id2"), _account("neko", "id3")]
        )
    assert [a.name for a in config.accounts] == ["net", "neko"]
    assert "keeping the first" in caplog.text


def test_legacy_default_account_not_in_accounts_falls_back_to_first(caplog):
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(
            accounts=[_account("net", "id1"), _account("neko", "id2")],
            default_account="ghost",
        )
    assert config.default_account == "net"
    assert "using 'net' instead" in caplog.text


def test_unknown_capability_key_is_dropped_with_warning(caplog):
    with caplog.at_level("WARNING", logger="outlook_mcp.config"):
        config = Config(
            accounts=[_account("net", "id1")],
            capability_accounts={"mailbox": "net", "mail": "net"},
        )
    assert config.capability_accounts == {"mail": "net"}
    assert "dropping them" in caplog.text


def test_unknown_routing_account_is_a_hard_error():
    """capability_accounts is a NEW field — no legacy config carries it, so a
    value naming an unknown account is a typo in new config. Falling back to
    the default account would serve the wrong mailbox silently; refuse."""
    with pytest.raises(ValueError, match="unknown accounts"):
        Config(
            accounts=[_account("net", "id1")],
            capability_accounts={"mail": "ghost"},
        )
