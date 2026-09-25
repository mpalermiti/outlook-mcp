"""Tests for config management."""

import json
import os
import subprocess
import sys

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


# ── OUTLOOK_MCP_CONFIG_DIR: one settings directory per process ──────────
#
# The directory constant is read once at import, so an override is exercised
# in a subprocess — the same way a second instance gets its own value.

_PROBE = (
    "from outlook_mcp import auth, config\n"
    "print(config.DEFAULT_CONFIG_DIR)\n"
    "print(auth._auth_record_path())\n"
    "print(config.Config().attachments_dir)\n"
)


def _paths_with_env(value: str | None) -> list[str]:
    env = dict(os.environ)
    if value is None:
        env.pop("OUTLOOK_MCP_CONFIG_DIR", None)
    else:
        env["OUTLOOK_MCP_CONFIG_DIR"] = value
    out = subprocess.run(
        [sys.executable, "-c", _PROBE],
        env=env,
        capture_output=True,
        text=True,
        check=True,
    )
    return out.stdout.strip().splitlines()


def test_config_dir_env_moves_every_setting_with_it(tmp_path):
    """config.json's directory, the auth record, and attachments move together."""
    custom = str(tmp_path / "instance-net")
    config_dir, record_path, attachments_dir = _paths_with_env(custom)

    assert config_dir == custom
    assert record_path == str(tmp_path / "instance-net" / "auth_record.json")
    assert attachments_dir == str(tmp_path / "instance-net" / "attachments")


def test_config_dir_env_empty_or_unset_keeps_the_default():
    """Empty string is not a directory — it must mean "default", not "".

    Compared against the expanded default, not this process's constant: a
    shell that exports the override would otherwise move the in-process
    constant while the probe (which drops the variable) still reports the
    default, and the test would fail for doing what it was told.
    """
    default = os.path.join(os.path.expanduser("~/.outlook-mcp"), "attachments")

    assert _paths_with_env(None)[0] == os.path.expanduser("~/.outlook-mcp")
    assert _paths_with_env("")[0] == os.path.expanduser("~/.outlook-mcp")
    assert _paths_with_env(None)[2] == default
    assert _paths_with_env("")[2] == default


def test_auth_record_follows_the_settings_directory_not_a_second_home(tmp_path):
    """The record path must derive from the config module constant — no second
    hardcode could ever disagree with it, because there is no second one."""
    custom = str(tmp_path / "instance-neko")
    record_path = _paths_with_env(custom)[1]

    assert "auth_record.json" in record_path
    assert record_path.startswith(custom)


# ── Legacy and unknown top-level keys load with a warning, never silently ──


def _write_config(config_dir, payload: dict) -> None:
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "config.json").write_text(json.dumps(payload))


def test_legacy_multi_account_fields_are_ignored_with_a_pointer(tmp_path, caplog):
    """`accounts` / `default_account` no longer do anything — the config still
    loads, and the warning names the replacement instead of leaving the
    operator wondering why switching accounts does nothing."""
    import logging

    config_dir = tmp_path / ".outlook-mcp"
    _write_config(
        config_dir,
        {
            "client_id": "id-1",
            "accounts": [{"name": "personal", "client_id": "id-2"}],
            "default_account": "personal",
        },
    )

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.client_id == "id-1"  # the rest of the config is intact
    warnings = [r.getMessage() for r in caplog.records]
    assert any("accounts" in w and "OUTLOOK_MCP_CONFIG_DIR" in w for w in warnings)
    assert any("default_account" in w and "OUTLOOK_MCP_CONFIG_DIR" in w for w in warnings)


def test_unknown_top_level_key_warns_instead_of_vanishing(tmp_path, caplog):
    """A typo'd key used to be dropped without a trace while the operator
    believed they had set it. It still loads — just not silently."""
    import logging

    config_dir = tmp_path / ".outlook-mcp"
    _write_config(config_dir, {"client_id": "id-1", "time_zone": "UTC"})

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.timezone == "UTC"  # the typo'd name never applied
    assert any("time_zone" in r.getMessage() for r in caplog.records)


def test_valid_config_produces_no_key_warnings(tmp_path, caplog):
    import logging

    config_dir = tmp_path / ".outlook-mcp"
    _write_config(config_dir, {"client_id": "id-1", "timezone": "Asia/Tokyo"})

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.timezone == "Asia/Tokyo"
    assert caplog.records == []
