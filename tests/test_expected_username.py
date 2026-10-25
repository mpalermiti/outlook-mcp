"""One instance may optionally refuse a different sign-in; every record here is synthetic."""

from unittest.mock import MagicMock

import pytest
from azure.identity import AuthenticationRecord
from pydantic import ValidationError

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config
from outlook_mcp.errors import OutlookMCPError


def record(username):
    return AuthenticationRecord(
        tenant_id="consumers",
        client_id="test-id",
        authority="https://login.microsoftonline.com/consumers",
        home_account_id="home-1",
        username=username,
    )


@pytest.mark.parametrize("username", ["other@example.com", ""])
def test_wrong_saved_user_is_refused_before_any_credential(tmp_path, monkeypatch, username):
    path = tmp_path / "auth_record.json"
    path.write_text(record(username).serialize())
    original = path.read_bytes()
    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: path)
    manager = AuthManager(Config(client_id="test-id", expected_username="owner@example.com"))
    make = MagicMock()
    monkeypatch.setattr(manager, "_make_credential", make)

    assert manager.try_cached_token() is False
    assert not manager.is_authenticated()
    make.assert_not_called()
    assert manager.startup_error.code == "account_mismatch"
    with pytest.raises(OutlookMCPError, match="expected_username"):
        manager.get_credential()
    assert path.read_bytes() == original


@pytest.mark.parametrize("expected", [None, "OWNER@Example.COM"])
def test_unset_or_case_insensitive_match_keeps_cached_refresh(monkeypatch, expected):
    saved = record("owner@example.com")
    monkeypatch.setattr(auth_module, "_load_auth_record", lambda: saved)
    manager = AuthManager(Config(client_id="test-id", expected_username=expected))
    make = MagicMock()
    monkeypatch.setattr(manager, "_make_credential", make)

    assert manager.try_cached_token() is True
    make.assert_called_once_with(auth_record=saved, silent=True)
    make.return_value.get_token.assert_called_once_with(auth_module.GRAPH_DEFAULT_SCOPE)


@pytest.mark.parametrize("username", ["other@example.com", None])
def test_wrong_new_user_does_not_replace_an_existing_record(tmp_path, monkeypatch, username):
    path = tmp_path / "auth_record.json"
    path.write_text(record("owner@example.com").serialize())
    original = path.read_bytes()
    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: path)
    manager = AuthManager(Config(client_id="test-id", expected_username="owner@example.com"))
    credential = MagicMock()
    credential._auth_record = record(username) if username is not None else None
    monkeypatch.setattr(manager, "_make_credential", MagicMock(return_value=credential))

    with pytest.raises(OutlookMCPError, match="expected_username"):
        manager.login_interactive()
    assert not manager.is_authenticated()
    assert path.read_bytes() == original


@pytest.mark.parametrize("expected", [None, "OWNER@Example.COM"])
def test_matching_new_sign_in_is_saved_and_used(tmp_path, monkeypatch, expected):
    path = tmp_path / "auth_record.json"
    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: path)
    manager = AuthManager(Config(client_id="test-id", expected_username=expected))
    credential = MagicMock()
    credential._auth_record = record("owner@example.com")
    monkeypatch.setattr(manager, "_make_credential", MagicMock(return_value=credential))

    manager.login_interactive()
    assert manager.get_credential() is credential
    assert auth_module._load_auth_record().username == "owner@example.com"


@pytest.mark.parametrize("value", ["", "  ", "owner\n@example.com", "owner\x00@example.com"])
def test_expected_username_refuses_blank_or_control_characters(value):
    with pytest.raises(ValidationError, match="expected_username"):
        Config(expected_username=value)


@pytest.mark.parametrize(
    "key", ["expectedUsername", "expected-user-name", "expected_usernmae", "expected_usename"]
)
def test_a_misspelt_identity_guard_is_not_ignored(key):
    with pytest.raises(ValidationError, match="expected_username"):
        Config(**{key: "owner@example.com"})


def test_status_shows_mismatch_remedy_without_either_username(tmp_path, monkeypatch, capsys):
    from outlook_mcp import cli

    path = tmp_path / "auth_record.json"
    path.write_text(record("other@example.com").serialize())
    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: path)
    monkeypatch.setattr(
        cli,
        "load_config",
        lambda: Config(client_id="test-id", expected_username="owner@example.com"),
    )
    make = MagicMock(side_effect=AssertionError("must refuse before credential construction"))
    monkeypatch.setattr(AuthManager, "_make_credential", make)
    cli.cmd_status()
    output = capsys.readouterr()
    text = output.out + output.err
    assert "not authenticated" in text and "expected_username" in text
    assert "outlook-mcp auth" in text
    assert "or correct expected_username if it is wrong" in text
    assert "do not change" in text
    assert "owner@example.com" not in text and "other@example.com" not in text
    make.assert_not_called()


@pytest.mark.parametrize("command", ["serve", "auth", "status"])
def test_exact_misspelling_stops_real_entry_points(tmp_path, command):
    import json
    import os
    import subprocess
    import sys
    from pathlib import Path

    config_dir = tmp_path / "settings"
    config_dir.mkdir()
    (config_dir / "config.json").write_text(
        json.dumps({"client_id": "test-id", "expected_usename": "owner@example.com"})
    )
    attempted_side_effect = tmp_path / "attempted-side-effect"
    # Install the barriers inside the child, before outlook_mcp imports either
    # Azure constructor. A broken config guard must not construct a credential,
    # touch OS token-cache persistence, resolve a host or make a connection.
    child = """
import os
import socket
from pathlib import Path

import azure.identity

def forbidden(*args, **kwargs):
    Path(os.environ["OUTLOOK_MCP_TEST_SIDE_EFFECT"]).write_text("attempted")
    raise SystemExit(97)

azure.identity.DeviceCodeCredential = forbidden
azure.identity.TokenCachePersistenceOptions = forbidden
socket.getaddrinfo = forbidden
socket.create_connection = forbidden
socket.socket.connect = forbidden
socket.socket.connect_ex = forbidden

from outlook_mcp.cli import main
main()
"""
    result = subprocess.run(
        [sys.executable, "-c", child, command],
        env={
            **os.environ,
            "HOME": str(tmp_path),
            "OUTLOOK_MCP_CONFIG_DIR": str(config_dir),
            "OUTLOOK_MCP_TEST_SIDE_EFFECT": str(attempted_side_effect),
            "PYTHONPATH": str(Path(__file__).resolve().parents[1] / "src"),
        },
        capture_output=True,
        text=True,
        timeout=15,
    )
    assert result.returncode == 1
    assert "misspelt" in result.stderr
    assert "expected_usename" in result.stderr and "expected_username" in result.stderr
    assert "owner@example.com" not in result.stdout + result.stderr
    assert "Traceback" not in result.stderr
    assert not attempted_side_effect.exists(), "credential, persistence or network was attempted"
