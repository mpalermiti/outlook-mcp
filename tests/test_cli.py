"""Tests for the CLI commands.

`cmd_status` reaches for the auth record, and the auth record lives in the
operator's real settings directory by default. These tests patch
``auth._auth_record_path`` to a tmp_path so the CLI is exercised without ever
touching the host's real ``~/.outlook-mcp/auth_record.json`` — a test that
reads the real record is order-dependent, host-dependent, and one refactor
away from deleting it.
"""


import pytest

from outlook_mcp import cli
from outlook_mcp.config import Config


@pytest.fixture(autouse=True)
def _record_in_tmp(tmp_path, monkeypatch):
    """Point the record path at an empty tmp dir for every test here."""
    from outlook_mcp import auth as auth_module

    monkeypatch.setattr(
        auth_module, "_auth_record_path", lambda: tmp_path / "auth_record.json"
    )


def test_status_without_config_exits_with_the_fix(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config())
    with pytest.raises(SystemExit) as exc:
        cli.cmd_status()
    assert exc.value.code == 1
    assert "client_id" in capsys.readouterr().out


def test_status_authenticated_flow_reads_only_the_patched_record(capsys, monkeypatch):
    """No record in the tmp dir -> 'not authenticated', and the real settings
    directory was never consulted."""
    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    cli.cmd_status()  # must not raise
    out = capsys.readouterr().out
    assert "not authenticated" in out
    assert "outlook-mcp auth" in out


def test_auth_without_config_exits_with_the_fix(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config())
    with pytest.raises(SystemExit) as exc:
        cli.cmd_auth()
    assert exc.value.code == 1
    assert "client_id" in capsys.readouterr().out


def test_logout_prints_keychain_guidance(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    cli.cmd_logout()  # must not raise, must not touch any files
    out = capsys.readouterr().out
    assert "Keychain" in out or "credential store" in out
