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


@pytest.mark.parametrize("command", [cli.cmd_auth, cli.cmd_status, cli.cmd_logout])
def test_an_unloadable_config_exits_with_the_repair_not_a_traceback(
    command, capsys, monkeypatch
):
    """Whatever the server's pre-run check catches, the CLI catches too."""
    monkeypatch.setattr(
        cli, "load_config", lambda: (_ for _ in ()).throw(OSError(13, "Permission denied"))
    )
    with pytest.raises(SystemExit) as exc:
        command()

    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "cannot start" in err
    assert "Traceback" not in err


def test_auth_reports_a_structured_failure_with_its_remedy(capsys, monkeypatch):
    """The unencrypted-cache refusal arrives as an OutlookMCPError whose text
    already says what to change — print that, not a traceback."""
    from unittest.mock import patch

    from outlook_mcp.errors import UnencryptedTokenCacheError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    with (
        patch(
            "outlook_mcp.auth.AuthManager.login_interactive",
            side_effect=UnencryptedTokenCacheError(),
        ),
        pytest.raises(SystemExit) as exc,
    ):
        cli.cmd_auth()

    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "allow_unencrypted_token_cache" in err
    assert "Traceback" not in err


def test_logout_removes_this_instances_record_and_says_what_stays(capsys, monkeypatch):
    from outlook_mcp import auth as auth_module

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    record = auth_module._auth_record_path()  # the autouse fixture's tmp path
    record.write_text("{}")

    cli.cmd_logout()

    assert not record.exists()  # the record this instance serves from is gone
    out = capsys.readouterr().out
    assert "auth_record.json" in out
    assert "outlook-mcp auth" in out  # next start asks for auth again
    # the OS cache entry is shared and stays — and the old advice ("remove
    # 'outlook-mcp' from Keychain Access") named an entry that never existed
    assert "Microsoft.Developer.IdentityService" in out
    assert "left in place" in out
    assert "remove 'outlook-mcp'" not in out
