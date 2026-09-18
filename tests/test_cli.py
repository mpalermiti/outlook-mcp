"""Tests for the CLI's multi-account surface and failure handling.

Review of #61 asked for two things the old CLI got wrong: `logout` printed
every account's cache name and claimed the whole server needed re-auth
(other accounts' records stayed, and the next serve would run on them), and
a config the validators reject crashed every command with a traceback —
`logout` included — instead of saying what to fix.
"""

import json

import pytest

from outlook_mcp import cli
from outlook_mcp.config import AccountConfig, Config


def _patch_config(monkeypatch, config: Config) -> None:
    monkeypatch.setattr(cli, "load_config", lambda: config)


def test_logout_prints_no_per_account_cache_names(capsys, monkeypatch):
    """One shared cache entry ('outlook-mcp'), not one per account — those do
    not exist on macOS — and no blanket 'server will require re-authentication'
    claim: only the logged-out account is affected."""
    _patch_config(
        monkeypatch,
        Config(
            accounts=[
                AccountConfig(name="net", client_id="id1-abcd"),
                AccountConfig(name="neko", client_id="id2-efgh"),
            ],
            capability_accounts={"mail": "net", "todo": "neko"},
        ),
    )
    cli.cmd_logout("neko")
    out = capsys.readouterr().out

    assert "outlook-mcp-neko" not in out  # per-account cache names are a macOS lie
    assert "outlook-mcp" in out  # the one real cache entry
    assert "does\nNOT fall back to another account" in out
    assert "Other accounts keep working" in out


def test_logout_without_accounts_keeps_legacy_guidance(capsys, monkeypatch):
    _patch_config(monkeypatch, Config(client_id="test-id"))
    cli.cmd_logout()
    out = capsys.readouterr().out

    assert "require re-authentication" in out


def test_invalid_config_exits_with_a_fix_not_a_traceback(tmp_path, capsys, monkeypatch):
    """A genuinely unparseable config prints which field is wrong and exits;
    every command stays usable, `logout` included (review of #61)."""
    config_dir = tmp_path / ".outlook-mcp"
    config_dir.mkdir()
    (config_dir / "config.json").write_text(json.dumps({"accounts": "not-a-list"}))
    monkeypatch.chdir(tmp_path)

    from outlook_mcp.config import load_config

    monkeypatch.setattr(cli, "load_config", lambda: load_config(config_dir=str(config_dir)))

    with pytest.raises(SystemExit) as exc:
        cli.cmd_logout()
    assert exc.value.code == 2
    out = capsys.readouterr().out
    assert "config.json is invalid" in out
    assert "Fix the file and run the command again" in out


def test_auth_with_unknown_account_name_exits_with_help(capsys, monkeypatch):
    _patch_config(
        monkeypatch,
        Config(accounts=[AccountConfig(name="net", client_id="id1-abcd")]),
    )
    with pytest.raises(SystemExit):
        cli.cmd_auth("ghost")
    out = capsys.readouterr().out
    assert "unknown account 'ghost'" in out
    assert "Configured accounts: net" in out
