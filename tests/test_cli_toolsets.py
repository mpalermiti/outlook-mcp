"""`outlook-mcp toolsets`: show the selection in force, or write config.json's `toolsets`.

Each test gets an empty settings directory of its own from conftest (also in
``OUTLOOK_MCP_CONFIG_DIR``), so the file written here is never the real one.
Writing sets that one key and leaves every other key, value and their order as
the person wrote them: no defaults are filled in, so an existing file does not
suddenly pin today's defaults.
"""

from __future__ import annotations

import json
import os
import stat
import sys
from pathlib import Path

import pytest

from outlook_mcp import cli, toolsets
from outlook_mcp.config import Config, load_config

DEFAULT_TOOLS = toolsets.loaded_tools(toolsets.resolve_groups(None, False))


@pytest.fixture(autouse=True)
def _no_selection_in_the_environment(monkeypatch):
    """CI runs the whole suite once with every group named in the variable."""
    monkeypatch.delenv("OUTLOOK_MCP_TOOLSETS", raising=False)


@pytest.fixture
def config_file() -> Path:
    return Path(cli.DEFAULT_CONFIG_DIR) / "config.json"


def _run(*args: str) -> int:
    try:
        cli.cmd_toolsets(list(args))
    except SystemExit as exc:
        return exc.code
    return 0


# ── Showing ───────────────────────────────────────────────────────────────


def test_showing_the_default(capsys):
    assert _run() == 0
    out = capsys.readouterr().out
    assert "classic" in out and "the default" in out
    assert f"{len(DEFAULT_TOOLS)} of {len(toolsets.TOOL_GROUPS)}" in out
    assert "account, admin, attachments" in out


def test_showing_the_file_selection(config_file, capsys):
    config_file.write_text(json.dumps({"toolsets": "todo"}), encoding="utf-8")
    assert _run() == 0
    out = capsys.readouterr().out
    assert 'config.json "toolsets"' in out
    assert "Groups:    account, todo" in out


def test_showing_the_environment_selection(config_file, capsys, monkeypatch):
    config_file.write_text(json.dumps({"toolsets": "todo"}), encoding="utf-8")
    monkeypatch.setenv("OUTLOOK_MCP_TOOLSETS", "calendar")
    assert _run() == 0
    out = capsys.readouterr().out
    assert "OUTLOOK_MCP_TOOLSETS" in out
    assert "Groups:    account, calendar" in out


def test_showing_names_an_unknown_name(capsys, monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_TOOLSETS", "mail,nothing_here")
    assert _run() == 0
    out = capsys.readouterr().out
    assert "'nothing_here'" in out and "not a tool group in this release" in out


def test_showing_a_broken_config_exits_with_the_repair(config_file, capsys):
    config_file.write_text("{nope", encoding="utf-8")
    assert _run() == 1
    assert "config file is invalid" in capsys.readouterr().err


# ── Writing ───────────────────────────────────────────────────────────────


def test_writing_creates_the_file_with_only_the_key(config_file, capsys):
    config_file.parent.rmdir()  # not even the directory yet

    assert _run("calendar,mail") == 0

    assert json.loads(config_file.read_text(encoding="utf-8")) == {"toolsets": "calendar,mail"}
    out = capsys.readouterr().out
    assert str(config_file) in out
    assert out.rstrip().endswith("Restart your MCP client to load the new tools.")


def test_the_intents_group_can_be_turned_on(config_file, capsys):
    """`classic,intents` is the default set plus outlook_find, the early preview."""
    config_file.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")

    assert _run("classic,intents") == 0

    written = json.loads(config_file.read_text(encoding="utf-8"))
    assert written == {"client_id": "c", "toolsets": "classic,intents"}
    capsys.readouterr()
    assert _run() == 0
    out = capsys.readouterr().out
    assert "intents" in out and "Unknown" not in out


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX modes")
def test_a_new_directory_is_0700_and_the_file_0600(config_file):
    config_file.parent.rmdir()
    assert _run("classic") == 0
    assert stat.S_IMODE(config_file.stat().st_mode) == 0o600
    assert stat.S_IMODE(config_file.parent.stat().st_mode) == 0o700


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX modes")
def test_an_existing_directory_keeps_the_mode_the_person_gave_it(config_file):
    config_file.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")
    config_file.chmod(0o644)
    os.chmod(config_file.parent, 0o750)
    assert _run("classic") == 0
    assert stat.S_IMODE(config_file.stat().st_mode) == 0o600
    assert stat.S_IMODE(config_file.parent.stat().st_mode) == 0o750


@pytest.mark.skipif(
    sys.platform == "win32" or os.geteuid() == 0, reason="POSIX modes, which root ignores"
)
def test_a_directory_locked_against_writing_is_not_unlocked(config_file, capsys):
    """Refused with the reason, not a traceback; the lock and the file stay as they were."""
    config_file.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")
    config_file.chmod(0o600)
    before = config_file.read_bytes()
    os.chmod(config_file.parent, 0o500)
    try:
        assert _run("mail") == 1
        assert stat.S_IMODE(config_file.parent.stat().st_mode) == 0o500
    finally:
        os.chmod(config_file.parent, 0o700)
    assert config_file.read_bytes() == before
    err = capsys.readouterr().err
    assert "Could not write" in err and "config.json was not changed" in err
    assert "Traceback" not in err


def test_a_write_that_fails_is_reported_without_a_traceback(config_file, capsys, monkeypatch):
    config_file.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")
    before = config_file.read_bytes()

    def full_disk(path, data):
        raise OSError(28, "No space left on device")

    monkeypatch.setattr(cli, "atomic_write", full_disk)
    assert _run("mail") == 1
    assert config_file.read_bytes() == before
    err = capsys.readouterr().err
    assert "No space left on device" in err and "config.json was not changed" in err


def test_a_value_json_cannot_write_back_is_refused(config_file, capsys):
    """`1e999` reads as infinity; writing it back would put `Infinity` in the file."""
    body = '{"client_id": "c", "note": 1e999}'
    config_file.write_text(body, encoding="utf-8")
    assert _run("mail") == 1
    assert config_file.read_text(encoding="utf-8") == body
    assert "Nothing was written" in capsys.readouterr().err


# ── A `toolsets` key of another type, written before the key existed ────────

BAD_VALUES = [["mail", "calendar"], 5, {"mail": True}]


@pytest.fixture
def bad_key_config(config_file, request):
    from outlook_mcp import config as config_mod

    request.getfixturevalue("monkeypatch").setattr(config_mod, "_toolsets_type_reported", False)

    def write(value):
        settings = {"client_id": "test-client", "read_only": True, "toolsets": value}
        config_file.write_text(json.dumps(settings), encoding="utf-8")
        return settings

    return write


@pytest.mark.parametrize("value", BAD_VALUES)
def test_showing_ignores_a_key_of_another_type(bad_key_config, capsys, value):
    bad_key_config(value)
    assert _run() == 0
    out = capsys.readouterr().out
    assert "classic, the default" in out
    assert f"{len(DEFAULT_TOOLS)} of {len(toolsets.TOOL_GROUPS)}" in out


@pytest.mark.parametrize("value", BAD_VALUES)
def test_the_command_replaces_a_key_of_another_type(bad_key_config, config_file, value):
    settings = bad_key_config(value)
    assert _run("classic") == 0
    written = json.loads(config_file.read_text(encoding="utf-8"))
    assert written == {**settings, "toolsets": "classic"}
    assert list(written) == list(settings)


@pytest.mark.parametrize("value", BAD_VALUES)
def test_status_and_auth_still_work_with_a_key_of_another_type(bad_key_config, capsys, value):
    from unittest.mock import patch

    bad_key_config(value)
    cli.cmd_status()
    out = capsys.readouterr().out
    assert "Client ID: test-cli..." in out and "read-only" in out
    assert "classic, the default" in out

    with patch("outlook_mcp.auth.AuthManager.login_interactive", lambda self: None):
        cli.cmd_auth()
    assert "Done." in capsys.readouterr().out


def test_writing_keeps_every_other_key_value_and_order(config_file):
    original = {
        "timezone": "Europe/London",
        "client_id": "test-client",
        "allow_categories": ["mail_drafts", "todo_write"],
        "read_only": False,
        "unknown_kept": {"nested": [1, 2.5, None]},
        "tenant_id": "consumers",
    }
    config_file.write_text(json.dumps(original), encoding="utf-8")

    assert _run("mail") == 0

    written = json.loads(config_file.read_text(encoding="utf-8"))
    assert list(written) == [*original, "toolsets"]
    assert written == {**original, "toolsets": "mail"}


def test_writing_replaces_an_existing_value_in_place(config_file):
    config_file.write_text(
        json.dumps({"client_id": "c", "toolsets": "todo", "timezone": "UTC"}), encoding="utf-8"
    )
    assert _run("classic") == 0
    written = json.loads(config_file.read_text(encoding="utf-8"))
    assert list(written.items()) == [
        ("client_id", "c"),
        ("toolsets", "classic"),
        ("timezone", "UTC"),
    ]


def test_writing_keeps_text_as_written_and_ends_with_a_newline(config_file):
    config_file.write_text(json.dumps({"timezone": "Europe/Zürich"}), encoding="utf-8")
    assert _run("mail") == 0
    text = config_file.read_text(encoding="utf-8")
    assert "Zürich" in text and text.endswith("}\n")
    assert (
        text
        == json.dumps(
            {"timezone": "Europe/Zürich", "toolsets": "mail"}, indent=2, ensure_ascii=False
        )
        + "\n"
    )


def test_names_are_normalised_and_may_be_given_as_several_arguments(config_file):
    assert _run(" Mail, CALENDAR", "todo", "mail") == 0
    assert json.loads(config_file.read_text(encoding="utf-8"))["toolsets"] == "mail,calendar,todo"


def test_what_was_written_loads_as_the_selection(config_file):
    assert _run("classic,documents") == 0
    config = load_config(cli.DEFAULT_CONFIG_DIR)
    assert toolsets.resolve_selection(None, config.toolsets) == (
        {"classic", "documents"},
        "config",
    )


@pytest.mark.parametrize("names", [("nothing_here",), ("mail,foo",), ("classic", "nothing_here")])
def test_an_unknown_name_is_refused_and_nothing_is_written(config_file, capsys, names):
    config_file.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")
    before = config_file.read_bytes()

    assert _run(*names) == 2

    assert config_file.read_bytes() == before
    err = capsys.readouterr().err
    assert "Nothing was written" in err
    assert "classic" in err and "mail" in err


@pytest.mark.parametrize("names", [("",), (" , ",)])
def test_an_empty_selection_is_refused(config_file, capsys, names):
    assert _run(*names) == 2
    assert not config_file.exists()
    assert "Nothing was written" in capsys.readouterr().err


def test_a_symlinked_config_is_refused_and_its_target_is_untouched(config_file, tmp_path, capsys):
    target = tmp_path / "elsewhere.json"
    target.write_text(json.dumps({"client_id": "c"}), encoding="utf-8")
    config_file.symlink_to(target)
    before = target.read_bytes()

    assert _run("mail") == 1

    assert target.read_bytes() == before and config_file.is_symlink()
    assert "symlink" in capsys.readouterr().err


def test_a_config_that_is_not_a_regular_file_is_refused(config_file, capsys):
    config_file.mkdir()
    assert _run("mail") == 1
    assert config_file.is_dir()
    assert "not a regular file" in capsys.readouterr().err


@pytest.mark.parametrize(
    "body",
    [
        "{nope",
        json.dumps({"read_only_consent": True}),
        '{"timezone": "UTC", "timezone": "Europe/London"}',
    ],
)
def test_a_config_that_will_not_load_exits_1_with_the_repair(config_file, capsys, body):
    config_file.write_text(body, encoding="utf-8")

    assert _run("mail") == 1

    assert config_file.read_text(encoding="utf-8") == body
    err = capsys.readouterr().err
    assert "config file is invalid" in err and "Traceback" not in err


def test_a_config_saved_in_a_code_page_is_refused_until_resaved(config_file, capsys, monkeypatch):
    """The loader reads it (as releases before #99 did); rewriting it is another matter."""
    from outlook_mcp import config as config_mod

    monkeypatch.setattr(config_mod.locale, "getpreferredencoding", lambda *a, **k: "cp1252")
    body = '{"timezone": "Europe/Paris", "client_id": "café"}'.encode("cp1252")
    config_file.write_bytes(body)

    assert _run("mail") == 1

    assert config_file.read_bytes() == body
    assert "UTF-8" in capsys.readouterr().err


def test_the_environment_note_never_prints_the_value(config_file, capsys, monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_TOOLSETS", "secretish,mail")
    assert _run("todo") == 0
    out = capsys.readouterr().out
    assert "OUTLOOK_MCP_TOOLSETS is set here and overrides this file." in out
    assert "secretish" not in out


@pytest.mark.parametrize("value", ["", "   "])
def test_a_blank_variable_does_not_override_so_there_is_no_note(capsys, monkeypatch, value):
    monkeypatch.setenv("OUTLOOK_MCP_TOOLSETS", value)
    assert _run("todo") == 0
    assert "overrides" not in capsys.readouterr().out


# ── status, and the dispatcher ─────────────────────────────────────────────


def test_status_shows_the_source_the_groups_and_the_count(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id", toolsets="todo"))
    cli.cmd_status()
    out = capsys.readouterr().out
    assert 'config.json "toolsets"' in out
    assert "Groups:    account, todo" in out
    todo = toolsets.loaded_tools({"todo"})
    assert f"{len(todo)} of {len(toolsets.TOOL_GROUPS)}" in out


def test_status_shows_the_default_with_rule_writes(capsys, monkeypatch):
    monkeypatch.setattr(
        cli, "load_config", lambda: Config(client_id="test-id", mailbox_settings_write=True)
    )
    cli.cmd_status()
    out = capsys.readouterr().out
    assert "classic" in out and "rules" in out
    with_rules = toolsets.loaded_tools(toolsets.resolve_groups(None, True))
    assert f"{len(with_rules)} of {len(toolsets.TOOL_GROUPS)}" in out


def test_the_command_is_dispatched_and_does_not_start_the_server(monkeypatch):
    calls = []
    monkeypatch.setattr(cli, "cmd_toolsets", lambda args: calls.append(args))
    monkeypatch.setattr(cli, "cmd_serve", lambda: calls.append("serve"))
    monkeypatch.setattr(sys, "argv", ["outlook-mcp", "toolsets", "mail", "todo"])

    cli.main()

    assert calls == [["mail", "todo"]]


def test_the_usage_lists_the_command(capsys):
    cli._print_usage()
    assert "toolsets" in capsys.readouterr().out
