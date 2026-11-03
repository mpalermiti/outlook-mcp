"""Tests for config management."""

import codecs
import json
import locale
import logging
import os
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest
from pydantic import ValidationError

from outlook_mcp.config import Config, config_repair_lines, load_config, save_config


def test_default_config():
    """Default config has sensible values."""
    config = Config()
    assert config.client_id is None
    assert config.tenant_id == "consumers"
    assert config.read_only is False
    assert config.timezone == "UTC"


def test_config_dir_created(tmp_path, monkeypatch):
    """Config directory is created on save."""
    config_dir = tmp_path / ".outlook-mcp"
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(config_dir))
    save_config(Config(), config_dir=str(config_dir))
    assert config_dir.exists()


@pytest.mark.skipif(
    sys.platform == "win32",
    reason="os.chmod on Windows honours only the read-only attribute, so 0o700 is not "
    "representable there (measured: 0o777). The POSIX guarantee is real and stays asserted "
    "where it holds (#85).",
)
def test_config_dir_is_restricted_to_its_owner(tmp_path):
    """The directory holds a token-adjacent file, so it is 0700 where modes are enforceable."""
    config_dir = tmp_path / ".outlook-mcp"
    save_config(Config(), config_dir=str(config_dir))
    assert oct(config_dir.stat().st_mode & 0o777) == "0o700"


@pytest.mark.skipif(
    sys.platform == "win32",
    reason="os.chmod on Windows honours only the read-only attribute, so 0o600 is not "
    "representable there (measured: 0o666). The POSIX guarantee is real and stays asserted "
    "where it holds (#85).",
)
def test_config_file_permissions(tmp_path):
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
    home_settings = os.path.abspath(os.path.expanduser("~/.outlook-mcp"))

    assert _paths_with_env(None)[0] == home_settings
    assert _paths_with_env("")[0] == home_settings
    assert _paths_with_env(None)[2] == os.path.join(home_settings, "attachments")
    assert _paths_with_env("")[2] == os.path.join(home_settings, "attachments")


def test_a_relative_override_is_anchored_not_followed_around(tmp_path):
    """`OUTLOOK_MCP_CONFIG_DIR=net` means the directory the launcher was in.

    Left relative, the value would name a different settings directory in
    the terminal that runs `outlook-mcp auth` and the client that starts the
    server — and the server would ask for re-authentication forever.
    """
    env = dict(os.environ)
    env["OUTLOOK_MCP_CONFIG_DIR"] = "net-instance"
    out = subprocess.run(
        [sys.executable, "-c", "from outlook_mcp import config; print(config.DEFAULT_CONFIG_DIR)"],
        env=env,
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        check=True,
    )

    assert out.stdout.strip() == str(tmp_path / "net-instance")


def test_load_refuses_an_override_that_is_a_file(tmp_path):
    """A settings path that exists but is a file fails at load, with the fix.

    Refused here rather than as a FileExistsError later, when part of a flow
    has already run. The ValueError arm is the one the server and CLI turn
    into a clean exit with this message.
    """
    not_a_dir = tmp_path / "settings.txt"
    not_a_dir.write_text("occupied")

    with pytest.raises(ValueError) as exc:
        load_config(config_dir=str(not_a_dir))

    assert "not a directory" in str(exc.value)
    assert "OUTLOOK_MCP_CONFIG_DIR" in str(exc.value)


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


@pytest.mark.skipif(
    sys.platform == "win32",
    reason="the repair asserts a POSIX mode, and os.chmod on Windows cannot set 0o600 (#85).",
)
def test_a_loose_mode_is_repaired_on_load(tmp_path):
    """The hardening is self-repairing where the bits apply: 0o644 on disk, 0o600 after a load."""
    config_dir = tmp_path / ".outlook-mcp"
    save_config(Config(), config_dir=str(config_dir))
    config_file = config_dir / "config.json"

    os.chmod(config_file, 0o644)
    assert oct(config_file.stat().st_mode & 0o777) == "0o644"  # premise

    load_config(config_dir=str(config_dir))
    assert oct(config_file.stat().st_mode & 0o777) == "0o600"


def test_the_repair_is_skipped_where_the_mode_bits_do_not_apply(tmp_path, monkeypatch):
    """On Windows the mode never reads back as 0o600, so the check could never converge.

    It re-chmodded the file on every single load instead — the defect in #85. This runs on
    every platform: the branch is driven by monkeypatching the expression the product
    actually evaluates, not by the host it happens to run on.
    """
    config_dir = tmp_path / ".outlook-mcp"
    save_config(Config(client_id="sentinel-from-disk"), config_dir=str(config_dir))
    config_file = config_dir / "config.json"

    os.chmod(config_file, 0o644)  # a no-op on Windows, which is the point
    assert oct(config_file.stat().st_mode & 0o777) != "0o600"  # premise, on both platforms

    calls = []
    monkeypatch.setattr(Path, "chmod", lambda self, mode, **kw: calls.append((self, mode)))
    monkeypatch.setattr(sys, "platform", "win32")

    # The sentinel, not a default: Config().tenant_id is "consumers" too, so asserting that
    # would have passed against a file that was never read at all. Twice, because the defect
    # is per load rather than on the first one.
    for _ in range(2):
        assert load_config(config_dir=str(config_dir)).client_id == "sentinel-from-disk"

    assert calls == []


def test_the_repair_still_runs_where_the_mode_bits_apply(tmp_path, monkeypatch):
    """The mirror of the test above: the guard must not disable the repair everywhere.

    Runs on every platform for the same reason, and is what fails if the guard is ever
    inverted — a direction the POSIX outcome test cannot demonstrate on Windows, where it
    is skipped.
    """
    config_dir = tmp_path / ".outlook-mcp"
    save_config(Config(), config_dir=str(config_dir))
    config_file = config_dir / "config.json"

    os.chmod(config_file, 0o644)  # a no-op on Windows; either way the mode is not 0o600
    assert oct(config_file.stat().st_mode & 0o777) != "0o600"  # premise, on both platforms

    calls = []
    monkeypatch.setattr(Path, "chmod", lambda self, mode, **kw: calls.append((self, mode)))
    monkeypatch.setattr(sys, "platform", "linux")

    load_config(config_dir=str(config_dir))

    assert calls == [(config_file, 0o600)]


# ── config.json is UTF-8, whatever the machine's locale (#99) ──
#
# The locale encoding is cp1252 on a typical Windows install and UTF-8 on macOS and Linux, so
# the first test below can only fail on a host whose locale is not UTF-8. The fallback tests
# monkeypatch the locale, and the EncodingWarning subprocess guard does not depend on it at all;
# those are the ones that can fail on a UTF-8 CI runner.

# attachments_dir must be absolute, so these encoding tests start their values at a drive
# on Windows and at the root elsewhere; what they test is the encoding, not the place.
_ABS = "C:/" if sys.platform == "win32" else "/"
_NON_ASCII_DIR = _ABS + "Users/Zoë/中文/att"


def _write_config_bytes(config_dir, raw: bytes) -> None:
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "config.json").write_bytes(raw)


def _pretend_the_locale_is(monkeypatch, encoding: str) -> None:
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: encoding)


def test_a_non_ascii_value_round_trips_as_utf8_on_disk(tmp_path):
    """Save then load keeps the value, and what is on disk is plain UTF-8 with no BOM.

    The bytes are asserted too, so a reader and writer sharing one wrong encoding cannot
    pass by agreeing with each other. `中文` is outside cp1252: before #99 the save itself
    raised UnicodeEncodeError on Windows.
    """
    config_dir = tmp_path / ".outlook-mcp"
    save_config(Config(client_id="x", attachments_dir=_NON_ASCII_DIR), config_dir=str(config_dir))

    raw = (config_dir / "config.json").read_bytes()
    assert not raw.startswith(codecs.BOM_UTF8)
    assert _NON_ASCII_DIR in raw.decode("utf-8")
    assert load_config(config_dir=str(config_dir)).attachments_dir == _NON_ASCII_DIR


def test_a_utf8_config_written_elsewhere_reads_the_same(tmp_path, monkeypatch, caplog):
    """A hand-edited or copied UTF-8 config means the same thing on a cp1252 machine.

    Before #99 this read `Zoë` as `ZoÃ«` on Windows — a different attachments directory,
    silently. The locale is pinned to cp1252 so the legacy fallback is armed: UTF-8 must still
    win, with no migration warning.
    """
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    payload = json.dumps({"attachments_dir": _NON_ASCII_DIR}, ensure_ascii=False)
    _write_config_bytes(config_dir, payload.encode("utf-8"))

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.attachments_dir == _NON_ASCII_DIR
    assert caplog.records == []


def test_a_utf8_bom_is_accepted(tmp_path):
    """Windows PowerShell 5.1's `Set-Content -Encoding utf8` writes a byte-order mark."""
    config_dir = tmp_path / ".outlook-mcp"
    _write_config_bytes(config_dir, codecs.BOM_UTF8 + b'{"client_id": "from-bom-file"}')

    assert load_config(config_dir=str(config_dir)).client_id == "from-bom-file"


def test_a_legacy_ansi_config_still_loads_and_asks_for_a_resave(tmp_path, monkeypatch, caplog):
    """Windows PowerShell 5.1's `Set-Content` writes the ANSI code page by default.

    Such a file with an accented value loaded correctly before #99, so it keeps loading, read
    in the machine's own code page exactly as before, and the warning asks for a re-save.
    """
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    _write_config_bytes(
        config_dir, f'{{"attachments_dir": "{_ABS}Users/Zoë/att"}}'.encode("cp1252")
    )

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.attachments_dir == _ABS + "Users/Zoë/att"
    warnings = [r.getMessage() for r in caplog.records]
    assert len(warnings) == 1
    assert "cp1252" in warnings[0]
    assert "UTF-8" in warnings[0]


def test_no_legacy_fallback_where_the_locale_is_already_utf8(tmp_path, monkeypatch):
    """Where the old code read UTF-8 too, the same bytes failed before and fail now."""
    _pretend_the_locale_is(monkeypatch, "UTF-8")
    config_dir = tmp_path / ".outlook-mcp"
    _write_config_bytes(
        config_dir, f'{{"attachments_dir": "{_ABS}Users/Zoë/att"}}'.encode("cp1252")
    )

    with pytest.raises(UnicodeDecodeError):
        load_config(config_dir=str(config_dir))


def test_a_utf16_config_is_refused_not_misread_as_ansi(tmp_path, monkeypatch):
    """`>` and `Out-File` in Windows PowerShell 5.1 write UTF-16LE with a BOM.

    Every byte of that is valid cp1252, so the fallback decodes it. The result, `ÿþ{` with
    NULs between the characters, can never validate. So the operator gets the UTF-8 error and
    its re-save remedy, where before #99 a Windows install reported invalid JSON.
    """
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    raw = codecs.BOM_UTF16_LE + '{"client_id": "x"}'.encode("utf-16-le")
    _write_config_bytes(config_dir, raw)

    with pytest.raises(UnicodeDecodeError):
        load_config(config_dir=str(config_dir))


def test_the_fallback_only_accepts_a_valid_config(tmp_path, monkeypatch):
    """Bytes that decode in the code page but do not validate get the UTF-8 error, not a
    validation error about text the operator never wrote in that encoding."""
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    raw = f'{{"attachments_dir": "{_ABS}Zoë", "allow_categories": ["not-a-category"]}}'
    raw = raw.encode("cp1252")
    _write_config_bytes(config_dir, raw)

    with pytest.raises(UnicodeDecodeError):
        load_config(config_dir=str(config_dir))


def test_an_unknown_locale_codec_means_no_fallback(tmp_path, monkeypatch):
    _pretend_the_locale_is(monkeypatch, "no-such-codec")
    config_dir = tmp_path / ".outlook-mcp"
    _write_config_bytes(config_dir, f'{{"attachments_dir": "{_ABS}Zoë"}}'.encode("cp1252"))

    with pytest.raises(UnicodeDecodeError):
        load_config(config_dir=str(config_dir))


def test_a_legacy_file_that_is_also_valid_utf8_is_read_as_utf8(tmp_path, monkeypatch, caplog):
    """The accepted limit of a UTF-8 config file, pinned so it is a decision and not a surprise.

    The cp1252 encoding of the literal text `ZoÃ«` is `5a 6f c3 ab`, which is also UTF-8 for
    `Zoë`. Bytes alone cannot tell the two apart, so it reads as UTF-8 with no warning. A strict
    UTF-8 reader with no fallback reads it identically; the fallback neither causes this case
    nor can detect it.
    """
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    raw = f'{{"attachments_dir": "{_ABS}ZoÃ«"}}'.encode("cp1252")
    # premise: the bytes are ambiguous
    assert raw.decode("utf-8") == f'{{"attachments_dir": "{_ABS}Zoë"}}'
    _write_config_bytes(config_dir, raw)

    with caplog.at_level(logging.WARNING, logger="outlook_mcp.config"):
        loaded = load_config(config_dir=str(config_dir))

    assert loaded.attachments_dir == _ABS + "Zoë"
    assert caplog.records == []


def test_an_undecodable_config_gets_the_utf8_remedy_not_the_permissions_one():
    """UnicodeDecodeError is a ValueError, and used to fall into the generic branch, which
    tells the operator to check ownership and permissions: the wrong fix."""
    decode_lines = config_repair_lines(
        UnicodeDecodeError("utf-8", b"\xff", 0, 1, "invalid start byte")
    )
    assert decode_lines[0].startswith("Cannot load the config file")
    assert any("UTF-8" in line for line in decode_lines)
    assert not any("readable and owned" in line for line in decode_lines)

    # The generic branch keeps its remedy for the causes it is true of.
    # EIO, not EACCES: OSError(13, ...) constructs a PermissionError, which is the symlink branch.
    os_lines = config_repair_lines(OSError(5, "Input/output error"))
    assert any("readable and owned" in line for line in os_lines)


_ENCODING_GUARD = textwrap.dedent(
    """
    import sys
    import warnings
    warnings.filterwarnings("error", category=EncodingWarning, module=r"outlook_mcp(\\.|$)")

    from azure.identity import AuthenticationRecord
    from outlook_mcp import auth
    from outlook_mcp.config import Config, load_config, save_config

    value = ("C:/" if sys.platform == "win32" else "/") + "Users/Zoë/中文/att"
    save_config(Config(client_id="x", attachments_dir=value))
    assert load_config().attachments_dir == value, "config did not round-trip"

    auth._save_auth_record(
        AuthenticationRecord("tenant", "client", "login.example", "home", "zoë@example.com")
    )
    # _load_auth_record swallows every exception and returns None, so an EncodingWarning
    # promoted to an error there is only visible through what it returns.
    loaded = auth._load_auth_record()
    assert loaded is not None and loaded.username == "zoë@example.com", "record did not load"
    """
)


def test_settings_files_never_use_the_locale_encoding(tmp_path):
    """Every settings read and write names its encoding — checked on any host.

    `-X warn_default_encoding` makes Python emit EncodingWarning wherever text I/O falls back
    to the locale encoding, whatever that encoding is, and the child promotes it to an error
    for outlook_mcp's own modules. That makes this the #99 test that fails on a UTF-8 Linux
    runner: drop `encoding=` from the config read, the shared writer or the auth record read
    and it reddens there too.
    """
    env = dict(os.environ)
    env["OUTLOOK_MCP_CONFIG_DIR"] = str(tmp_path / "settings")
    proc = subprocess.run(
        [sys.executable, "-X", "warn_default_encoding", "-c", _ENCODING_GUARD],
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )

    assert proc.returncode == 0, proc.stderr


# ── read_only_consent: a sign-in that asks for the read scopes only ──────
#
# `read_only` is a tool gate. This is the other half — what Microsoft is asked
# for at sign-in — and it has to be a separate, explicit key: narrowing the
# consent on `read_only` alone strands every write the day that flag is flipped.


def test_read_only_consent_is_off_unless_asked_for():
    assert Config().read_only_consent is False


def test_read_only_consent_without_read_only_is_refused():
    """A read-only sign-in cannot write, so a config that expects writes is a mistake.

    Left to load, every write tool would fail with a 403 whose hint points
    somewhere else entirely. Refusing at load names the two keys instead.
    """
    from pydantic import ValidationError

    with pytest.raises(ValidationError) as exc:
        Config(client_id="x", read_only_consent=True)

    message = str(exc.value)
    assert "read_only_consent: true needs read_only: true" in message


def test_read_only_consent_with_read_only_loads(tmp_path):
    (tmp_path / "config.json").write_text(
        '{"client_id": "x", "read_only": true, "read_only_consent": true}', encoding="utf-8"
    )

    loaded = load_config(config_dir=str(tmp_path))

    assert loaded.read_only is True
    assert loaded.read_only_consent is True


def test_the_repair_for_a_lone_read_only_consent_names_both_keys(tmp_path):
    from pydantic import ValidationError

    from outlook_mcp.config import config_repair_lines

    (tmp_path / "config.json").write_text(
        '{"client_id": "x", "read_only_consent": true}', encoding="utf-8"
    )
    with pytest.raises(ValidationError) as exc:
        load_config(config_dir=str(tmp_path))

    repair = "\n".join(config_repair_lines(exc.value))
    assert "read_only_consent: true needs read_only: true" in repair


# ── A misspelt safety key refuses to load ─────────────────────────────────
#
# Unknown keys are accepted with a warning, so a config written for a newer
# release still boots an older one. For three keys that is the wrong way round:
# `read_only`, `allow_categories` and `read_only_consent` all default to the
# open side, so `"readOnly": true` — ignored with a line on stderr nobody reads —
# ran a fully writable server while its operator believed it read-only. A key
# that is plainly one of those three, misspelt, is refused instead: the server
# does not start, and the repair names the key it meant.


@pytest.mark.parametrize(
    ("written", "meant"),
    [
        ("readOnly", "read_only"),
        ("read-only", "read_only"),
        ("readonly", "read_only"),
        ("READ_ONLY", "read_only"),
        ("read_onyl", "read_only"),
        ("allowCategories", "allow_categories"),
        ("allow_category", "allow_categories"),
        ("allow-categories", "allow_categories"),
        ("readOnlyConsent", "read_only_consent"),
        ("read_only_consnet", "read_only_consent"),
    ],
)
def test_a_misspelt_safety_key_refuses_to_load(tmp_path, written, meant):
    from pydantic import ValidationError

    (tmp_path / "config.json").write_text(
        json.dumps({"client_id": "x", written: True if "categor" not in written.lower() else []}),
        encoding="utf-8",
    )

    with pytest.raises(ValidationError) as exc:
        load_config(config_dir=str(tmp_path))

    repair = "\n".join(config_repair_lines(exc.value))
    assert repair.count(repr(written)) >= 1
    assert f"'{meant}'" in repair


@pytest.mark.parametrize("written", ["theme", "time_zone", "client-id", "accounts"])
def test_other_unknown_keys_still_load_with_a_warning(tmp_path, caplog, written):
    """Forward compatibility stays for everything that is not a safety key."""
    (tmp_path / "config.json").write_text(
        json.dumps({"client_id": "x", written: "anything"}), encoding="utf-8"
    )

    with caplog.at_level(logging.WARNING):
        loaded = load_config(config_dir=str(tmp_path))

    assert loaded.client_id == "x"
    assert written in caplog.text


def test_the_correctly_spelt_safety_keys_load(tmp_path):
    (tmp_path / "config.json").write_text(
        json.dumps(
            {
                "client_id": "x",
                "read_only": True,
                "read_only_consent": True,
                "allow_categories": ["calendar_write"],
            }
        ),
        encoding="utf-8",
    )

    loaded = load_config(config_dir=str(tmp_path))

    assert loaded.read_only is True
    assert loaded.read_only_consent is True
    assert loaded.allow_categories == ["calendar_write"]


class TestMailboxSettingsWriteConfig:
    def test_refused_when_paired_with_read_only_consent(self):
        """A key the sign-in would silently ignore is a config error, not a no-op."""
        with pytest.raises(ValidationError):
            Config(
                client_id="test",
                read_only=True,
                read_only_consent=True,
                mailbox_settings_write=True,
            )

    def test_allowed_with_plain_read_only(self):
        """read_only gates the tools; the token still asks for the write scopes."""
        cfg = Config(client_id="test", read_only=True, mailbox_settings_write=True)
        assert cfg.mailbox_settings_write is True

    def test_allowed_on_its_own(self):
        cfg = Config(client_id="test", mailbox_settings_write=True)
        assert cfg.mailbox_settings_write is True


class TestWarningsAreLoggedOnce:
    """The import-time peek reads the config; the start-up load reports it.

    Exactly one warning must reach the operator: an import-time read that warns
    printed every notice twice on every start, and `warn_unknown=False` only
    covered the unknown-key ones.
    """

    @staticmethod
    def _point_server_at(monkeypatch, config_dir):
        import outlook_mcp.server as server
        from outlook_mcp import config as config_mod

        real = config_mod.load_config
        monkeypatch.setattr(server, "load_config", lambda **kw: real(str(config_dir), **kw))
        return server, real

    def test_an_unknown_key_is_reported_once(self, tmp_path, caplog, monkeypatch):
        (tmp_path / "config.json").write_text(
            json.dumps({"client_id": "00000000-0000-0000-0000-000000000001", "unknown_key_xyz": 1}),
            encoding="utf-8",
        )
        server, real = self._point_server_at(monkeypatch, tmp_path)

        with caplog.at_level(logging.WARNING):
            assert server._mailbox_settings_write_is_on() is False
            real(str(tmp_path))

        hits = [r for r in caplog.records if "unknown_key_xyz" in r.getMessage()]
        assert len(hits) == 1, [r.getMessage() for r in caplog.records]

    def test_the_legacy_encoding_note_is_reported_once(self, tmp_path, caplog, monkeypatch):
        from outlook_mcp import config as config_mod

        # A cp1252 config, as Windows PowerShell 5.1 writes one: the accented
        # value only decodes in the code page, so the loader rescues it and warns.
        body = json.dumps(
            {"client_id": "00000000-0000-0000-0000-000000000001", "timezone": "Eur\u00f3pe/Paris"}
        )
        (tmp_path / "config.json").write_bytes(body.replace("\\u00f3", "\u00f3").encode("cp1252"))
        monkeypatch.setattr(config_mod.locale, "getpreferredencoding", lambda *a, **k: "cp1252")
        server, real = self._point_server_at(monkeypatch, tmp_path)

        with caplog.at_level(logging.WARNING):
            server._mailbox_settings_write_is_on()
            real(str(tmp_path))

        hits = [r for r in caplog.records if "not UTF-8" in r.getMessage()]
        assert len(hits) == 1, [r.getMessage() for r in caplog.records]


# ── attachments_dir: an absolute folder of its own ───────────────────────
#
# The attachment tools read and write every file under attachments_dir, so it
# must not be the home directory, a folder that contains it, or the settings
# directory (which holds config.json and the sign-in record). A subfolder of
# either is fine; the default is the `attachments` folder in the settings
# directory. A relative value would name a different folder in every working
# directory the server is started from.


@pytest.fixture
def home_and_settings(tmp_path, monkeypatch):
    """A home and a settings directory of the test's own, in separate trees."""
    from outlook_mcp import config as config_mod

    home = tmp_path / "users" / "me"
    settings = tmp_path / "apps" / "outlook-mcp"
    home.mkdir(parents=True)
    settings.mkdir(parents=True)
    monkeypatch.setenv("HOME", str(home))
    monkeypatch.setenv("USERPROFILE", str(home))  # what expanduser reads on Windows
    monkeypatch.setattr(config_mod, "DEFAULT_CONFIG_DIR", str(settings))
    return home, settings


def _attachments_value(which: str, home: Path, settings: Path, tmp_path: Path) -> str:
    return {
        "empty": "",
        "relative": "relative/dir",
        "tilde": "~",
        "home realpath": os.path.realpath(home),
        # The root of home's own drive: on Windows a bare `/` has no drive, so it
        # would only ever meet the "absolute" rule, never the home one.
        "filesystem root": home.anchor,
        "folder above home": str(home.parent),
        "settings dir": str(settings),
        "folder above settings": str(settings.parent),
        "home subfolder": "~/Downloads/x",
        "settings subfolder": str(settings / "attachments"),
        "elsewhere": str(tmp_path / "elsewhere"),
    }[which]


def _load_with_attachments_dir(tmp_path: Path, value: str):
    config_dir = tmp_path / "config-file"
    _write_config(config_dir, {"client_id": "private-client-sentinel", "attachments_dir": value})
    return load_config(config_dir=str(config_dir))


# Each refusal names the rule that fired, in the validator's own words.
_NOT_ABSOLUTE = "attachments_dir must be an absolute folder path"
_HOLDS_HOME = "attachments_dir must not be the home directory or a folder that contains it"
_HOLDS_SETTINGS = "attachments_dir must not be the settings directory or a folder that contains it"
_RULE = {"absolute": _NOT_ABSOLUTE, "home": _HOLDS_HOME, "settings": _HOLDS_SETTINGS}


@pytest.mark.parametrize(
    "which,rule",
    [
        ("empty", "absolute"),
        ("relative", "absolute"),
        ("tilde", "home"),
        ("home realpath", "home"),
        ("filesystem root", "home"),
        ("folder above home", "home"),
        ("settings dir", "settings"),
        ("folder above settings", "settings"),
    ],
)
def test_attachments_dir_that_is_not_a_folder_of_its_own_is_refused(
    tmp_path, home_and_settings, which, rule
):
    value = _attachments_value(which, *home_and_settings, tmp_path)

    with pytest.raises(ValidationError) as exc:
        _load_with_attachments_dir(tmp_path, value)

    repair = config_repair_lines(exc.value)
    assert repair[0] == "The config file is invalid — the server cannot start:"
    [line] = [line for line in repair if line.startswith("  attachments_dir: ")]
    assert _RULE[rule] in line
    assert "private-client-sentinel" not in "\n".join(repair)
    assert "private-client-sentinel" not in str(exc.value)


@pytest.mark.parametrize("which", ["home subfolder", "settings subfolder", "elsewhere"])
def test_attachments_dir_inside_home_settings_or_elsewhere_loads_unchanged(
    tmp_path, home_and_settings, which
):
    value = _attachments_value(which, *home_and_settings, tmp_path)

    loaded = _load_with_attachments_dir(tmp_path, value)

    assert loaded.attachments_dir == value
    # Checked without creating anything.
    home, settings = home_and_settings
    assert list(home.iterdir()) == [] and list(settings.iterdir()) == []
    assert not (tmp_path / "elsewhere").exists()


@pytest.mark.parametrize("target", ["home", "settings"])
def test_attachments_dir_symlinked_to_home_or_settings_is_refused(
    tmp_path, home_and_settings, target
):
    home, settings = home_and_settings
    link = tmp_path / "innocent-looking"
    try:
        link.symlink_to(home if target == "home" else settings, target_is_directory=True)
    except OSError:
        pytest.skip("symlink privilege unavailable")

    with pytest.raises(ValidationError, match=_RULE[target]):
        _load_with_attachments_dir(tmp_path, str(link))


@pytest.mark.parametrize("target", ["home", "settings"])
def test_attachments_dir_that_reaches_home_or_settings_by_another_route_is_refused(
    tmp_path, home_and_settings, target
):
    """A folder can hold home without being above it in the path.

    macOS shows the whole Data volume a second time under /System/Volumes/Data,
    so home is there too, though no parent of home is that folder. A folder
    with a link to home's parent inside it stands in for that here.
    """
    home, settings = home_and_settings
    parent = (home if target == "home" else settings).parent
    volume = tmp_path / "volume"
    volume.mkdir()
    try:
        (volume / parent.name).symlink_to(parent, target_is_directory=True)
    except OSError:
        pytest.skip("symlink privilege unavailable")

    with pytest.raises(ValidationError, match=_RULE[target]):
        _load_with_attachments_dir(tmp_path, str(volume))


def test_the_macos_data_volume_is_refused_because_home_is_on_it(tmp_path, home_and_settings):
    home, _ = home_and_settings
    data_volume = "/System/Volumes/Data"
    home_again = data_volume + os.path.realpath(home)
    if not (os.path.isdir(home_again) and os.path.samefile(home_again, home)):
        pytest.skip("no macOS Data volume holding home here")

    with pytest.raises(ValidationError, match=_HOLDS_HOME):
        _load_with_attachments_dir(tmp_path, data_volume)


def test_attachments_dir_spelt_in_another_case_is_still_home(tmp_path, home_and_settings):
    """On a case-insensitive filesystem, two spellings can name one directory."""
    home, _ = home_and_settings
    respelt = os.path.realpath(home).swapcase()
    if not (os.path.isdir(respelt) and os.path.samefile(respelt, home)):
        pytest.skip("this filesystem is case-sensitive")

    with pytest.raises(ValidationError, match=_HOLDS_HOME):
        _load_with_attachments_dir(tmp_path, respelt)


def test_the_default_attachments_dir_is_checked_too(tmp_path, home_and_settings):
    home, settings = home_and_settings
    no_file = str(tmp_path / "no-config-file")
    assert load_config(config_dir=no_file).attachments_dir == str(settings / "attachments")

    try:
        (settings / "attachments").symlink_to(home, target_is_directory=True)
    except OSError:
        pytest.skip("symlink privilege unavailable")
    with pytest.raises(ValidationError, match=_HOLDS_HOME):
        load_config(config_dir=no_file)


@pytest.mark.parametrize("value", ["/attachments", "\\attachments", "C:attachments"])
def test_on_windows_attachments_dir_needs_a_drive_or_share(monkeypatch, value):
    """A path rooted without a drive, or a drive without a root, depends on the current
    drive or directory; Python before 3.13 calls the first one absolute."""
    from outlook_mcp import config as config_mod

    monkeypatch.setattr(config_mod, "_WINDOWS", True)
    with pytest.raises(ValidationError, match="absolute"):
        Config(attachments_dir=value)


def test_on_windows_a_network_attachments_dir_is_judged_by_its_text(monkeypatch):
    """Resolving a share path connects to its host, so loading the config never does."""
    from outlook_mcp import config as config_mod

    share = "\\\\files.example\\share\\attachments"
    touched = []
    real_realpath, real_stat = os.path.realpath, os.stat

    def recording_realpath(path, *args, **kwargs):
        touched.append(("realpath", str(path)))
        return real_realpath(path, *args, **kwargs)

    def recording_stat(path, *args, **kwargs):
        touched.append(("stat", str(path)))
        return real_stat(path, *args, **kwargs)

    monkeypatch.setattr(config_mod, "_WINDOWS", True)
    with monkeypatch.context() as during:
        during.setattr(os.path, "realpath", recording_realpath)
        during.setattr(os, "stat", recording_stat)
        loaded = Config(attachments_dir=share)

    assert loaded.attachments_dir == share
    assert not [call for call in touched if call[1].startswith(("\\\\", "//"))]


async def test_a_refused_attachments_dir_boots_the_server_read_only_with_the_repair(
    tmp_path, home_and_settings, monkeypatch
):
    """The refusal takes the existing failure path: read-only, the repair on every call.

    The default folder is the hard case: the server's fail-safe substitute is built
    from defaults, so it must not be refused all over again.
    """
    from mcp.client import Client

    from outlook_mcp import server

    home, settings = home_and_settings
    try:
        (settings / "attachments").symlink_to(home, target_is_directory=True)
    except OSError:
        pytest.skip("symlink privilege unavailable")
    no_file = str(tmp_path / "no-config-file")
    monkeypatch.setattr(server, "load_config", lambda: load_config(config_dir=no_file))
    monkeypatch.setattr(
        server.AuthManager,
        "try_cached_token",
        lambda _: pytest.fail("a refused configuration must not try credentials"),
    )

    async with Client(server.mcp) as client:
        status = await client.call_tool("outlook_auth_status", {})
        whoami = await client.call_tool("outlook_whoami", {})

    payload = json.loads("".join(part.text for part in status.content if hasattr(part, "text")))
    assert payload["read_only"] is True and payload["authenticated"] is False
    assert "attachments_dir" in payload["action_required"]
    assert whoami.is_error is True
    assert "attachments_dir" in " ".join(p.text for p in whoami.content if hasattr(p, "text"))


# ── A key written twice in config.json is refused ────────────────────────
#
# JSON parsers keep one of the two values without a word, so the file does not
# say which one the operator meant — and the one that silently loses may be the
# restriction. The refusal names the key, never either value.


@pytest.mark.parametrize(
    "text",
    [
        '{"client_id": "x", "read_only": true, "read_only": false}',
        '{"client_id": "x", "read_only": true, "read\\u005fonly": false}',
        '{"client_id": "x", "theme": {"read_only": "private-first", '
        '"read_only": "private-second"}}',
        '{"client_id": "x", "theme": [{"a": 1}, {"read_only": "private-first", '
        '"read_only": "private-second"}]}',
    ],
    ids=["top-level", "escaped-spelling", "nested", "inside-a-list"],
)
def test_a_repeated_key_refuses_to_load_and_names_only_the_key(tmp_path, text):
    (tmp_path / "config.json").write_text(text, encoding="utf-8")

    with pytest.raises(ValidationError) as exc:
        load_config(config_dir=str(tmp_path))

    repair = "\n".join(config_repair_lines(exc.value))
    assert repair.startswith("The config file is invalid — the server cannot start:")
    assert "'read_only'" in repair and "config.json" in repair
    for value in ("private-first", "private-second", "true", "false", "True", "False"):
        assert value not in repair
        assert value not in str(exc.value)


def test_invalid_json_is_still_reported_by_validation(tmp_path):
    """The repeated-key check reads JSON first; malformed JSON keeps its old error."""
    (tmp_path / "config.json").write_text('{"client_id": "x",', encoding="utf-8")

    with pytest.raises(ValidationError) as exc:
        load_config(config_dir=str(tmp_path))

    assert exc.value.errors()[0]["type"] == "json_invalid"


def test_a_legacy_code_page_config_with_a_repeated_key_is_not_rescued(tmp_path, monkeypatch):
    _pretend_the_locale_is(monkeypatch, "cp1252")
    config_dir = tmp_path / ".outlook-mcp"
    raw = '{"timezone": "Zoë", "read_only": true, "read_only": false}'.encode("cp1252")
    _write_config_bytes(config_dir, raw)

    with pytest.raises(UnicodeDecodeError):
        load_config(config_dir=str(config_dir))
