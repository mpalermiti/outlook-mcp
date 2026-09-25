"""A host that cannot store a token safely must still get a working server.

Refusing to write a plaintext token cache is right. Killing the MCP process on
the way up is not: the client shows a dead server, the explanation goes to
stderr where no agent reads it, and the one line the operator needs — set
``allow_unencrypted_token_cache``, or install libsecret — never reaches them.

``lifespan`` already promised this in a comment ("if this fails, tools will
return an error telling the user to run `outlook-mcp auth`"). These tests hold
it to that promise, and to telling the truth about *which* remedy applies.
"""

import json
import os
import subprocess
import sys
from unittest.mock import MagicMock, patch

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config
from outlook_mcp.errors import (
    AuthRequiredError,
    ConfigLoadError,
    UnencryptedTokenCacheError,
)
from outlook_mcp.server import lifespan, outlook_auth_status


def _ctx(auth):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {"config": auth.config, "auth": auth}
    return ctx


@pytest.mark.asyncio
async def test_server_still_boots_when_the_token_cache_is_unwritable():
    with (
        patch("outlook_mcp.server.load_config", return_value=Config(client_id="x")),
        patch.object(
            AuthManager, "try_cached_token", side_effect=UnencryptedTokenCacheError()
        ),
    ):
        async with lifespan(MagicMock()) as state:
            assert state["auth"] is not None
            assert state["auth"].is_authenticated() is False


@pytest.mark.asyncio
async def test_every_tool_call_reports_the_real_remedy_not_re_run_auth():
    """`outlook-mcp auth` is the wrong advice here — it fails the same way."""
    auth = AuthManager(Config(client_id="x"))
    auth.startup_error = UnencryptedTokenCacheError()

    with pytest.raises(UnencryptedTokenCacheError) as exc:
        auth.get_credential()
    assert "allow_unencrypted_token_cache" in str(exc.value)


@pytest.mark.asyncio
async def test_auth_status_explains_why_rather_than_just_saying_no():
    auth = AuthManager(Config(client_id="x"))
    auth.startup_error = UnencryptedTokenCacheError()

    result = await outlook_auth_status(_ctx(auth))

    assert result["authenticated"] is False
    assert "allow_unencrypted_token_cache" in result["action_required"]


@pytest.mark.asyncio
async def test_an_ordinary_unauthenticated_host_is_unchanged():
    """The common case — no token yet — must still say "run outlook-mcp auth"."""
    auth = AuthManager(Config(client_id="x"))

    with pytest.raises(AuthRequiredError):
        auth.get_credential()

    result = await outlook_auth_status(_ctx(auth))
    assert result["action_required"] == (
        "Run `outlook-mcp auth` on the host to authenticate."
    )


# ── A config the server cannot load fails with the fix, never a traceback ──
#
# Two layers, two jobs. ``main`` validates the config before the transport
# starts and exits cleanly — those tests run the real entrypoint in a
# subprocess, because the property being held ("no exception group, no
# traceback, stdout stays the protocol channel") only exists on the path a
# real client takes. The lifespan keeps a backstop for reaching the server
# without going through ``main``: raising from inside the async task group
# is what produces the unreadable exception group, so the backstop degrades
# to a read-only boot that carries the repair on every tool call.


def _validation_error() -> Exception:
    """A real pydantic ValidationError, raised the way load_config raises it."""
    from pydantic import ValidationError

    try:
        Config.model_validate({"allow_categories": ["not-a-category"]})
    except ValidationError as exc:
        return exc
    raise AssertionError("expected a ValidationError")


@pytest.mark.asyncio
async def test_invalid_config_boots_read_only_with_the_fix(caplog):
    import logging

    with (
        patch("outlook_mcp.server.load_config", side_effect=_validation_error()),
        caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
    ):
        async with lifespan(MagicMock()) as state:
            assert state["config"].read_only is True  # fail-safe substitute
            with pytest.raises(ConfigLoadError) as exc:
                state["auth"].get_credential()
            assert "restart the server" in str(exc.value)

    messages = [r.getMessage() for r in caplog.records]
    assert any("allow_categories" in m for m in messages)  # names the field
    assert any("restart the server" in m for m in messages)  # names the fix


@pytest.mark.asyncio
async def test_symlinked_config_boots_the_same_way(caplog):
    import logging

    refusal = PermissionError("Refusing to load symlinked config: /x/config.json")
    with (
        patch("outlook_mcp.server.load_config", side_effect=refusal),
        caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
    ):
        async with lifespan(MagicMock()) as state:
            assert state["config"].read_only is True
            with pytest.raises(ConfigLoadError):
                state["auth"].get_credential()

    messages = [r.getMessage() for r in caplog.records]
    assert any("symlink" in m for m in messages)
    assert any("restart" in m for m in messages)


@pytest.mark.asyncio
async def test_unreadable_or_non_utf8_config_boots_the_same_way(caplog):
    """chmod/read failures are OSErrors and non-UTF-8 bytes are ValueErrors —
    both arms have to land in the same degraded boot, not escape the lifespan."""
    import logging

    failures = (
        OSError(13, "Permission denied"),
        UnicodeDecodeError("utf-8", b"\xff", 0, 1, "bad byte"),
    )
    for failure in failures:
        with (
            patch("outlook_mcp.server.load_config", side_effect=failure),
            caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
        ):
            async with lifespan(MagicMock()) as state:
                assert state["config"].read_only is True
                with pytest.raises(ConfigLoadError):
                    state["auth"].get_credential()

    assert any("cannot start" in r.getMessage() for r in caplog.records)


@pytest.mark.asyncio
async def test_auth_status_reports_the_config_failure_as_the_action():
    """The degraded boot has to say what to fix, not "run outlook-mcp auth"."""
    auth = AuthManager(Config(read_only=True))
    auth.startup_error = ConfigLoadError(
        OSError(13, "Permission denied"), "/settings"
    )

    result = await outlook_auth_status(_ctx(auth))

    assert result["authenticated"] is False
    assert "Fix /settings/config.json and restart the server." in result["action_required"]


# ── The real entrypoint: exit code 1, repair on stderr, stdout untouched ──

_ENTRY = "from outlook_mcp.server import main; main()"


def _run_server_entry(config_dir) -> subprocess.CompletedProcess[str]:
    env = dict(os.environ)
    env["OUTLOOK_MCP_CONFIG_DIR"] = str(config_dir)
    return subprocess.run(
        [sys.executable, "-c", _ENTRY],
        env=env,
        capture_output=True,
        text=True,
        timeout=60,
    )


def _assert_clean_exit(proc: subprocess.CompletedProcess[str], *expected: str):
    assert proc.returncode == 1
    assert proc.stdout == ""  # stdout is the protocol channel, not the log
    assert "Traceback" not in proc.stderr
    assert "Exception Group" not in proc.stderr
    for fragment in expected:
        assert fragment in proc.stderr


def test_main_exits_cleanly_on_an_invalid_config(tmp_path):
    config_dir = tmp_path / "broken"
    config_dir.mkdir()
    (config_dir / "config.json").write_text(
        json.dumps({"allow_categories": ["not-a-category"]})
    )

    _assert_clean_exit(
        _run_server_entry(config_dir),
        "The config file is invalid",
        "allow_categories",
        "restart the server",
    )


def test_main_exits_cleanly_on_a_non_utf8_config(tmp_path):
    config_dir = tmp_path / "binary"
    config_dir.mkdir()
    (config_dir / "config.json").write_bytes(b'\xff\xfe{"client_id": "x"}')

    _assert_clean_exit(
        _run_server_entry(config_dir),
        "Cannot load the config file",
        str(config_dir),  # the repair names the directory to check
    )


def test_main_exits_cleanly_on_a_symlinked_config(tmp_path):
    real = tmp_path / "real-config.json"
    real.write_text("{}")
    config_dir = tmp_path / "symlinked"
    config_dir.mkdir()
    link = config_dir / "config.json"
    try:
        link.symlink_to(real)
    except OSError:
        pytest.skip("this host cannot create symlinks")

    _assert_clean_exit(
        _run_server_entry(config_dir),
        "Cannot load the config file",
        "symlink",
    )


def test_main_exits_cleanly_when_the_settings_path_is_a_file(tmp_path):
    """An override pointing at an existing file is refused at load, with the
    repair — not later as a FileExistsError partway through a flow."""
    not_a_dir = tmp_path / "settings.txt"
    not_a_dir.write_text("occupied")

    _assert_clean_exit(
        _run_server_entry(not_a_dir),
        "Cannot load the config file",
        "not a directory",
        "OUTLOOK_MCP_CONFIG_DIR",
    )
