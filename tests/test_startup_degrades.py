"""A host that cannot store a token safely must still get a working server.

Refusing to write a plaintext token cache is right. Killing the MCP process on
the way up is not: the client shows a dead server, the explanation goes to
stderr where no agent reads it, and the one line the operator needs — set
``allow_unencrypted_token_cache``, or install libsecret — never reaches them.

``lifespan`` already promised this in a comment ("if this fails, tools will
return an error telling the user to run `outlook-mcp auth`"). These tests hold
it to that promise, and to telling the truth about *which* remedy applies.
"""

from unittest.mock import MagicMock, patch

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError, UnencryptedTokenCacheError
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


# ── A config the server cannot load exits with the fix, not a traceback ──


def _validation_error() -> Exception:
    """A real pydantic ValidationError, raised the way load_config raises it."""
    from pydantic import ValidationError

    try:
        Config.model_validate({"allow_categories": ["not-a-category"]})
    except ValidationError as exc:
        return exc
    raise AssertionError("expected a ValidationError")


@pytest.mark.asyncio
async def test_invalid_config_exits_naming_the_field_and_the_fix(caplog):
    import logging

    with (
        patch("outlook_mcp.server.load_config", side_effect=_validation_error()),
        caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
        pytest.raises(SystemExit) as exc,
    ):
        async with lifespan(MagicMock()):
            pass

    assert exc.value.code == 1
    messages = [r.getMessage() for r in caplog.records]
    assert any("allow_categories" in m for m in messages)  # names the field
    assert any("restart the server" in m for m in messages)  # names the fix


@pytest.mark.asyncio
async def test_symlinked_config_exits_with_the_reason_not_a_traceback(caplog):
    import logging

    refusal = PermissionError("Refusing to load symlinked config: /x/config.json")
    with (
        patch("outlook_mcp.server.load_config", side_effect=refusal),
        caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
        pytest.raises(SystemExit) as exc,
    ):
        async with lifespan(MagicMock()):
            pass

    assert exc.value.code == 1
    messages = [r.getMessage() for r in caplog.records]
    assert any("symlink" in m for m in messages)
    assert any("restart" in m for m in messages)


@pytest.mark.asyncio
async def test_unreadable_config_exits_the_same_way(caplog):
    """chmod and read failures are OSErrors too — same guidance, no crash."""
    import logging

    with (
        patch(
            "outlook_mcp.server.load_config",
            side_effect=OSError(13, "Permission denied"),
        ),
        caplog.at_level(logging.ERROR, logger="outlook_mcp.server"),
        pytest.raises(SystemExit) as exc,
    ):
        async with lifespan(MagicMock()):
            pass

    assert exc.value.code == 1
    assert any("cannot start" in r.getMessage() for r in caplog.records)
