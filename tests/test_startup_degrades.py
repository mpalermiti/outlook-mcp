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
