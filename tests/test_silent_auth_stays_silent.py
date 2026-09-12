"""The silent token path must never open an interactive flow.

``try_cached_token`` promises a token "without user interaction". It was calling
``get_token()`` on a ``DeviceCodeCredential`` built with azure-identity's
default ``disable_automatic_authentication=False``, so a cache miss did not
return False — it started a device-code flow and polled for a human until
``timeout`` (900s). Two consequences, both real:

- ``lifespan`` blocks for fifteen minutes on server startup instead of coming
  up unauthenticated and telling the agent to run ``outlook-mcp auth``.
- ``uv run pytest`` — documented as the offline unit suite — makes live calls
  to Microsoft on any machine that has ``~/.outlook-mcp/auth_record.json``,
  which is why this never showed up in CI.

The interactive flow belongs to ``login_interactive`` alone.
"""

from unittest.mock import MagicMock, patch

import pytest
from azure.identity import AuthenticationRequiredError

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config


def _kwargs_of(cred_cls):
    return cred_cls.call_args.kwargs


@pytest.fixture(autouse=True)
def _assume_an_encrypted_store():
    """Isolate the variable under test.

    These tests are about the interactive-flow flag, not about whether the host
    can encrypt. Left real, they fail on any Linux box without libsecret --
    including the GitHub runner -- for a reason that has nothing to do with what
    they assert.
    """
    with patch(
        "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False
    ):
        yield


def test_the_silent_path_disables_automatic_authentication():
    auth = AuthManager(Config(client_id="x"))
    with (
        patch("outlook_mcp.auth._load_auth_record", return_value=object()),
        patch("outlook_mcp.auth.DeviceCodeCredential") as cred_cls,
    ):
        auth.try_cached_token()
    assert _kwargs_of(cred_cls).get("disable_automatic_authentication") is True


def test_the_interactive_path_still_allows_the_prompt():
    """`outlook-mcp auth` exists to prompt; it must not inherit the muzzle."""
    auth = AuthManager(Config(client_id="x"))
    with (
        patch("outlook_mcp.auth.DeviceCodeCredential") as cred_cls,
        patch("outlook_mcp.auth._save_auth_record"),
    ):
        auth.login_interactive()
    assert _kwargs_of(cred_cls).get("disable_automatic_authentication") is not True


def test_a_cache_miss_reports_failure_instead_of_waiting_for_a_human():
    auth = AuthManager(Config(client_id="x"))
    cred = MagicMock()
    cred.get_token = MagicMock(
        side_effect=AuthenticationRequiredError(scopes=["https://graph.microsoft.com/.default"])
    )
    with (
        patch("outlook_mcp.auth._load_auth_record", return_value=object()),
        patch.object(AuthManager, "_make_credential", return_value=cred),
    ):
        assert auth.try_cached_token() is False
    assert auth.is_authenticated() is False


@pytest.mark.asyncio
async def test_server_startup_does_not_authenticate_interactively():
    """The regression that made the suite hang: lifespan must come up fast."""
    from outlook_mcp.server import lifespan

    auth_records = []

    def _spy(self, prompt_callback=None, auth_record=None, **kw):
        auth_records.append(kw.get("silent"))
        raise AuthenticationRequiredError(scopes=["s"])

    with (
        patch("outlook_mcp.server.load_config", return_value=Config(client_id="x")),
        patch("outlook_mcp.auth._load_auth_record", return_value=object()),
        patch.object(AuthManager, "_make_credential", _spy),
    ):
        async with lifespan(MagicMock()) as state:
            assert state["auth"].is_authenticated() is False
