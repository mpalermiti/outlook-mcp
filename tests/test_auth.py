"""Tests for auth module."""

import logging
from unittest.mock import patch

import pytest
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRecord, DeviceCodeCredential

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager, _unencrypted_fallback_will_be_used
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError, UnencryptedTokenCacheError

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


@pytest.fixture(autouse=True)
def _reset_unencrypted_warning_latch():
    """Reset the once-per-process warning latch between tests."""
    auth_module._warned_unencrypted_fallback = False
    yield
    auth_module._warned_unencrypted_fallback = False


def test_auth_manager_init():
    """AuthManager initializes with config."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.config is config
    assert auth.credential is None


def test_auth_scopes_default():
    """Default scopes include read-write."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.ReadWrite" in scopes
    assert "Mail.Send" in scopes
    assert "Calendars.ReadWrite" in scopes
    # offline_access is reserved — MSAL adds it automatically
    assert "offline_access" not in scopes


def test_auth_scopes_read_only():
    """Read-only mode uses read scopes."""
    config = Config(client_id="test-id", read_only=True)
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.Read" in scopes
    assert "Mail.ReadWrite" not in scopes
    assert "Mail.Send" not in scopes
    assert "Calendars.Read" in scopes
    assert "Calendars.ReadWrite" not in scopes


def test_auth_not_authenticated():
    """is_authenticated returns False before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False


def test_auth_get_credential_raises_when_not_authenticated():
    """get_credential raises AuthRequiredError before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    with pytest.raises(AuthRequiredError):
        auth.get_credential()


def test_login_interactive_requires_client_id():
    """login_interactive raises if client_id is not configured."""
    config = Config()  # No client_id
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="client_id"):
        auth.login_interactive()


def test_auth_record_round_trips_atomically(tmp_path, monkeypatch):
    """The record lands whole: same write pattern as config.json (temp file,
    fsync, rename), so a reader never sees a half-written record and no
    temp file outlives the save."""
    record = AuthenticationRecord(
        tenant_id="consumers",
        client_id="test-id",
        authority="https://login.microsoftonline.com/consumers",
        home_account_id="home-1",
        username="user@example.com",
    )
    monkeypatch.setattr(
        auth_module, "_auth_record_path", lambda: tmp_path / "auth_record.json"
    )

    auth_module._save_auth_record(record)

    loaded = auth_module._load_auth_record()
    assert loaded is not None
    assert loaded.home_account_id == record.home_account_id
    assert loaded.username == record.username
    assert list(tmp_path.glob("*.tmp")) == []  # the temp file became the record


def test_try_cached_token_returns_false_without_client_id():
    """try_cached_token returns False if client_id is not set."""
    config = Config()
    auth = AuthManager(config)
    assert auth.try_cached_token() is False


class TestUnencryptedFallbackDetection:
    """_unencrypted_fallback_will_be_used mirrors msal_extensions' check."""

    def test_macos_is_never_fallback(self):
        """macOS uses Keychain — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "darwin"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_windows_is_never_fallback(self):
        """Windows uses DPAPI — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "win32"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_with_gi_available(self):
        """Linux with PyGObject importable uses libsecret — no fallback."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=object()),
        ):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_without_gi_uses_fallback(self):
        """Linux without PyGObject (issue #7) triggers the fallback path."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=None),
        ):
            assert _unencrypted_fallback_will_be_used() is True


class TestUnencryptedFallbackWarning:
    """_make_credential emits a warning at most once when fallback is in use."""

    def test_warning_fires_once_when_fallback_active(self, caplog):
        """A single warning is logged on the first credential build."""
        config = Config(client_id="test-id", allow_unencrypted_token_cache=True)
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=True,
            ),
        ):
            auth._make_credential()
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert len(fallback_warnings) == 1
        assert fallback_warnings[0].levelno == logging.WARNING

    def test_no_warning_when_fallback_inactive(self, caplog):
        """No fallback warning is logged when encrypted storage is available."""
        config = Config(client_id="test-id")
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
        ):
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert fallback_warnings == []


class TestUnencryptedCacheIsOptIn:
    """Plaintext token caching must be a choice, never a silent fallback.

    ``allow_unencrypted_storage=True`` was unconditional, so on Linux without
    libsecret a reusable Microsoft Graph refresh token landed on disk in
    cleartext with nothing but a log line to say so — while SECURITY.md told
    readers tokens were "never in plain files". Either the storage is encrypted
    or the operator said in writing that it need not be.
    """

    def _options_used(self, auth):
        """Build a credential and return the TokenCachePersistenceOptions."""
        with patch("outlook_mcp.auth.DeviceCodeCredential") as cred_cls:
            auth._make_credential()
        return cred_cls.call_args.kwargs["cache_persistence_options"]

    def test_encrypted_storage_is_the_default(self):
        auth = AuthManager(Config(client_id="test-id"))
        with patch(
            "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False
        ):
            assert self._options_used(auth).allow_unencrypted_storage is False

    def test_refuses_to_build_a_credential_that_would_write_plaintext(self):
        auth = AuthManager(Config(client_id="test-id"))
        with (
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True
            ),
            pytest.raises(UnencryptedTokenCacheError) as exc,
        ):
            auth._make_credential()
        # Must name both ways out, or the operator is just stuck.
        assert "libsecret" in str(exc.value)
        assert "allow_unencrypted_token_cache" in str(exc.value)

    def test_opting_in_permits_the_plaintext_fallback(self):
        auth = AuthManager(
            Config(client_id="test-id", allow_unencrypted_token_cache=True)
        )
        with patch(
            "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True
        ):
            assert self._options_used(auth).allow_unencrypted_storage is True

    def test_opting_in_still_warns(self, caplog):
        auth = AuthManager(
            Config(client_id="test-id", allow_unencrypted_token_cache=True)
        )
        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True
            ),
        ):
            auth._make_credential()
        assert [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]

    def test_no_refusal_on_a_platform_that_encrypts(self):
        """macOS/Windows must be unaffected — this is a Linux-only failure mode."""
        auth = AuthManager(Config(client_id="test-id"))
        with patch.object(auth_module.sys, "platform", "darwin"):
            auth._make_credential()  # must not raise

    def test_cached_token_path_surfaces_the_config_error(self):
        """A misconfiguration must not read as an expired token.

        ``try_cached_token`` swallows refresh failures and returns False, which
        is right for a stale token and wrong for "this box cannot store one
        safely" — that loops the operator through `outlook-mcp auth` with no
        idea why.
        """
        auth = AuthManager(Config(client_id="test-id"))
        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=object()),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True
            ),
            pytest.raises(UnencryptedTokenCacheError),
        ):
            auth.try_cached_token()

    # azure-identity's own text, from _persistent_cache.py — raised lazily at
    # first token use, not at credential construction, where the persistent
    # cache is built.
    AZURE_REFUSAL = ValueError(
        "Cache encryption is impossible because libsecret dependencies are not "
        "installed or are unusable, for example because no display is available "
        '(as in an SSH session). The chained exception has more information. '
        'Specify "allow_unencrypted_storage=True" to store the cache unencrypted '
        "instead of raising this exception."
    )

    # A well-formed record, so the silent path gets past construction.
    RECORD = AuthenticationRecord(
        tenant_id="consumers",
        client_id="test-id",
        authority="https://login.microsoftonline.com/consumers",
        home_account_id="home-1",
        username="user@example.com",
    )

    def test_libsecret_installed_but_unusable_is_the_same_condition(self):
        """gi importable + no Secret Service: our eager check cannot see this.

        A display-less SSH session or a container hits azure's lazy refusal at
        ``get_token`` — and every token-acquisition method azure-identity
        exposes is wrapped by ``@wrap_exceptions``, which re-types the refusal
        into a ``ClientAuthenticationError`` before we see it. These tests
        drive the refusal through a REAL credential so the wrapping is on the
        path (a MagicMock short-circuits the decorator and proves nothing).
        Left untranslated it surfaces as a generic failure naming azure's
        kwarg, not the config key the operator actually sets.
        """
        auth = AuthManager(Config(client_id="test-id"))

        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=self.RECORD),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(
                DeviceCodeCredential, "_get_app", side_effect=self.AZURE_REFUSAL
            ),
            pytest.raises(UnencryptedTokenCacheError) as exc,
        ):
            auth.try_cached_token()
        assert "allow_unencrypted_token_cache" in str(exc.value)

    def test_the_cli_auth_path_translates_it_too(self):
        auth = AuthManager(Config(client_id="test-id"))

        with (
            patch.object(
                DeviceCodeCredential, "_get_app", side_effect=self.AZURE_REFUSAL
            ),
            pytest.raises(UnencryptedTokenCacheError),
        ):
            auth.login_interactive()

    def test_a_transient_auth_failure_does_not_flip_anything(self):
        """A wrapped non-refusal failure reads as "re-auth", never as the
        config error — the two remedies are different."""
        auth = AuthManager(Config(client_id="test-id"))
        transient = RuntimeError("connection reset while polling the device flow")

        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=self.RECORD),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(DeviceCodeCredential, "_get_app", side_effect=transient),
        ):
            assert auth.try_cached_token() is False  # not UnencryptedTokenCacheError

        # On the interactive path the same failure re-raises as the wrapped
        # ClientAuthenticationError azure-identity guarantees its callers.
        with (
            patch.object(DeviceCodeCredential, "_get_app", side_effect=transient),
            pytest.raises(ClientAuthenticationError) as exc,
        ):
            auth.login_interactive()
        assert auth_module._is_azure_unencrypted_refusal(exc.value) is False
