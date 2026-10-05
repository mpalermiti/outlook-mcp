"""Tests for auth module."""

import logging
from unittest.mock import patch

import pytest
from azure.core.credentials import AccessToken
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRecord, DeviceCodeCredential

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager, _unencrypted_fallback_will_be_used
from outlook_mcp.config import Config
from outlook_mcp.errors import (
    AuthRequiredError,
    StaleConsentError,
    UnencryptedTokenCacheError,
)

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


def test_read_only_config_still_gets_the_read_write_scopes():
    """read_only gates the tools, not the token: the config flag must not
    narrow what the first consent asks for — the scopes it omits could never
    be granted later without logging in again."""
    config = Config(client_id="test-id", read_only=True)
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert scopes == auth_module.SCOPES_READWRITE
    # offline_access is reserved — MSAL adds it automatically
    assert "offline_access" not in scopes


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
    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: tmp_path / "auth_record.json")

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
        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False):
            assert self._options_used(auth).allow_unencrypted_storage is False

    def test_refuses_to_build_a_credential_that_would_write_plaintext(self):
        auth = AuthManager(Config(client_id="test-id"))
        with (
            patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True),
            pytest.raises(UnencryptedTokenCacheError) as exc,
        ):
            auth._make_credential()
        # Must name both ways out, or the operator is just stuck.
        assert "libsecret" in str(exc.value)
        assert "allow_unencrypted_token_cache" in str(exc.value)

    def test_opting_in_permits_the_plaintext_fallback(self):
        auth = AuthManager(Config(client_id="test-id", allow_unencrypted_token_cache=True))
        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True):
            assert self._options_used(auth).allow_unencrypted_storage is True

    def test_opting_in_still_warns(self, caplog):
        auth = AuthManager(Config(client_id="test-id", allow_unencrypted_token_cache=True))
        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True),
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
            patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True),
            pytest.raises(UnencryptedTokenCacheError),
        ):
            auth.try_cached_token()

    # azure-identity's own text, from _persistent_cache.py — raised lazily at
    # first token use, not at credential construction, where the persistent
    # cache is built.
    AZURE_REFUSAL = ValueError(
        "Cache encryption is impossible because libsecret dependencies are not "
        "installed or are unusable, for example because no display is available "
        "(as in an SSH session). The chained exception has more information. "
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
            patch.object(DeviceCodeCredential, "_get_app", side_effect=self.AZURE_REFUSAL),
            pytest.raises(UnencryptedTokenCacheError) as exc,
        ):
            auth.try_cached_token()
        assert "allow_unencrypted_token_cache" in str(exc.value)

    def test_the_cli_auth_path_translates_it_too(self):
        auth = AuthManager(Config(client_id="test-id"))

        with (
            # Pin the environment probe: on a Linux runner without PyGObject
            # it returns True on its own, and the environment refusal would
            # satisfy this test before the injected refusal is ever reached.
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(DeviceCodeCredential, "_get_app", side_effect=self.AZURE_REFUSAL),
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
        # The fallback patch matters here too: a host with no encrypted store
        # (the Linux CI runners) would refuse at credential construction,
        # before the injected transient ever runs.
        with (
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(DeviceCodeCredential, "_get_app", side_effect=transient),
            pytest.raises(ClientAuthenticationError) as exc,
        ):
            auth.login_interactive()
        assert auth_module._is_azure_unencrypted_refusal(exc.value) is False

    def test_the_wrapped_refusal_matches_on_the_message_alone(self):
        """azure's wrapper embeds the original's whole text, cause included.

        The refusal arrives as ``ClientAuthenticationError("Authentication
        failed: <original>")`` with the original on ``__cause__`` — so the
        message arm is the complete matcher. A marker living only on the
        cause with a silent message does not occur on any real path and must
        not match.
        """
        original = self.AZURE_REFUSAL
        wrapped = ClientAuthenticationError(f"Authentication failed: {original}")
        wrapped.__cause__ = original

        assert auth_module._is_azure_unencrypted_refusal(wrapped) is True

        quiet = ClientAuthenticationError("Authentication failed: something else")
        quiet.__cause__ = original
        assert auth_module._is_azure_unencrypted_refusal(quiet) is False


# ── Which scopes each path asks for, and why they differ ──────────────────
#
# A first consent that asks only for .default can land a session with no
# delegated permissions on a personal account, and no scope can be redeemed
# from that session afterwards (AADSTS70000) — only logging in again. So the
# interactive flow consents the concrete read-write scopes — in read-only
# mode too, because read_only gates the tools, not the token, and a
# read-only consent could never be widened once the config flips — while
# silent refresh, serving from an already-issued session, redeems .default:
# the one scope every saved record can still serve.


class TestFirstConsentScopes:
    """The interactive flow consents the concrete read-write scopes, never
    .default — whatever read_only says."""

    def _consented_scopes(self, config: Config) -> tuple[str, ...]:
        """Run login_interactive against a mocked credential; return its scopes."""
        auth = AuthManager(config)
        with (
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch("outlook_mcp.auth.DeviceCodeCredential") as cred_cls,
            patch("outlook_mcp.auth._save_auth_record"),
        ):
            auth.login_interactive()
        return cred_cls.return_value.get_token.call_args.args

    def test_read_write_mode_consents_the_write_scopes(self):
        """The consent names Mail.ReadWrite & co. — and not .default, whose
        first-time consent is what strands an MSA session with nothing."""
        scopes = self._consented_scopes(Config(client_id="test-id"))
        assert scopes == tuple(auth_module.SCOPES_READWRITE)
        assert GRAPH_SCOPE not in scopes

    def test_read_only_mode_still_consents_the_write_scopes(self):
        """read_only gates the tools, not the token — the consent always
        names the full read-write set. A read-only first consent would strand
        writes forever: silent refresh redeems .default, which means exactly
        the consented set, so the write scopes could never be added once the
        config flips without logging in again."""
        scopes = self._consented_scopes(Config(client_id="test-id", read_only=True))
        assert scopes == tuple(auth_module.SCOPES_READWRITE)
        assert GRAPH_SCOPE not in scopes

    def test_read_only_consent_asks_for_the_read_scopes_only(self):
        """The explicit opt-in is the one thing that narrows the consent.

        `read_only` alone must not (the test above): flipping it later would
        strand every write. `read_only_consent` is a separate, deliberate
        statement — "this app must never hold write access" — and it is what
        makes a second, read-only app registration reachable at all.
        """
        scopes = self._consented_scopes(
            Config(client_id="test-id", read_only=True, read_only_consent=True)
        )
        assert scopes == tuple(auth_module.SCOPES_READONLY)
        assert GRAPH_SCOPE not in scopes

    def test_no_read_only_scope_can_write_or_send(self):
        for scope in auth_module.SCOPES_READONLY:
            assert "Write" not in scope, scope
            assert "Send" not in scope, scope

    def test_the_read_only_scopes_cover_every_read_the_write_set_covers(self):
        """Each read-write scope has its read counterpart, so no read tool loses its permission."""
        counterparts = [
            scope.replace(".ReadWrite", ".Read")
            for scope in auth_module.SCOPES_READWRITE
            if scope != "Mail.Send"
        ]
        assert auth_module.SCOPES_READONLY == counterparts
        assert auth_module.SCOPES_READONLY == [
            "Mail.Read",
            "Calendars.Read",
            "Contacts.Read",
            "Tasks.Read",
            "MailboxSettings.Read",
            "User.Read",
        ]


class _DefaultConsentCredential:
    """A record-pinned credential whose session consented `.default` alone.

    This is the stranded MSA session as an offline fixture: it serves
    `.default` — the one scope a `.default`-only consent can still redeem —
    and refuses every concrete scope exactly the way the live endpoint
    refuses it. Records saved before the consent fix have precisely these
    capabilities, which is why silent refresh must stay on `.default`.
    """

    def __init__(
        self,
        *,
        client_id: str,
        tenant_id: str,
        cache_persistence_options,
        timeout: int = 0,
        disable_automatic_authentication: bool = False,
        prompt_callback=None,
        authentication_record: AuthenticationRecord | None = None,
    ) -> None:
        self._record = authentication_record
        self.requested_scopes: list[tuple[str, ...]] = []

    def get_token(self, *scopes) -> AccessToken:
        self.requested_scopes.append(scopes)
        if scopes == (GRAPH_SCOPE,):
            return AccessToken("token-from-a-default-consent", 0)
        raise ClientAuthenticationError(
            "Authentication failed: AADSTS70000: The requested user must "
            "first sign-in and grant the client application access"
        )


class TestSilentRefreshScope:
    """Silent refresh redeems .default — the scope every record can serve."""

    def test_a_record_saved_under_a_default_consent_still_refreshes(self):
        """The regression pin for the consent change.

        A record whose session consented `.default` alone can redeem nothing
        but `.default`. Refreshing with the concrete consent list instead —
        the tempting "consistency" change — would brick every such record
        the moment its token expires, so the refresh is pinned here against
        a credential that refuses everything but `.default`, the way the
        live endpoint does.
        """
        auth = AuthManager(Config(client_id="test-id"))
        record = AuthenticationRecord(
            tenant_id="consumers",
            client_id="test-id",
            authority="https://login.microsoftonline.com/consumers",
            home_account_id="home-1",
            username="user@example.com",
        )
        with (
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch("outlook_mcp.auth._load_auth_record", return_value=record),
            patch("outlook_mcp.auth.DeviceCodeCredential", _DefaultConsentCredential),
        ):
            assert auth.try_cached_token() is True

        cred = auth.get_credential()
        assert isinstance(cred, _DefaultConsentCredential)
        assert cred.requested_scopes == [(GRAPH_SCOPE,)]


class TestStaleConsentRemedy:
    """An AADSTS70000 refresh failure must say "log in again", not "retry"."""

    RECORD = AuthenticationRecord(
        tenant_id="consumers",
        client_id="test-id",
        authority="https://login.microsoftonline.com/consumers",
        home_account_id="home-1",
        username="user@example.com",
    )

    # The refusal as it arrives from azure-identity: MSAL's text embedded in
    # the wrapped ClientAuthenticationError.
    AADSTS70000_REFUSAL = ClientAuthenticationError(
        "Authentication failed: AADSTS70000: The requested user must first "
        "sign-in and grant the client application access. The user must "
        "re-authenticate."
    )

    def test_the_dead_end_sets_the_log_in_again_error(self):
        """Named code, named exit: the error's own text suggests neither.

        The refusal is injected on a real credential's token path (the
        ``_get_app`` seam the unencrypted-refusal tests use) so it travels
        through the same wrapping a live failure would.
        """
        auth = AuthManager(Config(client_id="test-id"))
        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=self.RECORD),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(DeviceCodeCredential, "_get_app", side_effect=self.AADSTS70000_REFUSAL),
        ):
            assert auth.try_cached_token() is False

        assert auth.is_authenticated() is False
        assert isinstance(auth.startup_error, StaleConsentError)
        remedy = str(auth.startup_error)
        assert "AADSTS70000" in remedy
        assert "Log in again" in remedy
        # AADSTS70000 is Entra's generic invalid-grant — a revoked refresh
        # token reports it too — so the remedy hedges instead of asserting
        # the skipped-consent story as the only possibility.
        assert "most often" in remedy
        assert "revoked" in remedy

    def test_an_ordinary_stale_token_keeps_the_ordinary_remedy(self):
        """An expired token is refreshable — it must not claim the dead end.

        Two failure classes, two remedies: only the AADSTS70000 refusal gets
        the log-in-again error; anything else stays a plain stale token that
        `outlook-mcp auth` fixes.
        """
        auth = AuthManager(Config(client_id="test-id"))
        expired = ClientAuthenticationError(
            "Authentication failed: AADSTS7000215: Invalid grant. The token is expired or revoked."
        )
        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=self.RECORD),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch.object(DeviceCodeCredential, "_get_app", side_effect=expired),
        ):
            assert auth.try_cached_token() is False

        assert auth.startup_error is None  # AuthRequiredError on use, not this
        with pytest.raises(AuthRequiredError):
            auth.get_credential()


class TestSavedSignInBelongsToTheConfiguredApp:
    """A record saved for one app registration must not serve another.

    azure-identity uses the *record's* client id and ignores the one passed in,
    so after `client_id` is changed in config.json the old session kept being
    used, silently — and `outlook-mcp status` printed the new id next to
    "authenticated". Pointing the config at a read-only app without signing in
    again left the write-capable session in place.
    """

    @staticmethod
    def _record(client_id: str) -> AuthenticationRecord:
        return AuthenticationRecord(
            tenant_id="consumers",
            client_id=client_id,
            authority="https://login.microsoftonline.com/consumers",
            home_account_id="home-1",
            username="user@example.com",
        )

    def test_a_record_from_another_app_is_not_used(self):
        from outlook_mcp.errors import ClientIdMismatchError

        auth = AuthManager(Config(client_id="11111111-aaaa-4bbb-8ccc-000000000001"))
        record = self._record("99999999-aaaa-4bbb-8ccc-000000000009")
        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=record),
            patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False),
            patch("outlook_mcp.auth.DeviceCodeCredential") as cred_cls,
        ):
            assert auth.try_cached_token() is False

        cred_cls.assert_not_called()
        assert auth.is_authenticated() is False
        assert isinstance(auth.startup_error, ClientIdMismatchError)
        remedy = str(auth.startup_error)
        assert "outlook-mcp auth" in remedy
        assert "client_id" in remedy

    def test_the_same_app_written_in_another_case_is_the_same_app(self):
        auth = AuthManager(Config(client_id="ABCDEF00-AAAA-4BBB-8CCC-000000000001"))
        record = self._record("abcdef00-aaaa-4bbb-8ccc-000000000001")
        with (
            patch("outlook_mcp.auth._load_auth_record", return_value=record),
            patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False),
            patch("outlook_mcp.auth.DeviceCodeCredential", _DefaultConsentCredential),
        ):
            assert auth.try_cached_token() is True

        assert auth.startup_error is None
