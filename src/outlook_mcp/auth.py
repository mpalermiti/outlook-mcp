"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import importlib.util
import logging
import sys
from pathlib import Path

from azure.core.exceptions import AzureError, ClientAuthenticationError
from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)

from outlook_mcp.config import DEFAULT_CONFIG_DIR, ROUTING_CAPABILITIES, Config
from outlook_mcp.errors import (
    AuthRequiredError,
    OutlookMCPError,
    UnencryptedTokenCacheError,
)

logger = logging.getLogger(__name__)

# Process-local latch so the unencrypted-fallback warning fires at most
# once per run — _make_credential is called from both login_interactive
# and try_cached_token, often multiple times during startup.
_warned_unencrypted_fallback = False

# Display-only; token acquisition uses .default (GRAPH_DEFAULT_SCOPE below).
SCOPES_READWRITE = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
]

SCOPES_READONLY = [
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "Tasks.Read",
    "User.Read",
]

# ONE cache name for every account. It is tempting to give each account its
# own (outlook-mcp-<name>) — and on Windows (a file per name) and Linux
# (libsecret keyed by name) that works — but on macOS the Keychain item is
# always (Microsoft.Developer.IdentityService, MSALCache): the cache *name*
# only picks the signal file. Per-account names there mean `auth net` then
# `auth neko` leaves neko's cache empty, and its first save replaces the
# shared Keychain item with neko alone — net's next silent refresh fails
# (review of #61). MSAL caches are multi-account by design, and the
# per-account AuthenticationRecord below is what pins which identity a
# credential serves.
CACHE_NAME = "outlook-mcp"
AUTH_RECORD_FILE = "auth_record.json"


def _unencrypted_fallback_will_be_used() -> bool:
    """Return True if msal_extensions will fall back to plaintext caching.

    Mirrors msal_extensions' libsecret-availability check: macOS uses
    Keychain and Windows uses DPAPI, both always encrypted, so only
    Linux is at risk — and only when PyGObject/libsecret isn't
    importable in the current Python environment (the failure mode
    reported in #7 for `uv tool install`).

    This is only half the condition. libsecret can be importable and still
    unusable — no running Secret Service, as in a display-less SSH session or
    a container — which azure-identity discovers lazily at first token use.
    ``_is_azure_unencrypted_refusal`` below catches that half; a False here
    does not mean an encrypted cache is guaranteed.
    """
    if sys.platform != "linux":
        return False
    return importlib.util.find_spec("gi") is None


# azure-identity refuses to build a plaintext cache *lazily* — at first token
# use, not at credential construction — and only when libsecret is importable
# but unusable (a display-less SSH session, a container). The eager
# find_spec("gi") check above cannot see that case, so this is the second half
# of the same condition. Matched on azure's own wording from
# azure/identity/_persistent_cache.py.
_AZURE_UNENCRYPTED_MARKER = "allow_unencrypted_storage"


def _is_azure_unencrypted_refusal(exc: BaseException) -> bool:
    """True for azure-identity's "cache encryption is impossible" ValueError."""
    return isinstance(exc, ValueError) and _AZURE_UNENCRYPTED_MARKER in str(exc)


# The Graph SDK always requests .default scope internally, so we must
# acquire and cache tokens with the same scope to avoid cache misses
# that trigger interactive auth in the background.
GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"


def _auth_record_path(account: str | None = None) -> Path:
    """Auth record path — one per account; the unnamed one is the legacy default.

    The record is what pins which identity in the shared MSAL cache a
    credential serves, so it has to live in a per-account file.
    """
    name = AUTH_RECORD_FILE if account is None else f"auth_record-{account}.json"
    return Path(DEFAULT_CONFIG_DIR) / name


def _save_auth_record(record: AuthenticationRecord, account: str | None = None) -> None:
    """Persist AuthenticationRecord to disk."""
    path = _auth_record_path(account)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(record.serialize())
    path.chmod(0o600)


def _load_auth_record(account: str | None = None) -> AuthenticationRecord | None:
    """Load AuthenticationRecord from disk, or None if not found."""
    path = _auth_record_path(account)
    if not path.exists():
        return None
    try:
        return AuthenticationRecord.deserialize(path.read_text())
    except Exception:
        logger.warning("Failed to load auth record from %s", path)
        return None


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph.

    One credential per configured account, all reading from the single shared
    token cache (see CACHE_NAME): the per-account AuthenticationRecord pins
    which identity each credential serves. ``_active_account`` is the
    configured default and is never silently re-pointed — when its token is
    missing at startup, data capabilities fail closed naming it
    (AuthRequiredError), and only identity tools may fall back to another
    authenticated account (see ``_identity_fallback_account``).
    """

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._credentials: dict[str, DeviceCodeCredential] = {}
        self._active_account: str | None = config.default_account
        # Identity-only fallback, set ONLY by try_cached_token when the
        # configured default has no valid token but another account does.
        # Identity tools (capability None) may answer through it; data
        # capabilities never do — see resolve_capability_account.
        self._identity_fallback_account: str | None = None
        # Session-level routing overrides (capability -> account), set only via
        # switch_account(capability=...) — which allow_cross_account gates.
        self._routing_overrides: dict[str, str] = {}
        # Set when startup authentication failed for a reason the operator has
        # to fix in config rather than by running `outlook-mcp auth` — that
        # advice would just fail the same way. Surfaced by get_credential() and
        # get_account_credential() so the remedy reaches the agent on every
        # tool call, not only stderr.
        self.startup_error: OutlookMCPError | None = None

    @property
    def authenticated_accounts(self) -> list[str]:
        """Names of accounts holding a valid credential this run."""
        return list(self._credentials)

    @property
    def active_account(self) -> str | None:
        """The configured active account (None on single-account installs).

        Never silently re-pointed — it moves only via config or an explicit,
        gated switch_account. Auth surfaces report it so their remedies name
        the account that actually needs authenticating.
        """
        return self._active_account

    def get_scopes(self) -> list[str]:
        """Return individual scopes for display/consent purposes."""
        return SCOPES_READONLY if self.config.read_only else SCOPES_READWRITE

    def get_token_scopes(self) -> list[str]:
        """Return scopes for token acquisition — must match what the SDK requests."""
        return [GRAPH_DEFAULT_SCOPE]

    def is_authenticated(self) -> bool:
        """Check if the active account has a credential."""
        return self.credential is not None

    def _account_config(self, account: str | None) -> tuple[str | None, str]:
        """(client_id, tenant_id) for an account: its own, or the top-level values.

        ``account=None`` is the legacy single-account setup. Unknown names
        raise here — at login/routing time, not deep inside the SDK.
        """
        if account is None:
            return self.config.client_id, self.config.tenant_id
        for acc in self.config.accounts:
            if acc.name == account:
                return acc.client_id, acc.tenant_id
        raise ValueError(
            f"Account '{account}' is not configured. Accounts: "
            f"{[a.name for a in self.config.accounts]}"
        )

    def _make_credential(
        self,
        account: str | None = None,
        prompt_callback=None,
        auth_record: AuthenticationRecord | None = None,
        *,
        silent: bool = False,
    ) -> DeviceCodeCredential:
        """Create a DeviceCodeCredential for one account.

        Every account shares the one token cache (CACHE_NAME); the account's
        AuthenticationRecord — loaded by callers and passed as
        ``authentication_record`` — is what pins the identity. Without a
        record (interactive login) azure-identity asks the browser, and the
        caller saves the returned record for next time.

        ``silent=True`` forbids the interactive device-code flow. azure-identity
        defaults to allowing it, so a *cache miss* on what is supposed to be a
        silent refresh does not fail — it prints a code and polls for a human
        until ``timeout`` (900s). On the startup path that is a fifteen-minute
        hang where the honest answer is "not authenticated"; `get_token` raises
        ``AuthenticationRequiredError`` instead when this is set.
        """
        global _warned_unencrypted_fallback
        client_id, tenant_id = self._account_config(account)
        opted_in = self.config.allow_unencrypted_token_cache
        cache_options = TokenCachePersistenceOptions(
            name=CACHE_NAME,
            allow_unencrypted_storage=opted_in,
        )
        if _unencrypted_fallback_will_be_used() and not opted_in:
            # Stop here rather than hand msal_extensions a credential it can
            # only persist in cleartext. Silently doing it is what made
            # SECURITY.md's "never in plain files" untrue.
            raise UnencryptedTokenCacheError()
        if not _warned_unencrypted_fallback and _unencrypted_fallback_will_be_used():
            logger.warning(
                "Token cache will be stored unencrypted on disk: "
                "allow_unencrypted_token_cache is set and "
                "PyGObject/libsecret is not importable in this Python "
                "environment (common with `uv tool install` on Linux — "
                "the tool's isolated venv can't see system PyGObject). "
                "To get encrypted caching via libsecret/gnome-keyring, "
                "install the system packages "
                "(apt: `gnome-keyring libsecret-1-0 python3-gi`) and "
                "re-create the venv with `--system-site-packages`. See "
                "https://github.com/mpalermiti/outlook-mcp/issues/7."
            )
            _warned_unencrypted_fallback = True
        kwargs = {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "cache_persistence_options": cache_options,
            "timeout": 900,
        }
        if silent:
            kwargs["disable_automatic_authentication"] = True
        if prompt_callback:
            kwargs["prompt_callback"] = prompt_callback
        if auth_record:
            kwargs["authentication_record"] = auth_record
        return DeviceCodeCredential(**kwargs)

    def login_interactive(self, account: str | None = None) -> None:
        """Run the device code flow interactively in the terminal.

        Uses get_token() which respects the token cache — if a valid
        cached token exists, completes silently. Otherwise triggers the
        device code flow. Saves the AuthenticationRecord for silent
        token refresh by the MCP server.

        ``account`` picks which configured account to authenticate; with
        ``accounts`` configured and no argument it defaults to
        ``default_account``. Intended for CLI use
        (`outlook-mcp auth [account]`), not MCP tools.
        """
        if self.config.accounts:
            account = account or self.config.default_account
        client_id, _ = self._account_config(account)  # raises for unknown names
        if not client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json."
            )

        def _on_device_code(verification_uri: str, user_code: str, expires_on: object) -> None:
            print(f"Visit:  {verification_uri}")
            print(f"Code:   {user_code}")
            print()
            print("Waiting for you to complete sign-in in your browser...")

        cred = self._make_credential(account=account, prompt_callback=_on_device_code)
        # get_token() uses cache first, falls back to interactive.
        # Must use .default scope to match what the Graph SDK requests.
        try:
            cred.get_token(*self.get_token_scopes())
        except ValueError as exc:
            if _is_azure_unencrypted_refusal(exc):
                raise UnencryptedTokenCacheError() from exc
            raise

        # Save the auth record for silent refresh by the MCP server
        record = getattr(cred, "_auth_record", None)
        if record:
            _save_auth_record(record, account)

        if account is None:
            self.credential = cred
        else:
            self._credentials[account] = cred
            if account == self._active_account:
                self.credential = cred
                self._identity_fallback_account = None
        print(f"Authenticated successfully{' as ' + account if account else ''}.")

    def try_cached_token(self) -> bool:
        """Try to get tokens silently, for every configured account.

        Single-account installs (no ``accounts`` list) keep the legacy
        behavior. Returns True if any account obtained a valid token
        without user interaction. Used by the MCP server on startup and by
        `outlook-mcp status`.

        Fail-closed by design: when the active (default) account has no
        valid token, it stays the active account — data capabilities routed
        to it raise AuthRequiredError naming it, rather than silently
        running on whichever account happened to authenticate first. Only
        identity tools fall back, and only to another authenticated account,
        with a warning an operator can see.
        """
        if not self.config.accounts:
            return self._try_single_cached_token(None)

        assert self._active_account is not None  # set by config validation
        self._adopt_legacy_login(self._active_account)

        any_ok = False
        for acc in self.config.accounts:
            if self._try_single_cached_token(acc.name):
                any_ok = True

        if self._active_account not in self._credentials and self._credentials:
            fallback = next(iter(self._credentials))
            self._identity_fallback_account = fallback
            logger.warning(
                "Active account '%s' has no valid token. Data capabilities "
                "routed to it will fail closed until `outlook-mcp auth %s` "
                "is run; only identity tools fall back to '%s'.",
                self._active_account,
                self._active_account,
                fallback,
            )
        return any_ok

    def _adopt_legacy_login(self, account: str) -> None:
        """Carry a pre-multi-account login over to the default account.

        1.21 and earlier logged in once (auth_record.json) even when
        'accounts' was populated — the list did nothing, so that one login
        served everything. An upgrade that ignores it silently drops a
        working install's login: every account would start unauthenticated.
        Adopt the legacy record as the default account's (the identity that
        served everything yesterday), say so, and let the other accounts
        fail closed with their own `outlook-mcp auth <name>` remedy.
        """
        if _auth_record_path(account).exists() or not _auth_record_path(None).exists():
            return
        record = _load_auth_record(None)
        if record is None:
            return
        logger.warning(
            "Found a pre-multi-account login (auth_record.json) and no record "
            "for account '%s'. Adopting it as '%s'; if that is the wrong "
            "identity, run `outlook-mcp auth %s`.",
            account,
            account,
            account,
        )
        _save_auth_record(record, account)

    def _try_single_cached_token(self, account: str | None) -> bool:
        """The silent single-account path, parameterized by account."""
        client_id, _ = self._account_config(account)
        if not client_id:
            return False

        record = _load_auth_record(account)
        if record is None:
            return False

        try:
            cred = self._make_credential(account=account, auth_record=record, silent=True)
            cred.get_token(*self.get_token_scopes())
            if account is None:
                self.credential = cred
            else:
                self._credentials[account] = cred
                if account == self._active_account:
                    self.credential = cred
            return True
        except UnencryptedTokenCacheError:
            # Not a stale token — the environment cannot store one safely.
            # Swallowing it here sends the operator round the `outlook-mcp auth`
            # loop with no idea what to change.
            raise
        except ValueError as exc:
            if _is_azure_unencrypted_refusal(exc):
                raise UnencryptedTokenCacheError() from exc
            logger.warning(
                "Cached token refresh failed for account '%s' — re-run `outlook-mcp auth %s`.",
                account or "default",
                account or "",
            )
            return False
        except ClientAuthenticationError:
            # The token is genuinely unusable (expired beyond refresh, revoked,
            # or a cache miss in silent mode) — re-authenticating is the fix.
            logger.warning(
                "Cached token refresh failed for account '%s' — re-run `outlook-mcp auth %s`.",
                account or "default",
                account or "",
            )
            return False
        except AzureError as exc:
            # Transient: a 5xx or timeout from the token endpoint. Treat the
            # account as unauthenticated for this run — but say so honestly
            # rather than advising a re-auth that cannot help.
            logger.warning(
                "Transient error refreshing the token for account '%s' (%s); "
                "starting without it. Retry the request or restart the server "
                "once the network allows.",
                account or "default",
                type(exc).__name__,
            )
            return False

    def get_credential(self) -> DeviceCodeCredential:
        """Get the active account's credential, raising if not authenticated."""
        if self.credential is None:
            if self.startup_error is not None:
                raise self.startup_error
            if self.config.accounts and self._active_account:
                raise AuthRequiredError(self._active_account)
            raise AuthRequiredError()
        return self.credential

    def get_account_credential(self, account: str | None) -> DeviceCodeCredential:
        """Credential for an account by name; None means the legacy single account.

        Fail-closed: an account without a credential raises AuthRequiredError
        naming it — never a silent fallback to some other account's identity.
        ``startup_error`` is checked first (like get_credential) so a host that
        cannot store tokens safely reports the config fix, on every routed
        tool, instead of advising an auth command that would fail identically.
        """
        if account is None:
            return self.get_credential()
        if self.startup_error is not None:
            raise self.startup_error
        cred = self._credentials.get(account)
        if cred is not None:
            return cred
        names = [a.name for a in self.config.accounts]
        if account in names:
            raise AuthRequiredError(account)
        raise ValueError(f"Account '{account}' is not configured. Accounts: {names}")

    def resolve_capability_account(self, capability: str | None) -> str | None:
        """Which account serves a capability: override > config routing > active.

        Data capabilities always answer with *an account name* — if that
        account has no credential, get_account_credential raises
        AuthRequiredError naming it. They never substitute another
        authenticated account: that is how a config with `allow_cross_account
        = false` ended up writing mail as the wrong identity (review of #61).

        Identity (capability None) is the one exception: it may fall back to
        ``_identity_fallback_account`` — set only at startup, only when the
        configured default has no token — so whoami/auth_status still answer.
        """
        if capability is not None:
            override = self._routing_overrides.get(capability)
            if override is not None:
                return override
            configured = self.config.capability_accounts.get(capability)
            if configured is not None:
                return configured
            return self._active_account
        # Identity: the active account when it can answer. The one exception
        # is startup's identity-only fallback — active without a credential,
        # another account with one. After an explicit switch_account the
        # fallback is gone, so an unauthenticated active account raises
        # AuthRequiredError naming it rather than substituting again.
        if self._active_account is not None and (
            self._active_account in self._credentials or self._identity_fallback_account is None
        ):
            return self._active_account
        return self._identity_fallback_account

    def list_accounts(self) -> list[dict]:
        """Configured accounts with auth status.

        With allow_cross_account off, collapses to the active identity: the
        agent is meant to see one merged account, not a menu of accounts it
        cannot address.
        """
        if self.config.client_id and not self.config.accounts:
            return [
                {
                    "name": "default",
                    "client_id": self.config.client_id[:8] + "...",
                    "tenant_id": self.config.tenant_id,
                    "authenticated": self.credential is not None,
                    "active": True,
                }
            ]

        if not self.config.allow_cross_account:
            if self._active_account is None:
                return []
            return [
                {
                    "name": self._active_account,
                    "authenticated": self._active_account in self._credentials,
                    "active": True,
                }
            ]

        accounts = []
        for acc in self.config.accounts:
            accounts.append(
                {
                    "name": acc.name,
                    "client_id": acc.client_id[:8] + "...",
                    "tenant_id": acc.tenant_id,
                    "authenticated": acc.name in self._credentials,
                    "active": acc.name == self._active_account,
                }
            )
        return accounts

    def switch_account(self, name: str, capability: str | None = None) -> dict:
        """Switch the active account, or one capability's routing.

        Gated by allow_cross_account: with the gate closed there is nothing to
        switch to — the per-capability routing IS the mailbox the agent sees,
        and refusing here is the difference between "one merged account" and
        "one merged account with a door marked 'other accounts'".
        """
        if not self.config.allow_cross_account:
            raise ValueError(
                "Cross-account access is disabled (allow_cross_account=false). "
                "The per-capability account routing is the whole mailbox; to let "
                "the agent address other accounts' content, set "
                "allow_cross_account=true in ~/.outlook-mcp/config.json."
            )
        names = [a.name for a in self.config.accounts]
        if not names:
            raise ValueError("No multi-account config: 'accounts' is empty.")
        if name not in names:
            raise ValueError(f"Account '{name}' not found in config. Accounts: {names}")

        if capability is None:
            self._active_account = name
            self.credential = self._credentials.get(name)
            # An explicit switch gets honest results, not a fallback: if the
            # target has no credential, identity tools raise naming it.
            self._identity_fallback_account = None
            return {"status": "switched", "account": name}

        if capability not in ROUTING_CAPABILITIES:
            raise ValueError(
                f"Unknown capability '{capability}'. "
                f"Valid capabilities: {sorted(ROUTING_CAPABILITIES)}"
            )
        self._routing_overrides[capability] = name
        return {"status": "switched", "capability": capability, "account": name}

    def logout(self, account: str | None = None) -> dict[str, str]:
        """Clear in-memory credentials and one account's auth record.

        On a multi-account install a no-argument logout targets the default
        account; the other accounts' records stay on disk and their routed
        capabilities keep working on next start. Only the named account's
        in-memory credential and record file are removed — ``credential``
        (the active account's) is cleared only when the active account is
        the one being logged out.
        """
        if self.config.accounts and account is None:
            account = self.config.default_account
        if account is None:
            self.credential = None
            path = _auth_record_path()
            if path.exists():
                path.unlink()
            return {"status": "logged_out", "message": "Credentials cleared."}

        self._credentials.pop(account, None)
        if account == self._active_account:
            self.credential = None
        if account == self._identity_fallback_account:
            self._identity_fallback_account = None
        path = _auth_record_path(account)
        if path.exists():
            path.unlink()
        return {
            "status": "logged_out",
            "message": f"Credentials cleared for account '{account}'.",
        }
