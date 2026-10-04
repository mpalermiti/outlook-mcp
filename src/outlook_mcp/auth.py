"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import importlib.util
import logging
import sys
from pathlib import Path

from azure.core.exceptions import ClientAuthenticationError
from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)

from outlook_mcp.config import DEFAULT_CONFIG_DIR, Config, _ensure_dir, atomic_write
from outlook_mcp.errors import (
    AuthRequiredError,
    ClientIdMismatchError,
    OutlookMCPError,
    StaleConsentError,
    UnencryptedTokenCacheError,
)

logger = logging.getLogger(__name__)

# Process-local latch so the unencrypted-fallback warning fires at most
# once per run — _make_credential is called from both login_interactive
# and try_cached_token, often multiple times during startup.
_warned_unencrypted_fallback = False

# The concrete delegated scopes the first consent asks for — the full read-write
# set whatever read_only says: read_only gates the tools, not the token (#58),
# and the scopes a first consent omits can never be redeemed from the session
# it creates, so a consent narrowed by that flag could never be widened after
# the config flips. (SCOPES_READONLY below is the one, explicit, exception.)
# The FIRST consent names exactly these (login_interactive);
# every request afterwards — silent refresh and the Graph SDK's internal calls
# alike — uses .default, which on an already-consented session means precisely
# "the consented set". MSAL adds offline_access to its token requests itself,
# so it is deliberately absent.
SCOPES_READWRITE = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "MailboxSettings.Read",
    "User.Read",
]

# What `read_only_consent: true` asks for instead: the read half of each scope
# above, and no Mail.Send. This is what makes a second, read-only app
# registration usable — the route the README offers for a credential Microsoft
# itself limits. Asking that app for the write set either fails or grants it
# write access, and both defeat the point. It is its own config key, never
# inferred from read_only, for the reason given above; the config refuses the
# key without read_only, so the stranded-writes state cannot be configured.
SCOPES_READONLY = [
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "Tasks.Read",
    "MailboxSettings.Read",
    "User.Read",
]

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
# of the same condition. The refusal text is azure's own wording from
# azure/identity/_persistent_cache.py.
_AZURE_UNENCRYPTED_MARKER = "allow_unencrypted_storage"


def _is_azure_unencrypted_refusal(exc: BaseException) -> bool:
    """True for the "cache encryption is impossible" refusal, as it arrives.

    Every azure-identity token-acquisition path is wrapped in
    ``@wrap_exceptions``, which re-types anything that is not already a
    ``ClientAuthenticationError`` into one — message
    ``"Authentication failed: <original>"``, original exception on
    ``__cause__``. The persistent cache's refusal is a ``ValueError`` raised
    while building the cache inside those wrapped methods, so it never
    surfaces as a ``ValueError``: it arrives as a ``ClientAuthenticationError``
    whose message embeds azure's own wording — which is why matching the
    marker on the message alone is sufficient: the wrapper embeds the
    original's full text, cause included. The bare-``ValueError`` arm can
    only fire for a caller that bypassed a real credential.
    """
    if isinstance(exc, ValueError) and _AZURE_UNENCRYPTED_MARKER in str(exc):
        return True
    if not isinstance(exc, ClientAuthenticationError):
        return False
    return _AZURE_UNENCRYPTED_MARKER in str(exc)


# A session whose first consent went through .default alone can land with no
# delegated permissions on it, and redeeming any concrete scope from that
# session is refused with AADSTS70000 ("The requested user must first sign-in
# and grant the client application access"). AADSTS70000 is Entra's generic
# invalid-grant — a revoked refresh token reports it too — but every case it
# covers is a dead end for this process with the same exit: a fresh sign-in.
# The code embeds itself in the wrapped ClientAuthenticationError text like
# every MSAL error does, so matching the code on str(exc) is the whole check.
# The refusal's own wording reads like a retryable blip, which is why the
# remedy says "log in again" in so many words.
_AADSTS70000_MARKER = "AADSTS70000"


def _is_consent_dead_end(exc: BaseException) -> bool:
    """True for the "this session can never grant that scope" refusal."""
    return _AADSTS70000_MARKER in str(exc)


# The Graph SDK always requests .default scope internally, so silent refresh
# must acquire and cache tokens under the same scope — it is the cache key the
# SDK's calls land on. On a session that already consented concrete scopes,
# .default means exactly that consented set. Only the first consent must name
# the concrete scopes instead (see login_interactive).
GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"


def _auth_record_path() -> Path:
    # Derived from config.DEFAULT_CONFIG_DIR so the record sits next to
    # config.json wherever OUTLOOK_MCP_CONFIG_DIR has moved the settings
    # directory — this must stay the only place the record location is decided.
    return Path(DEFAULT_CONFIG_DIR) / AUTH_RECORD_FILE


def _save_auth_record(record: AuthenticationRecord) -> None:
    """Persist AuthenticationRecord to disk.

    Same atomic pattern as config.json (write-temp, fsync, 0600, rename):
    the record identifies the signed-in account, and a plain write_text
    leaves a world-readable window and a half-written file for anything
    that reads it concurrently.
    """
    path = _auth_record_path()
    # The settings directory is ours: create it 0700 like every other path
    # under it, not with the mkdir default that leaves group/other readable.
    _ensure_dir(str(path.parent))
    atomic_write(path, record.serialize())


def _load_auth_record() -> AuthenticationRecord | None:
    """Load AuthenticationRecord from disk, or None if not found."""
    path = _auth_record_path()
    if not path.exists():
        return None
    try:
        # Pairs with atomic_write's UTF-8. The record is ASCII JSON today (serialize() escapes
        # non-ASCII), so this is consistency with the writer, not a fix (#99).
        return AuthenticationRecord.deserialize(path.read_text(encoding="utf-8"))
    except Exception:
        logger.warning("Failed to load auth record from %s", path)
        return None


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        # Set when startup authentication failed for a reason the operator has
        # to fix in config rather than by running `outlook-mcp auth` — that
        # advice would just fail the same way. Surfaced by get_credential() so
        # the remedy reaches the agent on every tool call, not only stderr.
        self.startup_error: OutlookMCPError | None = None

    def get_scopes(self) -> list[str]:
        """Return the concrete delegated scopes the first consent asks for.

        The full read-write set, whatever ``read_only`` says: that flag gates
        the tools, not the token (#58). Which list this returns decides what
        the account is ever able to grant — the scopes a first consent omits
        cannot be redeemed from the session it creates, and because every
        later request renews via ``.default`` ("the consented set"), a consent
        narrowed by ``read_only`` could never be widened once that flag flips.

        ``read_only_consent`` is the explicit exception: the operator has said
        this app must never hold write access, and the config has already
        refused that key without ``read_only``.
        """
        return SCOPES_READONLY if self.config.read_only_consent else SCOPES_READWRITE

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def _make_credential(
        self,
        prompt_callback=None,
        auth_record: AuthenticationRecord | None = None,
        *,
        silent: bool = False,
    ) -> DeviceCodeCredential:
        """Create a DeviceCodeCredential with persistent cache.

        ``silent=True`` forbids the interactive device-code flow. azure-identity
        defaults to allowing it, so a *cache miss* on what is supposed to be a
        silent refresh does not fail — it prints a code and polls for a human
        until ``timeout`` (900s). On the startup path that is a fifteen-minute
        hang where the honest answer is "not authenticated"; `get_token` raises
        ``AuthenticationRequiredError`` instead when this is set.
        """
        global _warned_unencrypted_fallback
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
            "client_id": self.config.client_id,
            "tenant_id": self.config.tenant_id,
            "cache_persistence_options": cache_options,
            "timeout": 900,
        }
        if silent:
            kwargs["disable_automatic_authentication"] = True
        if prompt_callback:
            kwargs["prompt_callback"] = prompt_callback
        if auth_record:
            # azure-identity serves the record's identity, not this app's:
            # with an authentication_record it IGNORES the client_id passed
            # here and substitutes the record's own client_id and authority.
            # The tenant_id above still applies — only client_id is
            # overridden. That is what makes a per-instance record pin the
            # mailbox this credential talks to.
            kwargs["authentication_record"] = auth_record
        return DeviceCodeCredential(**kwargs)

    def login_interactive(self) -> None:
        """Run the device code flow interactively in the terminal.

        Always prompts: with no AuthenticationRecord to seed the credential,
        azure-identity's silent path cannot attempt a refresh, so the device
        code is printed and polled even when the cache holds a valid token.
        The flow still writes through to the persistent cache, and the saved
        AuthenticationRecord is what lets the MCP server refresh silently
        afterwards.

        Intended for CLI use (`outlook-mcp auth`), not MCP tools.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                f"client_id in {DEFAULT_CONFIG_DIR}/config.json."
            )

        def _on_device_code(verification_uri: str, user_code: str, expires_on: object) -> None:
            print(f"Visit:  {verification_uri}")
            print(f"Code:   {user_code}")
            print()
            print("Waiting for you to complete sign-in in your browser...")

        cred = self._make_credential(prompt_callback=_on_device_code)
        # get_token() consults the persistent cache first, but without a
        # record the silent path always misses (see the docstring above), so
        # this call is the interactive flow.
        #
        # The first consent must name the concrete delegated scopes, never
        # .default: on a personal account a .default-only first consent can
        # land a session with no delegated permissions, and no scope can be
        # redeemed from it afterwards (AADSTS70000) — only logging in again
        # fixes that. Once these scopes are consented, everything later uses
        # .default, which then means exactly this consented set. The list is
        # the read-write one whatever read_only says (see get_scopes): that
        # flag gates the tools, not the token. Only read_only_consent narrows it.
        try:
            cred.get_token(*self.get_scopes())
        except ClientAuthenticationError as exc:
            if _is_azure_unencrypted_refusal(exc):
                raise UnencryptedTokenCacheError() from exc
            raise

        # Save the auth record for silent refresh by the MCP server
        record = getattr(cred, "_auth_record", None)
        if record:
            _save_auth_record(record)

        self.credential = cred
        print("Authenticated successfully.")

    def try_cached_token(self) -> bool:
        """Try to get a token silently using a saved AuthenticationRecord.

        Returns True if a valid token was obtained without user interaction.
        Used by the MCP server on startup and by `outlook-mcp status`.
        """
        if not self.config.client_id:
            return False

        record = _load_auth_record()
        if record is None:
            return False

        # A record pins the app it was saved for: azure-identity uses the
        # record's client id and ignores the configured one (see
        # _make_credential). So a client_id changed in config.json has to be
        # caught here, or the old app's session — and whatever it was
        # consented for — goes on being used under the new id's name.
        saved_client_id = getattr(record, "client_id", None)
        if (
            isinstance(saved_client_id, str)
            and saved_client_id.lower() != self.config.client_id.lower()
        ):
            self.startup_error = ClientIdMismatchError(self.config.client_id, saved_client_id)
            logger.warning("%s", self.startup_error)
            return False

        try:
            cred = self._make_credential(auth_record=record, silent=True)
            # .default, not the consent list: this refreshes through the
            # saved record, and a record's session redeems .default however
            # it was consented — while a .default-only consent (the trap
            # login_interactive now avoids) redeems nothing else. The Graph
            # SDK asks under .default too, so this is also the cache key
            # that keeps its first call off the network.
            cred.get_token(GRAPH_DEFAULT_SCOPE)
            self.credential = cred
            return True
        except UnencryptedTokenCacheError:
            # Not a stale token — the environment cannot store one safely.
            # Swallowing it here sends the operator round the `outlook-mcp auth`
            # loop with no idea what to change.
            raise
        except Exception as exc:
            # Either the host cannot store a token safely — a config problem
            # with its own error, raised here — or this credential can no
            # longer serve its identity, which re-running `outlook-mcp auth`
            # actually fixes. Everything else is a stale-token-shaped failure
            # with the same remedy.
            if _is_azure_unencrypted_refusal(exc):
                raise UnencryptedTokenCacheError() from exc
            if _is_consent_dead_end(exc):
                # Name the dead end rather than offering the generic remedy:
                # "re-run auth" reads as an optional top-up, but no refresh
                # from this process can succeed again — the session cannot
                # grant what is being asked for. Surfaced as the startup
                # error so auth_status and every tool call carry it.
                self.startup_error = StaleConsentError()
                logger.warning("%s", self.startup_error)
                return False
            logger.warning("Cached token refresh failed — re-run `outlook-mcp auth`.")
            return False

    def get_credential(self) -> DeviceCodeCredential:
        """Get the current credential, raising if not authenticated."""
        if self.credential is None:
            if self.startup_error is not None:
                raise self.startup_error
            raise AuthRequiredError()
        return self.credential

    def logout(self) -> dict[str, str | bool]:
        """Clear in-memory credentials and remove this instance's auth record.

        The OS-level encrypted cache (Keychain item / DPAPI file / libsecret
        entry that azure-identity maintains) is shared across apps and is NOT
        touched — its tokens age out on their own. Callers report that part;
        this stays mechanical.
        """
        self.credential = None
        path = _auth_record_path()
        removed = path.exists()
        if removed:
            path.unlink()
        return {"status": "logged_out", "record_removed": removed}
