"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import logging
from pathlib import Path

from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)

from outlook_mcp.config import DEFAULT_CONFIG_DIR, Config
from outlook_mcp.errors import AuthRequiredError

logger = logging.getLogger(__name__)

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

CACHE_NAME = "outlook-mcp"
AUTH_RECORD_FILE = "auth_record.json"


def _auth_record_path() -> Path:
    return Path(DEFAULT_CONFIG_DIR) / AUTH_RECORD_FILE


def _save_auth_record(record: AuthenticationRecord) -> None:
    """Persist AuthenticationRecord to disk."""
    path = _auth_record_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(record.serialize())
    path.chmod(0o600)


def _load_auth_record() -> AuthenticationRecord | None:
    """Load AuthenticationRecord from disk, or None if not found."""
    path = _auth_record_path()
    if not path.exists():
        return None
    try:
        return AuthenticationRecord.deserialize(path.read_text())
    except Exception:
        logger.warning("Failed to load auth record from %s", path)
        return None


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._credentials: dict[str, DeviceCodeCredential] = {}
        self._active_account: str | None = config.default_account

    def get_scopes(self) -> list[str]:
        """Return the appropriate scopes based on config."""
        return SCOPES_READONLY if self.config.read_only else SCOPES_READWRITE

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def _make_credential(
        self,
        prompt_callback=None,
        auth_record: AuthenticationRecord | None = None,
    ) -> DeviceCodeCredential:
        """Create a DeviceCodeCredential with persistent cache."""
        cache_options = TokenCachePersistenceOptions(name=CACHE_NAME)
        kwargs = {
            "client_id": self.config.client_id,
            "tenant_id": self.config.tenant_id,
            "cache_persistence_options": cache_options,
        }
        if prompt_callback:
            kwargs["prompt_callback"] = prompt_callback
        if auth_record:
            kwargs["authentication_record"] = auth_record
        return DeviceCodeCredential(**kwargs)

    def login_interactive(self, scopes: list[str]) -> None:
        """Run the device code flow interactively in the terminal.

        Uses get_token() which respects the token cache — if a valid
        cached token exists, completes silently. Otherwise triggers the
        device code flow. Saves the AuthenticationRecord for silent
        token refresh by the MCP server.

        Intended for CLI use (`outlook-mcp auth`), not MCP tools.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json."
            )

        def _on_device_code(
            verification_uri: str, user_code: str, expires_on: object
        ) -> None:
            print(f"Visit:  {verification_uri}")
            print(f"Code:   {user_code}")
            print()
            print("Waiting for you to complete sign-in in your browser...")

        cred = self._make_credential(prompt_callback=_on_device_code)
        # get_token() uses cache first, falls back to interactive
        cred.get_token(*scopes)

        # Save the auth record for silent refresh by the MCP server
        record = getattr(cred, "_auth_record", None)
        if record:
            _save_auth_record(record)

        self.credential = cred
        print("Authenticated successfully.")

    def try_cached_token(self, scopes: list[str]) -> bool:
        """Try to get a token silently using a saved AuthenticationRecord.

        Returns True if a valid token was obtained without user interaction.
        Used by the MCP server on startup and by `outlook-mcp status`.
        """
        if not self.config.client_id:
            return False

        record = _load_auth_record()
        if record is None:
            return False

        try:
            cred = self._make_credential(auth_record=record)
            cred.get_token(*scopes)
            self.credential = cred
            return True
        except Exception:
            logger.warning("Cached token refresh failed — re-run `outlook-mcp auth`.")
            return False

    def get_credential(self) -> DeviceCodeCredential:
        """Get the current credential, raising if not authenticated."""
        if self.credential is None:
            raise AuthRequiredError()
        return self.credential

    def list_accounts(self) -> list[dict]:
        """List configured accounts with auth status."""
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
        if self.config.client_id and not self.config.accounts:
            accounts.append(
                {
                    "name": "default",
                    "client_id": self.config.client_id[:8] + "...",
                    "tenant_id": self.config.tenant_id,
                    "authenticated": self.credential is not None,
                    "active": True,
                }
            )
        return accounts

    def switch_account(self, name: str) -> dict:
        """Switch active account."""
        for acc in self.config.accounts:
            if acc.name == name:
                self._active_account = name
                if name in self._credentials:
                    self.credential = self._credentials[name]
                else:
                    self.credential = None
                return {"status": "switched", "account": name}
        raise ValueError(f"Account '{name}' not found in config")

    def logout(self) -> dict[str, str]:
        """Clear in-memory credentials and auth record."""
        self.credential = None
        path = _auth_record_path()
        if path.exists():
            path.unlink()
        return {"status": "logged_out", "message": "Credentials cleared."}
