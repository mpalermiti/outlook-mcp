"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import logging

from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions

from outlook_mcp.config import Config
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
        self, prompt_callback=None
    ) -> DeviceCodeCredential:
        """Create a DeviceCodeCredential with persistent cache."""
        cache_options = TokenCachePersistenceOptions(name=CACHE_NAME)
        return DeviceCodeCredential(
            client_id=self.config.client_id,
            tenant_id=self.config.tenant_id,
            cache_persistence_options=cache_options,
            prompt_callback=prompt_callback,
        )

    def login_interactive(self, scopes: list[str]) -> None:
        """Run the device code flow interactively in the terminal.

        Blocks until the user completes browser sign-in.
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
        # Blocks until auth completes or fails
        cred.get_token(*scopes)
        self.credential = cred
        print("Authenticated successfully.")

    def try_cached_token(self, scopes: list[str]) -> bool:
        """Try to get a token from cache without user interaction.

        Returns True if a valid cached token was found.
        Used by the MCP server on startup and by `outlook-mcp status`.
        """
        if not self.config.client_id:
            return False

        try:
            cred = self._make_credential()
            cred.get_token(*scopes)
            self.credential = cred
            return True
        except Exception:
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
        """Clear in-memory credentials."""
        self.credential = None
        return {"status": "logged_out", "message": "Credentials cleared."}
