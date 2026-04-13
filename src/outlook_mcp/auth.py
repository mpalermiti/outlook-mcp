"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import logging
import threading

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


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._account_email: str | None = None
        self._credentials: dict[str, DeviceCodeCredential] = {}  # name -> credential
        self._active_account: str | None = config.default_account

    def get_scopes(self) -> list[str]:
        """Return the appropriate scopes based on config."""
        return SCOPES_READONLY if self.config.read_only else SCOPES_READWRITE

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def login(self) -> dict[str, str]:
        """Start device code auth flow.

        Creates the credential, triggers token acquisition in a background
        thread, and waits for the device code callback to fire so we can
        return the verification URL and user code to the caller.

        Returns:
            Dict with 'status' and 'message' containing the device code URL
            and code for the user to complete sign-in.

        Raises:
            ValueError: If client_id is not configured.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json. See README for setup instructions."
            )

        cache_options = TokenCachePersistenceOptions(name="outlook-mcp")
        device_code_ready = threading.Event()
        self._device_code_message = ""

        # azure-identity prompt_callback receives (verification_uri, user_code, expires_on)
        def _on_device_code(verification_uri: str, user_code: str, expires_on: object) -> None:
            self._device_code_message = (
                f"To sign in, visit {verification_uri} and enter code: {user_code}"
            )
            device_code_ready.set()

        self.credential = DeviceCodeCredential(
            client_id=self.config.client_id,
            tenant_id=self.config.tenant_id,
            cache_persistence_options=cache_options,
            prompt_callback=_on_device_code,
        )

        scopes = self.get_scopes()
        self._auth_error: str | None = None

        def _acquire_token():
            try:
                self.credential.get_token(*scopes)
                logger.info("Device code auth completed successfully.")
            except Exception as e:
                self._auth_error = str(e)
                logger.exception("Device code auth failed.")
                device_code_ready.set()  # Unblock the wait

        # Start token acquisition in background — it blocks until user
        # completes browser sign-in, but the callback fires immediately
        # with the device code info.
        auth_thread = threading.Thread(target=_acquire_token, daemon=True)
        auth_thread.start()

        # Wait for the callback to fire (or timeout if token was cached)
        if device_code_ready.wait(timeout=15):
            if self._auth_error:
                return {
                    "status": "error",
                    "message": f"Authentication failed: {self._auth_error}",
                }
            return {
                "status": "login_started",
                "message": self._device_code_message,
            }

        # If we get here without the callback, the token was likely cached
        # — verify by checking if the thread completed without error
        auth_thread.join(timeout=5)
        if self._auth_error:
            return {
                "status": "error",
                "message": f"Authentication failed: {self._auth_error}",
            }

        return {
            "status": "authenticated",
            "message": "Already authenticated (cached token).",
        }

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
        # Also include the default single-account config if present
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
        """Clear stored credentials."""
        self.credential = None
        self._account_email = None
        return {"status": "logged_out", "message": "Credentials cleared."}
