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
    "offline_access",
]

SCOPES_READONLY = [
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "Tasks.Read",
    "User.Read",
    "offline_access",
]


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._account_email: str | None = None

    def get_scopes(self) -> list[str]:
        """Return the appropriate scopes based on config."""
        return SCOPES_READONLY if self.config.read_only else SCOPES_READWRITE

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def login(self) -> dict[str, str]:
        """Start device code auth flow.

        Returns:
            Dict with 'status' and 'message' for the agent to display.

        Raises:
            ValueError: If client_id is not configured.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json. See README for setup instructions."
            )

        cache_options = TokenCachePersistenceOptions(name="outlook-mcp")

        # azure-identity prompt_callback receives a single dict argument
        def _on_device_code(device_code_info: dict) -> None:
            self._device_code_message = device_code_info.get("message", "")

        self.credential = DeviceCodeCredential(
            client_id=self.config.client_id,
            tenant_id=self.config.tenant_id,
            cache_persistence_options=cache_options,
            prompt_callback=_on_device_code,
        )

        return {
            "status": "login_started",
            "message": "Device code authentication initiated. "
            "Complete the sign-in when prompted.",
        }

    def get_credential(self) -> DeviceCodeCredential:
        """Get the current credential, raising if not authenticated."""
        if self.credential is None:
            raise AuthRequiredError()
        return self.credential

    def logout(self) -> dict[str, str]:
        """Clear stored credentials."""
        self.credential = None
        self._account_email = None
        return {"status": "logged_out", "message": "Credentials cleared."}
