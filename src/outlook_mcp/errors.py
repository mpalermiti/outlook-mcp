"""Exception hierarchy for outlook-mcp."""


class OutlookMCPError(Exception):
    """Base exception for all outlook-mcp errors."""

    def __init__(self, code: str, message: str, action: str | None = None):
        self.code = code
        self.message = message
        self.action = action
        super().__init__(message)


class AuthRequiredError(OutlookMCPError):
    """Raised when a tool is called without authentication."""

    def __init__(self):
        super().__init__(
            "auth_required",
            "Not authenticated. No valid credential found.",
            "Call outlook_login to authenticate with your Microsoft account.",
        )


class ReadOnlyError(OutlookMCPError):
    """Raised when a write tool is called in read-only mode."""

    def __init__(self, tool_name: str):
        super().__init__(
            "read_only",
            f"Cannot use {tool_name} — server is in read-only mode.",
            "Set read_only to false in ~/.outlook-mcp/config.json to enable write operations.",
        )


class NotFoundError(OutlookMCPError):
    """Raised when a requested resource doesn't exist."""

    def __init__(self, resource: str, resource_id: str):
        super().__init__(
            "not_found",
            f"{resource} '{resource_id}' not found.",
            None,
        )


class GraphAPIError(OutlookMCPError):
    """Raised when the Graph API returns an error."""

    def __init__(self, status_code: int, error_code: str, message: str):
        action = None
        if status_code == 401:
            action = "Token may have expired. Try outlook_login to re-authenticate."
        elif status_code == 429:
            action = "Rate limited by Microsoft Graph. Wait a moment and retry."
        super().__init__(
            f"graph_api_{error_code}",
            message,
            action,
        )
        self.status_code = status_code
