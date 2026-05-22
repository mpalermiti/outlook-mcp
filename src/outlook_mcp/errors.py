"""Exception hierarchy for outlook-mcp."""

from __future__ import annotations


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


class PermissionDeniedError(OutlookMCPError):
    """Raised when a write tool is not in the user's allow_categories."""

    def __init__(self, tool_name: str, category: str):
        super().__init__(
            "permission_denied",
            f"Cannot use {tool_name} — category '{category}' is not in allow_categories.",
            (
                f"Add '{category}' to allow_categories in ~/.outlook-mcp/config.json, "
                "or unset allow_categories for full write access."
            ),
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
    """Raised when the Graph API returns an error.

    ``action`` is optional. If omitted, a default hint is picked from the
    legacy 401/429 table (preserved for back-compat with existing call
    sites). Pass ``action=<string>`` (or ``action=None`` explicitly to
    suppress) when constructing from ``wrap_graph_error`` — the wrapper
    selects a richer hint from a (status_code, error_code) table.
    """

    _SENTINEL = object()

    def __init__(
        self,
        status_code: int,
        error_code: str,
        message: str,
        action: str | None | object = _SENTINEL,
    ):
        if action is GraphAPIError._SENTINEL:
            # Legacy behavior: derive action from status_code only.
            action = None
            if status_code == 401:
                action = "Token may have expired. Try outlook_login to re-authenticate."
            elif status_code == 429:
                action = "Rate limited by Microsoft Graph. Wait a moment and retry."
        super().__init__(
            f"graph_api_{error_code}",
            message,
            action,  # type: ignore[arg-type]
        )
        self.status_code = status_code
        self.error_code = error_code


# ── Graph error wrapper ────────────────────────────────────


# Hint table keyed by (status_code, error_code).
# An error_code of ``None`` matches any error_code for that status_code.
_HINT_TABLE: dict[tuple[int, str | None], str] = {
    (401, None): "Token may have expired — run `outlook-mcp auth` on the host.",
    (403, "ErrorAccessDenied"): (
        "Endpoint may not be supported for this account type. "
        "See ROADMAP 'Investigated and not viable' for known dead-ends."
    ),
    (404, "ErrorItemNotFound"): (
        "Resource not found. The ID may be stale — re-list to get current IDs."
    ),
    (429, None): (
        "Rate limited by Microsoft Graph. "
        "Back off and retry; respect any Retry-After header."
    ),
    (503, None): (
        "Microsoft Graph is temporarily unavailable. Retry after a short delay."
    ),
}


def _lookup_hint(status_code: int | None, error_code: str | None) -> str | None:
    """Look up a recovery hint for a (status_code, error_code) pair.

    Prefers an exact (code, error_code) match; falls back to (code, None).
    Returns ``None`` when nothing matches.
    """
    if status_code is None:
        return None
    if error_code is not None:
        hit = _HINT_TABLE.get((status_code, error_code))
        if hit is not None:
            return hit
    return _HINT_TABLE.get((status_code, None))


def wrap_graph_error(exc: Exception) -> GraphAPIError:
    """Convert a Graph SDK exception into a structured ``GraphAPIError``.

    Catches the msgraph SDK's ``ODataError`` and its parent
    ``kiota_abstractions.api_error.APIError``. Extracts status code,
    error code, and message, and attaches a recovery hint when known.

    Raises ``TypeError`` if ``exc`` is not a recognized Graph SDK error —
    callers should pass through non-Graph exceptions unchanged.
    """
    # Lazy imports — these are heavy and only needed when an error actually
    # surfaces from the SDK.
    from kiota_abstractions.api_error import APIError

    try:
        from msgraph.generated.models.o_data_errors.o_data_error import (
            ODataError as _ODataError,
        )
        graph_types: tuple[type, ...] = (APIError, _ODataError)
    except ImportError:  # pragma: no cover — defensive
        graph_types = (APIError,)

    if not isinstance(exc, graph_types):
        raise TypeError(
            f"wrap_graph_error: not a Graph SDK error: {type(exc).__name__}"
        )

    status_code: int | None = getattr(exc, "response_status_code", None)

    error_code: str | None = None
    message: str | None = getattr(exc, "message", None)

    inner = getattr(exc, "error", None)
    if inner is not None:
        # ODataError.error is a MainError(code, message, ...)
        ec = getattr(inner, "code", None)
        if ec:
            error_code = ec
        em = getattr(inner, "message", None)
        if em:
            # MainError.message is generally the user-friendly one.
            message = em

    if not error_code:
        error_code = "UnknownError"
    if not message:
        message = f"Graph API error (status {status_code})"

    hint = _lookup_hint(status_code, error_code)

    return GraphAPIError(
        status_code if status_code is not None else 0,
        error_code,
        message,
        action=hint,
    )
