"""Exception hierarchy for outlook-mcp."""

from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError


class OutlookMCPError(ToolError):
    """Base exception for all outlook-mcp errors.

    Subclasses ``ToolError`` — the SDK's "anticipated failure" — because that is
    what decides whether the model reads our message or a bare ``Error executing
    tool <name>``. Every error in this module is a condition we saw coming and
    wrote a recovery hint for, so every one of them belongs on that side of the
    line. A genuine crash must keep inheriting from ``Exception`` and stay
    generic to the client. Guarded by ``tests/test_error_text_reaches_client.py``.
    """

    def __init__(self, code: str, message: str, action: str | None = None):
        self.code = code
        self.message = message
        self.action = action
        super().__init__(message)

    def __str__(self) -> str:
        """Message plus recovery hint — this string is what the model reads.

        ``action`` exists to tell an agent what to do next, so it has to be in
        the text the client receives, not only on the attribute.
        """
        return f"{self.message} {self.action}" if self.action else self.message


class AuthRequiredError(OutlookMCPError):
    """Raised when a tool is called without authentication."""

    def __init__(self):
        super().__init__(
            "auth_required",
            "Not authenticated. No valid credential found.",
            "Run `outlook-mcp auth` on the host to authenticate.",
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
                action = "Token may have expired — run `outlook-mcp auth` on the host."
            elif status_code == 429:
                action = "Rate limited by Microsoft Graph. Wait a moment and retry."
        super().__init__(
            f"graph_api_{error_code}",
            message,
            action,  # type: ignore[arg-type]
        )
        self.status_code = status_code
        self.error_code = error_code


class UnencryptedTokenCacheError(OutlookMCPError):
    """Raised when the token cache would be written in cleartext un-asked.

    On Linux without libsecret, msal_extensions falls back to a plaintext cache
    file. That fallback used to be enabled unconditionally, so a reusable Graph
    refresh token could land on disk in the clear with only a log line to mark
    it. Persisting a credential that way is a decision the operator gets to
    make, so the default is now to stop and say so.
    """

    def __init__(self):
        super().__init__(
            "unencrypted_token_cache",
            "Refusing to persist the token cache: this environment has no "
            "encrypted store (Linux without libsecret/gnome-keyring), so the "
            "cache would be written to disk in cleartext.",
            "Either install the system packages (apt: `gnome-keyring "
            "libsecret-1-0 python3-gi`) and re-create the venv with "
            "`--system-site-packages`, or accept plaintext storage by setting "
            '`"allow_unencrypted_token_cache": true` in '
            "~/.outlook-mcp/config.json. See "
            "https://github.com/mpalermiti/outlook-mcp/issues/7.",
        )


class UntrustedURLError(OutlookMCPError):
    """Raised when a URL that would receive a Graph token isn't a Graph URL.

    The delta tools hand their cursor straight back to Graph as a request URL,
    with the mailbox bearer token attached. That cursor is caller-held state and
    the caller is an agent that reads mail, so the cursor is untrusted input:
    left unchecked it redirects a live full-mailbox token to any host that can
    get a string in front of the model. Refusing is the only safe answer — there
    is no partial-trust version of "send the token somewhere else".
    """

    def __init__(self, source: str, url: str):
        shown = url[:120] if url else "(empty)"
        super().__init__(
            "untrusted_url",
            f"Refusing to send a Microsoft Graph token to a non-Graph URL "
            f"(from {source}): {shown!r}.",
            "Delta cursors must be https URLs on graph.microsoft.com. Discard "
            "this cursor and start a fresh sync by calling again with no "
            "delta_token.",
        )
        self.source = source


class ToolInputError(OutlookMCPError, ValueError):
    """A tool argument the caller can fix, raised as an anticipated failure.

    Tool modules raise plain ``ValueError`` for bad input and the server's
    ``_wrap_tool_errors`` converts it here, so the message reaches the model
    (SEP-1303: input validation errors are tool execution errors, not protocol
    errors). Still a ``ValueError``, so callers catching that are unaffected.
    """

    def __init__(self, message: str):
        super().__init__("invalid_input", message, None)


# ── Graph error wrapper ────────────────────────────────────


# Hint table keyed by (status_code, error_code).
# An error_code of ``None`` matches any error_code for that status_code.
_HINT_TABLE: dict[tuple[int, str | None], str] = {
    (401, None): "Token may have expired — run `outlook-mcp auth` on the host.",
    (403, "ErrorAccessDenied"): (
        "Endpoint may not be supported for this account type. See "
        "https://github.com/mpalermiti/outlook-mcp/blob/main/ROADMAP.md"
        "#investigated-and-not-viable for known dead-ends."
    ),
    (400, "ErrorPropertyValidationFailure"): (
        "Graph rejected a property but did not say which. On calendar writes this "
        "has two observed causes, both about time zones: a "
        "recurrence.range.recurrenceTimeZone naming a different zone than the "
        "event's own (omit it — Graph derives it from the event), or changing a "
        "series master's zone without re-sending its recurrence in the same call."
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
