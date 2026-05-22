"""Tests for the Graph SDK -> GraphAPIError wrapper."""

from __future__ import annotations

import pytest
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from outlook_mcp.errors import GraphAPIError, wrap_graph_error


def _make_odata_error(status_code: int, code: str | None, message: str | None) -> ODataError:
    """Build an ODataError fixture matching the shape Graph would deserialize.

    The Kiota request adapter populates ``response_status_code`` on the base
    APIError and ``error`` with a MainError that holds the user-facing
    ``code`` and ``message``.
    """
    inner = MainError()
    inner.code = code
    inner.message = message

    err = ODataError()
    err.response_status_code = status_code
    err.message = message  # APIError-level message; usually mirrors the inner one
    err.error = inner
    return err


def test_wrap_401_returns_auth_hint():
    """401 maps to the host-side `outlook-mcp auth` hint."""
    exc = _make_odata_error(401, "InvalidAuthenticationToken", "Token expired")
    wrapped = wrap_graph_error(exc)
    assert isinstance(wrapped, GraphAPIError)
    assert wrapped.status_code == 401
    assert wrapped.error_code == "InvalidAuthenticationToken"
    assert wrapped.action is not None
    assert "outlook-mcp auth" in wrapped.action


def test_wrap_403_access_denied_references_roadmap():
    """403/ErrorAccessDenied carries the unsupported-endpoint hint with ROADMAP pointer."""
    exc = _make_odata_error(403, "ErrorAccessDenied", "Access denied")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 403
    assert wrapped.error_code == "ErrorAccessDenied"
    assert wrapped.action is not None
    assert "ROADMAP" in wrapped.action


def test_wrap_403_unknown_subcode_has_no_hint():
    """403 with an error_code we don't recognize falls through to None."""
    exc = _make_odata_error(403, "SomeOtherDeny", "nope")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 403
    # No (403, None) entry, no (403, "SomeOtherDeny") entry.
    assert wrapped.action is None


def test_wrap_404_item_not_found_hint():
    """404/ErrorItemNotFound suggests re-listing for fresh IDs."""
    exc = _make_odata_error(404, "ErrorItemNotFound", "Not found")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 404
    assert wrapped.action is not None
    assert "stale" in wrapped.action.lower() or "re-list" in wrapped.action.lower()


def test_wrap_429_rate_limit_hint():
    """429 mentions Retry-After and backoff."""
    exc = _make_odata_error(429, "TooManyRequests", "Rate limited")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 429
    assert wrapped.action is not None
    assert "retry" in wrapped.action.lower()


def test_wrap_503_transient_hint():
    """503 mentions transient availability."""
    exc = _make_odata_error(503, "ServiceUnavailable", "Try again")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 503
    assert wrapped.action is not None
    assert "unavailable" in wrapped.action.lower() or "retry" in wrapped.action.lower()


def test_wrap_500_unknown_has_no_hint():
    """500 (or any code we don't have in the table) returns action=None."""
    exc = _make_odata_error(500, "InternalServerError", "boom")
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 500
    assert wrapped.error_code == "InternalServerError"
    assert wrapped.action is None


def test_wrap_non_graph_exception_raises_typeerror():
    """Passing a non-Graph exception trips a TypeError so the caller can re-raise."""
    with pytest.raises(TypeError):
        wrap_graph_error(ValueError("not a graph error"))


def test_wrap_preserves_message():
    """The wrapper carries through the human-friendly message from MainError."""
    exc = _make_odata_error(
        403,
        "ErrorAccessDenied",
        "ErrorAccessDenied: mailbox setting unsupported",
    )
    wrapped = wrap_graph_error(exc)
    assert "mailbox setting" in wrapped.message


def test_wrap_api_error_without_inner_error():
    """A bare APIError without an ``error`` payload still gets wrapped sanely."""
    from kiota_abstractions.api_error import APIError

    exc = APIError()
    exc.response_status_code = 429
    wrapped = wrap_graph_error(exc)
    assert wrapped.status_code == 429
    # Generic error_code fallback when none is provided.
    assert wrapped.error_code == "UnknownError"
    # 429 table entry still applies.
    assert wrapped.action is not None
    assert "retry" in wrapped.action.lower()
