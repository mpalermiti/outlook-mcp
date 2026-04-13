"""Tests for error hierarchy."""

from outlook_mcp.errors import (
    AuthRequiredError,
    GraphAPIError,
    NotFoundError,
    OutlookMCPError,
    ReadOnlyError,
)


def test_base_error_fields():
    """OutlookMCPError stores code, message, action."""
    err = OutlookMCPError("test_code", "test message", "test action")
    assert err.code == "test_code"
    assert err.message == "test message"
    assert err.action == "test action"
    assert str(err) == "test message"


def test_auth_required_error():
    """AuthRequiredError has correct defaults."""
    err = AuthRequiredError()
    assert err.code == "auth_required"
    assert "not authenticated" in err.message.lower()
    assert "outlook_login" in err.action
    assert isinstance(err, OutlookMCPError)


def test_read_only_error_includes_tool_name():
    """ReadOnlyError includes the tool name in message."""
    err = ReadOnlyError("outlook_send_message")
    assert "outlook_send_message" in err.message
    assert err.code == "read_only"
    assert "read_only" in err.action.lower() or "config" in err.action.lower()
    assert isinstance(err, OutlookMCPError)


def test_not_found_error():
    """NotFoundError includes resource type and ID."""
    err = NotFoundError("message", "AAMkAG123=")
    assert "message" in err.message
    assert "AAMkAG123=" in err.message
    assert err.code == "not_found"
    assert err.action is None
    assert isinstance(err, OutlookMCPError)


def test_graph_api_error_401():
    """GraphAPIError maps 401 to re-auth action."""
    err = GraphAPIError(401, "InvalidAuthenticationToken", "Token expired")
    assert err.status_code == 401
    assert err.code == "graph_api_InvalidAuthenticationToken"
    assert err.action is not None
    assert "re-authenticate" in err.action.lower() or "login" in err.action.lower()
    assert isinstance(err, OutlookMCPError)


def test_graph_api_error_429():
    """GraphAPIError maps 429 to rate limit action."""
    err = GraphAPIError(429, "TooManyRequests", "Rate limited")
    assert err.status_code == 429
    assert "rate limit" in err.action.lower() or "retry" in err.action.lower()


def test_graph_api_error_500():
    """GraphAPIError with unknown status has no action."""
    err = GraphAPIError(500, "InternalServerError", "Server error")
    assert err.status_code == 500
    assert err.action is None


def test_all_errors_are_exceptions():
    """All custom errors can be caught as Exception."""
    for err_class in [AuthRequiredError, ReadOnlyError, NotFoundError, GraphAPIError]:
        if err_class == AuthRequiredError:
            err = err_class()
        elif err_class == ReadOnlyError:
            err = err_class("test_tool")
        elif err_class == NotFoundError:
            err = err_class("resource", "id")
        else:
            err = err_class(400, "BadRequest", "Bad request")
        assert isinstance(err, Exception)
        assert isinstance(err, OutlookMCPError)
