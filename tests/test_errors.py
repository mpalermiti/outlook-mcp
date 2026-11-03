"""Tests for error hierarchy."""

import httpx

from outlook_mcp.errors import (
    AuthRequiredError,
    GraphAPIError,
    NotFoundError,
    OutlookMCPError,
    ReadOnlyError,
    left_in_drafts,
)


def test_base_error_fields():
    """OutlookMCPError stores code, message, action."""
    err = OutlookMCPError("test_code", "test message", "test action")
    assert err.code == "test_code"
    assert err.message == "test message"
    assert err.action == "test action"
    # str() is what the client receives, so the recovery hint has to be in it.
    assert str(err) == "test message test action"
    assert str(OutlookMCPError("c", "no hint here")) == "no hint here"


def test_auth_required_error():
    """AuthRequiredError has correct defaults."""
    err = AuthRequiredError()
    assert err.code == "auth_required"
    assert "not authenticated" in err.message.lower()
    assert "outlook-mcp auth" in err.action
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
    assert "outlook-mcp auth" in err.action
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


class TestLeftInDrafts:
    """The error for a reply or forward whose draft failed and could not be deleted."""

    def test_a_graph_error_keeps_its_status_message_and_hint_and_says_where_the_draft_is(self):
        original = GraphAPIError(400, "ErrorInvalidRequest", "Refused.", action="Fix it.")

        error = left_in_drafts(original, "AAMkDraft=")

        assert (error.status_code, error.error_code, error.message) == (
            400,
            "ErrorInvalidRequest",
            "Refused.",
        )
        assert error.action.startswith("Fix it. ")
        assert "may remain in Drafts" in error.action
        assert "The draft's id is AAMkDraft=." in error.action
        # A copy: the error passed in stays as it was, to be the cause.
        assert error is not original and original.action == "Fix it."

    def test_a_graph_error_with_no_hint_gets_only_the_drafts_note(self):
        error = left_in_drafts(GraphAPIError(400, "X", "Refused.", action=None), "D1")

        assert error.action.startswith("The new draft could not be deleted either")

    def test_an_sdk_error_is_converted_with_its_hint(self):
        from msgraph.generated.models.o_data_errors.main_error import MainError
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError

        sdk_error = ODataError()
        sdk_error.response_status_code = 429
        sdk_error.error = MainError(code="TooManyRequests", message="Slow down.")

        error = left_in_drafts(sdk_error, "D1")

        assert (error.status_code, error.message) == (429, "Slow down.")
        assert "may remain in Drafts" in error.action

    def test_any_other_failure_is_named_by_its_type_only(self):
        """Its text can carry the request URL."""
        url = "https://graph.microsoft.com/v1.0/me/messages/D1"

        error = left_in_drafts(httpx.ConnectError(url), "D1")

        assert (error.status_code, error.error_code) == (0, "DraftNotFinished")
        assert "ConnectError" in str(error) and "may remain in Drafts" in str(error)
        assert "graph.microsoft.com" not in str(error)
