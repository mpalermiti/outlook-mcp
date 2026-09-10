"""Guards that a tool's error text actually reaches the model.

The mock suite asserts on exception *types* (`pytest.raises(AuthRequiredError)`),
which is not the same thing as the client seeing the message. Under the 2.x SDK
those diverged: `MCPServer` forwards the text of an anticipated failure
(`ToolError`) but treats every other exception as a crash, replacing the text
with a bare ``Error executing tool <name>``. The whole error hierarchy —
including the ``action`` recovery hints — was landing on the model as that one
sentence while every type assertion stayed green.

So these tests go through a real client over the real decorator and assert on
the string the model reads, not on the exception the server raised.
"""

import pytest
from mcp.client import Client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mcp.errors import (
    AuthRequiredError,
    GraphAPIError,
    NotFoundError,
    OutlookMCPError,
    PermissionDeniedError,
    ReadOnlyError,
)
from outlook_mcp.server import _wrap_tool_errors


async def _text_the_model_sees(raiser) -> tuple[bool, str]:
    """Call a tool that raises, over a real client. Returns (is_error, text)."""
    server = MCPServer("error-text-guard", version="0")

    @server.tool()
    @_wrap_tool_errors
    async def failing_tool() -> dict:
        raise raiser()

    async with Client(server) as client:
        result = await client.call_tool("failing_tool", {})

    return result.is_error, result.content[0].text


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("raiser", "expected_fragments"),
    [
        pytest.param(
            AuthRequiredError,
            ["Not authenticated", "outlook-mcp auth"],
            id="auth_required",
        ),
        pytest.param(
            lambda: ReadOnlyError("outlook_send_message"),
            ["outlook_send_message", "read-only mode", "config.json"],
            id="read_only",
        ),
        pytest.param(
            lambda: PermissionDeniedError("outlook_delete_message", "mail"),
            ["outlook_delete_message", "allow_categories"],
            id="permission_denied",
        ),
        pytest.param(
            lambda: NotFoundError("Message", "AAMk123"),
            ["Message", "AAMk123", "not found"],
            id="not_found",
        ),
        pytest.param(
            lambda: GraphAPIError(429, "TooManyRequests", "Too many requests."),
            ["Too many requests", "Rate limited by Microsoft Graph"],
            id="graph_api_error_with_hint",
        ),
        pytest.param(
            lambda: ValueError("Invalid datetime: 'oops' — use ISO 8601."),
            ["Invalid datetime", "ISO 8601"],
            id="validation_error",
        ),
    ],
)
async def test_error_text_reaches_the_client(raiser, expected_fragments):
    """The message — and the recovery hint — arrive as content the model can act on."""
    is_error, text = await _text_the_model_sees(raiser)

    assert is_error is True
    for fragment in expected_fragments:
        assert fragment in text, f"{fragment!r} missing from client-visible text: {text!r}"


@pytest.mark.asyncio
async def test_graph_sdk_error_hint_reaches_the_client():
    """The real path: an ODataError becomes a GraphAPIError whose hint the model reads.

    The 403/`ErrorAccessDenied` hint points at the ROADMAP's dead-ends list, which
    is the difference between an agent retrying forever and an agent stopping.
    """
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    def _odata_403():
        inner = MainError()
        inner.code = "ErrorAccessDenied"
        inner.message = "Access is denied."

        err = ODataError()
        err.response_status_code = 403
        err.message = "Access is denied."
        err.error = inner
        return err

    is_error, text = await _text_the_model_sees(_odata_403)

    assert is_error is True
    assert "Access is denied" in text
    assert "ROADMAP" in text


@pytest.mark.asyncio
async def test_unexpected_exception_stays_generic():
    """A crash is not an anticipated failure: its text stays on the server."""

    class WeirdError(Exception):
        pass

    is_error, text = await _text_the_model_sees(
        lambda: WeirdError("internal detail that must not leak")
    )

    assert is_error is True
    assert "internal detail" not in text


def _all_subclasses(cls: type) -> set[type]:
    subs = set(cls.__subclasses__())
    return subs.union(*(_all_subclasses(s) for s in subs)) if subs else subs


def test_every_domain_error_is_an_anticipated_failure():
    """Drift guard: re-parenting the hierarchy away from ToolError re-breaks this."""
    assert issubclass(OutlookMCPError, ToolError)
    for subclass in _all_subclasses(OutlookMCPError):
        assert issubclass(subclass, ToolError), (
            f"{subclass.__name__} would reach the model as 'Error executing tool <name>'"
        )
