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

import logging
from contextlib import ExitStack
from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRequiredError
from mcp.client import Client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from outlook_mcp.errors import (
    AccountMismatchError,
    AuthRequiredError,
    GraphAPIError,
    NotFoundError,
    OutlookMCPError,
    PermissionDeniedError,
    ReadOnlyError,
)
from outlook_mcp.graph import GraphClient
from outlook_mcp.server import _wrap_tool_errors
from outlook_mcp.tools.mail_delta import list_inbox_delta
from outlook_mcp.tools.mail_read import read_messages


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
            AccountMismatchError,
            [
                "expected_username",
                "outlook-mcp auth",
                "or correct expected_username if it is wrong",
                "do not change",
            ],
            id="account_mismatch",
        ),
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
async def test_pinned_write_conflict_hint_reaches_the_client():
    """A 412 on a pinned write may follow a retry of a write that landed (#96).

    The hint the model reads must say so — "nothing was modified" there makes an
    agent report a successful update as a failure — and must send it to re-read
    before trying again.
    """
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    def _odata_412():
        inner = MainError()
        inner.code = "ErrorIrresolvableConflict"
        inner.message = "The change key passed in the request does not match."

        err = ODataError()
        err.response_status_code = 412
        err.error = inner
        return err

    is_error, text = await _text_the_model_sees(_odata_412)

    assert is_error is True
    assert "change key" in text
    assert "may already have been applied" in text
    assert "Re-read the item" in text
    assert "Nothing was modified" not in text


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


@pytest.mark.asyncio
async def test_a_send_answered_5xx_tells_the_model_it_may_have_gone():
    """A 5xx on a send leaves the outcome unknown (#131).

    The model must read that the message may already have been sent and where to
    check, not the 503 row's "retry": for a send, retrying blind is the duplicate.
    The error is the one ``call_without_replay`` raises from Graph's own 503.
    """
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    from outlook_mcp.errors import MESSAGE_MAY_HAVE_BEEN_SENT
    from outlook_mcp.throttle import call_without_replay

    async def _send(_once):
        inner = MainError()
        inner.code = "ServiceUnavailable"
        inner.message = "The service is temporarily unavailable."

        err = ODataError()
        err.response_status_code = 503
        err.error = inner
        raise err

    with pytest.raises(GraphAPIError) as caught:
        await call_without_replay(_send, outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT)

    is_error, text = await _text_the_model_sees(lambda: caught.value)

    assert is_error is True
    assert "temporarily unavailable" in text
    assert "may already have been sent" in text
    assert "Check Sent Items" in text
    assert "Retry after a short delay" not in text


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


# ── Failures that do not come from the Graph SDK ──────────────────────────
#
# Everything above raises an `OutlookMCPError` or an SDK `ODataError`. Three
# tool paths fail with neither: the raw-httpx delta and `$batch` calls
# (`httpx.HTTPStatusError`, and a reply that is not JSON), and token refresh
# (azure-identity's `ClientAuthenticationError`, raised the same way on the SDK
# path and the raw ones). These go through the real tool functions, so the
# exception is the one that path really raises, not one built to fit.

_DELTA_CURSOR = (
    "https://graph.microsoft.com/v1.0/me/mailFolders('inbox')/messages/delta"
    "?$deltatoken=CURSOR-SECRET-123"
)
_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch"


async def _text_the_model_sees_from(call) -> tuple[bool, str]:
    """Like ``_text_the_model_sees``, for a tool that awaits ``call()``."""
    server = MCPServer("error-text-guard", version="0")

    @server.tool()
    @_wrap_tool_errors
    async def failing_tool() -> dict:
        return await call()

    async with Client(server) as client:
        result = await client.call_tool("failing_tool", {})

    return result.is_error, result.content[0].text


def _fake_graph_client(get_token=None):
    client = MagicMock()
    client.credential.get_token = get_token or MagicMock(return_value=MagicMock(token="t"))
    return client


def _httpx_answering(method: str, response: httpx.Response) -> ExitStack:
    """Patch both raw-httpx call sites so every ``method`` returns ``response``."""
    fake = MagicMock()
    setattr(fake, method, AsyncMock(return_value=response))
    fake.__aenter__ = AsyncMock(return_value=fake)
    fake.__aexit__ = AsyncMock(return_value=False)

    stack = ExitStack()
    for module in ("outlook_mcp.tools._delta", "outlook_mcp.tools.mail_read"):
        stack.enter_context(patch(f"{module}.httpx.AsyncClient", return_value=fake))
    # The envelope retry on 429/503 sleeps between attempts; nothing to wait for here.
    stack.enter_context(patch("outlook_mcp.throttle.asyncio.sleep", new=AsyncMock()))
    return stack


def _delta_call(graph_client=None):
    return lambda: list_inbox_delta(graph_client or _fake_graph_client(), delta_token=_DELTA_CURSOR)


def _batch_call():
    return lambda: read_messages(_fake_graph_client(), ["AAMkAG1="])


@pytest.mark.asyncio
async def test_a_stale_delta_token_tells_the_agent_to_start_over(caplog):
    """A 410 on a stored cursor is the one delta failure with a known exit.

    `outlook_changes_since` resyncs on it by itself; the three standalone delta
    tools hand it to the agent, which until now read `Error executing tool …`
    and kept the dead cursor. The text must say what to do instead, and must not
    echo the request URL: that is the cursor, token and all. The server log must
    not either: the SDK logs a crash with its traceback, whose httpx message
    names the URL, while an anticipated failure is logged as its text alone.
    """
    caplog.set_level(logging.DEBUG)
    gone = httpx.Response(
        410,
        json={
            "error": {
                "code": "SyncStateNotFound",
                "message": "The sync state generation is not found.",
            }
        },
        request=httpx.Request("GET", _DELTA_CURSOR),
    )
    with _httpx_answering("get", gone):
        is_error, text = await _text_the_model_sees_from(_delta_call())

    assert is_error is True
    assert "The sync state generation is not found." in text
    assert "no delta_token" in text
    assert "CURSOR-SECRET-123" not in text
    assert "graph.microsoft.com" not in text
    assert "CURSOR-SECRET-123" not in caplog.text


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("status", "body", "expected"),
    [
        pytest.param(
            400,
            {"json": {"error": {"code": "BadRequest", "message": "Invalid batch payload."}}},
            ["Invalid batch payload."],
            id="graph_error_body",
        ),
        pytest.param(
            429,
            {
                "json": {
                    "error": {"code": "TooManyRequests", "message": "Please retry again later."}
                }
            },
            ["Please retry again later.", "Rate limited by Microsoft Graph"],
            id="still_throttled_after_retries",
        ),
        pytest.param(
            502,
            {"text": "<html><body>Bad Gateway</body></html>"},
            ["status 502"],
            id="error_body_that_is_not_json",
        ),
    ],
)
async def test_a_failed_batch_envelope_carries_graphs_message(status, body, expected):
    """`outlook_read_messages` raises on a non-2xx `$batch`; the model reads why."""
    failed = httpx.Response(status, request=httpx.Request("POST", _BATCH_URL), **body)
    with _httpx_answering("post", failed):
        is_error, text = await _text_the_model_sees_from(_batch_call())

    assert is_error is True
    for fragment in expected:
        assert fragment in text, f"{fragment!r} missing from client-visible text: {text!r}"
    assert "graph.microsoft.com" not in text


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("method", "call"),
    [
        pytest.param("get", _delta_call, id="delta"),
        pytest.param("post", _batch_call, id="batch"),
    ],
)
@pytest.mark.parametrize(
    "body",
    [
        pytest.param("<html><body>Sign in to continue</body></html>", id="html"),
        pytest.param('["not", "an", "object"]', id="json_that_is_not_an_object"),
    ],
)
async def test_a_reply_graph_could_not_have_sent_is_not_called_invalid_input(method, call, body):
    """A 200 Graph could not have sent is Graph's failure, not the caller's.

    `response.json()` raises `JSONDecodeError`, a `ValueError`, and the wrapper
    turns every `ValueError` into an input error, so the model read a JSON
    parser's complaint as if its own arguments were wrong. JSON that is not an
    object crashed on the first `.get()` instead.
    """
    url = _DELTA_CURSOR if method == "get" else _BATCH_URL
    page = httpx.Response(200, text=body, request=httpx.Request(method.upper(), url))
    with _httpx_answering(method, page):
        is_error, text = await _text_the_model_sees_from(call())

    assert is_error is True
    assert "not a JSON object" in text
    assert "Expecting value" not in text


_REFRESH_TOKEN_EXPIRED = ClientAuthenticationError(
    "Authentication failed: AADSTS700082: The refresh token has expired due to inactivity."
)
_INTERACTION_REQUIRED = AuthenticationRequiredError(["https://graph.microsoft.com/.default"])
_CONSENT_DEAD_END = ClientAuthenticationError(
    "Authentication failed: AADSTS70000: The requested user must first sign-in and grant "
    "the client application access."
)
# azure-identity's own wording (`_persistent_cache.py`), as its wrapper re-types it.
_NO_SAFE_CACHE = ClientAuthenticationError(
    "Authentication failed: Cache encryption is impossible because libsecret dependencies "
    "are not installed or are unusable, for example because no display is available (as in "
    "an SSH session). The chained exception has more information. Specify "
    '"allow_unencrypted_storage=True" to store the cache unencrypted instead of raising this '
    "exception."
)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("refusal", "expected"),
    [
        pytest.param(_REFRESH_TOKEN_EXPIRED, ["outlook-mcp auth"], id="refresh_token_expired"),
        pytest.param(_INTERACTION_REQUIRED, ["outlook-mcp auth"], id="interaction_required"),
        pytest.param(
            _CONSENT_DEAD_END, ["Log in again", "outlook-mcp auth"], id="consent_dead_end"
        ),
        pytest.param(_NO_SAFE_CACHE, ["cleartext", "libsecret"], id="no_safe_token_cache"),
    ],
)
@pytest.mark.parametrize("path", ["sdk", "raw"])
async def test_a_refresh_that_fails_mid_session_says_to_sign_in_again(refusal, expected, path):
    """The sign-in can die after startup: a revoked or idle-expired refresh token.

    azure-identity raises `ClientAuthenticationError` from `get_token`, on the
    SDK path (kiota's auth provider asks before sending) and the raw ones alike,
    and neither wraps it. It is not a Graph error, so the agent was told nothing
    and never learned the remedy. It now gets the one startup gives.
    """
    get_token = MagicMock(side_effect=refusal)
    if path == "sdk":
        sdk_client = GraphClient(MagicMock(get_token=get_token)).sdk_client

        def call():
            return sdk_client.me.messages.get()
    else:
        call = _delta_call(_fake_graph_client(get_token))

    is_error, text = await _text_the_model_sees_from(call)

    assert is_error is True
    for fragment in expected:
        assert fragment in text, f"{fragment!r} missing from client-visible text: {text!r}"
