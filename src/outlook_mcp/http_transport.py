"""Optional streamable-HTTP transport: ``OUTLOOK_MCP_TRANSPORT=streamable-http``.

For a server that runs on another machine or in a container, behind a reverse
proxy, or for an MCP client that cannot spawn a local process. stdio stays the
default, and ``server.main`` reads none of these variables unless HTTP was asked
for.

The app is the SDK's own ``streamable_http_app``, stateless and answering JSON:
no tool sends a request or notification back to the client mid-call (sampling,
elicitation, progress, logging), so a session would carry nothing. That is the
stateless mode of the released spec, where the server simply never issues an
``Mcp-Session-Id``; it does not depend on the 2026-07-28 release candidate.

The settings are environment variables, like ``OUTLOOK_MCP_CONFIG_DIR``, which
also keeps a token out of ``argv``. They are read and checked before anything
binds, so a refusal exits with its reason instead of after the port is open:

- one rule for every bind: no ``OUTLOOK_MCP_HTTP_TOKEN``, no server. Loopback
  is no exception, since a loopback port with no token lets any process running
  as that user read the mailbox with one HTTP request;
- every request must present the token as ``Authorization: Bearer <token>``
  or gets 401 before the SDK sees it;
- loopback (``127.0.0.1``) unless ``OUTLOOK_MCP_HTTP_HOST`` says otherwise. The
  server speaks plain HTTP, so on a bind other machines can reach the token
  crosses the network in the clear unless a proxy in front terminates TLS, and
  the refusal says so.

Nothing here logs, and uvicorn runs without an access log at WARNING, so no
token and no request URL reaches stderr or stdout.
"""

from __future__ import annotations

import hmac
import ipaddress
import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import uvicorn
    from mcp.server.mcpserver import MCPServer
    from starlette.types import ASGIApp, Receive, Scope, Send

TRANSPORT_ENV = "OUTLOOK_MCP_TRANSPORT"
HOST_ENV = "OUTLOOK_MCP_HTTP_HOST"
PORT_ENV = "OUTLOOK_MCP_HTTP_PORT"
TOKEN_ENV = "OUTLOOK_MCP_HTTP_TOKEN"

STDIO = "stdio"
STREAMABLE_HTTP = "streamable-http"

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8000

# RFC 6750's b64token alphabet, so a client can send the token as it is, and a
# floor that `secrets.token_urlsafe(32)` (43 characters) clears with room to spare.
_TOKEN_SHAPE = re.compile(r"[A-Za-z0-9\-._~+/]{32,}=*")
_TOKEN_RECIPE = 'python -c "import secrets; print(secrets.token_urlsafe(32))"'


@dataclass(frozen=True)
class HttpSettings:
    host: str = DEFAULT_HOST
    port: int = DEFAULT_PORT
    # Out of repr, so a settings object that ever reaches a log carries no token.
    token: str = field(default="", repr=False)


def settings_from_env(environ: Mapping[str, str]) -> HttpSettings | None:
    """Read the transport settings; ``None`` means stdio.

    Raises ``ValueError`` with the repair for anything unusable. The message
    never contains a token, not even a malformed one.
    """
    transport = environ.get(TRANSPORT_ENV, "").strip().lower()
    if transport in ("", STDIO):
        return None
    if transport != STREAMABLE_HTTP:
        raise ValueError(
            f"{TRANSPORT_ENV}={transport!r} is not a transport this server offers. "
            f"Use {STDIO!r} (the default) or {STREAMABLE_HTTP!r}."
        )

    host = environ.get(HOST_ENV, "").strip() or DEFAULT_HOST
    port = _port(environ.get(PORT_ENV, "").strip())
    token = _token(environ.get(TOKEN_ENV, ""))
    if not token:
        raise ValueError(_missing_token(host))
    return HttpSettings(host=host, port=port, token=token)


def _missing_token(host: str) -> str:
    """The refusal, naming what this particular bind would have exposed."""
    if _is_loopback(host):
        exposure = (
            "A token is required on every bind, loopback included: without one, any "
            "process on this machine could read the mailbox with one HTTP request."
        )
    else:
        exposure = (
            f"{HOST_ENV}={host} accepts connections from other machines: without a token, "
            f"anyone who can reach the port could read the mailbox. And this server speaks "
            f"plain HTTP, so the token crosses the network in the clear unless a reverse "
            f"proxy in front of it terminates TLS."
        )
    return (
        f"{TOKEN_ENV} must be set. {exposure} Every request then needs "
        f"'Authorization: Bearer <token>'. Generate one with: {_TOKEN_RECIPE}"
    )


def _port(value: str) -> int:
    if not value:
        return DEFAULT_PORT
    if re.fullmatch(r"[0-9]{1,5}", value) and 1 <= int(value) <= 65535:
        return int(value)
    raise ValueError(f"{PORT_ENV}={value!r} is not a TCP port. Use a number from 1 to 65535.")


def _token(value: str) -> str:
    token = value.strip()
    if token and not _TOKEN_SHAPE.fullmatch(token):
        raise ValueError(
            f"{TOKEN_ENV} is not usable. A token needs at least 32 characters from A-Z, "
            f"a-z, 0-9 and -._~+/ (what a bearer token may carry). Generate one with: "
            f"{_TOKEN_RECIPE}"
        )
    return token


def _is_loopback(host: str) -> bool:
    """Only a loopback address, or the name ``localhost``, counts.

    Any other name could resolve to anything, so it is treated as reachable.
    """
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


class BearerTokenMiddleware:
    """Pass an HTTP request on only if it presents the configured token.

    It wraps the whole app, so it stands in front of every path and of both
    protocol eras the SDK routes. A refusal is the same 401 whatever was wrong,
    and it says nothing about how close the token presented came.
    """

    def __init__(self, app: ASGIApp, token: str) -> None:
        self.app = app
        self._token = token.encode("ascii")

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] == "lifespan" or (scope["type"] == "http" and self._authorized(scope)):
            await self.app(scope, receive, send)
            return
        if scope["type"] == "http":
            await send(
                {
                    "type": "http.response.start",
                    "status": 401,
                    "headers": [
                        (b"www-authenticate", b"Bearer"),
                        (b"content-type", b"text/plain; charset=utf-8"),
                    ],
                }
            )
            await send({"type": "http.response.body", "body": b"Unauthorized"})
        # Anything else is never served: uvicorn runs without WebSocket support,
        # so nothing else should arrive, and returning unanswered refuses it.

    def _authorized(self, scope: Scope) -> bool:
        values = [value for name, value in scope["headers"] if name == b"authorization"]
        # Two credentials in one request is malformed (the header is not a
        # list), and picking either one would let the other ride along.
        if len(values) != 1:
            return False
        scheme, _, presented = values[0].partition(b" ")
        if scheme.lower() != b"bearer":
            return False
        presented = presented.strip(b" ")
        # In constant time, so the time taken does not say how much of it
        # matched. No token configured matches nothing, an empty one included.
        return bool(self._token) and hmac.compare_digest(presented, self._token)


def build_app(server: MCPServer, settings: HttpSettings) -> ASGIApp:
    """The SDK's streamable-HTTP app at ``/mcp``, always behind the token check.

    There is no unguarded variant: settings without a token, which
    ``settings_from_env`` never returns, get a gate that matches nothing.
    ``host`` goes to the SDK so that a loopback bind keeps its DNS-rebinding
    protection: a request whose ``Host`` is not a loopback name gets 421.
    """
    app = server.streamable_http_app(
        stateless_http=True,
        json_response=True,
        host=settings.host,
    )
    return BearerTokenMiddleware(app, settings.token)


def uvicorn_config(server: MCPServer, settings: HttpSettings) -> uvicorn.Config:
    import uvicorn

    return uvicorn.Config(
        build_app(server, settings),
        host=settings.host,
        port=settings.port,
        # The server's own level (WARNING), as the SDK's own runner passes it.
        # uvicorn's access log is INFO and carries each request line, query
        # string included; its handler is removed outright as well.
        log_level=server.settings.log_level.lower(),
        access_log=False,
        # There is no WebSocket endpoint, so no WebSocket implementation is
        # loaded even where one is installed: an upgrade request stays plain
        # HTTP and meets the token check like any other.
        ws="none",
    )


def serve(server: MCPServer, settings: HttpSettings) -> None:
    import uvicorn

    uvicorn.Server(uvicorn_config(server, settings)).run()
