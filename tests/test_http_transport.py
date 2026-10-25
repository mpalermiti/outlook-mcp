"""`OUTLOOK_MCP_TRANSPORT=streamable-http`: opt-in, and closed by default (#126).

stdio is how every client launches this server today, so the first thing pinned
here is that its path did not move: `main` still makes exactly the call it made
before, and the HTTP variables are not even read unless HTTP was asked for.

The rest pins the security model, which has one rule for every bind: no bearer
token, no server. Loopback is no exception, because a loopback port with no token
lets any process running as that user read the mailbox with one HTTP request.
Every request carries the token or gets 401 before it reaches the MCP layer; a
bind other machines can reach is told, as plainly as the refusal itself, that
plain HTTP sends the token in the clear; nothing logs the token or a request URL.
The round-trip tests run the production uvicorn config on an ephemeral port, so
they see what an HTTP client sees, not what a mock was told.
"""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import threading
import time

import httpx2
import pytest
import uvicorn
from mcp.client import Client
from mcp.client.streamable_http import streamable_http_client

from outlook_mcp import http_transport, server
from outlook_mcp.config import Config
from outlook_mcp.http_transport import (
    ALLOWED_HOSTS_ENV,
    HOST_ENV,
    PORT_ENV,
    TOKEN_ENV,
    TOKEN_FILE_ENV,
    TRANSPORT_ENV,
    HttpSettings,
    settings_from_env,
)

# Synthetic, and shaped like real ones: 32+ characters of the RFC 6750 alphabet.
TOKEN = "test-token-" + "a" * 32
WRONG = "wrong-token-" + "c" * 32

HTTP = {TRANSPORT_ENV: "streamable-http"}

INITIALIZE = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2025-11-25",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "0"},
    },
}


# ── stdio stays the default ─────────────────────────────────────────────


@pytest.fixture
def inherited_http_env(monkeypatch):
    # Reproduce a developer shell configured for a different HTTP instance.
    # The real entrypoint tests below must work regardless of that parent.
    monkeypatch.setenv(TOKEN_FILE_ENV, "/unused/inherited-token")
    monkeypatch.setenv(ALLOWED_HOSTS_ENV, "https://inherited.invalid")


@pytest.fixture
def no_transport_env(monkeypatch, inherited_http_env):
    for name in (TRANSPORT_ENV, HOST_ENV, PORT_ENV, TOKEN_ENV, TOKEN_FILE_ENV, ALLOWED_HOSTS_ENV):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setattr(server, "load_config", lambda: Config())


@pytest.mark.parametrize(
    "env",
    [
        {},
        {TRANSPORT_ENV: ""},
        {TRANSPORT_ENV: "stdio"},
        # The HTTP settings are not read at all on stdio, so values that would
        # be refused on HTTP change nothing here.
        {HOST_ENV: "0.0.0.0", PORT_ENV: "not-a-port", TOKEN_ENV: "short"},
    ],
    ids=["unset", "empty", "stdio", "http-vars-without-transport"],
)
def test_stdio_stays_the_default(env, no_transport_env, monkeypatch):
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    runs = []
    monkeypatch.setattr(server.mcp, "run", lambda *args, **kwargs: runs.append((args, kwargs)))
    monkeypatch.setattr(http_transport, "serve", lambda *_: pytest.fail("HTTP must not start"))

    server.main()

    # The exact call main made before HTTP existed.
    assert runs == [((), {"transport": "stdio"})]


def test_http_is_served_only_when_asked_for(no_transport_env, monkeypatch):
    monkeypatch.setenv(TRANSPORT_ENV, "streamable-http")
    monkeypatch.setenv(TOKEN_ENV, TOKEN)
    served = []
    monkeypatch.setattr(server.mcp, "run", lambda *_, **__: pytest.fail("stdio must not start"))
    monkeypatch.setattr(http_transport, "serve", lambda mcp, settings: served.append(settings))

    server.main()

    assert served == [HttpSettings(token=TOKEN)]


# ── Refused before anything binds ───────────────────────────────────────


@pytest.mark.parametrize(
    "env, named",
    [
        ({TRANSPORT_ENV: "http"}, TRANSPORT_ENV),
        (HTTP, TOKEN_ENV),
        ({**HTTP, HOST_ENV: "0.0.0.0"}, TOKEN_ENV),
        ({**HTTP, PORT_ENV: "eighty", TOKEN_ENV: TOKEN}, PORT_ENV),
        ({**HTTP, TOKEN_ENV: "short"}, TOKEN_ENV),
    ],
    ids=[
        "unknown-transport",
        "loopback-without-token",
        "public-bind-without-token",
        "bad-port",
        "bad-token",
    ],
)
def test_main_exits_with_the_reason_and_starts_nothing(
    env, named, no_transport_env, monkeypatch, capsys
):
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    monkeypatch.setattr(server.mcp, "run", lambda *_, **__: pytest.fail("stdio must not start"))
    monkeypatch.setattr(http_transport, "serve", lambda *_: pytest.fail("HTTP must not start"))

    with pytest.raises(SystemExit) as exc:
        server.main()

    assert exc.value.code == 1
    out, err = capsys.readouterr()
    assert out == ""
    assert named in err


LOOPBACK = ["127.0.0.1", "127.0.0.2", "localhost", "::1"]
PUBLIC = ["0.0.0.0", "::", "192.0.2.10", "2001:db8::1", "mcp.example.com"]


def test_defaults_are_loopback_and_port_8000():
    settings = settings_from_env({**HTTP, TOKEN_ENV: TOKEN})

    assert settings == HttpSettings(host="127.0.0.1", port=8000, token=TOKEN)


@pytest.mark.parametrize("host", [None, *LOOPBACK, *PUBLIC])
def test_every_bind_needs_a_token_loopback_included(host):
    env = {**HTTP} if host is None else {**HTTP, HOST_ENV: host}
    with pytest.raises(ValueError, match=TOKEN_ENV):
        settings_from_env(env)

    # ...and is accepted once one is set.
    assert settings_from_env({**env, TOKEN_ENV: TOKEN}).token == TOKEN


@pytest.mark.parametrize("host", LOOPBACK)
def test_the_loopback_refusal_names_the_local_risk(host):
    with pytest.raises(ValueError) as exc:
        settings_from_env({**HTTP, HOST_ENV: host})

    assert "any process on this machine" in str(exc.value)
    # Nothing crosses a network from a loopback bind; saying so would be noise.
    assert "in the clear" not in str(exc.value)


@pytest.mark.parametrize("host", PUBLIC)
def test_a_public_bind_refusal_says_plain_http_sends_the_token_in_the_clear(host):
    with pytest.raises(ValueError) as exc:
        settings_from_env({**HTTP, HOST_ENV: host})

    message = str(exc.value)
    assert f"{HOST_ENV}={host} accepts connections from other machines" in message
    assert "in the clear" in message
    assert "terminates TLS" in message


def test_surrounding_whitespace_is_not_part_of_the_token():
    assert settings_from_env({**HTTP, TOKEN_ENV: f" {TOKEN}\n"}).token == TOKEN


@pytest.mark.parametrize(
    "bad",
    [
        "x" * 31,
        "has space " + "x" * 32,
        'quote"' + "x" * 32,
        "ü" + "x" * 32,
        f"{TOKEN},{WRONG}",
    ],
    ids=["too-short", "space", "quote", "non-ascii", "a-list"],
)
def test_an_unusable_token_is_refused_without_being_echoed(bad):
    with pytest.raises(ValueError, match=TOKEN_ENV) as exc:
        settings_from_env({**HTTP, TOKEN_ENV: bad})

    assert bad not in str(exc.value)
    assert TOKEN not in str(exc.value)


def test_the_settings_never_print_their_tokens():
    assert TOKEN not in repr(settings_from_env({**HTTP, TOKEN_ENV: TOKEN}))


@pytest.mark.parametrize("port", ["eighty", "0", "65536", "-1", "80.5"])
def test_an_invalid_port_is_refused(port):
    with pytest.raises(ValueError, match=PORT_ENV):
        settings_from_env({**HTTP, PORT_ENV: port, TOKEN_ENV: TOKEN})


@pytest.mark.asyncio
async def test_an_app_built_without_a_token_refuses_every_request():
    """``settings_from_env`` never hands back settings without a token. Settings
    built by hand without one still get the gate, which then matches nothing:
    there is no unguarded app to fall back to."""
    sent = []
    app = http_transport.build_app(server.mcp, HttpSettings())

    async def send(message):
        sent.append(message)

    for presented in (f"Bearer {TOKEN}", "Bearer "):
        headers = [(b"authorization", presented.encode())]
        await app({"type": "http", "headers": headers}, None, send)

    assert isinstance(app, http_transport.BearerTokenMiddleware)
    assert [m["status"] for m in sent if "status" in m] == [401, 401]


# ── A real HTTP round trip ──────────────────────────────────────────────


@pytest.fixture(scope="module")
def base_url():
    """The production uvicorn config, on an ephemeral loopback port."""
    settings = HttpSettings(host="127.0.0.1", port=0, token=TOKEN)
    instance = uvicorn.Server(http_transport.uvicorn_config(server.mcp, settings))
    thread = threading.Thread(target=instance.run, daemon=True)
    thread.start()
    deadline = time.monotonic() + 30
    while not instance.started:
        if not thread.is_alive() or time.monotonic() > deadline:
            pytest.fail("the HTTP server did not start")
        time.sleep(0.05)
    port = instance.servers[0].sockets[0].getsockname()[1]
    yield f"http://127.0.0.1:{port}"
    instance.should_exit = True
    thread.join(timeout=10)


def _post(url: str, authorization: str | None = None, **headers: str) -> httpx2.Response:
    headers = {"Accept": "application/json, text/event-stream", **headers}
    if authorization is not None:
        headers["Authorization"] = authorization
    return httpx2.post(url, json=INITIALIZE, headers=headers, timeout=30)


@pytest.mark.parametrize(
    "authorization",
    [
        None,
        f"Bearer {WRONG}",
        f"Bearer {TOKEN}x",
        f"Bearer {TOKEN[:-1]}",
        f"Bearer {TOKEN},{WRONG}",
        "Bearer",
        f"Basic {TOKEN}",
        TOKEN,
    ],
    ids=["missing", "wrong", "longer", "shorter", "a-list", "empty", "basic", "no-scheme"],
)
def test_a_request_without_a_configured_bearer_gets_401(base_url, authorization):
    response = _post(f"{base_url}/mcp", authorization)

    assert response.status_code == 401
    assert response.headers["www-authenticate"] == "Bearer"


def test_the_gate_stands_in_front_of_every_path_and_protocol_era(base_url):
    # Paths the SDK does not route, including where an OAuth client looks first.
    for path in ("/", "/.well-known/oauth-protected-resource", "/mcp/"):
        assert httpx2.get(f"{base_url}{path}", timeout=30).status_code == 401
    # The SDK answers a 2026-07-28 request on its own path, whatever the flags.
    modern = _post(f"{base_url}/mcp", **{"MCP-Protocol-Version": "2026-07-28"})
    assert modern.status_code == 401
    # With the token, only /mcp is there.
    ok = {"Authorization": f"Bearer {TOKEN}"}
    assert httpx2.get(f"{base_url}/", headers=ok, timeout=30).status_code == 404


@pytest.mark.parametrize(
    "authorization",
    [f"bearer {TOKEN}", f"BEARER {TOKEN}", f"Bearer  {TOKEN}"],
    ids=["lower", "upper", "two-spaces"],
)
def test_the_scheme_is_case_insensitive_and_spacing_is_tolerated(base_url, authorization):
    # RFC 6750: the scheme is case-insensitive and is followed by one or more spaces.
    assert _post(f"{base_url}/mcp", authorization).status_code == 200


@pytest.mark.parametrize("first, second", [(TOKEN, WRONG), (WRONG, TOKEN)])
def test_two_authorization_headers_are_refused(base_url, first, second):
    headers = [
        ("Accept", "application/json, text/event-stream"),
        ("Authorization", f"Bearer {first}"),
        ("Authorization", f"Bearer {second}"),
    ]
    response = httpx2.post(f"{base_url}/mcp", json=INITIALIZE, headers=headers, timeout=30)

    assert response.status_code == 401


def test_a_loopback_bind_keeps_the_sdks_dns_rebinding_guard(base_url):
    # A page that rebinds its own name to 127.0.0.1 still sends its own Host.
    response = _post(f"{base_url}/mcp", f"Bearer {TOKEN}", Host="mcp.example.com")

    assert response.status_code == 421


@pytest.mark.asyncio
async def test_a_bind_other_machines_can_reach_accepts_the_proxys_host():
    """Behind a reverse proxy, ``Host`` is the public name.

    The SDK applies its loopback-only Host check to whatever ``host`` it is
    given, defaulting to 127.0.0.1, so the bind address has to reach it or every
    proxied request to a 0.0.0.0 bind gets 421. Driven in-process, so the test
    never listens on a public interface.
    """
    app = http_transport.build_app(
        server.mcp, HttpSettings(host="0.0.0.0", token=TOKEN, allowed_hosts=("mcp.example.com",))
    )
    sdk_app = app.app  # the lifespan belongs to the SDK's Starlette app inside the gate
    headers = {"Accept": "application/json, text/event-stream", "Authorization": f"Bearer {TOKEN}"}
    transport = httpx2.ASGITransport(app=app)
    async with sdk_app.router.lifespan_context(sdk_app):
        async with httpx2.AsyncClient(transport=transport, base_url="https://mcp.example.com") as c:
            response = await c.post("/mcp", json=INITIALIZE, headers=headers)

    assert response.status_code == 200


@pytest.mark.asyncio
async def test_only_http_requests_are_ever_passed_on():
    reached, sent = [], []

    async def app(scope, receive, send):
        reached.append(scope["type"])

    async def send(message):
        sent.append(message)

    gate = http_transport.BearerTokenMiddleware(app, TOKEN)
    headers = [(b"authorization", f"Bearer {TOKEN}".encode())]
    await gate({"type": "websocket", "headers": headers}, None, send)
    await gate({"type": "http", "headers": headers}, None, send)

    assert reached == ["http"]
    assert sent == []


def test_the_server_is_stateless_and_answers_json(base_url):
    response = _post(f"{base_url}/mcp", f"Bearer {TOKEN}")

    assert response.status_code == 200
    assert "mcp-session-id" not in response.headers
    assert response.headers["content-type"].startswith("application/json")
    assert response.json()["result"]["serverInfo"]["name"] == "outlook-mcp"


@pytest.mark.asyncio
async def test_the_right_bearer_lists_the_same_surface_as_stdio(base_url):
    async with httpx2.AsyncClient(headers={"Authorization": f"Bearer {TOKEN}"}) as http:
        transport = streamable_http_client(f"{base_url}/mcp", http_client=http)
        # "legacy" is the initialize handshake today's clients use.
        async with Client(transport, mode="legacy") as client:
            over_http = (await client.list_tools()).tools
    async with Client(server.mcp) as client:
        in_process = (await client.list_tools()).tools

    def dump(tools):
        return [t.model_dump(exclude_none=True, by_alias=True) for t in tools]

    # The surface test_tool_surface_budget.py measures, unchanged by the transport.
    assert over_http
    assert dump(over_http) == dump(in_process)


# ── The real entrypoint, in its own process ─────────────────────────────


def _free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return probe.getsockname()[1]


def test_the_entrypoint_serves_and_logs_no_token_or_request_url(tmp_path):
    """A fresh process, because this is about what a real start configures.

    pytest's log capture leaves the root logger with handlers before the
    server is imported, so in-process logging is not what an operator gets
    (see test_request_urls_are_not_logged). Everything the process writes to
    stdout or stderr is the log here.
    """
    port = _free_port()
    env = dict(os.environ)
    env.update(
        {
            "OUTLOOK_MCP_CONFIG_DIR": str(tmp_path),
            TRANSPORT_ENV: "streamable-http",
            PORT_ENV: str(port),
            TOKEN_ENV: TOKEN,
        }
    )
    env.pop(HOST_ENV, None)
    env.pop(TOKEN_FILE_ENV, None)
    env.pop(ALLOWED_HOSTS_ENV, None)
    proc = subprocess.Popen(
        [sys.executable, "-c", "from outlook_mcp.server import main; main()"],
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    url = f"http://127.0.0.1:{port}/mcp?q=find-me"
    try:
        deadline = time.monotonic() + 60
        while True:
            if proc.poll() is not None:
                pytest.fail(f"the server exited: {proc.communicate()[1]}")
            try:
                refused = _post(url, f"Bearer {WRONG}")
                break
            except httpx2.TransportError:
                if time.monotonic() > deadline:
                    pytest.fail("the server did not start listening")
                time.sleep(0.2)
        served = _post(url, f"Bearer {TOKEN}")
    finally:
        proc.terminate()
        out, err = proc.communicate(timeout=30)

    assert refused.status_code == 401
    assert served.status_code == 200
    logged = out + err
    for secret in (TOKEN, WRONG, "find-me", "/mcp", f"127.0.0.1:{port}"):
        assert secret not in logged


@pytest.mark.parametrize(
    "host",
    ["0.0.0.0", "::", "192.0.2.5", "mcp.example.com", "127.0.0.1", "127.0.0.2", "::1", "localhost"],
)
def test_public_bind_warning_is_emitted_without_the_token(monkeypatch, capsys, host):
    from unittest.mock import MagicMock

    monkeypatch.setattr(uvicorn, "Server", MagicMock())
    http_transport.serve(server.mcp, HttpSettings(host=host, token=TOKEN))
    output = capsys.readouterr().err
    assert ("unencrypted" in output) is (not http_transport._is_loopback(host))
    if not http_transport._is_loopback(host):
        assert "every tool result" in output and "TLS-terminating" in output
    assert TOKEN not in output


@pytest.mark.asyncio
@pytest.mark.parametrize("bind", ["127.0.0.2", "0.0.0.0", "::"])
@pytest.mark.parametrize("bad_header", ["Host", "Origin"])
async def test_explicit_transport_security_covers_every_bind(bind, bad_header):
    settings = HttpSettings(host=bind, token=TOKEN, allowed_hosts=("mcp.example.com",))
    app = http_transport.build_app(server.mcp, settings)
    headers = {
        "Accept": "application/json, text/event-stream",
        "Authorization": f"Bearer {TOKEN}",
        bad_header: "https://attacker.example" if bad_header == "Origin" else "attacker.example",
    }
    async with app.app.router.lifespan_context(app.app):
        async with httpx2.AsyncClient(
            transport=httpx2.ASGITransport(app=app), base_url="https://mcp.example.com"
        ) as client:
            response = await client.post("/mcp", json=INITIALIZE, headers=headers)
    assert response.status_code == (421 if bad_header == "Host" else 403)


@pytest.mark.parametrize("method", ["DELETE", "OPTIONS", "HEAD"])
def test_every_method_without_bearer_is_refused(base_url, method):
    assert httpx2.request(method, base_url + "/mcp", timeout=30).status_code == 401


def test_query_string_cannot_supply_the_bearer(base_url):
    response = httpx2.get(base_url + "/mcp?access_token=" + TOKEN, timeout=30)
    assert response.status_code == 401


def test_uvicorn_does_not_enable_access_logging_or_forwarded_headers(monkeypatch):
    monkeypatch.setenv("FORWARDED_ALLOW_IPS", "*")
    config = http_transport.uvicorn_config(server.mcp, HttpSettings(token=TOKEN))
    assert config.access_log is False
    assert config.proxy_headers is False
    assert config.server_header is False


def test_owner_only_token_file_is_accepted_and_never_represented(tmp_path):
    path = tmp_path / "token"
    path.write_text(TOKEN + "\n")
    path.chmod(0o600)
    settings = settings_from_env({**HTTP, http_transport.TOKEN_FILE_ENV: str(path)})
    assert settings.token == TOKEN
    assert TOKEN not in repr(settings)


@pytest.mark.parametrize(
    "case", ["both", "missing", "empty", "large", "bad", "directory", "permissions", "symlink"]
)
def test_unusable_token_file_refuses_without_echoing_path_or_content(tmp_path, case):
    path = tmp_path / "private-file-name"
    path.write_text(TOKEN)
    path.chmod(0o600)
    env = {**HTTP, http_transport.TOKEN_FILE_ENV: str(path)}
    if case == "both":
        env[TOKEN_ENV] = TOKEN
    elif case == "missing":
        path.unlink()
    elif case == "empty":
        path.write_text("")
    elif case == "large":
        path.write_text("X" * 4097)
    elif case == "bad":
        path.write_bytes(b"\xff")
    elif case == "directory":
        path.unlink()
        path.mkdir()
    elif case == "permissions":
        if os.name == "nt":
            pytest.skip("POSIX mode check; Windows requires operator-managed ACLs")
        path.chmod(0o640)
    elif case == "symlink":
        target = tmp_path / "target"
        path.rename(target)
        try:
            path.symlink_to(target)
        except OSError:
            pytest.skip("symlink privilege unavailable")
    with pytest.raises(ValueError, match="OUTLOOK_MCP_HTTP_TOKEN") as exc:
        settings_from_env(env)
    assert str(path) not in str(exc.value) and TOKEN not in str(exc.value)


@pytest.mark.parametrize("path", ["/mcp", "/mcp/"])
def test_authenticated_get_cannot_hold_an_empty_event_stream(base_url, path):
    response = httpx2.get(
        base_url + path,
        headers={"Authorization": f"Bearer {TOKEN}", "Accept": "text/event-stream"},
        timeout=2,
    )
    assert response.status_code == 405
    assert response.headers["allow"] == "POST"
    assert (
        httpx2.get(base_url + path, headers={"Accept": "text/event-stream"}, timeout=2).status_code
        == 401
    )


@pytest.mark.parametrize("bad", ["a" * 32, "abc123" * 10, "abcdefg" * 8 + "==="])
def test_obviously_low_variety_tokens_are_refused_without_echo(bad):
    with pytest.raises(ValueError) as exc:
        settings_from_env({**HTTP, TOKEN_ENV: bad})
    assert bad not in str(exc.value)


@pytest.mark.parametrize("host", ["0.0.0.0", "::"])
@pytest.mark.parametrize("source", ["env", "file"])
def test_wildcard_without_allowed_hosts_warns_before_server_creation(
    host, source, tmp_path, monkeypatch, capsys
):
    env = {**HTTP, HOST_ENV: host}
    path = tmp_path / "private-token-name"
    if source == "file":
        path.write_text(TOKEN)
        path.chmod(0o600)
        env[TOKEN_FILE_ENV] = str(path)
    else:
        env[TOKEN_ENV] = TOKEN
    warnings = []

    def create_server(config):
        # An operator needs the repair before the listener is started.
        warning = capsys.readouterr().err
        assert ALLOWED_HOSTS_ENV in warning and "421" in warning
        assert "Host" in warning and "TLS-terminating" in warning
        assert TOKEN not in warning and str(path) not in warning
        warnings.append(warning)
        from unittest.mock import MagicMock

        return MagicMock()

    monkeypatch.setattr(uvicorn, "Server", create_server)
    http_transport.serve(server.mcp, settings_from_env(env))
    assert len(warnings) == 1


def test_a_configured_wildcard_warns_only_about_plain_http(monkeypatch, capsys):
    from unittest.mock import MagicMock

    monkeypatch.setattr(uvicorn, "Server", MagicMock())
    settings = settings_from_env(
        {**HTTP, HOST_ENV: "0.0.0.0", TOKEN_ENV: TOKEN, ALLOWED_HOSTS_ENV: "mcp.example.com"}
    )
    http_transport.serve(server.mcp, settings)
    warning = capsys.readouterr().err
    assert "TLS-terminating" in warning
    assert "421" not in warning and ALLOWED_HOSTS_ENV not in warning


@pytest.mark.parametrize(
    "contents",
    [b"a" * 32, b"\xef\xbb\xbf" + TOKEN.encode(), b"KEY=" + TOKEN.encode(), b"", b"\xff"],
    ids=["low-variety", "byte-order-mark", "assignment", "empty", "non-ascii"],
)
def test_token_file_contents_get_their_own_repair_without_echo(tmp_path, contents):
    path = tmp_path / "private-token-name"
    path.write_bytes(contents)
    path.chmod(0o600)
    with pytest.raises(ValueError) as exc:
        settings_from_env({**HTTP, TOKEN_FILE_ENV: str(path)})
    message = str(exc.value)
    assert TOKEN_FILE_ENV in message and "contents" in message
    assert "32 characters" in message and "8 distinct characters" in message
    assert "permissions" not in message and "owned" not in message
    assert str(path) not in message and TOKEN not in message
    if contents:
        assert contents.decode("utf-8", errors="replace") not in message


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "host", ["localhost", "localhost:8443", "127.0.0.1", "127.0.0.1:8000", "[::1]", "[::1]:8000"]
)
async def test_nonstandard_loopback_bind_accepts_other_loopback_hosts_and_origins(host):
    app = http_transport.build_app(server.mcp, HttpSettings(host="127.0.0.2", token=TOKEN))
    headers = {
        "Accept": "application/json, text/event-stream",
        "Authorization": f"Bearer {TOKEN}",
        "Host": host,
        "Origin": "https://" + host,
    }
    async with app.app.router.lifespan_context(app.app):
        async with httpx2.AsyncClient(
            transport=httpx2.ASGITransport(app=app), base_url="http://127.0.0.2"
        ) as client:
            response = await client.post("/mcp", json=INITIALIZE, headers=headers)
    assert response.status_code == 200


def test_allowed_hosts_are_trimmed_lowercased_and_empty_entries_ignored():
    settings = settings_from_env(
        {
            **HTTP,
            TOKEN_ENV: TOKEN,
            ALLOWED_HOSTS_ENV: " MCP.Example.COM , , [2001:DB8::1]:443, 192.0.2.10:* , ",
        }
    )
    assert settings.allowed_hosts == ("mcp.example.com", "[2001:db8::1]:443", "192.0.2.10:*")


@pytest.mark.parametrize(
    "value",
    [
        "https://mcp.example.com",
        "mcp.example.com/path",
        "mcp.example.com:0",
        "mcp.example.com:65536",
        "[::1]:65536",
        "*.example.com",
        "::1",
        "mcp.example.com:*:*",
        "mcp.example.com\r\nattacker.example",
    ],
)
def test_malformed_allowed_hosts_refuse_startup_without_echo(value):
    with pytest.raises(ValueError, match=ALLOWED_HOSTS_ENV) as exc:
        settings_from_env({**HTTP, TOKEN_ENV: TOKEN, ALLOWED_HOSTS_ENV: value})
    assert value not in str(exc.value)


@pytest.mark.asyncio
@pytest.mark.parametrize("host", ["mcp.example.com", "192.0.2.10", "[2001:db8::1]"])
@pytest.mark.parametrize("port", ["", ":8443"])
async def test_wildcard_port_allows_default_https_host_and_explicit_ports(host, port):
    settings = settings_from_env(
        {
            **HTTP,
            HOST_ENV: "0.0.0.0",
            TOKEN_ENV: TOKEN,
            ALLOWED_HOSTS_ENV: host + ":*",
        }
    )
    app = http_transport.build_app(server.mcp, settings)
    headers = {
        "Accept": "application/json, text/event-stream",
        "Authorization": f"Bearer {TOKEN}",
        "Host": host + port,
        "Origin": "https://" + host + port,
    }
    async with app.app.router.lifespan_context(app.app):
        async with httpx2.AsyncClient(
            transport=httpx2.ASGITransport(app=app), base_url="https://mcp.example.com"
        ) as client:
            response = await client.post("/mcp", json=INITIALIZE, headers=headers)
            # Widening the port must not also widen the hostname or origin.
            wrong_host = await client.post(
                "/mcp", json=INITIALIZE, headers={**headers, "Host": "attacker.example"}
            )
            wrong_origin = await client.post(
                "/mcp", json=INITIALIZE, headers={**headers, "Origin": "https://attacker.example"}
            )
    assert response.status_code == 200
    assert wrong_host.status_code == 421
    assert wrong_origin.status_code == 403


@pytest.mark.skipif(not hasattr(os, "mkfifo"), reason="POSIX FIFO check")
def test_token_file_fifo_refuses_instead_of_blocking_startup(tmp_path):
    path = tmp_path / "private-fifo-name"
    os.mkfifo(path, mode=0o600)
    script = """
import sys
from outlook_mcp.http_transport import settings_from_env, TRANSPORT_ENV, TOKEN_FILE_ENV
try:
    settings_from_env({TRANSPORT_ENV: "streamable-http", TOKEN_FILE_ENV: sys.argv[1]})
except ValueError as error:
    print(error)
else:
    raise SystemExit("FIFO was accepted")
"""
    try:
        result = subprocess.run(
            [sys.executable, "-c", script, str(path)], capture_output=True, text=True, timeout=5
        )
    except subprocess.TimeoutExpired:
        pytest.fail("token-file startup blocked opening a FIFO instead of refusing it")
    assert result.returncode == 0
    assert TOKEN_FILE_ENV in result.stdout
    assert str(path) not in result.stdout + result.stderr


@pytest.mark.parametrize("path", ["/mcp", "/mcp/"])
def test_authenticated_delete_cannot_terminate_a_stateless_session(base_url, path):
    response = httpx2.delete(
        base_url + path,
        headers={"Authorization": f"Bearer {TOKEN}"},
        follow_redirects=True,
        timeout=2,
    )
    assert response.status_code == 405
