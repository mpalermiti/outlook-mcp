"""Shared test fixtures for outlook-mcp.

Two tiers here need real credentials and are skipped without them:

``-m integration``
    End-to-end plumbing smoke tests — does each tool return its documented
    response shape against real Graph.

``-m live``
    Query-shape regression guards. These exist because the default suite mocks
    the Graph client and therefore *cannot* observe a query that Graph rejects
    (400) or silently mis-evaluates (200 with wrong results). Every bug fixed
    in #30, #31 and the ``list_thread`` repair was invisible to 558 green mock
    tests. See ``tests/test_live_query_shape.py``.

``-m live_write``
    The only tier that mutates a real mailbox, added for #41. Some
    payloads can only be validated by Graph accepting them — a recurrence that
    is well-formed to the SDK still 400s if the range disagrees with the start
    — and a mock cannot see that. Kept as narrow as possible.

Tiers 1 and 2 are READ-ONLY: never add a fixture or test there that writes to
the mailbox — no sends, drafts, category or folder mutations, no deletes.

Rules for ``live_write``, which exists precisely because it breaks that:

* Double-gated — the marker is deselected by default *and* the tier skips
  unless ``OUTLOOK_MCP_LIVE_WRITE=1``. A cached token alone must never be
  enough to write to someone's calendar.
* Calendar, contacts, To Do and unsent drafts only — the surfaces whose writes
  have no outward side effect. Nothing here may send, and a stray send is not
  recoverable: a draft is created, read back and deleted, never sent
  (``tests/test_live_write_never_sends.py`` checks the tier's files for a send).
* No attendees, ever. A recurring invite emails real people on every
  occurrence.
* Bounded ranges only (``numbered``/``endDate``) — never ``noEnd``.
* Every created item is removed in a ``finally``, and subjects/names carry
  ``LIVE_WRITE_SUBJECT`` / ``LIVE_WRITE_NAME`` so anything a crash leaks is
  greppable in the UI.

Every other test is offline, and is held to it (#267) — see "Offline
isolation" below: it runs against an empty settings directory, and it fails if
it builds a real credential, opens the token cache, or connects off this
machine.
"""

import contextlib
import functools
import ipaddress
import os
import shutil
import socket
import tempfile
from pathlib import Path

import pytest

# Prefix for every item this tier creates, so an orphan left by a hard crash is
# obvious in the calendar UI and easy to search for.
LIVE_WRITE_SUBJECT = "[outlook-mcp live-write guard] safe to delete"
# Contacts and tasks have no subject; their names carry the same marker.
LIVE_WRITE_NAME = "LiveWriteGuard-SafeToDelete"


@pytest.fixture
def mock_graph_client():
    """Mock Microsoft Graph client for unit tests."""
    pass


# ── Live-credential fixtures (shared by the integration and live tiers) ──


@pytest.fixture(scope="session")
def real_config():
    """Load real config from ~/.outlook-mcp/config.json."""
    from outlook_mcp.config import load_config

    config = load_config()
    if not config.client_id:
        pytest.skip("No client_id configured — run the Azure AD app setup first")
    return config


@pytest.fixture(scope="session")
def real_auth(real_config):
    """AuthManager backed by the cached token, or skip.

    Uses ``try_cached_token`` — the same silent path the MCP server uses on
    startup. It must never trigger an interactive device-code prompt, since
    these run unattended.
    """
    from outlook_mcp.auth import AuthManager

    auth = AuthManager(real_config)
    if not auth.try_cached_token():
        pytest.skip("Not authenticated — run `outlook-mcp auth` on this host first")
    return auth


@pytest.fixture
def real_graph_client(real_auth):
    """Real Graph client built from the cached credential.

    Function-scoped on purpose: the underlying kiota/httpx transport binds to
    the running event loop, and pytest-asyncio gives each test a fresh one. A
    session-scoped client raises ``RuntimeError: Event loop is closed`` on the
    second test that uses it. ``real_config``/``real_auth`` stay session-scoped
    — they hold no loop-bound state, so the token is fetched once.
    """
    from outlook_mcp.graph import GraphClient

    return GraphClient(real_auth.get_credential())


# Every personal Microsoft account lives in this one well-known tenant; a work
# or school account carries its organisation's tenant instead. It is a public
# constant published by Microsoft, not a secret.
MSA_CONSUMER_TENANT = "9188040d-6c67-4c5b-b112-36a304b66dad"


def consumer_skip_reason(record) -> str | None:
    """Why this mailbox cannot answer a consumer-only assertion — None if it can.

    Split out from the fixture so it can be tested without a real token.
    """
    if record is None:
        return (
            "No auth record on disk, so the account type is unknown — this assertion "
            "is only meaningful on a personal Microsoft account"
        )
    if record.tenant_id != MSA_CONSUMER_TENANT:
        return (
            f"Work/school account (tenant {record.tenant_id}). This assertion is a claim "
            f"about personal Microsoft accounts (tenant {MSA_CONSUMER_TENANT}); Graph may "
            f"legitimately behave differently here, so a failure would say nothing about "
            f"the behaviour under test"
        )
    return None


@pytest.fixture(scope="session")
def consumer_mailbox_only(real_auth):
    """Skip unless the cached token belongs to a personal Microsoft account.

    Some live assertions are claims about *consumer* mailbox behaviour — Graph
    accepting a field and silently dropping it, say. On a work or school account
    Graph may honour the very thing the test says is unsupported, so the test
    fails correctly and tells us nothing about the claim.

    That is not hypothetical. The old ``is_online`` live pin (since replaced by
    ``TestOnlineMeetingsArePerAccount``, #70) reached two separate contributor
    PRs as a red test each author had to explain was not theirs, while passing
    on the maintainer's mailbox throughout. The tier already refuses to assume
    whose mailbox it is running against; this extends that to what *kind* of
    mailbox it is.

    The skip names the tenant it found, because a skip that does not say why it
    skipped is indistinguishable from coverage.
    """
    from outlook_mcp.auth import _load_auth_record

    reason = consumer_skip_reason(_load_auth_record())
    if reason:
        pytest.skip(reason)


@pytest.fixture
def live_write_config(real_config, tmp_path):
    """Real config for the write tier, or skip.

    Gates on an explicit environment opt-in beyond the marker: these tests
    create and delete real calendar events on whatever account the cached token
    belongs to.

    The config handed back is a sandboxed copy: `attachments_dir` is pointed
    at a per-test tmp_path so the attachment round-trip never stages files in
    the user's real settings directory (which a shell-exported
    `OUTLOOK_MCP_CONFIG_DIR` may have moved somewhere with other uses).
    """
    if os.environ.get("OUTLOOK_MCP_LIVE_WRITE") != "1":
        pytest.skip(
            "Write tier is opt-in — set OUTLOOK_MCP_LIVE_WRITE=1 to let it create and "
            "delete events, contacts, tasks and unsent drafts on the authenticated account"
        )
    if real_config.read_only:
        pytest.skip("Config is read_only — the write tier cannot run")

    # The real config may whitelist only some write categories (`allow_categories`
    # is a policy for the *server*, e.g. "my agent may not touch contacts"). The
    # tier's job is to exercise Graph, and the permission gate is unit-tested on
    # its own, so open the four surfaces this tier is allowed to write on an
    # in-process copy. ~/.outlook-mcp/config.json is never modified. Never
    # `mail_send`: nothing in this tier may send.
    update = {"attachments_dir": str(tmp_path / "attachments")}
    if real_config.allow_categories:
        from outlook_mcp.permissions import (
            CATEGORY_CALENDAR_WRITE,
            CATEGORY_CONTACTS_WRITE,
            CATEGORY_MAIL_DRAFTS,
            CATEGORY_TODO_WRITE,
        )

        needed = {
            CATEGORY_CALENDAR_WRITE,
            CATEGORY_CONTACTS_WRITE,
            CATEGORY_MAIL_DRAFTS,
            CATEGORY_TODO_WRITE,
        }
        update["allow_categories"] = sorted(set(real_config.allow_categories) | needed)
    return real_config.model_copy(update=update)


# ── Offline isolation (#267) ─────────────────────────────────────────────
#
# The offline suite used to reach the operator's real sign-in. Every test that
# opened the server — in process (`Client(mcp)`, the HTTP app) or in a child
# Python — ran the real lifespan, which read ~/.outlook-mcp, found the saved
# AuthenticationRecord and refreshed silently: the Keychain item azure-identity
# keeps the token cache in, then login.microsoftonline.com. On an interpreter
# that item does not trust, that was a password dialog per test; on one it
# trusts, a silent read of the real token.
#
# Two parts, for every test that has no live marker:
#
# 1. Isolation. The test gets an empty settings directory. In this process the
#    server lifespan loads its config from there and the auth record is looked
#    for there; child processes get it as OUTLOOK_MCP_CONFIG_DIR. No config, no
#    record, nothing to refresh. Proxy variables are cleared, so a stray request
#    connects straight to its real address, where part 2 can see it.
# 2. Guard. Building a real DeviceCodeCredential or MSAL client, opening a token
#    cache store, or connecting a socket off loopback fails the test with
#    SignInReached. A child process cannot be watched from here; part 1 is what
#    keeps it away from the sign-in.
#
# A test that needs a real credential object on purpose — to run azure-identity's
# own error wrapping, with `_get_app` replaced so the credential can reach
# neither MSAL nor the cache — is marked `real_credential`. The rest of the
# guard still applies to it.

LIVE_MARKERS = ("integration", "live", "live_write")

# Through a proxy every connection goes to the proxy's address, often loopback,
# and the socket check could not see where a request was really going.
PROXY_VARIABLES = (
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "ALL_PROXY",
    "http_proxy",
    "https_proxy",
    "all_proxy",
)

# msal_extensions' stores, one of which azure-identity opens for a persistent
# token cache: the Keychain (macOS), libsecret (Linux), DPAPI (Windows), and
# the plaintext file.
TOKEN_CACHE_STORES = (
    "KeychainPersistence",
    "LibsecretPersistence",
    "FilePersistenceWithDataProtection",
    "FilePersistence",
)


class SignInReached(pytest.fail.Exception):
    """An offline test reached the real sign-in, or the network.

    A pytest failure, which is a BaseException: the ``except Exception`` in
    ``AuthManager.try_cached_token`` cannot swallow it.
    """


# Breaches during the current test phase. The report hook below reads them, so a
# breach that something caught anyway (uvicorn's lifespan catches everything)
# still fails the test.
_breaches: list[str] = []


def is_offline(item) -> bool:
    """True unless the test carries a live marker, however it was applied."""
    return not any(item.get_closest_marker(name) for name in LIVE_MARKERS)


def is_loopback(host) -> bool:
    """True for localhost and the loopback addresses, IPv4-mapped included."""
    if host == "localhost":
        return True
    try:
        address = ipaddress.ip_address(str(host).split("%", 1)[0])  # drop an IPv6 zone
    except ValueError:
        return False  # any other host name
    mapped = getattr(address, "ipv4_mapped", None)  # ::ffff:127.0.0.1
    return (mapped or address).is_loopback


def _breach(what: str):
    message = (
        f"An offline test {what} (#267). The offline suite must never touch the real "
        "sign-in or the network. Mock auth instead: patch outlook_mcp.auth.DeviceCodeCredential, "
        "or, for the server lifespan, outlook_mcp.server.load_config or "
        "AuthManager.try_cached_token. A test that needs Graph belongs in a live tier."
    )
    _breaches.append(message)
    raise SignInReached(message)


def _refuse(what: str):
    def refused(*args, **kwargs):
        _breach(what)

    return refused


def _loopback_only(connect):
    @functools.wraps(connect)
    def guarded(sock, address):
        if sock.family in (socket.AF_INET, socket.AF_INET6) and not is_loopback(address[0]):
            _breach(f"connected a socket to {address[0]!r}, off this machine")
        return connect(sock, address)

    return guarded


def _install_guard(monkeypatch, *, real_credential_allowed: bool) -> None:
    import azure.identity
    import msal
    import msal_extensions

    if not real_credential_allowed:
        monkeypatch.setattr(
            azure.identity.DeviceCodeCredential,
            "__init__",
            _refuse("built a real DeviceCodeCredential"),
        )
    monkeypatch.setattr(msal.ClientApplication, "__init__", _refuse("built an MSAL client"))
    for store in TOKEN_CACHE_STORES:
        monkeypatch.setattr(
            getattr(msal_extensions, store),
            "__init__",
            _refuse(f"opened the token cache (msal_extensions.{store})"),
        )
    monkeypatch.setattr(socket.socket, "connect", _loopback_only(socket.socket.connect))
    monkeypatch.setattr(socket.socket, "connect_ex", _loopback_only(socket.socket.connect_ex))


@contextlib.contextmanager
def _offline(item):
    from outlook_mcp import auth, config, server

    settings = Path(tempfile.mkdtemp(prefix="outlook-mcp-offline-"))
    monkeypatch = pytest.MonkeyPatch()
    try:
        monkeypatch.setenv(config.CONFIG_DIR_ENV, str(settings))
        monkeypatch.setattr(
            server, "load_config", functools.partial(config.load_config, str(settings))
        )
        monkeypatch.setattr(auth, "_auth_record_path", lambda: settings / auth.AUTH_RECORD_FILE)
        for name in PROXY_VARIABLES:
            monkeypatch.delenv(name, raising=False)
        _install_guard(
            monkeypatch,
            real_credential_allowed=item.get_closest_marker("real_credential") is not None,
        )
        yield
    finally:
        monkeypatch.undo()
        _breaches.clear()
        shutil.rmtree(settings, ignore_errors=True)


@pytest.hookimpl(wrapper=True)
def pytest_runtest_protocol(item, nextitem):
    """Isolate and guard each offline test, its fixtures' setup and teardown included.

    A hook, not an autouse fixture: a module-scoped fixture is set up before any
    function-scoped one, and test_http_transport.py's server fixture runs the
    lifespan.
    """
    if not is_offline(item):
        return (yield)
    with _offline(item):
        return (yield)


def hold_breaches_against(report, breaches: list[str], *, raised: bool) -> None:
    """Fail a test phase in which the guard fired, whatever became of the exception."""
    found = "\n\n".join(dict.fromkeys(breaches))
    if not report.failed:
        report.outcome = "failed"
        report.longrepr = found
    elif not raised:
        report.sections.append(("offline isolation (#267)", found))


@pytest.hookimpl(wrapper=True)
def pytest_runtest_makereport(item, call):
    report = yield
    if _breaches:
        raised = call.excinfo is not None and call.excinfo.errisinstance(SignInReached)
        hold_breaches_against(report, _breaches, raised=raised)
        _breaches.clear()
    return report


@contextlib.contextmanager
def expect_breach():
    """For the guard's own tests: the block must breach, and is not held against the test."""
    with pytest.raises(SignInReached) as caught:
        yield caught
    _breaches.clear()
