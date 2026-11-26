"""Shared test fixtures for outlook-mcp.

Four tiers here reach a real mailbox. Each is deselected by default and runs
against the project's own throwaway test accounts — a main account and its pair —
signed in only inside the container ``scripts/live-tests.sh`` starts
(RELEASING.md). Nothing in the repository names them: the run takes their
addresses from the environment (``OUTLOOK_MCP_TEST_ACCOUNT``,
``OUTLOOK_MCP_TEST_PAIR``), and a live test refuses to start unless the signed-in
account, read from ``/me`` and the saved sign-in record, is one of them
(``tests/live_guard.py``, "Whose mailbox" below). The one exception is
deliberate: ``OUTLOOK_MCP_LIVE_OWN_MAILBOX=1`` lets the read-only ``live`` tier
run against whatever account is signed in, as an optional smoke check of a real,
lived-in mailbox.

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
    Mutates the test mailbox, added for #41. Some payloads can only be
    validated by Graph accepting them — a recurrence that is well-formed to the
    SDK still 400s if the range disagrees with the start — and a mock cannot
    see that. Kept as narrow as possible.

``-m live_send``
    The only tier that sends: real mail and invitations from the main test
    account to its pair, and nobody else. See "Sending" below.

Tiers 1 and 2 are READ-ONLY: never add a fixture or test there that writes to
the mailbox — no sends, drafts, category or folder mutations, no deletes. They
are held to it: while one runs, any request but a read is refused before it
leaves (``live_guard.ReadOnly``).

Rules for ``live_write``, which exists precisely because it breaks that:

* Double-gated — the marker is deselected by default *and* the tier skips
  unless ``OUTLOOK_MCP_LIVE_WRITE=1``. A cached token alone must never be
  enough to write to someone's calendar.
* Calendar, contacts, To Do and unsent drafts only — the surfaces whose writes
  have no outward side effect. Nothing here may send, and a stray send is not
  recoverable: a draft is created, read back and deleted, never sent
  (``tests/test_live_tiers_never_send.py`` checks the tier's files for a send,
  and every request a live_write test makes is refused before it leaves if it
  would send: ``live_guard.NeverSends``, see "What may leave" below).
* No attendees, ever. A recurring invite emails real people on every
  occurrence.
* Bounded ranges only (``numbered``/``endDate``) — never ``noEnd``.
* Every created item is removed in a ``finally``, and subjects/names carry
  ``LIVE_WRITE_SUBJECT`` / ``LIVE_WRITE_NAME`` so anything a crash leaks is
  greppable in the UI (``scripts/live-tests.sh`` lists any after the run).

Sending (``live_send``), the one place a send is allowed:

* Triple-gated: the marker, ``OUTLOOK_MCP_LIVE_SEND=1``, and a send budget for
  the run (``OUTLOOK_MCP_LIVE_SEND_BUDGET``; ``scripts/live-tests.sh live_send
  --sends N`` sets both). It runs only as the main test account.
* Recipients: only the pair. Every request the test makes is checked before
  it leaves (``live_guard.SendGuard``): each recipient of a ``sendMail``, each
  attendee of an invitation, the attendees of a meeting it updates or cancels.
  A send whose recipients the request does not name (a reply, a draft's send)
  is refused, as is changing or deleting any event the tier did not create.
* Counted: a test reserves the sends it will make before making any, and
  skips if the run's budget cannot cover them, so a test never stops halfway
  with an invitation it cannot cancel.
* Tests that need the pair to answer (reply, accept) skip unless
  ``OUTLOOK_MCP_TEST_PAIR_CAN_SEND=1``: a new Outlook.com account cannot send
  until its owner verifies it.
* Everything it sends or receives carries ``LIVE_SEND_SUBJECT`` and is deleted
  from both mailboxes afterwards (``tests/live_mailbox.py``).

Every other test is offline, and is held to it (#267) — see "Offline
isolation" below: it runs against an empty settings directory of its own, and
it fails if it builds a real credential, opens the token cache, or reaches the
network.
"""

import contextlib
import functools
import inspect
import ipaddress
import os
import shutil
import socket
import sys
import tempfile

import pytest

from tests.live_guard import (  # the markers are re-exported for the live test files
    LIVE_MARKERS,
    LIVE_SEND_SUBJECT,  # noqa: F401
    LIVE_WRITE_NAME,  # noqa: F401
    LIVE_WRITE_SUBJECT,  # noqa: F401
    MAX_SEND_BUDGET,
    NeverSends,
    ReadOnly,
    SendGuard,
    account_refusal,
    guarding,
    identity_of,
    live_tiers_of,
    own_mailbox_allows,
    refusal,
    settings_from,
)


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
        pytest.skip(NOT_SIGNED_IN)
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


# ── Whose mailbox (tests/live_guard.py) ─────────────────────────────────
#
# Checked in pytest_runtest_setup, before any fixture runs, so a live test that
# builds its own client is held to it too. First, from the environment alone and
# before any sign-in is read: is a test account named, or has the owner opted in
# to the read-only `live` tier on their own mailbox? Then, once per run, which
# account the cached token belongs to (/me and the saved sign-in record).

NOT_SIGNED_IN = (
    "Not authenticated — the settings directory this run uses holds no valid sign-in. "
    "Live tiers run in the container scripts/live-tests.sh starts, where each test "
    "account's signed-in settings are mounted."
)
GRAPH_ME = "https://graph.microsoft.com/v1.0/me"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# The signed-in identity, looked up once per run: {"who": address-or-None} or {"skip": reason}.
_identity: dict[str, str | None] = {}


def _signed_in_identity() -> str | None:
    """The address the cached token belongs to, or None if the sources disagree."""
    if "skip" in _identity:
        pytest.skip(_identity["skip"])
    if "who" in _identity:
        return _identity["who"]

    import httpx

    from outlook_mcp.auth import AuthManager, _load_auth_record
    from outlook_mcp.config import load_config

    config = load_config()
    auth = AuthManager(config)
    if not config.client_id or not auth.try_cached_token():
        _identity["skip"] = NOT_SIGNED_IN
        pytest.skip(NOT_SIGNED_IN)
    token = auth.get_credential().get_token(GRAPH_SCOPE).token
    response = httpx.get(
        GRAPH_ME,
        params={"$select": "userPrincipalName,mail"},
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    # Fail closed: a /me that did not answer leaves the account unknown, which
    # refuses. The saved record alone never stands in for it.
    try:
        me = response.json() if response.is_success else {}
    except ValueError:
        me = {}
    record = _load_auth_record()
    _identity["who"] = identity_of(me, getattr(record, "username", None))
    return _identity["who"]


def _refuse_to_run(item, reason: str) -> None:
    item.session.shouldstop = reason  # one refusal is enough to say it; stop the run
    pytest.fail(reason, pytrace=False)


@pytest.hookimpl(tryfirst=True)
def pytest_runtest_setup(item):
    tiers = live_tiers_of(item)
    if not tiers:
        return
    settings = settings_from(os.environ)
    reason = refusal(tiers, settings)
    if reason:
        _refuse_to_run(item, reason)
    if own_mailbox_allows(tiers, settings):
        return
    reason = account_refusal(_signed_in_identity(), tiers, settings)
    if reason:
        _refuse_to_run(item, reason)


# ── What may leave (tests/live_guard.py) ──────────────────────────────────
#
# For the whole of each live test, fixtures included, every HTTP request the
# process makes (the SDK's, a tool's own client's, a raw httpx or requests call)
# is held to the tier's guard before it leaves: `live` and `integration`, the
# own-mailbox smoke check among them, may only read; `live_write` may change the
# mailbox but not send; `live_send` sends only to the pair, counted.


@pytest.fixture(autouse=True)
def _what_may_leave(request):
    tiers = live_tiers_of(request.node)
    if not tiers:
        yield
        return
    label = ", ".join(sorted(tiers))
    if "live_send" in tiers:
        guard = request.getfixturevalue("send_guard")
    elif "live_write" in tiers:
        guard = NeverSends(tier=label)
    else:
        guard = ReadOnly(tier=label)
    with guarding(guard):
        yield


# ── The id family switch (intents/ids.py) ─────────────────────────────────


@pytest.fixture(autouse=True)
def _a_classic_install_unless_a_test_says_otherwise(monkeypatch):
    """Every test starts with the immutable-id switch off and no id family learned.

    The server turns the switch on at import when its selection loads an intent
    tool, so CI's every-group run would otherwise send the preference from every
    classic tool and change the headers their tests pin. A test of the switch, or
    of a process that loads intents, sets it itself or runs a fresh process.
    """
    from outlook_mcp.intents import ids

    monkeypatch.setattr(ids, "IMMUTABLE_IDS", False)
    monkeypatch.setattr(ids, "ID_FAMILY", ids.IdFamily())


# ── The send tier ─────────────────────────────────────────────────────────

# Every SendGuard this run made, for the summary line.
_send_guards: list[SendGuard] = []


@pytest.fixture(scope="session")
def send_guard():
    """The run's send budget and recipient rule, or skip."""
    settings = settings_from(os.environ)
    if not settings.send:
        pytest.skip(
            "Send tier is opt-in — scripts/live-tests.sh live_send --sends N sets "
            "OUTLOOK_MCP_LIVE_SEND=1 and the run's send budget"
        )
    if settings.send_budget is None:
        pytest.skip(
            "No send budget for this run — set OUTLOOK_MCP_LIVE_SEND_BUDGET to a count "
            f"from 0 to {MAX_SEND_BUDGET}"
        )
    guard = SendGuard(pair=settings.pair, budget=settings.send_budget)
    _send_guards.append(guard)
    return guard


@pytest.fixture
def sends(send_guard):
    """The run's send guard; whatever the test reserved and did not use is released."""
    yield send_guard
    send_guard.release()


@pytest.fixture
def send_graph_client(real_auth, send_guard):
    """A Graph client for the send tier. Like every request a live_send test makes,
    its requests are held to the send guard (`_what_may_leave`)."""
    from outlook_mcp.graph import GraphClient

    return GraphClient(real_auth.get_credential())


@pytest.fixture
def send_config(real_config, tmp_path):
    """Real config for the send tier, with sending and invitations allowed in process.

    As with `live_write_config`, the settings file is never modified.
    """
    if real_config.read_only:
        pytest.skip("Config is read_only — the send tier cannot run")
    update = {"attachments_dir": str(tmp_path / "attachments")}
    if real_config.allow_categories:
        from outlook_mcp.permissions import CATEGORY_CALENDAR_WRITE, CATEGORY_MAIL_SEND

        needed = {CATEGORY_CALENDAR_WRITE, CATEGORY_MAIL_SEND}
        update["allow_categories"] = sorted(set(real_config.allow_categories) | needed)
    return real_config.model_copy(update=update)


@pytest.fixture(scope="session")
def pair_mailbox():
    """`tests.live_mailbox`, once the pair account is known to be mounted and reachable."""
    settings = settings_from(os.environ)
    if not settings.pair_config_dir or not os.path.isdir(settings.pair_config_dir):
        pytest.skip(
            "The pair account's settings are not mounted "
            "(OUTLOOK_MCP_TEST_PAIR_CONFIG_DIR; scripts/live-tests.sh mounts them)"
        )
    from tests import live_mailbox

    # Before any test sends: a send the test cannot then find, or clean up after,
    # is a send wasted. Outlook.com can lock a new account outright.
    try:
        live_mailbox.run("pair", "whoami", timeout=120)
    except live_mailbox.MailboxError as exc:
        pytest.skip(f"The pair account cannot be reached, so nothing is sent to it: {exc}")
    return live_mailbox


def pytest_terminal_summary(terminalreporter):
    for guard in _send_guards:
        terminalreporter.write_line(
            f"live_send: {guard.used} real send(s) this run, budget {guard.budget}"
            + (f" — {'; '.join(guard.log)}" if guard.log else "")
        )


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
# Three parts:
#
# 1. The run. When the -m expression selects no live tier (the default run),
#    pytest_configure points OUTLOOK_MCP_CONFIG_DIR at an empty temporary
#    directory before outlook_mcp is first imported. Nothing later is early
#    enough: config.DEFAULT_CONFIG_DIR is fixed at import, and server.py reads
#    the config at import to choose its tools. A run that selects a live tier
#    keeps the real directory, because the live fixtures need it.
# 2. Each test without a live marker gets a fresh, empty settings directory of
#    its own. Every DEFAULT_CONFIG_DIR the package holds and every function
#    default bound to it (load_config, save_config) point there, and so does
#    OUTLOOK_MCP_CONFIG_DIR for child processes. No config, no auth record,
#    nothing to refresh, and nothing a test writes outlives it. Proxy variables
#    are cleared, so a stray request goes straight to its real address, where
#    part 3 can see it.
# 3. The guard. Building a real DeviceCodeCredential or MSAL client, opening a
#    token cache store, looking a host name up in DNS, or connecting or sending
#    to an address off loopback fails the test with SignInReached. A child
#    process cannot be watched from here; parts 1 and 2 keep it away.
#
# A test that needs a real credential object on purpose — to run azure-identity's
# own error wrapping, with `_get_app` replaced so the credential can reach
# neither MSAL nor the cache — is marked `real_credential`. The rest of the
# guard still applies to it.

# outlook_mcp.config.CONFIG_DIR_ENV, spelled out: importing outlook_mcp here
# would fix the settings directory before part 1 can move it.
CONFIG_DIR_ENV = "OUTLOOK_MCP_CONFIG_DIR"

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


# The run's empty settings directory (part 1), or None when a live tier is selected.
run_settings: str | None = None

# Breaches during the current test phase. The report hook below reads them, so a
# breach that something caught anyway (uvicorn's lifespan catches everything)
# still fails the test.
_breaches: list[str] = []


def is_offline(item) -> bool:
    """True unless the test carries a live marker, however it was applied."""
    return not any(item.get_closest_marker(name) for name in LIVE_MARKERS)


def _carrying(marker: str):
    return lambda name, /, **kwargs: name == marker


def selects_a_live_tier(markexpr: str) -> bool:
    """Whether a run with this -m expression can select a test from a live tier.

    Uses pytest's own parser for -m (private: there is no public one), so the
    answer is the one pytest's deselection gives.
    """
    if not markexpr.strip():
        return True  # no -m at all: every test, the live tiers included
    from _pytest.mark.expression import Expression

    expression = Expression.compile(markexpr)
    return any(expression.evaluate(_carrying(marker)) for marker in LIVE_MARKERS)


def pytest_configure(config):
    """Part 1: an offline run never sees the real settings directory, even at import."""
    global run_settings
    if selects_a_live_tier(config.getoption("markexpr")):
        return
    previous = os.environ.get(CONFIG_DIR_ENV)
    run_settings = tempfile.mkdtemp(prefix="outlook-mcp-offline-run-")
    os.environ[CONFIG_DIR_ENV] = run_settings

    def restore():
        if previous is None:
            os.environ.pop(CONFIG_DIR_ENV, None)
        else:
            os.environ[CONFIG_DIR_ENV] = previous
        shutil.rmtree(run_settings, ignore_errors=True)

    config.add_cleanup(restore)


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


def needs_dns(host) -> bool:
    """True for a host name a lookup would send to DNS: not empty, localhost or an address."""
    if host is None:
        return False
    if isinstance(host, bytes):
        host = host.decode("ascii", "replace")
    if host in ("", "localhost"):
        return False
    try:
        ipaddress.ip_address(host.split("%", 1)[0])
    except ValueError:
        return True
    return False


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


def _sends_to_loopback_only(sendto):
    @functools.wraps(sendto)
    def guarded(sock, *args):
        address = args[-1]  # sendto(data, address) or sendto(data, flags, address)
        if sock.family in (socket.AF_INET, socket.AF_INET6) and not is_loopback(address[0]):
            _breach(f"sent a datagram to {address[0]!r}, off this machine")
        return sendto(sock, *args)

    return guarded


def _no_name_lookups(getaddrinfo):
    @functools.wraps(getaddrinfo)
    def guarded(host, *args, **kwargs):
        if needs_dns(host):
            _breach(f"looked up {host!r} in DNS")
        return getaddrinfo(host, *args, **kwargs)

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
    monkeypatch.setattr(socket.socket, "sendto", _sends_to_loopback_only(socket.socket.sendto))
    monkeypatch.setattr(socket, "getaddrinfo", _no_name_lookups(socket.getaddrinfo))


def _move_settings(set_attribute, old: str, new: str) -> None:
    """Point each DEFAULT_CONFIG_DIR the package holds, and each function default
    bound to one, from `old` to `new`."""
    for name, module in list(sys.modules.items()):
        if module is None or not (name == "outlook_mcp" or name.startswith("outlook_mcp.")):
            continue
        if getattr(module, "DEFAULT_CONFIG_DIR", None) == old:
            set_attribute(module, "DEFAULT_CONFIG_DIR", new)
        for value in list(vars(module).values()):
            defaults = value.__defaults__ if inspect.isfunction(value) else None
            if defaults and any(isinstance(d, str) and d == old for d in defaults):
                moved = tuple(new if isinstance(d, str) and d == old else d for d in defaults)
                set_attribute(value, "__defaults__", moved)


@contextlib.contextmanager
def _offline(item):
    from outlook_mcp import config

    settings = tempfile.mkdtemp(prefix="outlook-mcp-offline-")
    monkeypatch = pytest.MonkeyPatch()
    try:
        # Part 2: this process, and child processes through the environment.
        _move_settings(monkeypatch.setattr, config.DEFAULT_CONFIG_DIR, settings)
        monkeypatch.setenv(CONFIG_DIR_ENV, settings)
        for name in PROXY_VARIABLES:
            monkeypatch.delenv(name, raising=False)
        _install_guard(
            monkeypatch,
            real_credential_allowed=item.get_closest_marker("real_credential") is not None,
        )
        yield
    finally:
        monkeypatch.undo()
        # A module first imported during the test took the test's directory as its
        # own; hand it the restored one, as if it had been imported outside.
        _move_settings(setattr, settings, config.DEFAULT_CONFIG_DIR)
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
    """For the guard's own tests: the block must breach, and only that breach is forgiven."""
    with pytest.raises(SignInReached) as caught:
        yield caught
    _breaches.remove(caught.value.msg)
