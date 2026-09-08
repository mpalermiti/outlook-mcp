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
* Calendar only. No mail: nothing here may send, and a stray send is not
  recoverable.
* No attendees, ever. A recurring invite emails real people on every
  occurrence.
* Bounded ranges only (``numbered``/``endDate``) — never ``noEnd``.
* Every created item is removed in a ``finally``, and subjects carry
  ``LIVE_WRITE_SUBJECT`` so anything a crash leaks is greppable in the UI.
"""

import os

import pytest

# Prefix for every item this tier creates, so an orphan left by a hard crash is
# obvious in the calendar UI and easy to search for.
LIVE_WRITE_SUBJECT = "[outlook-mcp live-write guard] safe to delete"


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


@pytest.fixture
def live_write_config(real_config):
    """Real config for the write tier, or skip.

    Gates on an explicit environment opt-in beyond the marker: these tests
    create and delete real calendar events on whatever account the cached token
    belongs to.
    """
    if os.environ.get("OUTLOOK_MCP_LIVE_WRITE") != "1":
        pytest.skip(
            "Write tier is opt-in — set OUTLOOK_MCP_LIVE_WRITE=1 to let it "
            "create and delete events on the authenticated calendar"
        )
    if real_config.read_only:
        pytest.skip("Config is read_only — the write tier cannot run")
    return real_config
