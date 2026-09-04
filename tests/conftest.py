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

Both tiers are READ-ONLY. Never add a fixture or test here that writes to the
mailbox — no sends, drafts, category or folder mutations, no deletes.
"""

import pytest


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
    if not auth.try_cached_token(auth.get_scopes()):
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
