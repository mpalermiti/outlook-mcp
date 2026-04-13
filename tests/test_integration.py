"""Integration smoke test — requires real Azure AD app and prior outlook_login.

Run with: uv run pytest -m integration -v
Skipped by default in CI.
"""

import pytest

# Mark all tests in this module as integration tests
pytestmark = pytest.mark.integration


@pytest.fixture
def real_config():
    """Load real config from ~/.outlook-mcp/config.json."""
    from outlook_mcp.config import load_config

    config = load_config()
    if not config.client_id:
        pytest.skip("No client_id configured — run Azure AD app setup first")
    return config


@pytest.fixture
def real_auth(real_config):
    """Get real AuthManager with existing credentials."""
    from outlook_mcp.auth import AuthManager

    auth = AuthManager(real_config)
    # Try to create credential with cached token
    try:
        auth.login()
        auth.get_credential()
    except Exception:
        pytest.skip("Not authenticated — run outlook_login first")
    return auth


@pytest.fixture
def real_graph_client(real_auth):
    """Get real Graph client."""
    from outlook_mcp.graph import GraphClient

    return GraphClient(real_auth.get_credential())


@pytest.mark.asyncio
async def test_list_inbox_smoke(real_graph_client):
    """List inbox returns valid response shape."""
    from outlook_mcp.tools.mail_read import list_inbox

    result = await list_inbox(real_graph_client.sdk_client, count=1)
    assert "messages" in result
    assert "count" in result
    assert isinstance(result["messages"], list)


@pytest.mark.asyncio
async def test_list_events_smoke(real_graph_client, real_config):
    """List events returns valid response shape."""
    from outlook_mcp.tools.calendar_read import list_events

    result = await list_events(
        real_graph_client.sdk_client, days=1, timezone=real_config.timezone
    )
    assert "events" in result
    assert "count" in result
    assert isinstance(result["events"], list)


@pytest.mark.asyncio
async def test_list_folders_smoke(real_graph_client):
    """List folders returns valid response shape."""
    from outlook_mcp.tools.mail_read import list_folders

    result = await list_folders(real_graph_client.sdk_client)
    assert "folders" in result
    assert "count" in result
    assert isinstance(result["folders"], list)
    # Every Outlook account has at least inbox
    assert result["count"] > 0


@pytest.mark.asyncio
async def test_list_contacts_smoke(real_graph_client):
    """List contacts returns valid response shape."""
    from outlook_mcp.tools.contacts import list_contacts

    result = await list_contacts(real_graph_client.sdk_client)
    assert "contacts" in result
    assert "count" in result
    assert isinstance(result["contacts"], list)


@pytest.mark.asyncio
async def test_list_task_lists_smoke(real_graph_client):
    """List task lists returns valid response shape."""
    from outlook_mcp.tools.todo import list_task_lists

    result = await list_task_lists(real_graph_client.sdk_client)
    assert "task_lists" in result
    assert "count" in result
    assert isinstance(result["task_lists"], list)


@pytest.mark.asyncio
async def test_list_drafts_smoke(real_graph_client):
    """List drafts returns valid response shape."""
    from outlook_mcp.tools.mail_drafts import list_drafts

    result = await list_drafts(real_graph_client.sdk_client)
    assert "drafts" in result
    assert "count" in result
    assert isinstance(result["drafts"], list)
