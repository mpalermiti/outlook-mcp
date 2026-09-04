"""Integration smoke tests — require a real Azure AD app and a cached token.

Run with: uv run pytest -m integration -v
Auto-skipped without credentials (so CI, which has none, skips them).

These check that each tool returns its documented response SHAPE. They are not
query-shape guards — see tests/test_live_query_shape.py for those.

READ-ONLY: every call here must be a read. Never add a write to this module.

Fixtures (real_config / real_auth / real_graph_client) live in conftest.py and
are shared with the live tier.
"""

import pytest

# Mark all tests in this module as integration tests
pytestmark = pytest.mark.integration


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
    # NB: the key is "messages", not "drafts" — this assertion said "drafts"
    # and never failed because the whole tier was silently skipping.
    assert "messages" in result
    assert "count" in result
    assert isinstance(result["messages"], list)
