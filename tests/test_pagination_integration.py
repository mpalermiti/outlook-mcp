"""Tests for pagination integration with list tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.pagination import encode_cursor
from outlook_mcp.tools.calendar_read import list_events
from outlook_mcp.tools.mail_read import list_folders, list_inbox, search_mail


def _make_folder_mock(messages, next_link=None):
    """Build a mock Graph client for folder-based message queries."""
    response = MagicMock(value=messages, odata_next_link=next_link)
    messages_obj = MagicMock()
    messages_obj.get = AsyncMock(return_value=response)
    folder_obj = MagicMock()
    folder_obj.messages = messages_obj
    client = MagicMock()
    client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=folder_obj)
    return client


def _make_search_mock(messages, next_link=None):
    """Build a mock Graph client for search queries (no folder)."""
    response = MagicMock(value=messages, odata_next_link=next_link)
    client = MagicMock()
    client.me.messages.get = AsyncMock(return_value=response)
    return client


class TestListInboxPagination:
    @pytest.mark.asyncio
    async def test_no_cursor_backward_compatible(self):
        """list_inbox without cursor works as before."""
        mock_client = _make_folder_mock([])
        result = await list_inbox(mock_client)
        assert result["cursor"] is None
        assert result["has_more"] is False

    @pytest.mark.asyncio
    async def test_with_cursor(self):
        """list_inbox with cursor applies skip from cursor."""
        cursor = encode_cursor(50)
        mock_client = _make_folder_mock([])
        result = await list_inbox(mock_client, cursor=cursor)
        assert result["cursor"] is None
        # Verify $skip was set from cursor
        call_kwargs = (
            mock_client.me.mail_folders.by_mail_folder_id
            .return_value.messages.get.call_args
        )
        query_params = call_kwargs.kwargs.get("query_params") or call_kwargs[1].get("query_params")
        assert query_params["$skip"] == 50

    @pytest.mark.asyncio
    async def test_returns_cursor_when_has_more(self):
        """Response includes cursor when there's a next page."""
        mock_client = _make_folder_mock(
            [],
            next_link="https://graph.microsoft.com/v1.0/me/messages?$skip=25&$top=25",
        )
        result = await list_inbox(mock_client)
        assert result["has_more"] is True
        assert result["cursor"] is not None

    @pytest.mark.asyncio
    async def test_invalid_cursor_raises(self):
        """Invalid cursor raises ValueError."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await list_inbox(mock_client, cursor="garbage")

    @pytest.mark.asyncio
    async def test_cursor_overrides_skip(self):
        """When cursor is provided, manual skip param is ignored."""
        cursor = encode_cursor(75)
        mock_client = _make_folder_mock([])
        await list_inbox(mock_client, skip=10, cursor=cursor)
        call_kwargs = (
            mock_client.me.mail_folders.by_mail_folder_id
            .return_value.messages.get.call_args
        )
        query_params = call_kwargs.kwargs.get("query_params") or call_kwargs[1].get("query_params")
        # Cursor's skip (75) should be used, not the manual skip (10)
        assert query_params["$skip"] == 75


class TestSearchMailPagination:
    @pytest.mark.asyncio
    async def test_search_returns_cursor_key(self):
        """search_mail response includes cursor field."""
        mock_client = _make_search_mock([])
        result = await search_mail(mock_client, query="test")
        assert "cursor" in result
        assert result["cursor"] is None

    @pytest.mark.asyncio
    async def test_search_returns_cursor_when_has_more(self):
        """search_mail includes cursor when next page exists."""
        mock_client = _make_search_mock(
            [],
            next_link="https://graph.microsoft.com/v1.0/me/messages?$skip=25&$top=25",
        )
        result = await search_mail(mock_client, query="test")
        assert result["has_more"] is True
        assert result["cursor"] is not None

    @pytest.mark.asyncio
    async def test_search_with_cursor(self):
        """search_mail with cursor applies skip."""
        cursor = encode_cursor(50)
        mock_client = _make_search_mock([])
        await search_mail(mock_client, query="test", cursor=cursor)
        call_kwargs = mock_client.me.messages.get.call_args
        query_params = call_kwargs.kwargs.get("query_params") or call_kwargs[1].get("query_params")
        assert query_params["$skip"] == 50

    @pytest.mark.asyncio
    async def test_search_invalid_cursor_raises(self):
        """search_mail rejects invalid cursor."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await search_mail(mock_client, query="test", cursor="garbage")


class TestListFoldersPagination:
    @pytest.mark.asyncio
    async def test_folders_returns_cursor_key(self):
        """list_folders response includes cursor field."""
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )
        result = await list_folders(mock_client)
        assert "cursor" in result
        assert result["cursor"] is None

    @pytest.mark.asyncio
    async def test_folders_returns_cursor_when_has_more(self):
        """list_folders includes cursor when next page exists."""
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=MagicMock(
                value=[],
                odata_next_link="https://graph.microsoft.com/v1.0/me/mailFolders?$skip=10&$top=10",
            )
        )
        result = await list_folders(mock_client)
        assert result["has_more"] is True
        assert result["cursor"] is not None


class TestListEventsPagination:
    @pytest.mark.asyncio
    async def test_events_returns_cursor_key(self):
        """list_events response includes cursor field."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )
        result = await list_events(mock_client, timezone="UTC")
        assert "cursor" in result
        assert result["cursor"] is None

    @pytest.mark.asyncio
    async def test_events_returns_cursor_when_has_more(self):
        """list_events includes cursor when next page exists."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(
                value=[],
                odata_next_link="https://graph.microsoft.com/v1.0/me/calendarView?$skip=50&$top=50",
            )
        )
        result = await list_events(mock_client, timezone="UTC")
        assert result["has_more"] is True
        assert result["cursor"] is not None

    @pytest.mark.asyncio
    async def test_events_with_cursor(self):
        """list_events with cursor applies skip."""
        cursor = encode_cursor(100)
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )
        await list_events(mock_client, timezone="UTC", cursor=cursor)
        call_kwargs = mock_client.me.calendar_view.get.call_args
        query_params = call_kwargs.kwargs.get("query_params") or call_kwargs[1].get("query_params")
        assert query_params["$skip"] == 100

    @pytest.mark.asyncio
    async def test_events_invalid_cursor_raises(self):
        """list_events rejects invalid cursor."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError):
            await list_events(mock_client, timezone="UTC", cursor="garbage")
