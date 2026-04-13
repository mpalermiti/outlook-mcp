"""Tests for user tools: whoami, list_calendars."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools.user import list_calendars, whoami


class TestWhoami:
    @pytest.mark.asyncio
    async def test_whoami_returns_profile(self):
        """whoami returns display_name, email, id from /me."""
        mock_user = MagicMock()
        mock_user.id = "user123"
        mock_user.display_name = "Michael"
        mock_user.mail = "michael@outlook.com"
        mock_user.user_principal_name = "michael@outlook.com"

        mock_client = MagicMock()
        mock_client.me.get = AsyncMock(return_value=mock_user)

        result = await whoami(mock_client)
        assert result["display_name"] == "Michael"
        assert result["email"] == "michael@outlook.com"
        assert result["id"] == "user123"

    @pytest.mark.asyncio
    async def test_whoami_falls_back_to_upn(self):
        """whoami uses user_principal_name when mail is None."""
        mock_user = MagicMock()
        mock_user.id = "user456"
        mock_user.display_name = "Test User"
        mock_user.mail = None
        mock_user.user_principal_name = "testuser@outlook.com"

        mock_client = MagicMock()
        mock_client.me.get = AsyncMock(return_value=mock_user)

        result = await whoami(mock_client)
        assert result["email"] == "testuser@outlook.com"

    @pytest.mark.asyncio
    async def test_whoami_handles_empty_fields(self):
        """whoami handles None display_name gracefully."""
        mock_user = MagicMock()
        mock_user.id = "user789"
        mock_user.display_name = None
        mock_user.mail = None
        mock_user.user_principal_name = None

        mock_client = MagicMock()
        mock_client.me.get = AsyncMock(return_value=mock_user)

        result = await whoami(mock_client)
        assert result["display_name"] == ""
        assert result["email"] == ""
        assert result["id"] == "user789"


class TestListCalendars:
    @pytest.mark.asyncio
    async def test_list_calendars_returns_list(self):
        """list_calendars returns structured calendar list."""
        mock_cal = MagicMock()
        mock_cal.id = "cal1"
        mock_cal.name = "Calendar"
        mock_cal.color = MagicMock(value="auto")
        mock_cal.is_default_calendar = True
        mock_cal.can_edit = True

        mock_client = MagicMock()
        mock_client.me.calendars.get = AsyncMock(return_value=MagicMock(value=[mock_cal]))

        result = await list_calendars(mock_client)
        assert result["count"] == 1
        assert result["calendars"][0]["id"] == "cal1"
        assert result["calendars"][0]["name"] == "Calendar"
        assert result["calendars"][0]["is_default"] is True
        assert result["calendars"][0]["can_edit"] is True

    @pytest.mark.asyncio
    async def test_list_calendars_empty(self):
        """list_calendars returns empty list when no calendars."""
        mock_client = MagicMock()
        mock_client.me.calendars.get = AsyncMock(return_value=MagicMock(value=[]))

        result = await list_calendars(mock_client)
        assert result["count"] == 0
        assert result["calendars"] == []
