"""Tests for mailbox settings tools: get_timezone, set_timezone."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.mailbox_settings import get_timezone, set_timezone

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)
_CFG_ALLOW_TZ = Config(client_id="test", allow_categories=["mailbox_settings"])
_CFG_ALLOW_OTHER = Config(client_id="test", allow_categories=["mail_send"])


class TestGetTimezone:
    async def test_returns_timezone_string(self):
        """get_timezone returns the time_zone field from MailboxSettings."""
        settings = MagicMock()
        settings.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_timezone(mock_client)
        assert result == {"timezone": "America/Los_Angeles"}
        mock_client.me.mailbox_settings.get.assert_called_once()

    async def test_returns_empty_string_when_none(self):
        """get_timezone returns empty string when time_zone is None."""
        settings = MagicMock()
        settings.time_zone = None

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_timezone(mock_client)
        assert result == {"timezone": ""}


class TestSetTimezone:
    async def test_patches_mailbox_settings_and_returns_updated(self):
        """set_timezone PATCHes with MailboxSettings(time_zone=...) and returns echoed value."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG)
        assert result == {"status": "updated", "timezone": "America/Los_Angeles"}

        mock_client.me.mailbox_settings.patch.assert_called_once()
        body = mock_client.me.mailbox_settings.patch.call_args.args[0]
        assert body.time_zone == "America/Los_Angeles"

    async def test_falls_back_to_input_when_echo_missing(self):
        """set_timezone returns input value when patch returns None or no time_zone."""
        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_timezone(mock_client, timezone="Pacific Standard Time", config=_CFG)
        assert result == {"status": "updated", "timezone": "Pacific Standard Time"}

    async def test_rejects_empty_string(self):
        """set_timezone raises ValueError on empty string."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await set_timezone(mock_client, timezone="", config=_CFG)

    async def test_rejects_whitespace_only(self):
        """set_timezone raises ValueError on whitespace-only string."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await set_timezone(mock_client, timezone="   ", config=_CFG)

    @pytest.mark.asyncio
    async def test_rejects_control_characters(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock()
        with pytest.raises(ValueError, match="control"):
            await set_timezone(client, "America/\x00Los_Angeles", config=_CFG)

    async def test_raises_read_only(self):
        """set_timezone raises ReadOnlyError when config.read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG_RO)

    async def test_raises_permission_denied_when_not_in_whitelist(self):
        """set_timezone raises PermissionDeniedError when category not in allow_categories."""
        mock_client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG_ALLOW_OTHER)

    async def test_succeeds_when_category_in_whitelist(self):
        """set_timezone succeeds when 'mailbox_settings' is in allow_categories."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(
            mock_client, timezone="America/Los_Angeles", config=_CFG_ALLOW_TZ
        )
        assert result["status"] == "updated"
        assert result["timezone"] == "America/Los_Angeles"

    async def test_succeeds_with_empty_allow_categories(self):
        """set_timezone succeeds when allow_categories is empty (fully open)."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG)
        assert result["status"] == "updated"
