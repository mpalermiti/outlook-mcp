"""Tests for admin tools: list_categories, get_mail_tips."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools.admin import get_mail_tips, list_categories


class TestListCategories:
    @pytest.mark.asyncio
    async def test_list_categories_returns_with_colors(self):
        """list_categories returns categories with color values."""
        mock_cat = MagicMock()
        mock_cat.id = "cat1"
        mock_cat.display_name = "Blue Category"
        mock_cat.color = MagicMock(value="preset0")

        mock_client = MagicMock()
        mock_client.me.outlook.master_categories.get = AsyncMock(
            return_value=MagicMock(value=[mock_cat])
        )

        result = await list_categories(mock_client)
        assert result["count"] == 1
        assert result["categories"][0]["id"] == "cat1"
        assert result["categories"][0]["display_name"] == "Blue Category"
        assert result["categories"][0]["color"] == "preset0"

    @pytest.mark.asyncio
    async def test_list_categories_empty(self):
        """list_categories returns empty list when no categories."""
        mock_client = MagicMock()
        mock_client.me.outlook.master_categories.get = AsyncMock(return_value=MagicMock(value=[]))

        result = await list_categories(mock_client)
        assert result["count"] == 0
        assert result["categories"] == []


class TestGetMailTips:
    @pytest.mark.asyncio
    async def test_get_mail_tips_validates_emails(self):
        """get_mail_tips rejects invalid email addresses."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid email"):
            await get_mail_tips(mock_client, ["not-an-email"])

    @pytest.mark.asyncio
    async def test_get_mail_tips_returns_per_recipient(self):
        """get_mail_tips returns tips for each recipient."""
        mock_tip = MagicMock()
        mock_tip.email_address = MagicMock(address="boss@company.com")
        mock_tip.is_moderated = False
        mock_tip.delivery_restricted = False
        mock_tip.max_message_size = 36700160
        mock_tip.automatic_replies = MagicMock(message="I'm OOO")

        mock_client = MagicMock()
        mock_client.me.get_mail_tips.post = AsyncMock(return_value=MagicMock(value=[mock_tip]))

        result = await get_mail_tips(mock_client, ["boss@company.com"])
        assert len(result["tips"]) == 1
        tip = result["tips"][0]
        assert tip["email"] == "boss@company.com"
        assert tip["is_moderated"] is False
        assert tip["is_delivery_restricted"] is False
        assert tip["max_message_size"] == 36700160
        assert tip["out_of_office_message"] == "I'm OOO"

    @pytest.mark.asyncio
    async def test_get_mail_tips_multiple_recipients(self):
        """get_mail_tips handles multiple recipients."""
        mock_tip1 = MagicMock()
        mock_tip1.email_address = MagicMock(address="alice@company.com")
        mock_tip1.is_moderated = False
        mock_tip1.delivery_restricted = False
        mock_tip1.max_message_size = 36700160
        mock_tip1.automatic_replies = MagicMock(message="")

        mock_tip2 = MagicMock()
        mock_tip2.email_address = MagicMock(address="bob@company.com")
        mock_tip2.is_moderated = True
        mock_tip2.delivery_restricted = False
        mock_tip2.max_message_size = 10485760
        mock_tip2.automatic_replies = MagicMock(message="On vacation")

        mock_client = MagicMock()
        mock_client.me.get_mail_tips.post = AsyncMock(
            return_value=MagicMock(value=[mock_tip1, mock_tip2])
        )

        result = await get_mail_tips(mock_client, ["alice@company.com", "bob@company.com"])
        assert len(result["tips"]) == 2
        assert result["tips"][0]["email"] == "alice@company.com"
        assert result["tips"][1]["is_moderated"] is True
        assert result["tips"][1]["out_of_office_message"] == "On vacation"
