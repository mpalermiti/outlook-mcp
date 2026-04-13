"""Tests for mail triage tools: move, delete, flag, categorize, mark_read."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_triage import (
    categorize_message,
    delete_message,
    flag_message,
    mark_read,
    move_message,
)


class TestMoveMessage:
    async def test_move_validates_folder_and_calls_post(self):
        """move_message validates folder name and calls move.post()."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.move.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await move_message(mock_client, message_id="AAMkAG123=", folder="inbox")
        assert result["status"] == "moved"
        assert result["folder"] == "inbox"
        msg_builder.move.post.assert_called_once()

    async def test_move_rejects_invalid_message_id(self):
        """move_message rejects message IDs with invalid characters."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await move_message(mock_client, message_id="<script>alert(1)</script>", folder="inbox")

    async def test_move_raises_read_only(self):
        """move_message raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await move_message(
                mock_client, message_id="AAMkAG123=", folder="inbox", read_only=True
            )


class TestDeleteMessage:
    async def test_soft_delete_moves_to_deleted_items(self):
        """delete_message with permanent=False moves to deleteditems folder."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.move.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await delete_message(mock_client, message_id="AAMkAG123=", permanent=False)
        assert result["status"] == "moved"
        assert result["folder"] == "deleteditems"
        msg_builder.move.post.assert_called_once()

    async def test_hard_delete_calls_delete(self):
        """delete_message with permanent=True calls .delete()."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.delete = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await delete_message(mock_client, message_id="AAMkAG123=", permanent=True)
        assert result["status"] == "permanently_deleted"
        msg_builder.delete.assert_called_once()

    async def test_delete_raises_read_only(self):
        """delete_message raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_message(mock_client, message_id="AAMkAG123=", read_only=True)


class TestFlagMessage:
    async def test_flag_sets_status(self):
        """flag_message patches with the correct flag status."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await flag_message(mock_client, message_id="AAMkAG123=", status="flagged")
        assert result["status"] == "flagged"
        assert result["flag_status"] == "flagged"
        msg_builder.patch.assert_called_once()

    async def test_flag_rejects_invalid_status(self):
        """flag_message rejects status values not in the allowed enum."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="flag status must be one of"):
            await flag_message(mock_client, message_id="AAMkAG123=", status="invalid")

    async def test_flag_raises_read_only(self):
        """flag_message raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await flag_message(
                mock_client, message_id="AAMkAG123=", status="flagged", read_only=True
            )


class TestCategorizeMessage:
    async def test_categorize_sets_categories(self):
        """categorize_message patches with the categories list."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        categories = ["Red Category", "Blue Category"]
        result = await categorize_message(
            mock_client, message_id="AAMkAG123=", categories=categories
        )
        assert result["status"] == "categorized"
        assert result["categories"] == categories
        msg_builder.patch.assert_called_once()

    async def test_categorize_rejects_invalid_message_id(self):
        """categorize_message rejects invalid message IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await categorize_message(
                mock_client, message_id="bad id with spaces!", categories=["Red"]
            )


class TestMarkRead:
    async def test_mark_read_patches_is_read(self):
        """mark_read patches isRead on the message."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await mark_read(mock_client, message_id="AAMkAG123=", is_read=True)
        assert result["status"] == "updated"
        assert result["is_read"] is True
        msg_builder.patch.assert_called_once()

    async def test_mark_unread(self):
        """mark_read with is_read=False marks the message as unread."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await mark_read(mock_client, message_id="AAMkAG123=", is_read=False)
        assert result["status"] == "updated"
        assert result["is_read"] is False

    async def test_mark_read_raises_read_only(self):
        """mark_read raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await mark_read(
                mock_client, message_id="AAMkAG123=", is_read=True, read_only=True
            )
