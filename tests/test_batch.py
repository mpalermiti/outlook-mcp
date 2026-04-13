"""Tests for batch triage tool."""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.batch import batch_triage


class TestBatchTriageValidation:
    async def test_read_only_raises(self):
        """batch_triage raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="move",
                value="inbox",
                read_only=True,
            )

    async def test_max_20_messages(self):
        """batch_triage rejects more than 20 message IDs."""
        mock_client = MagicMock()
        ids = [f"AAMkAG{i}=" for i in range(21)]
        with pytest.raises(ValueError, match="Maximum 20"):
            await batch_triage(mock_client, message_ids=ids, action="move", value="inbox")

    async def test_exactly_20_messages_allowed(self):
        """batch_triage accepts exactly 20 message IDs."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.move.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        ids = [f"AAMkAG{i}=" for i in range(20)]
        result = await batch_triage(mock_client, message_ids=ids, action="move", value="inbox")
        assert result["success_count"] == 20

    async def test_invalid_action_raises(self):
        """batch_triage rejects unknown action names."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="action must be one of"):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="delete",
                value="inbox",
            )

    async def test_invalid_message_id_raises(self):
        """batch_triage rejects message IDs with invalid characters."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await batch_triage(
                mock_client,
                message_ids=["<script>alert(1)</script>"],
                action="move",
                value="inbox",
            )

    async def test_move_validates_folder_name(self):
        """batch_triage with action=move validates the folder name."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="move",
                value="<bad folder>",
            )


class TestBatchTriageMove:
    async def test_move_calls_individual_move(self):
        """batch_triage action=move calls move_message for each ID."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.move.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="move",
            value="archive",
        )
        assert result["success_count"] == 2
        assert result["failure_count"] == 0
        assert len(result["results"]) == 2
        assert all(r["status"] == "success" for r in result["results"])


class TestBatchTriageFlag:
    async def test_flag_calls_individual_flag(self):
        """batch_triage action=flag calls flag_message for each ID."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="flag",
            value="flagged",
        )
        assert result["success_count"] == 2
        assert result["failure_count"] == 0


class TestBatchTriageCategorize:
    async def test_categorize_splits_comma_separated(self):
        """batch_triage action=categorize splits comma-separated categories."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        with patch(
            "outlook_mcp.tools.batch.categorize_message", new_callable=AsyncMock
        ) as mock_cat:
            mock_cat.return_value = {"status": "categorized", "categories": ["Red", "Blue"]}
            result = await batch_triage(
                mock_client,
                message_ids=["AAMkAG1="],
                action="categorize",
                value="Red, Blue",
            )
            mock_cat.assert_called_once_with(mock_client, "AAMkAG1=", ["Red", "Blue"])
            assert result["success_count"] == 1


class TestBatchTriageMarkRead:
    async def test_mark_read_true(self):
        """batch_triage action=mark_read with value 'true' marks as read."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        with patch("outlook_mcp.tools.batch.mark_read", new_callable=AsyncMock) as mock_mr:
            mock_mr.return_value = {"status": "updated", "is_read": True}
            result = await batch_triage(
                mock_client,
                message_ids=["AAMkAG1="],
                action="mark_read",
                value="true",
            )
            mock_mr.assert_called_once_with(mock_client, "AAMkAG1=", True)
            assert result["success_count"] == 1

    async def test_mark_read_false(self):
        """batch_triage action=mark_read with value 'false' marks as unread."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.patch = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        with patch("outlook_mcp.tools.batch.mark_read", new_callable=AsyncMock) as mock_mr:
            mock_mr.return_value = {"status": "updated", "is_read": False}
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG1="],
                action="mark_read",
                value="false",
            )
            mock_mr.assert_called_once_with(mock_client, "AAMkAG1=", False)


class TestBatchTriageErrorHandling:
    async def test_per_item_failure_captured(self):
        """Individual failures are captured without aborting the batch."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        # First call succeeds, second raises
        msg_builder.move.post = AsyncMock(side_effect=[None, Exception("Not found")])
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="move",
            value="inbox",
        )
        assert result["success_count"] == 1
        assert result["failure_count"] == 1
        assert result["results"][0]["status"] == "success"
        assert result["results"][1]["status"] == "error"
        assert "Not found" in result["results"][1]["error"]

    async def test_all_failures_counted(self):
        """When all items fail, failure_count equals total."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.move.post = AsyncMock(side_effect=Exception("Graph error"))
        mock_client.me.messages.by_message_id.return_value = msg_builder

        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2=", "AAMkAG3="],
            action="move",
            value="inbox",
        )
        assert result["success_count"] == 0
        assert result["failure_count"] == 3
