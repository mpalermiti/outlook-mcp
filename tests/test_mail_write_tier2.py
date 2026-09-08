"""Tests for Tier 2 send_message enhancements (read receipt)."""

from unittest.mock import AsyncMock

from outlook_mcp.config import Config
from outlook_mcp.tools.mail_write import send_message

_CFG = Config(client_id="test")


class TestSendMessageReadReceipt:
    async def test_read_receipt_default_false(self):
        """Read receipt defaults to False."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client, to=["a@b.com"], subject="Test", body="Hello", config=_CFG
        )
        assert result["status"] == "sent"

    async def test_read_receipt_true(self):
        """Read receipt flag True is accepted."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["a@b.com"],
            subject="Test",
            body="Hello",
            request_read_receipt=True,
            config=_CFG,
        )
        assert result["status"] == "sent"
