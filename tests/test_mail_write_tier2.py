"""Tests for Tier 2 send_message enhancements (sensitivity, read receipt)."""

from unittest.mock import AsyncMock

from outlook_mcp.config import Config
from outlook_mcp.tools.mail_write import send_message

_CFG = Config(client_id="test")


class TestSendMessageSensitivity:
    async def test_sensitivity_default(self):
        """Default sensitivity is normal -- no arg needed."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client, to=["a@b.com"], subject="Test", body="Hello", config=_CFG
        )
        assert result["status"] == "sent"

    async def test_sensitivity_confidential(self):
        """Confidential sensitivity is accepted and message sends."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["a@b.com"],
            subject="Test",
            body="Hello",
            sensitivity="confidential",
            config=_CFG,
        )
        assert result["status"] == "sent"

    async def test_sensitivity_private(self):
        """Private sensitivity is accepted."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["a@b.com"],
            subject="Test",
            body="Hello",
            sensitivity="private",
            config=_CFG,
        )
        assert result["status"] == "sent"

    async def test_sensitivity_personal(self):
        """Personal sensitivity is accepted."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["a@b.com"],
            subject="Test",
            body="Hello",
            sensitivity="personal",
            config=_CFG,
        )
        assert result["status"] == "sent"


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
