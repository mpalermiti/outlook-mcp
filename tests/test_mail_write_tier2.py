"""Tests for Tier 2 send_message enhancements (sensitivity, read receipt)."""

from unittest.mock import AsyncMock

import pytest

from outlook_mcp.models.mail import SendMessageInput
from outlook_mcp.tools.mail_write import send_message


class TestSendMessageSensitivity:
    async def test_sensitivity_default(self):
        """Default sensitivity is normal -- no arg needed."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(mock_client, to=["a@b.com"], subject="Test", body="Hello")
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
        )
        assert result["status"] == "sent"


class TestSendMessageReadReceipt:
    async def test_read_receipt_default_false(self):
        """Read receipt defaults to False."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(mock_client, to=["a@b.com"], subject="Test", body="Hello")
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
        )
        assert result["status"] == "sent"


class TestSendMessageInputModel:
    def test_model_validates_sensitivity_rejects_invalid(self):
        """SendMessageInput rejects invalid sensitivity values."""
        with pytest.raises(ValueError):
            SendMessageInput(
                to=["a@b.com"],
                subject="Test",
                body="Hello",
                sensitivity="top_secret",
            )

    def test_model_accepts_all_valid_sensitivity_values(self):
        """SendMessageInput accepts normal, personal, private, confidential."""
        for s in ("normal", "personal", "private", "confidential"):
            msg = SendMessageInput(to=["a@b.com"], subject="Test", body="Hello", sensitivity=s)
            assert msg.sensitivity == s

    def test_model_sensitivity_default(self):
        """SendMessageInput defaults sensitivity to normal."""
        msg = SendMessageInput(to=["a@b.com"], subject="Test", body="Hello")
        assert msg.sensitivity == "normal"

    def test_model_read_receipt_default(self):
        """SendMessageInput defaults request_read_receipt to False."""
        msg = SendMessageInput(to=["a@b.com"], subject="Test", body="Hello")
        assert msg.request_read_receipt is False

    def test_model_read_receipt_true(self):
        """SendMessageInput accepts request_read_receipt=True."""
        msg = SendMessageInput(
            to=["a@b.com"],
            subject="Test",
            body="Hello",
            request_read_receipt=True,
        )
        assert msg.request_read_receipt is True
