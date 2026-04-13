"""Tests for mail write tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_write import forward, reply, send_message


class TestSendMessage:
    async def test_send_validates_emails(self):
        """send_message validates to addresses and calls send_mail.post()."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client, to=["valid@test.com"], subject="Test", body="Hello"
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()

    async def test_send_rejects_invalid_email(self):
        """send_message rejects invalid email addresses."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError):
            await send_message(
                mock_client, to=["not-an-email"], subject="Test", body="Hello"
            )

    async def test_send_raises_read_only(self):
        """send_message raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await send_message(
                mock_client, to=["a@b.com"], subject="Test", body="Hello", read_only=True
            )

    async def test_send_with_cc_bcc(self):
        """send_message passes cc and bcc recipients."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["to@test.com"],
            subject="Test",
            body="Hello",
            cc=["cc@test.com"],
            bcc=["bcc@test.com"],
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()


class TestReply:
    async def test_reply_calls_reply_post(self):
        """reply calls reply.post() for single reply."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder
        result = await reply(mock_client, message_id="AAMkAG123=", body="Thanks!")
        assert result["status"] == "replied"
        assert result["reply_all"] is False
        msg_builder.reply.post.assert_called_once()

    async def test_reply_all_calls_reply_all_post(self):
        """reply with reply_all=True calls reply_all.post()."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply_all.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder
        result = await reply(
            mock_client, message_id="AAMkAG123=", body="Thanks!", reply_all=True
        )
        assert result["reply_all"] is True
        msg_builder.reply_all.post.assert_called_once()

    async def test_reply_raises_read_only(self):
        """reply raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await reply(mock_client, message_id="AAMkAG123=", body="Thanks!", read_only=True)


class TestForward:
    async def test_forward_validates_to(self):
        """forward validates recipient addresses and calls forward.post()."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.forward.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder
        result = await forward(mock_client, message_id="AAMkAG123=", to=["a@b.com"])
        assert result["status"] == "forwarded"
        msg_builder.forward.post.assert_called_once()

    async def test_forward_raises_read_only(self):
        """forward raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await forward(
                mock_client, message_id="AAMkAG123=", to=["a@b.com"], read_only=True
            )
