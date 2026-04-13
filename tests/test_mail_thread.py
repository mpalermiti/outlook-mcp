"""Tests for mail thread tools: list_thread, copy_message."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_thread import copy_message, list_thread


def _make_mock_message(**overrides):
    """Factory for mock Graph SDK message objects."""
    msg = MagicMock()
    msg.id = overrides.get("id", "msg1")
    msg.subject = overrides.get("subject", "Thread message")
    msg.from_ = MagicMock()
    msg.from_.email_address = MagicMock()
    msg.from_.email_address.address = overrides.get("from_email", "a@b.com")
    msg.from_.email_address.name = overrides.get("from_name", "Alice")
    msg.received_date_time = overrides.get("received", "2026-04-12T10:00:00Z")
    msg.is_read = overrides.get("is_read", True)
    msg.importance = MagicMock(value=overrides.get("importance", "normal"))
    msg.body_preview = overrides.get("body_preview", "Preview")
    msg.has_attachments = overrides.get("has_attachments", False)
    msg.categories = overrides.get("categories", [])
    msg.flag = MagicMock()
    msg.flag.flag_status = MagicMock(value=overrides.get("flag", "notFlagged"))
    msg.conversation_id = overrides.get("conversation_id", "conv123")
    return msg


class TestListThread:
    async def test_validates_conversation_id(self):
        """list_thread rejects invalid conversation IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await list_thread(mock_client, conversation_id="<script>alert(1)</script>")

    async def test_rejects_empty_conversation_id(self):
        """list_thread rejects empty conversation ID."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="must not be empty"):
            await list_thread(mock_client, conversation_id="")

    async def test_returns_messages_in_thread(self):
        """list_thread returns formatted messages with count."""
        mock_msg1 = _make_mock_message(id="msg1", subject="Re: Hello")
        mock_msg2 = _make_mock_message(id="msg2", subject="Re: Re: Hello")

        mock_client = MagicMock()
        mock_client.me.messages.get = AsyncMock(
            return_value=MagicMock(value=[mock_msg1, mock_msg2], odata_next_link=None)
        )

        result = await list_thread(mock_client, conversation_id="conv123")
        assert result["count"] == 2
        assert len(result["messages"]) == 2
        assert result["messages"][0]["id"] == "msg1"
        assert result["messages"][1]["id"] == "msg2"
        assert result["messages"][0]["subject"] == "Re: Hello"

    async def test_passes_filter_and_orderby(self):
        """list_thread passes correct $filter and $orderby to Graph API."""
        mock_client = MagicMock()
        mock_client.me.messages.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        await list_thread(mock_client, conversation_id="conv123", count=10)

        call_kwargs = mock_client.me.messages.get.call_args
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert "conversationId eq 'conv123'" in qp.filter
        assert qp.orderby == ["receivedDateTime asc"]
        assert qp.top == 10


class TestCopyMessage:
    async def test_copy_validates_message_id(self):
        """copy_message rejects invalid message IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await copy_message(mock_client, message_id="<bad>", folder="inbox")

    async def test_copy_validates_folder(self):
        """copy_message rejects invalid folder names."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await copy_message(mock_client, message_id="AAMkAG123=", folder="bad folder!")

    async def test_copy_raises_read_only(self):
        """copy_message raises ReadOnlyError when read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await copy_message(mock_client, message_id="AAMkAG123=", folder="inbox", read_only=True)

    async def test_copy_calls_post_and_returns_status(self):
        """copy_message calls copy.post() and returns success dict."""
        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.copy.post = AsyncMock()

        result = await copy_message(mock_client, message_id="AAMkAG123=", folder="archive")
        assert result["status"] == "copied"
        assert result["folder"] == "archive"
        mock_client.me.messages.by_message_id.return_value.copy.post.assert_called_once()
