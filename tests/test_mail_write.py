"""Tests for mail write tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_write import forward, reply, send_message

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


class TestSendMessage:
    async def test_send_validates_emails(self):
        """send_message validates to addresses and calls send_mail.post()."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client, to=["valid@test.com"], subject="Test", body="Hello",
            config=_CFG,
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()

    async def test_send_rejects_invalid_email(self):
        """send_message rejects invalid email addresses."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError):
            await send_message(
                mock_client, to=["not-an-email"], subject="Test", body="Hello",
                config=_CFG,
            )

    async def test_send_raises_read_only(self):
        """send_message raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await send_message(
                mock_client, to=["a@b.com"], subject="Test", body="Hello",
                config=_CFG_RO,
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
            config=_CFG,
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()

    async def test_send_sets_reply_to(self):
        """send_message populates Message.reply_to when reply_to is provided."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["to@test.com"],
            subject="Test",
            body="Hello",
            reply_to=["alias@test.com", "team@test.com"],
            config=_CFG,
        )
        assert result["status"] == "sent"
        request_body = mock_client.me.send_mail.post.call_args.args[0]
        reply_to_list = request_body.message.reply_to
        assert reply_to_list is not None
        assert [r.email_address.address for r in reply_to_list] == [
            "alias@test.com",
            "team@test.com",
        ]

    async def test_send_no_reply_to_leaves_field_unset(self):
        """send_message does not set Message.reply_to when reply_to is omitted."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        await send_message(
            mock_client, to=["to@test.com"], subject="Test", body="Hello", config=_CFG,
        )
        request_body = mock_client.me.send_mail.post.call_args.args[0]
        assert request_body.message.reply_to is None

    async def test_send_rejects_invalid_reply_to(self):
        """send_message rejects malformed reply_to addresses."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError):
            await send_message(
                mock_client,
                to=["to@test.com"],
                subject="Test",
                body="Hello",
                reply_to=["not-an-email"],
                config=_CFG,
            )


class TestReply:
    async def test_reply_calls_reply_post(self):
        """reply calls reply.post() for single reply."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder
        result = await reply(mock_client, message_id="AAMkAG123=", body="Thanks!", config=_CFG)
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
            mock_client, message_id="AAMkAG123=", body="Thanks!", reply_all=True,
            config=_CFG,
        )
        assert result["reply_all"] is True
        msg_builder.reply_all.post.assert_called_once()

    async def test_reply_raises_read_only(self):
        """reply raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await reply(mock_client, message_id="AAMkAG123=", body="Thanks!", config=_CFG_RO)


class TestForward:
    async def test_forward_validates_to(self):
        """forward validates recipient addresses and calls forward.post()."""
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.forward.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder
        result = await forward(
            mock_client, message_id="AAMkAG123=", to=["a@b.com"], config=_CFG,
        )
        assert result["status"] == "forwarded"
        msg_builder.forward.post.assert_called_once()

    async def test_forward_raises_read_only(self):
        """forward raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await forward(
                mock_client, message_id="AAMkAG123=", to=["a@b.com"], config=_CFG_RO,
            )


class TestReplyContentType:
    """#41 shape found by the dead-parameter guard: is_html was accepted and ignored.

    Graph's reply action takes `comment` as plain text only. Markup has to ride
    on the action's `message` field as an ItemBody, so setting `comment` for an
    HTML reply silently strips it.
    """

    async def test_plain_reply_uses_comment(self):
        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        await reply(mock_client, message_id="AAMkAG123=", body="Thanks!", config=_CFG)

        posted = msg_builder.reply.post.call_args[0][0]
        assert posted.comment == "Thanks!"
        assert posted.message is None

    async def test_html_reply_rides_on_message_body(self):
        from msgraph.generated.models.body_type import BodyType

        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        await reply(
            mock_client,
            message_id="AAMkAG123=",
            body="<b>Thanks!</b>",
            is_html=True,
            config=_CFG,
        )

        posted = msg_builder.reply.post.call_args[0][0]
        assert posted.message.body.content == "<b>Thanks!</b>"
        assert posted.message.body.content_type is BodyType.Html
        # Setting both would duplicate the text in the sent reply.
        assert posted.comment is None

    async def test_html_reply_all_rides_on_message_body(self):
        from msgraph.generated.models.body_type import BodyType

        mock_client = MagicMock()
        msg_builder = MagicMock()
        msg_builder.reply_all.post = AsyncMock()
        mock_client.me.messages.by_message_id.return_value = msg_builder

        await reply(
            mock_client,
            message_id="AAMkAG123=",
            body="<i>All</i>",
            reply_all=True,
            is_html=True,
            config=_CFG,
        )

        posted = msg_builder.reply_all.post.call_args[0][0]
        assert posted.message.body.content_type is BodyType.Html
        assert posted.comment is None
