"""Tests for mail write tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import GraphAPIError, ReadOnlyError
from outlook_mcp.tools.mail_write import forward, reply, send_message

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


class TestSendMessage:
    async def test_send_validates_emails(self):
        """send_message validates to addresses and calls send_mail.post()."""
        mock_client = AsyncMock()
        mock_client.me.send_mail.post = AsyncMock()
        result = await send_message(
            mock_client,
            to=["valid@test.com"],
            subject="Test",
            body="Hello",
            config=_CFG,
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()

    async def test_send_rejects_invalid_email(self):
        """send_message rejects invalid email addresses."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError):
            await send_message(
                mock_client,
                to=["not-an-email"],
                subject="Test",
                body="Hello",
                config=_CFG,
            )

    async def test_send_raises_read_only(self):
        """send_message raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await send_message(
                mock_client,
                to=["a@b.com"],
                subject="Test",
                body="Hello",
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
            mock_client,
            to=["to@test.com"],
            subject="Test",
            body="Hello",
            config=_CFG,
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


def _thread_client(content="<html><body><div>quoted</div></body></html>", html=True):
    """A client whose message M1 offers createReply/createReplyAll/createForward, and
    whose new draft R1 can be patched, sent and deleted. Returns (client, source, draft)."""
    from msgraph.generated.models.body_type import BodyType

    created = MagicMock(id="R1")
    created.body = MagicMock(content=content, content_type=BodyType.Html if html else BodyType.Text)
    source = MagicMock()
    for action in ("create_reply", "create_reply_all", "create_forward"):
        getattr(source, action).post = AsyncMock(return_value=created)
    draft = MagicMock()
    draft.patch = AsyncMock()
    draft.delete = AsyncMock()
    draft.send.post = AsyncMock()
    builders = {"AAMkAG123=": source, "R1": draft}
    client = MagicMock()
    client.me.messages.by_message_id = MagicMock(side_effect=lambda i: builders[i])
    return client, source, draft


class TestReply:
    async def test_reply_creates_a_reply_draft_patches_it_and_sends_it(self):
        client, source, draft = _thread_client()
        result = await reply(client, message_id="AAMkAG123=", body="Thanks!", config=_CFG)
        assert result["status"] == "replied"
        assert result["reply_all"] is False
        source.create_reply.post.assert_called_once()
        source.create_reply_all.post.assert_not_called()
        draft.patch.assert_called_once()
        draft.send.post.assert_called_once()
        draft.delete.assert_not_called()

    async def test_reply_all_creates_a_reply_all_draft(self):
        client, source, draft = _thread_client()
        result = await reply(
            client,
            message_id="AAMkAG123=",
            body="Thanks!",
            reply_all=True,
            config=_CFG,
        )
        assert result["reply_all"] is True
        source.create_reply_all.post.assert_called_once()
        source.create_reply.post.assert_not_called()
        draft.send.post.assert_called_once()

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
            mock_client,
            message_id="AAMkAG123=",
            to=["a@b.com"],
            config=_CFG,
        )
        assert result["status"] == "forwarded"
        msg_builder.forward.post.assert_called_once()

    async def test_forward_raises_read_only(self):
        """forward raises ReadOnlyError in read-only mode."""
        mock_client = AsyncMock()
        with pytest.raises(ReadOnlyError):
            await forward(
                mock_client,
                message_id="AAMkAG123=",
                to=["a@b.com"],
                config=_CFG_RO,
            )


class TestReplyContentType:
    """Where the reply's text goes. Graph's `comment` joins a multi-line text's lines
    with spaces on an HTML message, and `message.body` replaces the quoted original
    (checked against an Exchange Online mailbox), so neither is used: the draft is
    created empty and the text is patched in above the quote.

    #41 shape found by the dead-parameter guard first: is_html was accepted and
    ignored until 1.18.0.
    """

    async def test_plain_reply_is_converted_and_patched_above_the_quote(self):
        from msgraph.generated.models.body_type import BodyType

        client, source, draft = _thread_client()

        await reply(client, message_id="AAMkAG123=", body="Thanks!\nSam", config=_CFG)

        created = source.create_reply.post.call_args[0][0]
        assert created.comment is None
        assert created.message is None
        patched = draft.patch.call_args[0][0]
        assert patched.body.content_type is BodyType.Html
        assert (
            patched.body.content
            == "<html><body><p>Thanks!<br>Sam</p><div>quoted</div></body></html>"
        )

    async def test_html_reply_is_patched_above_the_quote_as_written(self):
        from msgraph.generated.models.body_type import BodyType

        client, source, draft = _thread_client()

        await reply(
            client,
            message_id="AAMkAG123=",
            body="<b>Thanks!</b>",
            is_html=True,
            config=_CFG,
        )

        created = source.create_reply.post.call_args[0][0]
        # message.body would have replaced the quote; comment cannot carry markup.
        assert created.message is None
        assert created.comment is None
        patched = draft.patch.call_args[0][0]
        assert patched.body.content == "<html><body><b>Thanks!</b><div>quoted</div></body></html>"
        assert patched.body.content_type is BodyType.Html

    async def test_a_plain_text_thread_gets_a_plain_text_reply(self):
        from msgraph.generated.models.body_type import BodyType

        client, _, draft = _thread_client(content="\r\n> quoted\r\n", html=False)

        await reply(client, message_id="AAMkAG123=", body="Thanks!\nSam", config=_CFG)

        patched = draft.patch.call_args[0][0]
        assert patched.body.content_type is BodyType.Text
        assert patched.body.content == "Thanks!\r\nSam\r\n\r\n> quoted\r\n"

    async def test_an_empty_body_sends_the_draft_unpatched(self):
        client, _, draft = _thread_client()

        await reply(client, message_id="AAMkAG123=", body="", config=_CFG)

        draft.patch.assert_not_called()
        draft.send.post.assert_called_once()

    async def test_a_failed_patch_deletes_the_draft_and_sends_nothing(self):
        client, _, draft = _thread_client()
        draft.patch = AsyncMock(side_effect=RuntimeError("patch refused"))
        draft.delete = AsyncMock(side_effect=RuntimeError("delete refused"))

        with pytest.raises(RuntimeError, match="patch refused"):
            await reply(client, message_id="AAMkAG123=", body="x", config=_CFG)

        draft.delete.assert_called_once()
        draft.send.post.assert_not_called()

    async def test_a_send_that_fails_outside_graph_leaves_the_draft(self):
        """A dropped connection says nothing about whether the send happened, so the
        agent is told what a 5xx tells it, with the draft's id, not handed a crash."""
        client, _, draft = _thread_client()
        draft.send.post = AsyncMock(side_effect=OSError("connection reset"))

        with pytest.raises(GraphAPIError) as caught:
            await reply(client, message_id="AAMkAG123=", body="x", config=_CFG)

        draft.delete.assert_not_called()
        text = str(caught.value)
        assert "OSError" in text and "connection reset" not in text, text
        assert "Check Sent Items" in text and "still in Drafts" in text, text
        assert "The draft's id is R1." in text, text


class TestForwardComment:
    async def test_a_comment_goes_through_a_forward_draft_with_the_recipients(self):
        client, source, draft = _thread_client()

        result = await forward(
            client, message_id="AAMkAG123=", to=["a@b.com"], comment="FYI\nSam", config=_CFG
        )

        assert result == {"status": "forwarded"}
        created = source.create_forward.post.call_args[0][0]
        assert [r.email_address.address for r in created.message.to_recipients] == ["a@b.com"]
        assert created.comment is None
        assert created.to_recipients is None
        assert "<p>FYI<br>Sam</p>" in draft.patch.call_args[0][0].body.content
        draft.send.post.assert_called_once()
        source.forward.post.assert_not_called()

    async def test_recipients_are_validated_before_any_call(self):
        client, source, _ = _thread_client()

        with pytest.raises(ValueError):
            await forward(client, message_id="AAMkAG123=", to=["nope"], comment="x", config=_CFG)

        source.create_forward.post.assert_not_called()
