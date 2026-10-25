"""Tests for mail draft tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from msgraph.generated.users.item.messages.item.create_forward import (
    create_forward_post_request_body as fwd,
)
from msgraph.generated.users.item.messages.item.create_reply import (
    create_reply_post_request_body as rep,
)
from msgraph.generated.users.item.messages.item.create_reply_all import (
    create_reply_all_post_request_body as rep_all,
)

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_drafts import (
    create_draft,
    delete_draft,
    list_drafts,
    send_draft,
    update_draft,
)
from tests.test_write_payloads_reach_the_wire import assert_on_wire

_CFG = Config(client_id="test")

_WIRE_NAMES = {
    "to_recipients": "toRecipients",
    "cc_recipients": "ccRecipients",
    "bcc_recipients": "bccRecipients",
    "reply_to": "replyTo",
}
_CFG_RO = Config(client_id="test", read_only=True)


def _make_mock_message(**overrides):
    """Factory for mock Graph SDK message objects."""
    msg = MagicMock()
    msg.id = overrides.get("id", "AAMkAG123=")
    msg.subject = overrides.get("subject", "Draft Subject")
    msg.from_ = MagicMock()
    msg.from_.email_address = MagicMock()
    msg.from_.email_address.address = overrides.get("from_email", "me@test.com")
    msg.from_.email_address.name = overrides.get("from_name", "Me")
    msg.received_date_time = overrides.get("received", "2026-04-12T10:00:00Z")
    msg.is_read = overrides.get("is_read", True)
    msg.importance = MagicMock(value=overrides.get("importance", "normal"))
    msg.body_preview = overrides.get("body_preview", "Draft preview...")
    msg.has_attachments = overrides.get("has_attachments", False)
    msg.categories = overrides.get("categories", [])
    msg.flag = MagicMock()
    msg.flag.flag_status = MagicMock(value=overrides.get("flag", "notFlagged"))
    msg.conversation_id = overrides.get("conversation_id", "conv456")
    return msg


class TestListDrafts:
    async def test_list_drafts_returns_summaries(self):
        """list_drafts returns paginated message summaries from drafts folder."""
        mock_msg = _make_mock_message()
        response = MagicMock(value=[mock_msg], odata_next_link=None)

        messages_obj = MagicMock()
        messages_obj.get = AsyncMock(return_value=response)
        folder_obj = MagicMock()
        folder_obj.messages = messages_obj

        client = MagicMock()
        client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=folder_obj)

        result = await list_drafts(client, count=25)

        client.me.mail_folders.by_mail_folder_id.assert_called_with("drafts")
        assert result["count"] == 1
        assert result["messages"][0]["subject"] == "Draft Subject"
        assert result["messages"][0]["id"] == "AAMkAG123="
        assert result["next_cursor"] is None

    async def test_list_drafts_with_cursor(self):
        """list_drafts respects cursor for pagination."""
        response = MagicMock(value=[], odata_next_link=None)
        messages_obj = MagicMock()
        messages_obj.get = AsyncMock(return_value=response)
        folder_obj = MagicMock()
        folder_obj.messages = messages_obj

        client = MagicMock()
        client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=folder_obj)

        # Encode a cursor for skip=25
        from outlook_mcp.pagination import encode_cursor

        cursor = encode_cursor(25)
        result = await list_drafts(client, count=10, cursor=cursor)
        assert result["count"] == 0

    async def test_list_drafts_has_more(self):
        """list_drafts returns next_cursor when there are more results."""
        mock_msg = _make_mock_message()
        next_link = "https://graph.microsoft.com/v1.0/me/mailFolders/drafts/messages?$skip=25"
        response = MagicMock(value=[mock_msg], odata_next_link=next_link)

        messages_obj = MagicMock()
        messages_obj.get = AsyncMock(return_value=response)
        folder_obj = MagicMock()
        folder_obj.messages = messages_obj

        client = MagicMock()
        client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=folder_obj)

        result = await list_drafts(client, count=1)
        assert result["next_cursor"] is not None


class TestCreateDraft:
    async def test_create_draft_validates_emails(self):
        """create_draft validates to addresses and creates message in drafts."""
        created_msg = MagicMock()
        created_msg.id = "AAMkNewDraft="

        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=created_msg)

        result = await create_draft(
            client,
            to=["recipient@test.com"],
            subject="Test Draft",
            body="Hello",
            config=_CFG,
        )

        assert result["status"] == "created"
        assert result["draft_id"] == "AAMkNewDraft="
        client.me.messages.post.assert_called_once()

    async def test_create_draft_rejects_invalid_email(self):
        """create_draft raises ValueError for invalid email addresses."""
        client = MagicMock()
        with pytest.raises(ValueError):
            await create_draft(
                client,
                to=["not-an-email"],
                subject="Test",
                body="Hello",
                config=_CFG,
            )

    async def test_create_draft_with_cc_bcc(self):
        """create_draft passes cc and bcc recipients."""
        created_msg = MagicMock()
        created_msg.id = "AAMkDraftCC="

        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=created_msg)

        result = await create_draft(
            client,
            to=["to@test.com"],
            subject="CC Test",
            body="Hello",
            cc=["cc@test.com"],
            bcc=["bcc@test.com"],
            config=_CFG,
        )

        assert result["status"] == "created"
        assert result["draft_id"] == "AAMkDraftCC="

    async def test_create_draft_raises_read_only(self):
        """create_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await create_draft(
                client,
                to=["a@b.com"],
                subject="Test",
                body="Hello",
                config=_CFG_RO,
            )

    async def test_create_draft_sets_reply_to(self):
        """create_draft populates Message.reply_to when reply_to is provided."""
        created_msg = MagicMock()
        created_msg.id = "AAMkDraftReplyTo="

        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=created_msg)

        await create_draft(
            client,
            to=["to@test.com"],
            subject="Reply-To Test",
            body="Hello",
            reply_to=["alias@test.com"],
            config=_CFG,
        )

        posted_msg = client.me.messages.post.call_args.args[0]
        assert posted_msg.reply_to is not None
        assert [r.email_address.address for r in posted_msg.reply_to] == ["alias@test.com"]

    async def test_create_draft_rejects_invalid_reply_to(self):
        """create_draft rejects malformed reply_to addresses."""
        client = MagicMock()
        with pytest.raises(ValueError):
            await create_draft(
                client,
                to=["to@test.com"],
                subject="Test",
                body="Hello",
                reply_to=["bogus"],
                config=_CFG,
            )

    async def test_create_draft_sets_deferred_send_property(self):
        """create_draft attaches PR_DEFERRED_SEND_TIME extended property."""
        created_msg = MagicMock()
        created_msg.id = "AAMkDeferred="

        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=created_msg)

        result = await create_draft(
            client,
            to=["to@test.com"],
            subject="Scheduled",
            body="Hello",
            deferred_send_datetime="2026-05-06T08:00:00Z",
            config=_CFG,
        )

        posted_msg = client.me.messages.post.call_args.args[0]
        assert posted_msg.single_value_extended_properties is not None
        assert len(posted_msg.single_value_extended_properties) == 1
        prop = posted_msg.single_value_extended_properties[0]
        assert prop.id == "SystemTime 0x3FEF"
        assert prop.value == "2026-05-06T08:00:00Z"
        assert result["deferred_send_datetime"] == "2026-05-06T08:00:00Z"

    async def test_create_draft_normalizes_timezone_offset(self):
        """create_draft normalizes a TZ-offset datetime to UTC Z form."""
        created_msg = MagicMock()
        created_msg.id = "AAMkDeferredTZ="

        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=created_msg)

        # 10:00 in Helsinki (DST, +03:00) -> 07:00 UTC
        result = await create_draft(
            client,
            to=["to@test.com"],
            subject="Scheduled",
            body="Hello",
            deferred_send_datetime="2026-05-06T10:00:00+03:00",
            config=_CFG,
        )

        assert result["deferred_send_datetime"] == "2026-05-06T07:00:00Z"
        prop = client.me.messages.post.call_args.args[0].single_value_extended_properties[0]
        assert prop.value == "2026-05-06T07:00:00Z"

    async def test_create_draft_rejects_malformed_deferred_send(self):
        """create_draft rejects non-ISO deferred_send_datetime values."""
        client = MagicMock()
        with pytest.raises(ValueError):
            await create_draft(
                client,
                to=["to@test.com"],
                subject="Test",
                body="Hello",
                deferred_send_datetime="tomorrow at 8",
                config=_CFG,
            )


class TestUpdateDraft:
    async def test_update_draft_patches_partial(self):
        """update_draft sends PATCH with only provided fields."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await update_draft(
            client,
            draft_id="AAMkAG123=",
            subject="Updated Subject",
            config=_CFG,
        )

        assert result["status"] == "updated"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.patch.assert_called_once()

    async def test_update_draft_body_defaults_to_text(self):
        """update_draft sends body as Text by default."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        await update_draft(
            client,
            draft_id="AAMkAG123=",
            body="plain text",
            config=_CFG,
        )

        sent_msg = msg_builder.patch.call_args[0][0]
        from msgraph.generated.models.body_type import BodyType

        assert sent_msg.body.content == "plain text"
        assert sent_msg.body.content_type == BodyType.Text

    async def test_update_draft_body_html_when_is_html(self):
        """update_draft sends body as Html when is_html=True (required for
        drafts originally composed as HTML in the Outlook UI — Text PATCH on
        an HTML draft triggers MapiSetProperties / ErrorAccessDenied)."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        await update_draft(
            client,
            draft_id="AAMkAG123=",
            body="<p>hello</p>",
            is_html=True,
            config=_CFG,
        )

        sent_msg = msg_builder.patch.call_args[0][0]
        from msgraph.generated.models.body_type import BodyType

        assert sent_msg.body.content == "<p>hello</p>"
        assert sent_msg.body.content_type == BodyType.Html

    async def test_update_draft_validates_graph_id(self):
        """update_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await update_draft(
                client,
                draft_id="bad id with spaces!",
                subject="Test",
                config=_CFG,
            )

    async def test_update_draft_raises_read_only(self):
        """update_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await update_draft(
                client,
                draft_id="AAMkAG123=",
                subject="Test",
                config=_CFG_RO,
            )

    async def test_update_draft_patches_reply_to(self):
        """update_draft sets reply_to on the PATCH payload when provided."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        await update_draft(
            client,
            draft_id="AAMkAG123=",
            reply_to=["alias@test.com", "team@test.com"],
            config=_CFG,
        )

        patched_msg = msg_builder.patch.call_args.args[0]
        assert patched_msg.reply_to is not None
        assert [r.email_address.address for r in patched_msg.reply_to] == [
            "alias@test.com",
            "team@test.com",
        ]

    async def test_update_draft_clears_reply_to_with_empty_list(self):
        """update_draft with reply_to=[] clears the field on the draft."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        await update_draft(
            client,
            draft_id="AAMkAG123=",
            reply_to=[],
            config=_CFG,
        )

        patched_msg = msg_builder.patch.call_args.args[0]
        assert patched_msg.reply_to == []

    async def test_update_draft_rejects_invalid_reply_to(self):
        """update_draft rejects malformed reply_to addresses."""
        client = MagicMock()
        with pytest.raises(ValueError):
            await update_draft(
                client,
                draft_id="AAMkAG123=",
                reply_to=["bogus"],
                config=_CFG,
            )

    async def test_update_draft_sets_deferred_send_property(self):
        """update_draft attaches PR_DEFERRED_SEND_TIME on the PATCH payload."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await update_draft(
            client,
            draft_id="AAMkAG123=",
            deferred_send_datetime="2026-05-06T08:00:00Z",
            config=_CFG,
        )

        patched_msg = msg_builder.patch.call_args.args[0]
        props = patched_msg.single_value_extended_properties
        assert props is not None and len(props) == 1
        assert props[0].id == "SystemTime 0x3FEF"
        assert props[0].value == "2026-05-06T08:00:00Z"
        assert result["deferred_send_datetime"] == "2026-05-06T08:00:00Z"

    async def test_update_draft_clears_deferred_send_with_empty_string(self):
        """update_draft with deferred_send_datetime='' clears the property."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.patch = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await update_draft(
            client,
            draft_id="AAMkAG123=",
            deferred_send_datetime="",
            config=_CFG,
        )

        patched_msg = msg_builder.patch.call_args.args[0]
        props = patched_msg.single_value_extended_properties
        assert props is not None and len(props) == 1
        assert props[0].id == "SystemTime 0x3FEF"
        assert props[0].value == ""
        assert result["deferred_send_datetime"] == ""

    async def test_update_draft_rejects_malformed_deferred_send(self):
        """update_draft rejects non-ISO deferred_send_datetime values."""
        client = MagicMock()
        with pytest.raises(ValueError):
            await update_draft(
                client,
                draft_id="AAMkAG123=",
                deferred_send_datetime="not-a-date",
                config=_CFG,
            )


class TestSendDraft:
    async def test_send_draft_calls_send_endpoint(self):
        """send_draft POSTs to /me/messages/{id}/send."""
        msg_builder = MagicMock()
        msg_builder.send.post = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await send_draft(client, draft_id="AAMkAG123=", config=_CFG)

        assert result["status"] == "sent"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.send.post.assert_called_once()

    async def test_send_draft_validates_graph_id(self):
        """send_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await send_draft(client, draft_id="bad id!!!", config=_CFG)

    async def test_send_draft_raises_read_only(self):
        """send_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await send_draft(client, draft_id="AAMkAG123=", config=_CFG_RO)


class TestDeleteDraft:
    async def test_delete_draft_calls_delete(self):
        """delete_draft DELETEs /me/messages/{id}."""
        msg_builder = MagicMock()
        msg_builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        msg_builder.delete = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await delete_draft(client, draft_id="AAMkAG123=", config=_CFG)

        assert result["status"] == "deleted"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.delete.assert_called_once()

    async def test_delete_draft_validates_graph_id(self):
        """delete_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await delete_draft(client, draft_id="bad id!!!", config=_CFG)

    async def test_delete_draft_raises_read_only(self):
        """delete_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_draft(client, draft_id="AAMkAG123=", config=_CFG_RO)


# ── The draft tools only touch drafts ────────────────────────────────────────
# A draft is addressed by message id, and every message in the mailbox has one.
# Without a check, `outlook_delete_draft` was a permanent delete for any message
# under the `mail_drafts` category — the one the README calls "drafts only".


def _message_builder(*, is_draft):
    """A message builder whose GET answers `$select=isDraft` the way Graph would."""
    builder = MagicMock()
    builder.get = AsyncMock(return_value=MagicMock(is_draft=is_draft))
    builder.patch = AsyncMock()
    builder.delete = AsyncMock()
    return builder


def _client_for(builder):
    client = MagicMock()
    client.me.messages.by_message_id = MagicMock(return_value=builder)
    return client


class TestDraftToolsRefuseNonDrafts:
    async def test_delete_draft_refuses_a_received_message(self):
        builder = _message_builder(is_draft=False)

        with pytest.raises(ValueError, match="not a draft"):
            await delete_draft(_client_for(builder), draft_id="AAMkInboxMsg=", config=_CFG)

        builder.delete.assert_not_called()

    async def test_update_draft_refuses_a_received_message(self):
        builder = _message_builder(is_draft=False)

        with pytest.raises(ValueError, match="not a draft"):
            await update_draft(
                _client_for(builder), draft_id="AAMkInboxMsg=", subject="rewritten", config=_CFG
            )

        builder.patch.assert_not_called()

    async def test_a_message_graph_does_not_call_a_draft_is_refused(self):
        """Fail closed: only an explicit `isDraft: true` lets the delete through."""
        builder = _message_builder(is_draft=None)

        with pytest.raises(ValueError, match="not a draft"):
            await delete_draft(_client_for(builder), draft_id="AAMkAG123=", config=_CFG)

        builder.delete.assert_not_called()

    async def test_the_check_asks_graph_for_isdraft_and_nothing_else(self):
        builder = _message_builder(is_draft=True)

        await delete_draft(_client_for(builder), draft_id="AAMkAG123=", config=_CFG)

        request_config = builder.get.call_args.kwargs["request_configuration"]
        assert request_config.query_parameters.select == ["isDraft"]
        builder.delete.assert_called_once()

    async def test_read_only_still_refuses_before_any_graph_call(self):
        builder = _message_builder(is_draft=True)

        with pytest.raises(ReadOnlyError):
            await delete_draft(_client_for(builder), draft_id="AAMkAG123=", config=_CFG_RO)

        builder.get.assert_not_called()


# ── Drafts in a thread: reply, reply-all, forward (#143) ─────────────────────
# Graph's createReply / createReplyAll / createForward make the draft in the
# thread with the original quoted below. Nothing in this section may send.

# What a createReply draft's body looks like: the quoted original sits below an
# <hr> and Outlook's header block, and the reply goes above both.
_QUOTED_DRAFT_HTML = (
    '<html><head><meta charset="utf-8"></head><body dir="ltr">'
    '<div id="appendonsend"></div><hr style="display:inline-block;width:98%">'
    '<div id="divRplyFwdMsg"><b>From:</b> Sender &lt;sender@example.com&gt;</div>'
    "<div><p>The <b>original</b> message.</p></div></body></html>"
)


def _recipient(address):
    return MagicMock(email_address=MagicMock(address=address))


def _thread_draft(*, content=_QUOTED_DRAFT_HTML, html=True, to=("sender@example.com",), cc=()):
    """The Message Graph returns from createReply/createReplyAll/createForward."""
    from msgraph.generated.models.body_type import BodyType

    draft = MagicMock()
    draft.id = "AAMkThreadDraft="
    draft.body = MagicMock(content=content, content_type=BodyType.Html if html else BodyType.Text)
    draft.to_recipients = [_recipient(a) for a in to]
    draft.cc_recipients = [_recipient(a) for a in cc]
    return draft


def _thread_client(draft=None):
    """A client whose source message offers the three actions, and whose new
    draft can be patched and deleted. Returns (client, source, new_draft)."""
    draft = draft or _thread_draft()
    source = MagicMock()
    for action in ("create_reply", "create_reply_all", "create_forward"):
        getattr(source, action).post = AsyncMock(return_value=draft)
    new_draft = MagicMock()
    new_draft.patch = AsyncMock()
    new_draft.delete = AsyncMock()
    builders = {"AAMkOriginal=": source, draft.id: new_draft}
    client = MagicMock()
    client.me.messages.by_message_id = MagicMock(side_effect=lambda i: builders[i])
    client.me.messages.post = AsyncMock()
    return client, source, new_draft


class TestCreateDraftInThread:
    async def test_reply_creates_the_draft_with_create_reply_then_patches_the_body(self):
        from msgraph.generated.models.body_type import BodyType

        client, source, new_draft = _thread_client()

        result = await create_draft(
            client,
            to=None,
            subject=None,
            body="Thanks, see you then.",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        sent = source.create_reply.post.call_args.args[0]
        assert isinstance(sent, rep.CreateReplyPostRequestBody)
        # Neither channel: `comment` loses the text's line breaks on an HTML
        # message, and `message.body` replaces the quoted original.
        assert sent.comment is None
        assert sent.message.body is None
        source.create_reply_all.post.assert_not_called()
        source.create_forward.post.assert_not_called()
        client.me.messages.post.assert_not_called()
        new_draft.patch.assert_called_once()
        patched = new_draft.patch.call_args.args[0]
        assert patched.body.content_type == BodyType.Html
        assert patched.body.content == _QUOTED_DRAFT_HTML.replace(
            '<body dir="ltr">', '<body dir="ltr"><p>Thanks, see you then.</p>'
        )
        new_draft.delete.assert_not_called()
        assert result == {
            "status": "created",
            "draft_id": "AAMkThreadDraft=",
            "to": ["sender@example.com"],
            "cc": [],
        }

    async def test_reply_all_uses_create_reply_all(self):
        draft = _thread_draft(to=("sender@example.com",), cc=("other@example.com",))
        client, source, _ = _thread_client(draft)

        result = await create_draft(
            client,
            to=None,
            subject=None,
            body="Agreed.",
            reply_to_message_id="AAMkOriginal=",
            reply_all=True,
            config=_CFG,
        )

        source.create_reply_all.post.assert_called_once()
        source.create_reply.post.assert_not_called()
        assert isinstance(
            source.create_reply_all.post.call_args.args[0], rep_all.CreateReplyAllPostRequestBody
        )
        assert source.create_reply_all.post.call_args.args[0].comment is None
        assert result["cc"] == ["other@example.com"]

    @pytest.mark.parametrize(
        ("field", "attribute"),
        [
            ("to", "to_recipients"),
            ("cc", "cc_recipients"),
            ("bcc", "bcc_recipients"),
            ("reply_to", "reply_to"),
        ],
    )
    async def test_an_empty_recipient_list_clears_what_graph_prefilled(self, field, attribute):
        """`cc=[]` on a reply-all must remove them, not leave Graph's list alone.

        Graph prefills to/cc on a reply-all draft, so treating an empty list as
        "not mentioned" kept recipients the caller had explicitly removed — and
        sending that draft mailed them.
        """
        client, source, _ = _thread_client(_thread_draft())

        kwargs = {"to": None, field: []} if field != "to" else {"to": []}
        await create_draft(
            client,
            subject=None,
            body="Agreed.",
            reply_to_message_id="AAMkOriginal=",
            reply_all=True,
            config=_CFG,
            **kwargs,
        )

        sent = source.create_reply_all.post.call_args.args[0].message
        assert getattr(sent, attribute) == []
        # On the wire too: Kiota does emit an empty collection, so the clear is what
        # Graph receives rather than an attribute that never serialized.
        assert_on_wire(sent, f'"{_WIRE_NAMES[attribute]}": []')

    @pytest.mark.parametrize("field", ["to", "cc", "bcc", "reply_to"])
    async def test_an_omitted_recipient_list_leaves_graphs_alone(self, field):
        """Omitted is not empty: leaving the field out keeps the thread's recipients."""
        client, source, _ = _thread_client(_thread_draft())

        await create_draft(
            client,
            to=None,
            subject=None,
            body="Agreed.",
            reply_to_message_id="AAMkOriginal=",
            reply_all=True,
            config=_CFG,
        )

        sent = source.create_reply_all.post.call_args.args[0].message
        assert not getattr(sent, "cc_recipients", None)
        assert not getattr(sent, "bcc_recipients", None)
        assert not getattr(sent, "reply_to", None)

    async def test_forward_uses_create_forward_with_the_recipients(self):
        client, source, new_draft = _thread_client(_thread_draft(to=("fwd@example.com",)))

        result = await create_draft(
            client,
            to=["fwd@example.com"],
            subject=None,
            body="FYI",
            forward_message_id="AAMkOriginal=",
            config=_CFG,
        )

        sent = source.create_forward.post.call_args.args[0]
        assert isinstance(sent, fwd.CreateForwardPostRequestBody)
        assert [r.email_address.address for r in sent.message.to_recipients] == ["fwd@example.com"]
        assert sent.comment is None
        assert "<p>FYI</p>" in new_draft.patch.call_args.args[0].body.content
        source.create_reply.post.assert_not_called()
        assert result["to"] == ["fwd@example.com"]

    async def test_a_forward_without_recipients_is_refused_before_any_call(self):
        client, source, _ = _thread_client()

        for to in (None, []):
            with pytest.raises(ValueError, match="forward needs `to`"):
                await create_draft(
                    client,
                    to=to,
                    subject=None,
                    body="FYI",
                    forward_message_id="AAMkOriginal=",
                    config=_CFG,
                )

        source.create_forward.post.assert_not_called()

    async def test_reply_and_forward_together_are_refused(self):
        client, source, _ = _thread_client()

        with pytest.raises(ValueError, match="not both"):
            await create_draft(
                client,
                to=["fwd@example.com"],
                subject=None,
                body="x",
                reply_to_message_id="AAMkOriginal=",
                forward_message_id="AAMkOriginal=",
                config=_CFG,
            )

        source.create_reply.post.assert_not_called()
        source.create_forward.post.assert_not_called()

    async def test_reply_all_without_a_message_to_reply_to_is_refused(self):
        client, source, _ = _thread_client()

        for forward_id in (None, "AAMkOriginal="):
            with pytest.raises(ValueError, match="reply_all"):
                await create_draft(
                    client,
                    to=["a@example.com"],
                    subject="s",
                    body="x",
                    reply_all=True,
                    forward_message_id=forward_id,
                    config=_CFG,
                )

        client.me.messages.post.assert_not_called()
        source.create_forward.post.assert_not_called()

    async def test_a_new_draft_still_needs_to_and_subject(self):
        client, _, _ = _thread_client()

        with pytest.raises(ValueError, match="new draft needs `to` and `subject`"):
            await create_draft(client, to=None, subject="s", body="x", config=_CFG)
        with pytest.raises(ValueError, match="new draft needs `to` and `subject`"):
            await create_draft(client, to=["a@example.com"], subject=None, body="x", config=_CFG)

        client.me.messages.post.assert_not_called()

    async def test_a_new_draft_with_an_empty_to_list_is_still_created(self):
        """As before: `to=[]` was always a valid new draft, just one with no recipients."""
        client, _, _ = _thread_client()
        client.me.messages.post = AsyncMock(return_value=MagicMock(id="AAMkNew="))

        result = await create_draft(client, to=[], subject="s", body="x", config=_CFG)

        assert result == {"status": "created", "draft_id": "AAMkNew="}

    async def test_the_message_id_is_validated_before_any_call(self):
        client, source, _ = _thread_client()

        with pytest.raises(ValueError, match="invalid characters"):
            await create_draft(
                client,
                to=None,
                subject=None,
                body="x",
                reply_to_message_id="not an id",
                config=_CFG,
            )

        source.create_reply.post.assert_not_called()

    async def test_read_only_refuses_before_any_call(self):
        client, source, _ = _thread_client()

        with pytest.raises(ReadOnlyError):
            await create_draft(
                client,
                to=None,
                subject=None,
                body="x",
                reply_to_message_id="AAMkOriginal=",
                config=_CFG_RO,
            )

        source.create_reply.post.assert_not_called()

    async def test_fields_passed_with_a_reply_go_on_its_message(self):
        client, source, _ = _thread_client()

        await create_draft(
            client,
            to=["to@example.com"],
            subject="New subject",
            body="x",
            cc=["cc@example.com"],
            bcc=["bcc@example.com"],
            importance="high",
            reply_to=["alias@example.com"],
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        message = source.create_reply.post.call_args.args[0].message
        assert message.subject == "New subject"
        assert [r.email_address.address for r in message.to_recipients] == ["to@example.com"]
        assert [r.email_address.address for r in message.cc_recipients] == ["cc@example.com"]
        assert [r.email_address.address for r in message.bcc_recipients] == ["bcc@example.com"]
        assert [r.email_address.address for r in message.reply_to] == ["alias@example.com"]
        assert message.importance.value == "high"

    async def test_an_empty_to_clears_the_recipients_graph_would_have_chosen(self):
        """`to=[]` is a request to clear, and omitting the list is the way to accept
        Graph's choice. This used to treat `[]` as "not mentioned", which is how a
        reply-all kept recipients the caller had removed.
        """
        client, source, _ = _thread_client()

        await create_draft(
            client,
            to=[],
            subject=None,
            body="x",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        message = source.create_reply.post.call_args.args[0].message
        assert message.subject is None
        assert message.to_recipients == []
        assert message.cc_recipients is None

    async def test_an_omitted_to_leaves_the_recipients_to_graph(self):
        client, source, _ = _thread_client()

        await create_draft(
            client,
            to=None,
            subject=None,
            body="x",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        message = source.create_reply.post.call_args.args[0].message
        assert message.to_recipients is None


class TestHtmlReplyDraft:
    async def test_html_goes_above_the_quote_by_patching_the_draft(self):
        from msgraph.generated.models.body_type import BodyType

        client, source, new_draft = _thread_client()

        result = await create_draft(
            client,
            to=None,
            subject=None,
            body="<p>Thanks, <b>see you then</b>.</p>",
            is_html=True,
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        sent = source.create_reply.post.call_args.args[0]
        # comment is plain text only, and Graph refuses comment + message.body.
        assert sent.comment is None
        assert sent.message.body is None
        patched = new_draft.patch.call_args.args[0]
        assert patched.body.content_type == BodyType.Html
        assert patched.body.content == _QUOTED_DRAFT_HTML.replace(
            '<body dir="ltr">', '<body dir="ltr"><p>Thanks, <b>see you then</b>.</p>'
        )
        new_draft.delete.assert_not_called()
        assert result["draft_id"] == "AAMkThreadDraft="

    async def test_html_forward_is_patched_the_same_way(self):
        client, source, new_draft = _thread_client()

        await create_draft(
            client,
            to=["fwd@example.com"],
            subject=None,
            body="<p>FYI</p>",
            is_html=True,
            forward_message_id="AAMkOriginal=",
            config=_CFG,
        )

        assert source.create_forward.post.call_args.args[0].comment is None
        content = new_draft.patch.call_args.args[0].body.content
        assert content.startswith('<html><head><meta charset="utf-8"></head><body dir="ltr"><p>FYI')


class TestPlainTextReplyDraft:
    """Plain text in a reply or forward draft keeps its paragraphs, line breaks and
    lists. Graph's `comment` would join its lines with spaces on an HTML message."""

    _TEXT = (
        "First paragraph line one.\nFirst paragraph line two.\n\nSecond paragraph.\n\n"
        "- bullet one\n- bullet two"
    )

    async def test_on_an_html_original_it_is_patched_in_as_html_above_the_quote(self):
        from msgraph.generated.models.body_type import BodyType

        client, source, new_draft = _thread_client()

        await create_draft(
            client,
            to=None,
            subject=None,
            body=self._TEXT,
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        assert source.create_reply.post.call_args.args[0].comment is None
        patched = new_draft.patch.call_args.args[0]
        assert patched.body.content_type == BodyType.Html
        assert patched.body.content == _QUOTED_DRAFT_HTML.replace(
            '<body dir="ltr">',
            '<body dir="ltr"><p>First paragraph line one.<br>First paragraph line two.</p>'
            "<p>Second paragraph.</p><ul><li>bullet one</li><li>bullet two</li></ul>",
        )

    async def test_on_a_plain_text_original_it_stays_text_above_the_quote(self):
        from msgraph.generated.models.body_type import BodyType

        quote = "\r\n________________\r\nFrom: Sender <sender@example.com>\r\n\r\nOriginal.\r\n"
        client, _, new_draft = _thread_client(_thread_draft(content=quote, html=False))

        await create_draft(
            client,
            to=None,
            subject=None,
            body=self._TEXT,
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        patched = new_draft.patch.call_args.args[0]
        assert patched.body.content_type == BodyType.Text
        assert patched.body.content == self._TEXT.replace("\n", "\r\n") + "\r\n" + quote

    async def test_an_empty_body_patches_nothing(self):
        client, source, new_draft = _thread_client()

        await create_draft(
            client,
            to=None,
            subject=None,
            body="",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        source.create_reply.post.assert_called_once()
        new_draft.patch.assert_not_called()


class TestDeferredReplyDraft:
    async def test_the_deferred_send_time_is_patched_onto_the_reply(self):
        client, source, new_draft = _thread_client()

        result = await create_draft(
            client,
            to=None,
            subject=None,
            body="x",
            deferred_send_datetime="2026-12-01T09:00:00Z",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        # Not on createReply's `message`: that action does not document extended
        # properties, and the PATCH path is the one update_draft already uses.
        assert (
            source.create_reply.post.call_args.args[0].message.single_value_extended_properties
            is None
        )
        new_draft.patch.assert_called_once()
        patched = new_draft.patch.call_args.args[0]
        assert "<p>x</p>" in patched.body.content
        (prop,) = patched.single_value_extended_properties
        assert prop.id == "SystemTime 0x3FEF"
        assert prop.value == "2026-12-01T09:00:00Z"
        assert result["deferred_send_datetime"] == "2026-12-01T09:00:00Z"

    async def test_html_and_deferred_share_one_patch(self):
        client, _, new_draft = _thread_client()

        await create_draft(
            client,
            to=None,
            subject=None,
            body="<p>x</p>",
            is_html=True,
            deferred_send_datetime="2026-12-01T09:00:00Z",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        new_draft.patch.assert_called_once()
        patched = new_draft.patch.call_args.args[0]
        assert "<p>x</p>" in patched.body.content
        assert patched.single_value_extended_properties[0].value == "2026-12-01T09:00:00Z"

    async def test_a_failed_patch_deletes_the_new_draft_and_raises(self):
        client, _, new_draft = _thread_client()
        new_draft.patch = AsyncMock(side_effect=RuntimeError("patch refused"))

        with pytest.raises(RuntimeError, match="patch refused"):
            await create_draft(
                client,
                to=None,
                subject=None,
                body="x",
                deferred_send_datetime="2026-12-01T09:00:00Z",
                reply_to_message_id="AAMkOriginal=",
                config=_CFG,
            )

        new_draft.delete.assert_called_once()

    async def test_the_patch_error_is_reported_even_if_the_delete_fails(self):
        client, _, new_draft = _thread_client()
        new_draft.patch = AsyncMock(side_effect=RuntimeError("patch refused"))
        new_draft.delete = AsyncMock(side_effect=RuntimeError("delete refused"))

        with pytest.raises(RuntimeError, match="patch refused"):
            await create_draft(
                client,
                to=None,
                subject=None,
                body="<p>x</p>",
                is_html=True,
                reply_to_message_id="AAMkOriginal=",
                config=_CFG,
            )

        new_draft.delete.assert_called_once()
