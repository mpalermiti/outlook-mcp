"""Tests for mail draft tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_drafts import (
    create_draft,
    delete_draft,
    list_drafts,
    send_draft,
    update_draft,
)

_CFG = Config(client_id="test")
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
