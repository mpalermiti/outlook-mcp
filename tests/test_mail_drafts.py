"""Tests for mail draft tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_drafts import (
    create_draft,
    delete_draft,
    list_drafts,
    send_draft,
    update_draft,
)


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
                read_only=True,
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
        )

        assert result["status"] == "updated"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.patch.assert_called_once()

    async def test_update_draft_validates_graph_id(self):
        """update_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await update_draft(
                client,
                draft_id="bad id with spaces!",
                subject="Test",
            )

    async def test_update_draft_raises_read_only(self):
        """update_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await update_draft(
                client,
                draft_id="AAMkAG123=",
                subject="Test",
                read_only=True,
            )


class TestSendDraft:
    async def test_send_draft_calls_send_endpoint(self):
        """send_draft POSTs to /me/messages/{id}/send."""
        msg_builder = MagicMock()
        msg_builder.send.post = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await send_draft(client, draft_id="AAMkAG123=")

        assert result["status"] == "sent"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.send.post.assert_called_once()

    async def test_send_draft_validates_graph_id(self):
        """send_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await send_draft(client, draft_id="bad id!!!")

    async def test_send_draft_raises_read_only(self):
        """send_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await send_draft(client, draft_id="AAMkAG123=", read_only=True)


class TestDeleteDraft:
    async def test_delete_draft_calls_delete(self):
        """delete_draft DELETEs /me/messages/{id}."""
        msg_builder = MagicMock()
        msg_builder.delete = AsyncMock()

        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(return_value=msg_builder)

        result = await delete_draft(client, draft_id="AAMkAG123=")

        assert result["status"] == "deleted"
        assert result["draft_id"] == "AAMkAG123="
        msg_builder.delete.assert_called_once()

    async def test_delete_draft_validates_graph_id(self):
        """delete_draft rejects invalid draft IDs."""
        client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await delete_draft(client, draft_id="bad id!!!")

    async def test_delete_draft_raises_read_only(self):
        """delete_draft raises ReadOnlyError in read-only mode."""
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_draft(client, draft_id="AAMkAG123=", read_only=True)
