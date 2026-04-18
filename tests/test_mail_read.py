"""Tests for mail read tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools.mail_read import (
    _format_message_summary,
    list_folders,
    list_inbox,
    read_message,
    search_mail,
)


def _make_mock_message(**overrides):
    """Factory for mock Graph SDK message objects."""
    msg = MagicMock()
    msg.id = overrides.get("id", "AAMkAG123=")
    msg.subject = overrides.get("subject", "Test Subject")
    msg.from_ = MagicMock()
    msg.from_.email_address = MagicMock()
    msg.from_.email_address.address = overrides.get("from_email", "sender@test.com")
    msg.from_.email_address.name = overrides.get("from_name", "Sender")
    msg.received_date_time = overrides.get("received", "2026-04-12T10:00:00Z")
    msg.is_read = overrides.get("is_read", False)
    msg.importance = MagicMock(value=overrides.get("importance", "normal"))
    msg.body_preview = overrides.get("body_preview", "Preview text...")
    msg.has_attachments = overrides.get("has_attachments", False)
    msg.categories = overrides.get("categories", [])
    msg.flag = MagicMock()
    msg.flag.flag_status = MagicMock(value=overrides.get("flag", "notFlagged"))
    msg.conversation_id = overrides.get("conversation_id", "conv123")
    msg.inference_classification = MagicMock(
        value=overrides.get("classification", "focused")
    )
    # Body for read_message
    msg.body = MagicMock()
    msg.body.content = overrides.get("body_content", "<p>Hello</p>")
    msg.body.content_type = overrides.get("content_type", "html")
    msg.to_recipients = overrides.get("to_recipients", [])
    msg.cc_recipients = overrides.get("cc_recipients", [])
    msg.attachments = overrides.get("attachments", [])
    return msg


def _make_folder_mock(messages, next_link=None):
    """Build a mock Graph client for folder-based message queries.

    The Graph SDK uses a synchronous fluent chain:
        client.me.mail_folders.by_mail_folder_id("inbox").messages.get(...)
    Only .get() is async. The intermediate methods return sync objects.
    """
    response = MagicMock(value=messages, odata_next_link=next_link)
    messages_obj = MagicMock()
    messages_obj.get = AsyncMock(return_value=response)
    folder_obj = MagicMock()
    folder_obj.messages = messages_obj
    client = MagicMock()
    client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=folder_obj)
    return client


def _make_message_mock(msg):
    """Build a mock Graph client for single message reads.

    Chain: client.me.messages.by_message_id("id").get()
    Only .get() is async.
    """
    msg_obj = MagicMock()
    msg_obj.get = AsyncMock(return_value=msg)
    client = MagicMock()
    client.me.messages.by_message_id = MagicMock(return_value=msg_obj)
    return client


def _make_search_mock(messages, next_link=None):
    """Build a mock Graph client for search queries (no folder).

    Chain: client.me.messages.get(...)
    """
    response = MagicMock(value=messages, odata_next_link=next_link)
    client = MagicMock()
    client.me.messages.get = AsyncMock(return_value=response)
    return client


class TestFormatMessageSummary:
    def test_formats_basic_message(self):
        """_format_message_summary extracts all expected fields."""
        msg = _make_mock_message()
        result = _format_message_summary(msg)
        assert result["id"] == "AAMkAG123="
        assert result["subject"] == "Test Subject"
        assert result["from_email"] == "sender@test.com"
        assert result["from_name"] == "Sender"
        assert result["is_read"] is False
        assert result["importance"] == "normal"
        assert result["flag"] == "notFlagged"
        assert result["conversation_id"] == "conv123"

    def test_handles_missing_from(self):
        """Gracefully handles message with no from field."""
        msg = _make_mock_message()
        msg.from_ = None
        result = _format_message_summary(msg)
        assert result["from_email"] == ""
        assert result["from_name"] == ""


class TestListInbox:
    @pytest.mark.asyncio
    async def test_list_inbox_returns_messages(self):
        """list_inbox returns structured message list."""
        mock_message = _make_mock_message()
        mock_client = _make_folder_mock([mock_message])

        result = await list_inbox(mock_client, folder="inbox", count=25)
        assert result["count"] == 1
        assert result["messages"][0]["subject"] == "Test Subject"
        assert result["messages"][0]["from_email"] == "sender@test.com"
        assert result["messages"][0]["is_read"] is False

    @pytest.mark.asyncio
    async def test_list_inbox_validates_count(self):
        """Count is clamped to 1-100."""
        mock_client = _make_folder_mock([])
        result = await list_inbox(mock_client, count=200)
        assert result["count"] == 0

    @pytest.mark.asyncio
    async def test_list_inbox_validates_dates(self):
        """Date params are validated via validate_datetime."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid datetime"):
            await list_inbox(mock_client, after="not-a-date")

    @pytest.mark.asyncio
    async def test_list_inbox_has_more(self):
        """has_more is True when odata_next_link is present."""
        mock_client = _make_folder_mock(
            [_make_mock_message()],
            next_link="https://graph.microsoft.com/v1.0/next",
        )
        result = await list_inbox(mock_client, count=1)
        assert result["has_more"] is True

    @pytest.mark.asyncio
    async def test_list_inbox_classification_filter_adds_odata(self):
        """classification="focused" adds inferenceClassification filter to query."""
        mock_client = _make_folder_mock([])
        await list_inbox(mock_client, classification="focused")

        call_kwargs = (
            mock_client.me.mail_folders.by_mail_folder_id.return_value
            .messages.get.call_args
        )
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert qp.filter is not None
        assert "inferenceClassification eq 'focused'" in qp.filter

    @pytest.mark.asyncio
    async def test_list_inbox_classification_other(self):
        """classification="other" is allowed."""
        mock_client = _make_folder_mock([])
        await list_inbox(mock_client, classification="other")

        call_kwargs = (
            mock_client.me.mail_folders.by_mail_folder_id.return_value
            .messages.get.call_args
        )
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert "inferenceClassification eq 'other'" in qp.filter

    @pytest.mark.asyncio
    async def test_list_inbox_classification_invalid(self):
        """Invalid classification raises ValueError."""
        mock_client = _make_folder_mock([])
        with pytest.raises(ValueError, match="classification"):
            await list_inbox(mock_client, classification="bogus")

    @pytest.mark.asyncio
    async def test_list_inbox_classification_combines_with_other_filters(self):
        """classification filter combines with unread_only via 'and'."""
        mock_client = _make_folder_mock([])
        await list_inbox(mock_client, unread_only=True, classification="focused")

        call_kwargs = (
            mock_client.me.mail_folders.by_mail_folder_id.return_value
            .messages.get.call_args
        )
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert "isRead eq false" in qp.filter
        assert "inferenceClassification eq 'focused'" in qp.filter
        assert " and " in qp.filter


class TestReadMessage:
    @pytest.mark.asyncio
    async def test_read_message_text_format(self):
        """read_message returns text body by default."""
        mock_msg = _make_mock_message(body_content="<p>Hello World</p>")
        mock_client = _make_message_mock(mock_msg)
        result = await read_message(mock_client, "AAMkAG123=", format="text")
        assert result["body"] != ""
        assert result["body_html"] is None

    @pytest.mark.asyncio
    async def test_read_message_html_format(self):
        """read_message with format=html returns body_html."""
        mock_msg = _make_mock_message(body_content="<p>Hello</p>")
        mock_client = _make_message_mock(mock_msg)
        result = await read_message(mock_client, "AAMkAG123=", format="html")
        assert result["body_html"] == "<p>Hello</p>"
        assert result["body"] == ""

    @pytest.mark.asyncio
    async def test_read_message_validates_id(self):
        """read_message rejects invalid message IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await read_message(mock_client, "bad id with spaces!")


class TestSearchMail:
    @pytest.mark.asyncio
    async def test_search_sanitizes_query(self):
        """Search query is sanitized before sending to Graph."""
        mock_client = _make_search_mock([])
        result = await search_mail(mock_client, query='test" OR (hack)')
        assert result["count"] == 0

    @pytest.mark.asyncio
    async def test_search_with_folder(self):
        """Search within a specific folder uses folder endpoint."""
        mock_client = _make_folder_mock([])
        result = await search_mail(mock_client, query="test", folder="inbox")
        assert result["count"] == 0
        mock_client.me.mail_folders.by_mail_folder_id.assert_called_with("inbox")


def _make_folder(folder_id: str, name: str, *, total=0, unread=0, parent_id=None, children=0):
    """Build a MagicMock that stands in for a Graph MailFolder entity."""
    f = MagicMock()
    f.id = folder_id
    f.display_name = name
    f.total_item_count = total
    f.unread_item_count = unread
    f.parent_folder_id = parent_id
    f.child_folder_count = children
    return f


class TestListFolders:
    @pytest.mark.asyncio
    async def test_list_folders_returns_folders(self):
        """list_folders returns folder list with counts, parent_id, child_count."""
        mock_folder = _make_folder(
            "folder123", "Inbox", total=42, unread=5, parent_id="root", children=2
        )

        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=MagicMock(value=[mock_folder], odata_next_link=None)
        )

        result = await list_folders(mock_client)
        assert result["count"] == 1
        assert result["folders"][0]["name"] == "Inbox"
        assert result["folders"][0]["total"] == 42
        assert result["folders"][0]["unread"] == 5
        assert result["folders"][0]["parent_id"] == "root"
        assert result["folders"][0]["child_count"] == 2

    @pytest.mark.asyncio
    async def test_list_folders_recursive_walks_subfolders(self):
        """recursive=True returns the full folder tree via BFS walk."""
        inbox = _make_folder("inbox_id", "Inbox", parent_id="root", children=0)
        receipts = _make_folder("receipts_id", "Receipts", parent_id="root", children=1)
        domains = _make_folder("domains_id", "Domains", parent_id="receipts_id", children=0)

        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=MagicMock(value=[inbox, receipts], odata_next_link=None)
        )
        child_builder = MagicMock()
        child_builder.child_folders.get = AsyncMock(
            return_value=MagicMock(value=[domains], odata_next_link=None)
        )
        mock_client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=child_builder)

        result = await list_folders(mock_client, recursive=True)
        names = {f["name"] for f in result["folders"]}
        assert names == {"Inbox", "Receipts", "Domains"}
        assert result["count"] == 3
        assert result["has_more"] is False
        # Only the folder with children triggers a child_folders fetch.
        mock_client.me.mail_folders.by_mail_folder_id.assert_called_once_with("receipts_id")

    @pytest.mark.asyncio
    async def test_list_folders_recursive_paginates_child_folders(self):
        """Child-folder listings follow odata_next_link so >10 subfolders are returned."""
        inbox = _make_folder("inbox_id", "Inbox", parent_id="root", children=15)

        page1 = [_make_folder(f"c{i}", f"Child{i}", parent_id="inbox_id") for i in range(10)]
        page2 = [_make_folder(f"c{i}", f"Child{i}", parent_id="inbox_id") for i in range(10, 15)]

        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=MagicMock(value=[inbox], odata_next_link=None)
        )
        child_builder = MagicMock()
        # First page has nextLink; with_url(...)  returns a builder whose get() returns page2.
        page1_response = MagicMock(value=page1, odata_next_link="https://next-page")
        page2_response = MagicMock(value=page2, odata_next_link=None)
        next_page_builder = MagicMock()
        next_page_builder.get = AsyncMock(return_value=page2_response)
        child_builder.child_folders.get = AsyncMock(return_value=page1_response)
        child_builder.child_folders.with_url = MagicMock(return_value=next_page_builder)
        mock_client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=child_builder)

        result = await list_folders(mock_client, recursive=True)
        names = {f["name"] for f in result["folders"]}
        # Inbox + 15 children
        assert len(result["folders"]) == 16
        assert "Inbox" in names
        assert {f"Child{i}" for i in range(15)}.issubset(names)
        child_builder.child_folders.with_url.assert_called_once_with("https://next-page")
