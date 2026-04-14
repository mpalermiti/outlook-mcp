"""Tests for mail attachment tools."""

import base64
import os
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_attachments import (
    download_attachment,
    list_attachments,
    send_with_attachments,
)

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


class TestListAttachments:
    async def test_list_returns_metadata(self):
        """list_attachments returns attachment metadata and count."""
        mock_att = MagicMock()
        mock_att.id = "att1"
        mock_att.name = "report.pdf"
        mock_att.size = 1024
        mock_att.content_type = "application/pdf"

        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.attachments.get = AsyncMock(
            return_value=MagicMock(value=[mock_att])
        )

        result = await list_attachments(mock_client, message_id="AAMkAG123=")
        assert result["count"] == 1
        assert len(result["attachments"]) == 1
        att = result["attachments"][0]
        assert att["id"] == "att1"
        assert att["name"] == "report.pdf"
        assert att["size"] == 1024
        assert att["content_type"] == "application/pdf"

    async def test_list_empty_attachments(self):
        """list_attachments returns empty list when no attachments."""
        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.attachments.get = AsyncMock(
            return_value=MagicMock(value=[])
        )

        result = await list_attachments(mock_client, message_id="AAMkAG123=")
        assert result["count"] == 0
        assert result["attachments"] == []

    async def test_list_validates_message_id(self):
        """list_attachments rejects invalid message IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await list_attachments(mock_client, message_id="")

    async def test_list_multiple_attachments(self):
        """list_attachments handles multiple attachments."""
        att1 = MagicMock()
        att1.id = "att1"
        att1.name = "a.pdf"
        att1.size = 100
        att1.content_type = "application/pdf"
        att2 = MagicMock()
        att2.id = "att2"
        att2.name = "b.png"
        att2.size = 200
        att2.content_type = "image/png"

        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.attachments.get = AsyncMock(
            return_value=MagicMock(value=[att1, att2])
        )

        result = await list_attachments(mock_client, message_id="AAMkAG123=")
        assert result["count"] == 2
        assert result["attachments"][0]["name"] == "a.pdf"
        assert result["attachments"][1]["name"] == "b.png"


class TestDownloadAttachment:
    async def test_download_returns_base64(self):
        """download_attachment returns base64 content when no save_path."""
        mock_att = MagicMock()
        mock_att.id = "att1"
        mock_att.name = "report.pdf"
        mock_att.size = 1024
        mock_att.content_type = "application/pdf"
        mock_att.content_bytes = b"file content bytes"

        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.attachments.by_attachment_id.return_value.get = AsyncMock(  # noqa: E501
            return_value=mock_att
        )

        result = await download_attachment(
            mock_client, message_id="AAMkAG123=", attachment_id="att1"
        )
        expected_b64 = base64.b64encode(b"file content bytes").decode("utf-8")
        assert result["content_base64"] == expected_b64
        assert result["name"] == "report.pdf"
        assert result["content_type"] == "application/pdf"
        assert result["size"] == 1024

    async def test_download_writes_to_path(self, tmp_path):
        """download_attachment writes bytes to file when save_path given."""
        mock_att = MagicMock()
        mock_att.id = "att1"
        mock_att.name = "report.pdf"
        mock_att.size = 18
        mock_att.content_type = "application/pdf"
        mock_att.content_bytes = b"file content bytes"

        mock_client = MagicMock()
        mock_client.me.messages.by_message_id.return_value.attachments.by_attachment_id.return_value.get = AsyncMock(  # noqa: E501
            return_value=mock_att
        )

        save_path = str(tmp_path / "report.pdf")
        result = await download_attachment(
            mock_client,
            message_id="AAMkAG123=",
            attachment_id="att1",
            save_path=save_path,
        )
        assert result["saved_to"] == save_path
        assert os.path.isfile(save_path)
        with open(save_path, "rb") as f:
            assert f.read() == b"file content bytes"

    async def test_download_validates_message_id(self):
        """download_attachment rejects invalid message ID."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await download_attachment(mock_client, message_id="", attachment_id="att1")

    async def test_download_validates_attachment_id(self):
        """download_attachment rejects invalid attachment ID."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await download_attachment(mock_client, message_id="AAMkAG123=", attachment_id="")

    async def test_download_rejects_path_traversal(self):
        """download_attachment rejects save_path with .. traversal."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="traversal"):
            await download_attachment(
                mock_client,
                message_id="AAMkAG123=",
                attachment_id="att1",
                save_path="/tmp/../etc/passwd",
            )


class TestSendWithAttachments:
    async def test_send_small_file(self, tmp_path):
        """send_with_attachments sends small files as inline base64."""
        # Create a small test file (under 3MB)
        test_file = tmp_path / "small.txt"
        test_file.write_bytes(b"hello world")

        mock_client = MagicMock()
        mock_client.me.send_mail.post = AsyncMock()

        result = await send_with_attachments(
            mock_client,
            to=["test@example.com"],
            subject="Test",
            body="See attached",
            attachment_paths=[str(test_file)],
            config=_CFG,
        )
        assert result["status"] == "sent"
        assert result["attachment_count"] == 1
        mock_client.me.send_mail.post.assert_called_once()

    async def test_send_raises_read_only(self, tmp_path):
        """send_with_attachments raises ReadOnlyError in read-only mode."""
        test_file = tmp_path / "small.txt"
        test_file.write_bytes(b"hello")

        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await send_with_attachments(
                mock_client,
                to=["a@b.com"],
                subject="Test",
                body="Hi",
                attachment_paths=[str(test_file)],
                config=_CFG_RO,
            )

    async def test_send_rejects_invalid_email(self, tmp_path):
        """send_with_attachments validates email addresses."""
        test_file = tmp_path / "small.txt"
        test_file.write_bytes(b"hello")

        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await send_with_attachments(
                mock_client,
                to=["not-an-email"],
                subject="Test",
                body="Hi",
                attachment_paths=[str(test_file)],
                config=_CFG,
            )

    async def test_send_rejects_missing_file(self):
        """send_with_attachments raises FileNotFoundError for missing files."""
        mock_client = MagicMock()
        with pytest.raises(FileNotFoundError):
            await send_with_attachments(
                mock_client,
                to=["a@b.com"],
                subject="Test",
                body="Hi",
                attachment_paths=["/nonexistent/file.txt"],
                config=_CFG,
            )

    async def test_send_large_file_uses_upload_session(self, tmp_path):
        """send_with_attachments uses upload session for files over 3MB."""
        # Create a file just over 3MB
        large_file = tmp_path / "large.bin"
        large_file.write_bytes(b"x" * (3 * 1024 * 1024 + 1))

        mock_client = MagicMock()
        # Mock draft message creation
        mock_draft = MagicMock()
        mock_draft.id = "draft123"
        mock_client.me.messages.post = AsyncMock(return_value=mock_draft)

        # Mock upload session creation
        mock_session = MagicMock()
        mock_session.upload_url = "https://graph.microsoft.com/upload/session"
        mock_client.me.messages.by_message_id.return_value.attachments.create_upload_session.post = AsyncMock(  # noqa: E501
            return_value=mock_session
        )

        # Mock the send for the draft
        mock_client.me.messages.by_message_id.return_value.send.post = AsyncMock()

        with patch("outlook_mcp.tools.mail_attachments._upload_large_file", new_callable=AsyncMock):
            result = await send_with_attachments(
                mock_client,
                to=["a@b.com"],
                subject="Large file",
                body="See attached",
                attachment_paths=[str(large_file)],
                config=_CFG,
            )
        assert result["status"] == "sent"
        assert result["attachment_count"] == 1

    async def test_send_with_cc_bcc(self, tmp_path):
        """send_with_attachments passes cc and bcc recipients."""
        test_file = tmp_path / "small.txt"
        test_file.write_bytes(b"hello")

        mock_client = MagicMock()
        mock_client.me.send_mail.post = AsyncMock()

        result = await send_with_attachments(
            mock_client,
            to=["to@test.com"],
            subject="Test",
            body="Hi",
            attachment_paths=[str(test_file)],
            cc=["cc@test.com"],
            bcc=["bcc@test.com"],
            config=_CFG,
        )
        assert result["status"] == "sent"
        mock_client.me.send_mail.post.assert_called_once()
