"""Tests for To Do task attachment tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.todo_attachments import (
    delete_task_attachment,
    download_task_attachment,
    list_task_attachments,
    upload_task_attachment,
)


def _cfg(tmp_path):
    return Config(client_id="test", attachments_dir=str(tmp_path / "att"))


def _mock_attachment(att_id="att1", name="handbook.pdf", size=1234, content_type="application/pdf"):
    mock = MagicMock()
    mock.id = att_id
    mock.name = name
    mock.size = size
    mock.content_type = content_type
    mock.last_modified_date_time = "2026-09-15T10:00:00Z"
    return mock


def _build_mock_client(attachments=None, upload_url="https://upload.example/session"):
    """Build a mock Graph client wired for task-attachment operations."""
    if attachments is None:
        attachments = [_mock_attachment()]

    mock_client = MagicMock()

    # Default task list for _resolve_list_id
    default_list = MagicMock()
    default_list.id = "list1"
    default_list.is_owner = True
    default_list.wellknown_list_name = MagicMock(value="defaultList")
    mock_client.me.todo.lists.get = AsyncMock(return_value=MagicMock(value=[default_list]))

    # .../tasks/{taskId}/attachments/{attId}/$value and DELETE
    mock_attachment_item = MagicMock()
    mock_attachment_item.content = MagicMock()
    mock_attachment_item.content.get = AsyncMock(return_value=b"%PDF-fake-bytes")
    mock_attachment_item.delete = AsyncMock()

    mock_attachments = MagicMock()
    mock_attachments.get = AsyncMock(return_value=MagicMock(value=attachments))
    mock_attachments.by_attachment_base_id = MagicMock(return_value=mock_attachment_item)
    mock_attachments.create_upload_session = MagicMock()
    mock_attachments.create_upload_session.post = AsyncMock(
        return_value=MagicMock(upload_url=upload_url)
    )

    mock_task_item = MagicMock()
    mock_task_item.attachments = mock_attachments

    mock_tasks = MagicMock()
    mock_tasks.by_todo_task_id = MagicMock(return_value=mock_task_item)

    mock_list_item = MagicMock()
    mock_list_item.tasks = mock_tasks

    mock_client.me.todo.lists.by_todo_task_list_id = MagicMock(return_value=mock_list_item)

    return mock_client


def _attachments_of(client):
    task_item = (
        client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id
    )
    return task_item.return_value.attachments


# --- list_task_attachments ---


class TestListTaskAttachments:
    async def test_lists_attachments(self):
        client = _build_mock_client(
            attachments=[
                _mock_attachment(),
                _mock_attachment("att2", "notes.txt", 20, "text/plain"),
            ]
        )

        result = await list_task_attachments(client, task_id="task1")

        assert result["count"] == 2
        assert result["attachments"][0]["id"] == "att1"
        assert result["attachments"][0]["name"] == "handbook.pdf"
        assert result["attachments"][0]["size"] == 1234
        assert result["attachments"][0]["content_type"] == "application/pdf"
        assert result["attachments"][0]["last_modified"] == "2026-09-15T10:00:00Z"

    async def test_empty_attachments(self):
        client = _build_mock_client(attachments=[])

        result = await list_task_attachments(client, task_id="task1")

        assert result["attachments"] == []
        assert result["count"] == 0


# --- download_task_attachment ---


class TestDownloadTaskAttachment:
    async def test_download_writes_content(self, tmp_path):
        client = _build_mock_client()
        config = _cfg(tmp_path)

        result = await download_task_attachment(
            client,
            task_id="task1",
            attachment_id="att1",
            save_path="handbook.pdf",
            config=config,
        )

        saved = tmp_path / "att" / "handbook.pdf"
        assert saved.read_bytes() == b"%PDF-fake-bytes"
        assert result["saved_to"] == str(saved)
        assert result["size"] == len(b"%PDF-fake-bytes")
        _attachments_of(client).by_attachment_base_id.assert_called_with("att1")

    async def test_download_rejects_escape_path(self, tmp_path):
        """save_path outside attachments_dir is refused, not written."""
        client = _build_mock_client()
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="outside the permitted directory"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path=str(tmp_path / "escape.pdf"),
                config=config,
            )


# --- upload_task_attachment ---


class TestUploadTaskAttachment:
    @pytest.fixture
    def source_file(self, tmp_path):
        f = tmp_path / "att" / "upload.bin"
        f.parent.mkdir(parents=True)
        f.write_bytes(b"\x00" * 2048)
        return f

    async def test_upload_creates_session_and_puts_chunks(self, tmp_path, source_file, monkeypatch):
        client = _build_mock_client()
        config = _cfg(tmp_path)
        put = AsyncMock()
        monkeypatch.setattr("outlook_mcp.tools.todo_attachments._upload_large_file", put)

        result = await upload_task_attachment(
            client, task_id="task1", file_path=str(source_file), config=config
        )

        assert result["status"] == "uploaded"
        assert result["name"] == "upload.bin"
        assert result["size"] == 2048

        from msgraph.generated.models.attachment_info import AttachmentInfo
        from msgraph.generated.models.attachment_type import AttachmentType

        body = _attachments_of(client).create_upload_session.post.call_args.args[0]
        info = body.attachment_info
        assert isinstance(info, AttachmentInfo)
        assert info.attachment_type is AttachmentType.File
        assert info.name == "upload.bin"
        assert info.size == 2048
        assert info.content_type == "application/octet-stream"

        put.assert_awaited_once_with("https://upload.example/session", str(source_file), 2048)

    async def test_upload_rejects_oversize(self, tmp_path, source_file, monkeypatch):
        client = _build_mock_client()
        config = _cfg(tmp_path)
        monkeypatch.setattr("os.path.getsize", lambda p: 25 * 1024 * 1024 + 1)

        with pytest.raises(ValueError, match="25 MB"):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(source_file), config=config
            )
        _attachments_of(client).create_upload_session.post.assert_not_called()

    async def test_upload_rejects_empty_file(self, tmp_path):
        client = _build_mock_client()
        config = _cfg(tmp_path)
        empty = tmp_path / "att" / "empty.bin"
        empty.parent.mkdir(parents=True)
        empty.write_bytes(b"")

        with pytest.raises(ValueError, match="empty"):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(empty), config=config
            )

    async def test_upload_rejects_missing_file(self, tmp_path):
        client = _build_mock_client()
        config = _cfg(tmp_path)

        with pytest.raises(FileNotFoundError):
            await upload_task_attachment(
                client, task_id="task1", file_path="ghost.bin", config=config
            )

    async def test_upload_rejects_escape_path(self, tmp_path):
        client = _build_mock_client()
        config = _cfg(tmp_path)
        outside = tmp_path / "outside.bin"
        outside.write_bytes(b"x")

        with pytest.raises(ValueError, match="outside the permitted directory"):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(outside), config=config
            )

    async def test_upload_read_only(self, tmp_path, source_file):
        client = _build_mock_client()
        config = Config(client_id="test", read_only=True, attachments_dir=str(tmp_path / "att"))

        with pytest.raises(ReadOnlyError):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(source_file), config=config
            )


# --- delete_task_attachment ---


class TestDeleteTaskAttachment:
    async def test_delete_task_attachment(self):
        client = _build_mock_client()
        config = Config(client_id="test")

        result = await delete_task_attachment(
            client, task_id="task1", attachment_id="att1", config=config
        )

        assert result["status"] == "deleted"
        _attachments_of(client).by_attachment_base_id.assert_called_with("att1")
        _attachments_of(client).by_attachment_base_id.return_value.delete.assert_called_once()

    async def test_delete_read_only(self):
        client = _build_mock_client()
        config = Config(client_id="test", read_only=True)

        with pytest.raises(ReadOnlyError):
            await delete_task_attachment(
                client, task_id="task1", attachment_id="att1", config=config
            )
