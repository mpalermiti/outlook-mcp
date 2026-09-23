"""Tests for To Do task attachment tools.

Upload goes through the inline base64 POST (consumer mailboxes 404 the
upload-session endpoint — verified live), so these tests assert the
``taskFileAttachment`` payload that reaches ``attachments.post`` and treat the
SDK's base64 handling as the SDK's problem.
"""

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


def _entity(**attrs):
    """A mock attachment entity. Attributes are assigned, not passed to the
    constructor — MagicMock(name=...) names the mock and leaves `.name` an
    auto-vivified child, which blows up in sanitize_output."""
    entity = MagicMock()
    for key, value in attrs.items():
        setattr(entity, key, value)
    return entity


def _mock_attachment(att_id="att1", name="handbook.pdf", size=1234, content_type="application/pdf"):
    mock = MagicMock()
    mock.id = att_id
    mock.name = name
    mock.size = size
    mock.content_type = content_type
    mock.last_modified_date_time = "2026-09-15T10:00:00Z"
    return mock


def _build_mock_client(attachments=None, post_result=None, attachment_entity=None):
    """Build a mock Graph client wired for task-attachment operations.

    ``attachment_entity`` is what an entity GET on one attachment returns
    (download path); defaults to a stub with contentBytes set.
    """
    if attachments is None:
        attachments = [_mock_attachment()]
    if post_result is None:
        post_result = _mock_attachment(att_id="newatt1", name="upload.bin", size=2048)
    if attachment_entity is None:
        # Built attribute-by-attribute: MagicMock(name=...) names the *mock*,
        # it does not set a `.name` attribute.
        attachment_entity = MagicMock()
        attachment_entity.id = "att1"
        attachment_entity.name = "handbook.pdf"
        attachment_entity.content_type = "application/pdf"
        attachment_entity.content_bytes = b"%PDF-fake-bytes"

    mock_client = MagicMock()

    # Default task list for _resolve_list_id (odata_next_link spelled out —
    # a MagicMock auto-vivifies a truthy one the pagination would chase).
    default_list = MagicMock()
    default_list.id = "list1"
    default_list.is_owner = True
    default_list.wellknown_list_name = MagicMock(value="defaultList")
    mock_client.me.todo.lists.get = AsyncMock(
        return_value=MagicMock(value=[default_list], odata_next_link=None)
    )

    # Entity GET (download) and DELETE on one attachment
    mock_attachment_item = MagicMock()
    mock_attachment_item.get = AsyncMock(return_value=attachment_entity)
    mock_attachment_item.delete = AsyncMock()

    mock_attachments = MagicMock()
    mock_attachments.get = AsyncMock(
        return_value=MagicMock(value=attachments, odata_next_link=None)
    )
    mock_attachments.post = AsyncMock(return_value=post_result)
    mock_attachments.by_attachment_base_id = MagicMock(return_value=mock_attachment_item)

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
        assert result["has_more"] is False
        assert result["next_cursor"] is None
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

    async def test_paginates_with_top_and_cursor(self):
        """$top reaches the wire; a nextLink becomes has_more/next_cursor —
        the shape every other list tool already has."""
        client = _build_mock_client()
        _attachments_of(client).get = AsyncMock(
            return_value=MagicMock(
                value=[_mock_attachment()],
                odata_next_link="https://graph.microsoft.com/v1.0/me/todo/lists/"
                "list1/tasks/task1/attachments?$skip=25",
            )
        )

        result = await list_task_attachments(client, task_id="task1", count=25)

        qp = _attachments_of(client).get.call_args.kwargs[
            "request_configuration"
        ].query_parameters
        assert qp.top == 25
        assert result["has_more"] is True
        assert result["next_cursor"] is not None

    async def test_last_modified_datetime_is_normalized(self):
        """kiota hands back a datetime; the listing must emit ISO 8601, not
        str(datetime)'s '2026-09-15 10:00:00+00:00'."""
        from datetime import datetime, timezone

        att = _mock_attachment()
        att.last_modified_date_time = datetime(2026, 9, 15, 10, 0, tzinfo=timezone.utc)
        client = _build_mock_client(attachments=[att])

        result = await list_task_attachments(client, task_id="task1")

        assert result["attachments"][0]["last_modified"] == "2026-09-15T10:00:00+00:00"


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

    async def test_download_zero_byte_attachment_writes_zero_byte_file(self, tmp_path):
        """A 0-byte attachment is honest data, not a crash: the entity GET
        returns no contentBytes and the file lands as 0 bytes without a
        TypeError or a truncated leftover."""
        client = _build_mock_client(
            attachment_entity=_entity(
                id="att1", name="empty.bin", content_type="application/octet-stream",
                content_bytes=None,
            )
        )
        config = _cfg(tmp_path)

        result = await download_task_attachment(
            client,
            task_id="task1",
            attachment_id="att1",
            save_path="empty.bin",
            config=config,
        )

        saved = tmp_path / "att" / "empty.bin"
        assert saved.read_bytes() == b""
        assert result["size"] == 0

    async def test_download_does_not_truncate_destination_before_bytes(self, tmp_path):
        """The old bug: open(wb) truncated the target first, so a failed fetch
        zeroed a staged file. Now the existing file survives a fetch failure."""
        client = _build_mock_client()
        _attachments_of(client).by_attachment_base_id.return_value.get = AsyncMock(
            side_effect=RuntimeError("graph blew up")
        )
        config = _cfg(tmp_path)
        base = tmp_path / "att"
        base.mkdir(parents=True)
        staged = base / "staged.pdf"
        staged.write_bytes(b"precious staged bytes")

        with pytest.raises(RuntimeError, match="graph blew up"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="staged.pdf",
                config=config,
            )

        assert staged.read_bytes() == b"precious staged bytes"

    async def test_download_replaces_atomically(self, tmp_path):
        """Success replaces the destination in one step and leaves no temp
        files behind."""
        client = _build_mock_client()
        config = _cfg(tmp_path)
        base = tmp_path / "att"
        base.mkdir(parents=True)
        dst = base / "handbook.pdf"
        dst.write_bytes(b"previous contents")

        await download_task_attachment(
            client,
            task_id="task1",
            attachment_id="att1",
            save_path="handbook.pdf",
            config=config,
        )

        assert dst.read_bytes() == b"%PDF-fake-bytes"
        leftovers = [p.name for p in base.iterdir() if p.name.startswith(".download-")]
        assert leftovers == []

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

    async def test_download_none_entity_is_an_error_not_a_crash(self, tmp_path):
        client = _build_mock_client(attachment_entity=None)
        _attachments_of(client).by_attachment_base_id.return_value.get = AsyncMock(
            return_value=None
        )
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="no attachment"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="x.bin",
                config=config,
            )

    async def test_download_missing_bytes_with_declared_size_is_an_error(
        self, tmp_path
    ):
        """The old bug: contentBytes absent -> `or b""` wrote a 0-byte file
        under the trusted name and reported success. The entity's own size is
        the witness — bytes and declared size must not contradict each other,
        and a staged destination file must survive the refusal."""
        client = _build_mock_client(
            attachment_entity=_entity(
                id="att1",
                name="handbook.pdf",
                content_type="application/pdf",
                content_bytes=None,
                size=1234,
            )
        )
        config = _cfg(tmp_path)
        base = tmp_path / "att"
        base.mkdir(parents=True)
        staged = base / "handbook.pdf"
        staged.write_bytes(b"precious staged bytes")

        with pytest.raises(ValueError, match="declares 1234 bytes"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="handbook.pdf",
                config=config,
            )

        assert staged.read_bytes() == b"precious staged bytes"

    async def test_download_entity_without_content_bytes_property_is_a_clear_error(
        self, tmp_path
    ):
        """A response without @odata.type deserializes as AttachmentBase,
        which has no content_bytes attribute — that must be a ValueError, not
        a bare AttributeError whose text never reaches the model."""
        from msgraph.generated.models.attachment_base import AttachmentBase

        entity = AttachmentBase()
        entity.id = "att1"
        entity.name = "handbook.pdf"
        entity.content_type = "application/pdf"
        entity.size = 1234

        client = _build_mock_client(attachment_entity=entity)
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="no contentBytes property"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="handbook.pdf",
                config=config,
            )

    async def test_download_bytes_where_size_says_zero_is_an_error(self, tmp_path):
        """The contradiction in the other direction: a payload the entity
        claims is empty is not written either."""
        client = _build_mock_client(
            attachment_entity=_entity(
                id="att1",
                name="handbook.pdf",
                content_type="application/pdf",
                content_bytes=b"%PDF-fake-bytes",
                size=0,
            )
        )
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="contradicts itself"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="handbook.pdf",
                config=config,
            )

    async def test_download_declared_size_zero_with_no_bytes_writes_empty(
        self, tmp_path
    ):
        """A 0-byte attachment carries size=0 and no contentBytes — honest
        data, and the witness check must not block it."""
        client = _build_mock_client(
            attachment_entity=_entity(
                id="att1",
                name="empty.bin",
                content_type="application/octet-stream",
                content_bytes=None,
                size=0,
            )
        )
        config = _cfg(tmp_path)

        result = await download_task_attachment(
            client,
            task_id="task1",
            attachment_id="att1",
            save_path="empty.bin",
            config=config,
        )

        assert (tmp_path / "att" / "empty.bin").read_bytes() == b""
        assert result["size"] == 0

    async def test_download_rejects_the_directory_itself_before_any_fetch(
        self, tmp_path
    ):
        """save_path resolving to attachments_dir (or '.') passed the
        confinement check and sent dirname() one level above the fence —
        mkstemp would land next to config.json. Refused before the Graph
        call now."""
        client = _build_mock_client()
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="is a directory"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path=".",
                config=config,
            )

        _attachments_of(client).by_attachment_base_id.return_value.get.assert_not_called()

    async def test_download_rejects_missing_parent_directory_before_any_fetch(
        self, tmp_path
    ):
        """sub/x.pdf with sub/ absent used to fail at mkstemp *after* the
        fetch, with a FileNotFoundError the model never sees. ValueError,
        before the fetch."""
        client = _build_mock_client()
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="does not exist"):
            await download_task_attachment(
                client,
                task_id="task1",
                attachment_id="att1",
                save_path="sub/x.pdf",
                config=config,
            )

        _attachments_of(client).by_attachment_base_id.return_value.get.assert_not_called()


# --- upload_task_attachment ---


class TestUploadTaskAttachment:
    @pytest.fixture
    def source_file(self, tmp_path):
        f = tmp_path / "att" / "upload.bin"
        f.parent.mkdir(parents=True)
        f.write_bytes(b"\x00" * 2048)
        return f

    async def test_upload_posts_inline_task_file_attachment(self, tmp_path, source_file):
        client = _build_mock_client()
        config = _cfg(tmp_path)

        result = await upload_task_attachment(
            client, task_id="task1", file_path=str(source_file), config=config
        )

        assert result["status"] == "attached"
        assert result["attachment_id"] == "newatt1"
        assert result["name"] == "upload.bin"
        assert result["size"] == 2048

        from msgraph.generated.models.task_file_attachment import TaskFileAttachment

        payload = _attachments_of(client).post.call_args.args[0]
        assert isinstance(payload, TaskFileAttachment), (
            f"Graph expects a typed TaskFileAttachment, got {type(payload).__name__}"
        )
        assert payload.odata_type == "#microsoft.graph.taskFileAttachment"
        assert payload.name == "upload.bin"
        assert payload.size == 2048
        assert payload.content_bytes == b"\x00" * 2048
        assert payload.content_type == "application/octet-stream"

    async def test_upload_rejects_oversize(self, tmp_path):
        """One byte over the 20 MiB ceiling is refused before any network I/O
        — a real file, because the guard reads os.stat, not a mocked size."""
        client = _build_mock_client()
        config = _cfg(tmp_path)
        oversize = tmp_path / "att" / "oversize.bin"
        oversize.parent.mkdir(parents=True)
        oversize.write_bytes(b"\x00" * (20 * 1024 * 1024 + 1))

        with pytest.raises(ValueError, match="20 MiB"):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(oversize), config=config
            )
        _attachments_of(client).post.assert_not_called()

    async def test_upload_rejects_empty_file(self, tmp_path):
        client = _build_mock_client()
        config = _cfg(tmp_path)
        empty = tmp_path / "att" / "empty.bin"
        empty.parent.mkdir(parents=True)
        empty.write_bytes(b"")

        with pytest.raises(ValueError, match="1 byte"):
            await upload_task_attachment(
                client, task_id="task1", file_path=str(empty), config=config
            )

    async def test_upload_missing_file_error_carries_the_path(self, tmp_path):
        """ValueError, not FileNotFoundError: only ValueError text survives
        _wrap_tool_errors, and the path is what the model needs to fix."""
        client = _build_mock_client()
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="ghost.bin"):
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

    async def test_upload_null_response_id_is_an_error(self, tmp_path, source_file):
        """An empty 201/204 must raise (the attachment may exist server-side),
        not AttributeError — a blind retry would attach a duplicate."""
        client = _build_mock_client(post_result=_entity(id=None, name=None))
        config = _cfg(tmp_path)

        with pytest.raises(ValueError, match="do not retry"):
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
