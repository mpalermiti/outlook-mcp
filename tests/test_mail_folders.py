"""Tests for mail folder management tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_folders import create_folder, delete_folder, rename_folder

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _mock_folder(folder_id: str = "folder123", name: str = "My Folder") -> MagicMock:
    """Create a mock folder object."""
    folder = MagicMock()
    folder.id = folder_id
    folder.display_name = name
    return folder


class TestCreateFolder:
    async def test_create_top_level(self):
        """create_folder creates a top-level folder via mail_folders.post()."""
        mock_client = MagicMock()
        mock_client.me.mail_folders.post = AsyncMock(return_value=_mock_folder())

        result = await create_folder(mock_client, name="My Folder", config=_CFG)

        assert result == {"id": "folder123", "name": "My Folder"}
        mock_client.me.mail_folders.post.assert_called_once()

    async def test_create_nested(self):
        """create_folder with parent_folder creates a child folder."""
        mock_client = MagicMock()
        child_mock = MagicMock()
        child_mock.child_folders.post = AsyncMock(return_value=_mock_folder())
        mock_client.me.mail_folders.by_mail_folder_id.return_value = child_mock

        result = await create_folder(
            mock_client, name="Sub Folder", parent_folder="AAMkAGparent123=", config=_CFG,
        )

        assert result == {"id": "folder123", "name": "My Folder"}
        mock_client.me.mail_folders.by_mail_folder_id.assert_called_once_with("AAMkAGparent123=")
        child_mock.child_folders.post.assert_called_once()

    async def test_create_raises_read_only(self):
        """create_folder raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await create_folder(mock_client, name="Test", config=_CFG_RO)


class TestRenameFolder:
    async def test_rename_validates_id_and_patches(self):
        """rename_folder validates the folder ID and calls patch()."""
        mock_client = MagicMock()
        folder_builder = MagicMock()
        folder_builder.patch = AsyncMock(return_value=_mock_folder(name="Renamed Folder"))
        mock_client.me.mail_folders.by_mail_folder_id.return_value = folder_builder

        result = await rename_folder(
            mock_client, folder_id="AAMkAGfolder123=", name="Renamed Folder", config=_CFG,
        )

        assert result == {"id": "folder123", "name": "Renamed Folder"}
        mock_client.me.mail_folders.by_mail_folder_id.assert_called_once_with("AAMkAGfolder123=")
        folder_builder.patch.assert_called_once()

    async def test_rename_rejects_invalid_id(self):
        """rename_folder rejects invalid folder IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await rename_folder(mock_client, folder_id="", name="New Name", config=_CFG)

    async def test_rename_raises_read_only(self):
        """rename_folder raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await rename_folder(
                mock_client, folder_id="AAMkAG123=", name="New", config=_CFG_RO,
            )


class TestDeleteFolder:
    async def test_delete_validates_id_and_deletes(self):
        """delete_folder validates the ID and calls delete()."""
        mock_client = MagicMock()
        folder_builder = MagicMock()
        folder_builder.delete = AsyncMock()
        mock_client.me.mail_folders.by_mail_folder_id.return_value = folder_builder

        result = await delete_folder(mock_client, folder_id="AAMkAGfolder456=", config=_CFG)

        assert result == {"status": "deleted"}
        mock_client.me.mail_folders.by_mail_folder_id.assert_called_once_with("AAMkAGfolder456=")
        folder_builder.delete.assert_called_once()

    async def test_delete_rejects_well_known_inbox(self):
        """delete_folder refuses to delete the inbox folder."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Cannot delete well-known folder"):
            await delete_folder(mock_client, folder_id="inbox", config=_CFG)

    async def test_delete_rejects_well_known_drafts(self):
        """delete_folder refuses to delete the drafts folder."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Cannot delete well-known folder"):
            await delete_folder(mock_client, folder_id="drafts", config=_CFG)

    async def test_delete_rejects_invalid_id(self):
        """delete_folder rejects invalid folder IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await delete_folder(mock_client, folder_id="", config=_CFG)

    async def test_delete_raises_read_only(self):
        """delete_folder raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_folder(mock_client, folder_id="AAMkAG123=", config=_CFG_RO)
