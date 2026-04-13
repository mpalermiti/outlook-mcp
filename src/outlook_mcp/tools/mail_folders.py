"""Mail folder management tools: create, rename, delete."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.validation import WELL_KNOWN_FOLDERS, validate_graph_id


def _check_read_only(read_only: bool, tool_name: str) -> None:
    """Raise ReadOnlyError if server is in read-only mode."""
    if read_only:
        raise ReadOnlyError(tool_name)


async def create_folder(
    graph_client: Any,
    name: str,
    parent_folder: str | None = None,
    read_only: bool = False,
) -> dict:
    """Create a mail folder.

    Creates a top-level folder under /me/mailFolders, or a child folder
    under the specified parent_folder.
    """
    _check_read_only(read_only, "outlook_create_folder")

    from msgraph.generated.models.mail_folder import MailFolder

    body = MailFolder()
    body.display_name = name

    if parent_folder:
        parent_folder = validate_graph_id(parent_folder)
        result = await graph_client.me.mail_folders.by_mail_folder_id(
            parent_folder
        ).child_folders.post(body)
    else:
        result = await graph_client.me.mail_folders.post(body)

    return {"id": result.id, "name": result.display_name}


async def rename_folder(
    graph_client: Any,
    folder_id: str,
    name: str,
    read_only: bool = False,
) -> dict:
    """Rename a mail folder.

    Validates the folder ID and patches the display name.
    """
    _check_read_only(read_only, "outlook_rename_folder")
    folder_id = validate_graph_id(folder_id)

    from msgraph.generated.models.mail_folder import MailFolder

    body = MailFolder()
    body.display_name = name

    result = await graph_client.me.mail_folders.by_mail_folder_id(folder_id).patch(body)

    return {"id": result.id, "name": result.display_name}


async def delete_folder(
    graph_client: Any,
    folder_id: str,
    read_only: bool = False,
) -> dict:
    """Delete a mail folder.

    Validates the folder ID. Refuses to delete well-known folders
    (inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox).
    """
    _check_read_only(read_only, "outlook_delete_folder")

    # Reject well-known folders before validation — they're valid names
    # but must never be deleted.
    if folder_id.lower() in WELL_KNOWN_FOLDERS:
        raise ValueError(
            f"Cannot delete well-known folder '{folder_id}'. "
            "Only user-created folders can be deleted."
        )

    folder_id = validate_graph_id(folder_id)

    await graph_client.me.mail_folders.by_mail_folder_id(folder_id).delete()

    return {"status": "deleted"}
