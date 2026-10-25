"""Resolve folder references (well-known names, display names, paths or IDs) to Graph IDs.

Graph's folder endpoints accept either a canonical well-known name (e.g. `inbox`,
`junkemail`) or a full Graph folder ID, but not display names like "Junk Email" or
user-created names like "TLDR". This resolver lets callers pass any of those forms
and returns something the Graph API will accept.

A path names a nested folder by where it sits: `Inbox/Travel/Flights`. A "/"
inside a folder's own name is written "//", so `Inbox/Q1//Q2` is the folder
"Q1/Q2" under Inbox. Not a backslash escape: JSON reads a backslash before a
"/" as a plain "/", so a client encoding the arguments would drop it.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from typing import Any

from outlook_mcp.pagination import build_request_config
from outlook_mcp.validation import WELL_KNOWN_FOLDERS, sanitize_output, validate_graph_id

# How many candidate paths an ambiguity error lists before it just counts.
_MAX_CANDIDATE_PATHS = 5


class _NoSuchFolderError(ValueError):
    """A path's first part names no folder: none at the top of the tree, or, for a
    path of one escaped name, none anywhere.

    The one refusal after which a reference shaped like an id is tried as one. A
    path whose first folder exists is a path, however long and id-like its parts.
    """


async def fetch_all_top_level_folders(graph_client: Any, top: int = 100) -> list[Any]:
    """Fetch every top-level mail folder, following `@odata.nextLink` pagination."""
    from msgraph.generated.users.item.mail_folders.mail_folders_request_builder import (
        MailFoldersRequestBuilder,
    )

    config = build_request_config(
        MailFoldersRequestBuilder.MailFoldersRequestBuilderGetQueryParameters,
        {"$top": top},
    )
    response = await graph_client.me.mail_folders.get(request_configuration=config)
    collected = list(response.value) if response and response.value else []
    while getattr(response, "odata_next_link", None):
        response = await graph_client.me.mail_folders.with_url(response.odata_next_link).get()
        collected.extend(list(response.value) if response and response.value else [])
    return collected


async def fetch_all_child_folders(graph_client: Any, parent_id: str, top: int = 100) -> list[Any]:
    """Fetch every child folder of `parent_id`, following pagination."""
    from msgraph.generated.users.item.mail_folders.item.child_folders import (
        child_folders_request_builder as cf,
    )

    builder = graph_client.me.mail_folders.by_mail_folder_id(parent_id).child_folders
    config = build_request_config(
        cf.ChildFoldersRequestBuilder.ChildFoldersRequestBuilderGetQueryParameters,
        {"$top": top},
    )
    response = await builder.get(request_configuration=config)
    collected = list(response.value) if response and response.value else []
    while getattr(response, "odata_next_link", None):
        response = await builder.with_url(response.odata_next_link).get()
        collected.extend(list(response.value) if response and response.value else [])
    return collected


def folder_path_part(name: str) -> str:
    """A folder's display name as one part of a path: each "/" doubled."""
    return name.replace("/", "//")


def split_folder_path(ref: str) -> list[str]:
    """Split a path into display names: "/" separates, "//" is a "/" in a name.

    Read left to right, so "a///b" is "a/" then "b". A name that begins with
    "/" can therefore only be a path's first part; anywhere else it needs its id.
    Each part is stripped, and an empty part is refused.
    """
    parts: list[str] = []
    current: list[str] = []
    i = 0
    while i < len(ref):
        if ref.startswith("//", i):
            current.append("/")
            i += 2
        elif ref[i] == "/":
            parts.append("".join(current))
            current = []
            i += 1
        else:
            current.append(ref[i])
            i += 1
    parts.append("".join(current))
    parts = [part.strip() for part in parts]
    if not all(parts):
        raise ValueError(
            f"Folder path '{ref}' has an empty part. Separate folders with a single '/', "
            "and write a '/' inside a folder name as '//'."
        )
    return parts


async def walk_folder_tree(
    graph_client: Any, top_level: list[Any]
) -> AsyncIterator[tuple[Any, str]]:
    """Breadth-first over the tree below ``top_level``: each folder with its path."""
    queue = [(f, folder_path_part(f.display_name or "")) for f in top_level]
    while queue:
        folder, path = queue.pop(0)
        yield folder, path
        if (getattr(folder, "child_folder_count", 0) or 0) > 0:
            children = await fetch_all_child_folders(graph_client, folder.id)
            queue.extend((c, f"{path}/{folder_path_part(c.display_name or '')}") for c in children)


async def resolve_folder_id(graph_client: Any, folder_ref: str) -> str:
    """Resolve a folder reference to a Graph-acceptable identifier.

    Accepts (in priority order):
    - Canonical well-known names: `inbox`, `drafts`, `sentitems`, `deleteditems`,
      `junkemail`, `archive`, `outbox`.
    - Display aliases for well-known folders: `Inbox`, `Sent Items`, `Junk Email`,
      `Deleted Items`, `Drafts`, `Archive`, `Outbox` (case-insensitive; whitespace
      ignored).
    - Microsoft Graph folder IDs (base64-ish strings passed through after syntactic
      validation).
    - Display names of user-created folders (case-insensitive, looked up via
      `/me/mailFolders`, then the subfolders below it).
    - Paths, when the reference contains a "/": `Inbox/Travel/Flights`, with a
      "/" inside a name written "//". The first part is a well-known name or a
      top-level folder; each later part is one of the previous folder's
      children, so a path costs one listing per level below the top.

    Returns either a canonical well-known name or a Graph folder ID — both of which
    Graph's `by_mail_folder_id(...)` builder accepts.

    Raises ValueError if the reference is empty, ambiguous, or not found.
    """
    if not folder_ref or not folder_ref.strip():
        raise ValueError("Folder reference must not be empty")

    trimmed = folder_ref.strip()

    normalized = trimmed.lower().replace(" ", "")
    if normalized in WELL_KNOWN_FOLDERS:
        return normalized

    if "/" in trimmed:
        # Outside the try: an id has no empty part, so a path split_folder_path
        # refuses is refused as a path.
        parts = split_folder_path(trimmed)
        try:
            return await _resolve_path(graph_client, parts)
        except _NoSuchFolderError:
            # An id may contain "/" (the validator allows it), and until paths
            # existed a "/" meant an id. One whose first part names no folder
            # still goes through as an id; a path that went wrong below a folder
            # that exists keeps the resolver's "No folder X in Y".
            if _could_be_graph_id(trimmed):
                return trimmed
            raise

    if _looks_like_graph_id(trimmed):
        return validate_graph_id(trimmed)

    return await _lookup_display_name(graph_client, trimmed)


def _looks_like_graph_id(value: str) -> bool:
    """Heuristic: Graph folder IDs are long base64-url strings.

    Real folder IDs are 100+ chars; shorter alphanumeric strings like "TLDR"
    or "Receipts" are almost certainly display names. The heuristic triggers
    on either (a) base64 padding/variant characters (`=`, `+`) that can't
    appear in folder display names, or (b) length >= 40 (no realistic display
    name is that long). A reference with a "/" is a path, tried before this.
    """
    if any(c in value for c in "=+"):
        return True
    return len(value) >= 40


def _could_be_graph_id(value: str) -> bool:
    """A reference with the shape of a Graph id: 40+ characters of its alphabet."""
    if len(value) < 40:
        return False
    try:
        validate_graph_id(value)
    except ValueError:
        return False
    return True


def _only_match(
    folders: list[Any],
    name: str,
    where: str,
    hint: str = "",
    missing: type[ValueError] = ValueError,
) -> str:
    """The id of the one folder in ``folders`` called ``name`` (case-insensitive).

    ``missing`` is what is raised when no folder is called that.
    """
    target = name.lower()
    matches = [f for f in folders if f.display_name and f.display_name.lower() == target]
    if len(matches) == 1:
        return matches[0].id
    if len(matches) > 1:
        raise ValueError(
            f"Folder '{name}' {where} is ambiguous ({len(matches)} matches). "
            "Pass a Graph folder ID instead."
        )
    raise missing(
        f"No folder '{name}' {where}.{hint} Use outlook_list_folders(recursive=True) to see "
        "each folder's path, or pass a Graph folder ID."
    )


async def _resolve_path(graph_client: Any, parts: list[str]) -> str:
    """Walk a split path down from the top of the tree, one listing per level."""
    if len(parts) == 1:
        # A single name with an escaped "/" in it: looked up like any name.
        return await _lookup_display_name(graph_client, parts[0], missing=_NoSuchFolderError)

    first = parts[0].lower().replace(" ", "")
    if first in WELL_KNOWN_FOLDERS:
        folder_id = first
    else:
        folder_id = _only_match(
            await fetch_all_top_level_folders(graph_client),
            parts[0],
            "at the top level",
            " A path starts at the top of the folder tree, e.g. 'Inbox/Travel'.",
            missing=_NoSuchFolderError,
        )

    for depth in range(1, len(parts)):
        parent = "/".join(folder_path_part(p) for p in parts[:depth])
        children = await fetch_all_child_folders(graph_client, folder_id)
        folder_id = _only_match(children, parts[depth], f"in '{parent}'")
    return folder_id


async def _lookup_display_name(
    graph_client: Any, display_name: str, missing: type[ValueError] = ValueError
) -> str:
    """Look up a folder by display name (case-insensitive), walking subfolders.

    Prefers a top-level match when one exists. Falls back to a BFS walk through
    subfolders so names like "Domains" nested under "Receipts" still resolve.
    ``missing`` is what is raised when no folder has the name.
    """
    target = display_name.lower()

    top_level = await fetch_all_top_level_folders(graph_client)

    top_matches = [f for f in top_level if f.display_name and f.display_name.lower() == target]
    if len(top_matches) == 1:
        return top_matches[0].id
    if len(top_matches) > 1:
        raise ValueError(
            f"Folder name '{display_name}' is ambiguous "
            f"({len(top_matches)} top-level matches). Pass a Graph folder ID instead."
        )

    matches: list[tuple[Any, str]] = []
    async for f, path in walk_folder_tree(graph_client, top_level):
        if f.display_name and f.display_name.lower() == target:
            matches.append((f, path))

    if not matches:
        raise missing(
            f"Folder '{display_name}' not found. "
            "Use outlook_list_folders(recursive=True) to see the full folder tree, "
            "or pass a Graph folder ID."
        )

    if len(matches) > 1:
        shown = ", ".join(f"'{sanitize_output(p)}'" for _, p in matches[:_MAX_CANDIDATE_PATHS])
        if len(matches) > _MAX_CANDIDATE_PATHS:
            shown += f" and {len(matches) - _MAX_CANDIDATE_PATHS} more"
        raise ValueError(
            f"Folder name '{display_name}' is ambiguous "
            f"({len(matches)} matches across tree: {shown}). "
            "Pass one of those paths, or a Graph folder ID."
        )

    return matches[0][0].id
