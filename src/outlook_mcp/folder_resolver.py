"""Resolve folder references (well-known names, display names, or IDs) to Graph IDs.

Graph's folder endpoints accept either a canonical well-known name (e.g. `inbox`,
`junkemail`) or a full Graph folder ID, but not display names like "Junk Email" or
user-created names like "TLDR". This resolver lets callers pass any of those forms
and returns something the Graph API will accept.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.validation import WELL_KNOWN_FOLDERS, validate_graph_id


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
    - Display names of user-created top-level folders (case-insensitive, looked up
      via `/me/mailFolders`).

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

    if _looks_like_graph_id(trimmed):
        return validate_graph_id(trimmed)

    return await _lookup_display_name(graph_client, trimmed)


def _looks_like_graph_id(value: str) -> bool:
    """Heuristic: Graph folder IDs are long base64-url strings.

    Real folder IDs are 100+ chars; shorter alphanumeric strings like "TLDR"
    or "Receipts" are almost certainly display names. The heuristic triggers
    on either (a) base64 padding/variant characters (`=`, `+`, `/`) that can't
    appear in folder display names, or (b) length >= 40 (no realistic display
    name is that long).
    """
    if any(c in value for c in "=+/"):
        return True
    return len(value) >= 40


async def _lookup_display_name(graph_client: Any, display_name: str) -> str:
    """Look up a folder by display name (case-insensitive), walking subfolders.

    Prefers a top-level match when one exists. Falls back to a BFS walk through
    subfolders so names like "Domains" nested under "Receipts" still resolve.
    """
    target = display_name.lower()

    response = await graph_client.me.mail_folders.get()
    top_level = list(response.value) if response and response.value else []

    top_matches = [f for f in top_level if f.display_name and f.display_name.lower() == target]
    if len(top_matches) == 1:
        return top_matches[0].id
    if len(top_matches) > 1:
        raise ValueError(
            f"Folder name '{display_name}' is ambiguous "
            f"({len(top_matches)} top-level matches). Pass a Graph folder ID instead."
        )

    matches: list[Any] = []
    queue: list[Any] = list(top_level)
    while queue:
        f = queue.pop(0)
        if f.display_name and f.display_name.lower() == target:
            matches.append(f)
        if (getattr(f, "child_folder_count", 0) or 0) > 0:
            child_resp = await graph_client.me.mail_folders.by_mail_folder_id(
                f.id
            ).child_folders.get()
            queue.extend(list(child_resp.value) if child_resp and child_resp.value else [])

    if not matches:
        raise ValueError(
            f"Folder '{display_name}' not found. "
            "Use outlook_list_folders(recursive=True) to see the full folder tree, "
            "or pass a Graph folder ID."
        )

    if len(matches) > 1:
        raise ValueError(
            f"Folder name '{display_name}' is ambiguous "
            f"({len(matches)} matches across tree). Pass a Graph folder ID instead."
        )

    return matches[0].id
