"""Mail delta tool: ``list_inbox_delta``.

Wraps ``GET /me/mailFolders/{folder}/messages/delta``.

First call (no ``delta_token``): full snapshot of the folder. Subsequent
calls (with ``delta_token``): only messages changed since that token. See
``_delta.py`` for the cursor/cap semantics that are shared with the
calendar and contacts delta tools.

Tokens are *not* persisted server-side. The caller (typically an agent
running on a schedule) stores the returned ``delta_token`` and passes it
back on the next call. Pattern matches the existing pagination ``cursor``.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import quote, urljoin

from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.tools._delta import GRAPH_BASE, fetch_delta_pages, format_delta_item
from outlook_mcp.validation import sanitize_output


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _addr_pair(node: dict | None) -> tuple[str, str]:
    """Extract (email, name) from a Graph ``recipient`` JSON node."""
    if not node:
        return "", ""
    ea = node.get("emailAddress") or {}
    return ea.get("address") or "", ea.get("name") or ""


def _flag_status(node: dict | None) -> str:
    if not node:
        return "notFlagged"
    return node.get("flagStatus") or "notFlagged"


def _format_message_delta(raw: dict) -> dict:
    """Build the wire-shape message summary from a raw Graph JSON message.

    Mirrors ``mail_read._format_message_summary`` field-for-field so
    callers don't have to branch on whether a given dict came from a
    delta call or a regular list call. Extra ``@`` annotations on the
    Graph response (e.g. ``@odata.etag``) are ignored.
    """
    from_email, from_name = _addr_pair(raw.get("from"))

    return {
        "id": raw.get("id"),
        "subject": sanitize_output(raw.get("subject") or "(no subject)"),
        "from_email": from_email,
        "from_name": sanitize_output(from_name),
        "received": raw.get("receivedDateTime") or "",
        "is_read": bool(raw.get("isRead")),
        "importance": raw.get("importance") or "normal",
        "preview": sanitize_output(raw.get("bodyPreview") or ""),
        "has_attachments": bool(raw.get("hasAttachments")),
        "categories": list(raw.get("categories") or []),
        "flag": _flag_status(raw.get("flag")),
        "conversation_id": raw.get("conversationId") or "",
        "classification": raw.get("inferenceClassification") or "",
    }


async def list_inbox_delta(
    graph_client: Any,
    folder: str = "inbox",
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List inbox (or any folder) changes since the last delta token.

    Args:
        graph_client: A ``GraphClient`` instance (delta tools need the
            ``credential`` attribute to mint raw bearer tokens — the SDK
            client is only used for folder name resolution on the first
            call).
        folder: Folder display name, well-known name, or Graph ID.
            Defaults to ``"inbox"``. Ignored when ``delta_token`` is set;
            the token already encodes the folder.
        page_size: Items per Graph page (1-100). Mapped to ``$top``.
        delta_token: Opaque cursor from a previous call. ``None`` on the
            first call.

    Returns:
        ``{messages, delta_token, has_more}`` where:

        - ``messages`` is the list of changed-message summaries. Live
          items use the same shape as ``outlook_list_inbox`` plus an
          ``is_deleted: False`` field. Tombstones are
          ``{id, is_deleted: True}`` — *no other fields* on a delete.
        - ``delta_token`` is either a new opaque cursor (continue
          syncing later) or ``None`` (no change since the previous
          token — the caller should reuse it).
        - ``has_more`` is ``True`` when the per-call safety cap stopped
          us mid-sync (the caller should pass ``delta_token`` back
          immediately to keep draining), ``False`` when the round is
          complete.
    """
    page_size = _clamp(page_size, 1, 100)

    sdk_client = graph_client.sdk_client
    credential = graph_client.credential

    initial_url = ""
    if not delta_token:
        # Only resolve the folder name on the first call — subsequent
        # calls reuse the deltaLink/nextLink URL Graph handed us, which
        # already contains the folder ID.
        folder_id = await resolve_folder_id(sdk_client, folder)
        initial_url = urljoin(
            GRAPH_BASE,
            f"me/mailFolders/{quote(folder_id, safe='')}/messages/delta",
        )
        initial_url = f"{initial_url}?$top={page_size}"

    raw_items, next_token, has_more = await fetch_delta_pages(
        credential,
        initial_url=initial_url,
        delta_token=delta_token,
        page_size=page_size,
    )

    messages = [format_delta_item(item, _format_message_delta) for item in raw_items]

    return {
        "messages": messages,
        "delta_token": next_token,
        "has_more": has_more,
    }
