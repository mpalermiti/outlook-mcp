"""Mail thread tools: list_thread, copy_message."""

from __future__ import annotations

import re
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.pagination import build_request_config
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.tools.mail_read import _format_message_summary
from outlook_mcp.validation import sanitize_output, validate_graph_id


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


# Heuristic markers used to detect the start of quoted prior-message content.
# Order matters only for documentation; we match any of them per line.
_QUOTE_MARKERS = (
    # "On Mon, May 5, 2026 at 10:00 AM, Alice <alice@example.com> wrote:"
    re.compile(r"^\s*On .+ wrote:\s*$"),
    # "From: Alice <alice@example.com>" (RFC-style header preamble)
    re.compile(r"^\s*From: .+", re.IGNORECASE),
    # Outlook-style horizontal divider
    re.compile(r"^\s*-{3,}\s*Original Message\s*-{3,}\s*$", re.IGNORECASE),
)


def _strip_quoted_text(body: str) -> str:
    """Drop everything from the first quoted-reply marker onward.

    Heuristic: split on lines starting with ``On ... wrote:``, ``From: ...``,
    or ``----- Original Message -----``. Keep only the content above the first
    matching line. Trailing whitespace is trimmed. If no marker matches, the
    body is returned unchanged.
    """
    if not body:
        return body
    lines = body.splitlines()
    for i, line in enumerate(lines):
        for marker in _QUOTE_MARKERS:
            if marker.match(line):
                return "\n".join(lines[:i]).rstrip()
    return body


async def list_thread(
    graph_client: Any,
    conversation_id: str,
    count: int = 50,
    concise: bool = False,
) -> dict:
    """List messages in a conversation thread, chronological order.

    concise: when True, strip quoted prior-message text from each message's
    ``preview`` using a simple heuristic — any line matching ``On ... wrote:``,
    ``From: ...``, or ``----- Original Message -----`` and everything after
    is dropped. Also drops ``categories`` and the per-message preview becomes
    just the leading (top-posted) reply. Default False preserves the existing
    shape.
    """
    conversation_id = validate_graph_id(conversation_id)
    count = _clamp(count, 1, 100)

    query_params = {
        "$filter": f"conversationId eq '{conversation_id}'",
        "$orderby": "receivedDateTime asc",
        "$top": count,
        "$select": (
            "id,subject,from,receivedDateTime,isRead,importance,"
            "bodyPreview,hasAttachments,categories,flag,conversationId"
        ),
    }

    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    req_config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.messages.get(request_configuration=req_config)

    raw_messages = list(response.value or [])
    messages = [_format_message_summary(m) for m in raw_messages]
    if concise:
        for summary, raw in zip(messages, raw_messages):
            # Apply the heuristic to the raw body_preview *before* the
            # newline-flattening sanitizer would erase the structural cues
            # (`\nOn ... wrote:\n`) the heuristic relies on.
            raw_preview = raw.body_preview or ""
            stripped = _strip_quoted_text(raw_preview)
            summary["preview"] = sanitize_output(stripped)
            summary.pop("categories", None)

    return {
        "messages": messages,
        "count": len(messages),
    }


async def copy_message(
    graph_client: Any,
    message_id: str,
    folder: str,
    *,
    config: Config,
) -> dict:
    """Copy a message to a folder."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_copy_message")
    message_id = validate_graph_id(message_id)
    folder = await resolve_folder_id(graph_client, folder)

    from msgraph.generated.users.item.messages.item.copy.copy_post_request_body import (
        CopyPostRequestBody,
    )

    request_body = CopyPostRequestBody()
    request_body.destination_id = folder

    await graph_client.me.messages.by_message_id(message_id).copy.post(request_body)
    return {"status": "copied", "folder": folder}
