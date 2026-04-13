"""Batch triage tool: process up to 20 messages in one call."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.mail_triage import (
    categorize_message,
    flag_message,
    mark_read,
    move_message,
)
from outlook_mcp.validation import validate_folder_name, validate_graph_id

_VALID_ACTIONS = ("move", "flag", "categorize", "mark_read")


def _check_read_only(read_only: bool, tool_name: str) -> None:
    if read_only:
        raise ReadOnlyError(tool_name)


async def batch_triage(
    graph_client: Any,
    message_ids: list[str],
    action: str,
    value: str,
    read_only: bool = False,
) -> dict:
    """Triage up to 20 messages in a single tool call.

    Loops through individual triage operations with per-item error handling.
    """
    _check_read_only(read_only, "outlook_batch_triage")

    if len(message_ids) > 20:
        raise ValueError("Maximum 20 messages per batch (Graph API limit)")

    if action not in _VALID_ACTIONS:
        raise ValueError(f"action must be one of {_VALID_ACTIONS}; got {action!r}")

    # Validate all IDs upfront before processing any
    for mid in message_ids:
        validate_graph_id(mid)

    # Validate folder name upfront for move action
    if action == "move":
        validate_folder_name(value)

    results: list[dict] = []
    for mid in message_ids:
        try:
            if action == "move":
                await move_message(graph_client, mid, value)
            elif action == "flag":
                await flag_message(graph_client, mid, value)
            elif action == "categorize":
                categories = [c.strip() for c in value.split(",")]
                await categorize_message(graph_client, mid, categories)
            elif action == "mark_read":
                await mark_read(graph_client, mid, value.lower() == "true")
            results.append({"id": mid, "status": "success"})
        except Exception as e:
            results.append({"id": mid, "status": "error", "error": str(e)})

    success_count = sum(1 for r in results if r["status"] == "success")
    return {
        "results": results,
        "success_count": success_count,
        "failure_count": len(results) - success_count,
    }
