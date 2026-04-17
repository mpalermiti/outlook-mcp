"""Batch triage tool: process up to 20 messages in one call."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.tools.mail_triage import (
    categorize_message,
    flag_message,
    mark_read,
    move_message,
)
from outlook_mcp.validation import validate_graph_id

_VALID_ACTIONS = ("move", "flag", "categorize", "mark_read")


async def batch_triage(
    graph_client: Any,
    message_ids: list[str],
    action: str,
    value: str,
    *,
    config: Config,
) -> dict:
    """Triage up to 20 messages in a single tool call.

    Loops through individual triage operations with per-item error handling.
    """
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_batch_triage")

    if len(message_ids) > 20:
        raise ValueError("Maximum 20 messages per batch (Graph API limit)")

    if action not in _VALID_ACTIONS:
        raise ValueError(f"action must be one of {_VALID_ACTIONS}; got {action!r}")

    # Validate all IDs upfront before processing any
    for mid in message_ids:
        validate_graph_id(mid)

    # Resolve folder upfront for move action — converts display names to IDs
    # once, so move_message's internal resolution short-circuits on each call.
    if action == "move":
        value = await resolve_folder_id(graph_client, value)

    results: list[dict] = []
    for mid in message_ids:
        try:
            if action == "move":
                await move_message(graph_client, mid, value, config=config)
            elif action == "flag":
                await flag_message(graph_client, mid, value, config=config)
            elif action == "categorize":
                categories = [c.strip() for c in value.split(",")]
                await categorize_message(graph_client, mid, categories, config=config)
            elif action == "mark_read":
                await mark_read(graph_client, mid, value.lower() == "true", config=config)
            results.append({"id": mid, "status": "success"})
        except Exception as e:
            results.append({"id": mid, "status": "error", "error": str(e)})

    success_count = sum(1 for r in results if r["status"] == "success")
    return {
        "results": results,
        "success_count": success_count,
        "failure_count": len(results) - success_count,
    }
