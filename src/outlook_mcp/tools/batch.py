"""Batch triage tool: process up to 20 messages in a single Graph /$batch call."""

from __future__ import annotations

import json
from typing import Any

from kiota_abstractions.method import Method
from kiota_abstractions.request_information import RequestInformation
from opentelemetry import trace

from outlook_mcp.config import Config
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.validation import validate_graph_id

_VALID_ACTIONS = ("move", "flag", "categorize", "mark_read")
_VALID_FLAG_STATUSES = ("flagged", "complete", "notFlagged")


def _build_subrequest(index: int, message_id: str, action: str, value: Any) -> dict:
    """Map a (message_id, action, value) tuple to a Graph batch sub-request."""
    request_id = str(index)
    if action == "move":
        return {
            "id": request_id,
            "method": "POST",
            "url": f"/me/messages/{message_id}/move",
            "body": {"destinationId": value},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "flag":
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{message_id}",
            "body": {"flag": {"flagStatus": value}},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "categorize":
        categories = [c.strip() for c in value.split(",")] if isinstance(value, str) else value
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{message_id}",
            "body": {"categories": categories},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "mark_read":
        is_read = value.lower() == "true" if isinstance(value, str) else bool(value)
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{message_id}",
            "body": {"isRead": is_read},
            "headers": {"Content-Type": "application/json"},
        }
    raise ValueError(f"action must be one of {_VALID_ACTIONS}; got {action!r}")


async def batch_triage(
    graph_client: Any,
    message_ids: list[str],
    action: str,
    value: str,
    *,
    config: Config,
) -> dict:
    """Triage up to 20 messages in a single Graph /$batch request.

    Collapses N sequential HTTP calls into one round-trip to Graph — typically
    10-20× faster than per-message calls for large triage operations.
    """
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_batch_triage")

    if not message_ids:
        raise ValueError("message_ids must not be empty")
    if len(message_ids) > 20:
        raise ValueError("Maximum 20 messages per batch (Graph API limit)")
    if action not in _VALID_ACTIONS:
        raise ValueError(f"action must be one of {_VALID_ACTIONS}; got {action!r}")

    for mid in message_ids:
        validate_graph_id(mid)

    if action == "flag" and value not in _VALID_FLAG_STATUSES:
        raise ValueError(
            f"flag value must be one of {_VALID_FLAG_STATUSES}; got {value!r}"
        )

    resolved_value: Any = value
    if action == "move":
        resolved_value = await resolve_folder_id(graph_client, value)

    subrequests = [
        _build_subrequest(i, mid, action, resolved_value)
        for i, mid in enumerate(message_ids)
    ]
    batch_body = {"requests": subrequests}

    req = RequestInformation()
    req.url = "https://graph.microsoft.com/v1.0/$batch"
    req.http_method = Method.POST
    req.headers.try_add("Content-Type", "application/json")
    req.headers.try_add("Accept", "application/json")
    req.content = json.dumps(batch_body).encode("utf-8")

    tracer = trace.get_tracer("outlook_mcp.batch_triage")
    with tracer.start_as_current_span("batch_triage") as span:
        response = await graph_client.request_adapter.get_http_response_message(req, span)
    payload = response.json()

    id_to_message = {str(i): mid for i, mid in enumerate(message_ids)}
    results_by_id: dict[str, dict] = {}
    for sub in payload.get("responses", []):
        sub_id = str(sub.get("id"))
        mid = id_to_message.get(sub_id)
        if mid is None:
            continue
        status = sub.get("status", 0)
        if 200 <= status < 300:
            results_by_id[mid] = {"id": mid, "status": "success"}
        else:
            body = sub.get("body") or {}
            err = body.get("error") or {}
            message = err.get("message") or body.get("error_description") or f"HTTP {status}"
            results_by_id[mid] = {"id": mid, "status": "error", "error": message}

    results = [
        results_by_id.get(mid, {"id": mid, "status": "error", "error": "no response from batch"})
        for mid in message_ids
    ]
    success_count = sum(1 for r in results if r["status"] == "success")
    return {
        "results": results,
        "success_count": success_count,
        "failure_count": len(results) - success_count,
    }
