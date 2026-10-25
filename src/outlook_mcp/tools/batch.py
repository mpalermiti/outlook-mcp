"""Batch triage tool: process up to 20 messages in a single Graph /$batch call."""

from __future__ import annotations

import json
from typing import Any
from urllib.parse import quote

import httpx
from kiota_abstractions.api_error import APIError
from kiota_abstractions.method import Method
from kiota_abstractions.request_information import RequestInformation
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from outlook_mcp.config import Config
from outlook_mcp.errors import wrap_graph_error
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.throttle import retry_throttled_subrequests
from outlook_mcp.validation import validate_graph_id

_VALID_ACTIONS = ("move", "flag", "categorize", "mark_read")
_VALID_FLAG_STATUSES = ("flagged", "complete", "notFlagged")

# `value` for mark_read. Anything else used to mean False, so a value such as
# "yes" marked every message unread; now it is refused before anything is sent.
_MARK_READ_VALUES = {"true": True, "read": True, "false": False, "unread": False}

# The error mapping every msgraph request builder passes: a non-2xx envelope is
# raised as `ODataError` carrying Graph's code and message, which
# `_wrap_tool_errors` turns into the same `GraphAPIError` any SDK call produces.
_ERROR_MAPPING = {"XXX": ODataError}

_NOT_ANSWERED = (
    "The retried request did not complete or its reply could not be read, so it is not "
    "known whether this message was changed. Re-read it before retrying."
)


def _failed_round(requests: list[dict], reason: str) -> dict:
    """A retry round that failed, as per-message errors.

    An earlier round has already changed messages by then, so raising would
    hide which ones; only the re-sent requests are reported as failed.
    """
    return {
        "responses": [
            {"id": r["id"], "status": 0, "body": {"error": {"message": reason}}} for r in requests
        ]
    }


def _parse_mark_read(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    is_read = _MARK_READ_VALUES.get(str(value).strip().lower())
    if is_read is None:
        raise ValueError(
            f"mark_read value must be 'true' or 'false' ('read' and 'unread' also work); "
            f"got {value!r}"
        )
    return is_read


def _build_subrequest(index: int, message_id: str, action: str, value: Any) -> dict:
    """Map a (message_id, action, value) tuple to a Graph batch sub-request.

    The id is percent-encoded into the URL, exactly as the typed SDK encodes it
    on every other call and as `mail_read`'s own `$batch` does: an id may carry
    `/` or `+`, and unencoded a `/` would split it into two path segments.
    """
    request_id = str(index)
    if action == "move":
        return {
            "id": request_id,
            "method": "POST",
            "url": f"/me/messages/{quote(message_id, safe='')}/move",
            "body": {"destinationId": value},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "flag":
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{quote(message_id, safe='')}",
            "body": {"flag": {"flagStatus": value}},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "categorize":
        categories = [c.strip() for c in value.split(",")] if isinstance(value, str) else value
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{quote(message_id, safe='')}",
            "body": {"categories": categories},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "mark_read":
        is_read = _parse_mark_read(value)
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{quote(message_id, safe='')}",
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
        raise ValueError(f"flag value must be one of {_VALID_FLAG_STATUSES}; got {value!r}")

    resolved_value: Any = value
    if action == "move":
        resolved_value = await resolve_folder_id(graph_client, value)

    subrequests = [
        _build_subrequest(i, mid, action, resolved_value) for i, mid in enumerate(message_ids)
    ]

    rounds_sent = 0

    async def post_batch(requests: list[dict]) -> dict:
        nonlocal rounds_sent
        is_retry = rounds_sent > 0
        rounds_sent += 1

        req = RequestInformation()
        req.url = "https://graph.microsoft.com/v1.0/$batch"
        req.http_method = Method.POST
        req.headers.try_add("Content-Type", "application/json")
        req.headers.try_add("Accept", "application/json")
        req.content = json.dumps({"requests": requests}).encode("utf-8")
        try:
            # kiota retries a throttled envelope (429/503/504) itself; a refused
            # one is raised through `_ERROR_MAPPING` in Graph's own words.
            content = await graph_client.request_adapter.send_primitive_async(
                req, "bytes", _ERROR_MAPPING
            )
        except (APIError, httpx.HTTPError) as exc:
            if not is_retry:
                raise
            reason = str(wrap_graph_error(exc)) if isinstance(exc, APIError) else _NOT_ANSWERED
            return _failed_round(requests, reason)
        try:
            return json.loads(content or b"")
        except ValueError:  # JSONDecodeError, or bytes that are not UTF-8
            if not is_retry:
                raise
            return _failed_round(requests, _NOT_ANSWERED)

    # Graph answers the envelope with 200 even when it throttled sub-requests:
    # each one is checked against the mailbox's limits on its own and can come
    # back 429 with its own Retry-After, and the SDK's retry handler never sees
    # those. Re-send just those, as `outlook_read_messages` does.
    responses_by_id = await retry_throttled_subrequests(post_batch, subrequests)

    id_to_message = {str(i): mid for i, mid in enumerate(message_ids)}
    results_by_id: dict[str, dict] = {}
    for sub_id, sub in responses_by_id.items():
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
