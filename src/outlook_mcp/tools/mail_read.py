"""Mail read tools: list_inbox, read_message, read_messages, search_mail, list_folders."""

from __future__ import annotations

import json
from typing import Any
from urllib.parse import quote

import httpx

from outlook_mcp.folder_resolver import (
    fetch_all_child_folders,
    fetch_all_top_level_folders,
    resolve_folder_id,
)
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_datetime,
    validate_email,
    validate_graph_id,
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
GRAPH_TOKEN_SCOPE = "https://graph.microsoft.com/.default"
BATCH_URL = GRAPH_BASE + "$batch"
MAX_BATCH_SIZE = 20


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


# Fields stripped from the message-summary dict when ``concise=True``.
# Kept in sync with the docstrings on ``outlook_list_inbox`` /
# ``outlook_search_mail``.
_CONCISE_SUMMARY_DROP = ("preview", "categories")


def _concise_summary(summary: dict) -> dict:
    """Return a copy of a message-summary dict with bulky fields stripped."""
    return {k: v for k, v in summary.items() if k not in _CONCISE_SUMMARY_DROP}


def _format_message_summary(msg: Any) -> dict:
    """Convert Graph SDK message to summary dict.

    Module-level helper — also imported by thread tools (Tier 2).
    """
    from_addr = ""
    from_name = ""
    if msg.from_ and msg.from_.email_address:
        from_addr = msg.from_.email_address.address or ""
        from_name = msg.from_.email_address.name or ""

    flag_status = "notFlagged"
    if msg.flag and msg.flag.flag_status:
        flag_status = (
            msg.flag.flag_status.value
            if hasattr(msg.flag.flag_status, "value")
            else str(msg.flag.flag_status)
        )

    importance = "normal"
    if msg.importance:
        importance = (
            msg.importance.value if hasattr(msg.importance, "value") else str(msg.importance)
        )

    classification = ""
    ic = getattr(msg, "inference_classification", None)
    if ic:
        classification = ic.value if hasattr(ic, "value") else str(ic)

    return {
        "id": msg.id,
        "subject": sanitize_output(msg.subject or "(no subject)"),
        "from_email": from_addr,
        "from_name": sanitize_output(from_name),
        "received": str(msg.received_date_time or ""),
        "is_read": bool(msg.is_read),
        "importance": importance,
        "preview": sanitize_output(msg.body_preview or ""),
        "has_attachments": bool(msg.has_attachments),
        "categories": list(msg.categories or []),
        "flag": flag_status,
        "conversation_id": msg.conversation_id or "",
        "classification": classification,
    }


VALID_CLASSIFICATIONS = {"focused", "other"}


async def list_inbox(
    graph_client: Any,
    folder: str = "inbox",
    count: int = 25,
    unread_only: bool = False,
    from_address: str | None = None,
    after: str | None = None,
    before: str | None = None,
    skip: int = 0,
    cursor: str | None = None,
    classification: str | None = None,
    concise: bool = False,
) -> dict:
    """List messages in a folder.

    classification: filter by Focused Inbox classification — "focused" or "other".
    None means no filter (both).

    concise: when True, drop ``preview`` and ``categories`` from each message
    (smaller payload for triage scans). Default False preserves the existing shape.
    """
    count = _clamp(count, 1, 100)
    folder = await resolve_folder_id(graph_client, folder)

    query_params = apply_pagination({}, count, cursor)
    query_params["$orderby"] = "receivedDateTime desc"
    query_params["$select"] = (
        "id,subject,from,receivedDateTime,isRead,importance,"
        "bodyPreview,hasAttachments,categories,flag,conversationId,inferenceClassification"
    )

    # If cursor provided, it already set $skip — ignore the manual skip param
    if not cursor and skip:
        query_params["$skip"] = skip

    # Build filter with validated inputs
    filters = []
    if unread_only:
        filters.append("isRead eq false")
    if from_address:
        validate_email(from_address)
        safe_from = from_address.replace("'", "''")
        filters.append(f"from/emailAddress/address eq '{safe_from}'")
    if after:
        safe_after = validate_datetime(after)
        filters.append(f"receivedDateTime ge {safe_after}")
    if before:
        safe_before = validate_datetime(before)
        filters.append(f"receivedDateTime le {safe_before}")
    if classification is not None:
        if classification not in VALID_CLASSIFICATIONS:
            raise ValueError(
                f"Invalid classification '{classification}'. "
                f"Must be one of: {sorted(VALID_CLASSIFICATIONS)}"
            )
        filters.append(f"inferenceClassification eq '{classification}'")

    if filters:
        query_params["$filter"] = " and ".join(filters)

    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    req_config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.mail_folders.by_mail_folder_id(folder).messages.get(
        request_configuration=req_config
    )

    messages = [_format_message_summary(m) for m in (response.value or [])]
    if concise:
        messages = [_concise_summary(m) for m in messages]

    return {
        "messages": messages,
        "count": len(messages),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }


def _format_read_message_from_sdk(
    msg: Any, format: str, concise: bool, deferred_value: str | None, include_deferred_send: bool
) -> dict:
    """Build the ``read_message`` wire shape from an SDK ``Message`` object.

    Module-private — extracted from ``read_message`` so the bulk-read tool
    (``read_messages``) can share the same body / preview / sanitization
    logic via the ``_format_read_message_from_raw`` adapter below.
    """
    from_addr = ""
    from_name = ""
    if msg.from_ and msg.from_.email_address:
        from_addr = msg.from_.email_address.address or ""
        from_name = msg.from_.email_address.name or ""

    to_list = []
    for r in msg.to_recipients or []:
        if r.email_address:
            to_list.append({
                "name": sanitize_output(r.email_address.name or ""),
                "email": r.email_address.address or "",
            })

    cc_list = []
    for r in msg.cc_recipients or []:
        if r.email_address:
            cc_list.append({
                "name": sanitize_output(r.email_address.name or ""),
                "email": r.email_address.address or "",
            })

    body_text = ""
    body_html = None
    if msg.body:
        content = msg.body.content or ""
        if format in ("html", "full"):
            body_html = content
        if format in ("text", "full"):
            body_text = sanitize_output(content, multiline=True)

    attachments = []
    for att in msg.attachments or []:
        attachments.append({
            "id": att.id,
            "name": sanitize_output(att.name or ""),
            "size": att.size or 0,
        })

    importance = "normal"
    if msg.importance and hasattr(msg.importance, "value"):
        importance = msg.importance.value

    flag_status = "notFlagged"
    if msg.flag and msg.flag.flag_status and hasattr(msg.flag.flag_status, "value"):
        flag_status = msg.flag.flag_status.value

    result = {
        "id": msg.id,
        "subject": sanitize_output(msg.subject or "(no subject)"),
        "from_email": from_addr,
        "from_name": sanitize_output(from_name),
        "to": to_list,
        "cc": cc_list,
        "received": str(msg.received_date_time or ""),
        "body": body_text,
        "body_html": body_html,
        "is_read": bool(msg.is_read),
        "importance": importance,
        "has_attachments": bool(msg.has_attachments),
        "attachments": attachments,
        "categories": list(msg.categories or []),
        "flag": flag_status,
        "conversation_id": msg.conversation_id or "",
    }

    if concise:
        # Surface a compact preview drawn from whichever body content we have.
        preview_source = msg.body.content if msg.body and msg.body.content else ""
        result.pop("body", None)
        result.pop("body_html", None)
        result["body_preview"] = _make_body_preview(preview_source)

    if include_deferred_send:
        result["deferred_send_datetime"] = deferred_value

    return result


def _format_read_message_from_raw(
    raw: dict, format: str, concise: bool, include_deferred_send: bool
) -> dict:
    """Build the ``read_message`` wire shape from a raw Graph JSON dict.

    Used by the bulk-read tool (``read_messages``) which goes through
    Graph's ``$batch`` endpoint via raw httpx — the SDK ``Message`` object
    isn't available there. Output is byte-identical to
    ``_format_read_message_from_sdk`` for the same input message.

    Mirrors ``mail_delta._format_message_delta`` in spirit (dict-shape
    Graph response → outlook-mcp wire shape).
    """
    fa = raw.get("from") or {}
    from_ea = fa.get("emailAddress") or {}
    from_addr = from_ea.get("address") or ""
    from_name = from_ea.get("name") or ""

    to_list = []
    for r in raw.get("toRecipients") or []:
        ea = (r or {}).get("emailAddress") or {}
        to_list.append({
            "name": sanitize_output(ea.get("name") or ""),
            "email": ea.get("address") or "",
        })

    cc_list = []
    for r in raw.get("ccRecipients") or []:
        ea = (r or {}).get("emailAddress") or {}
        cc_list.append({
            "name": sanitize_output(ea.get("name") or ""),
            "email": ea.get("address") or "",
        })

    body = raw.get("body") or {}
    content = body.get("content") or ""
    body_text = ""
    body_html = None
    if format in ("html", "full"):
        body_html = content
    if format in ("text", "full"):
        body_text = sanitize_output(content, multiline=True)

    attachments = []
    for att in raw.get("attachments") or []:
        attachments.append({
            "id": att.get("id"),
            "name": sanitize_output(att.get("name") or ""),
            "size": att.get("size") or 0,
        })

    flag_status = "notFlagged"
    flag_node = raw.get("flag")
    if isinstance(flag_node, dict):
        flag_status = flag_node.get("flagStatus") or "notFlagged"

    result = {
        "id": raw.get("id"),
        "subject": sanitize_output(raw.get("subject") or "(no subject)"),
        "from_email": from_addr,
        "from_name": sanitize_output(from_name),
        "to": to_list,
        "cc": cc_list,
        "received": str(raw.get("receivedDateTime") or ""),
        "body": body_text,
        "body_html": body_html,
        "is_read": bool(raw.get("isRead")),
        "importance": raw.get("importance") or "normal",
        "has_attachments": bool(raw.get("hasAttachments")),
        "attachments": attachments,
        "categories": list(raw.get("categories") or []),
        "flag": flag_status,
        "conversation_id": raw.get("conversationId") or "",
    }

    if concise:
        result.pop("body", None)
        result.pop("body_html", None)
        result["body_preview"] = _make_body_preview(content)

    if include_deferred_send:
        from outlook_mcp.tools.mail_drafts import _PR_DEFERRED_SEND_TIME_ID

        # Graph normalizes the property tag's hex segment to lowercase
        # in the response (matches the SDK path in ``read_message``).
        target = _PR_DEFERRED_SEND_TIME_ID.lower()
        deferred_value: str | None = None
        for prop in raw.get("singleValueExtendedProperties") or []:
            if ((prop or {}).get("id") or "").lower() == target:
                deferred_value = prop.get("value")
                break
        result["deferred_send_datetime"] = deferred_value

    return result


def _make_body_preview(content: str, limit: int = 200) -> str:
    """Compact a body to a single-line ``body_preview`` of at most ``limit`` chars.

    Strips HTML to plain text via the same sanitizer used elsewhere, then
    collapses whitespace (newlines, tabs, runs of spaces) into single spaces.
    """
    plain = sanitize_output(content or "", multiline=True)
    collapsed = " ".join(plain.split())
    return collapsed[:limit]


async def read_message(
    graph_client: Any,
    message_id: str,
    format: str = "text",
    include_deferred_send: bool = False,
    concise: bool = False,
) -> dict:
    """Read a single message by ID.

    Pass ``include_deferred_send=True`` to surface the
    PR_DEFERRED_SEND_TIME extended property (the value Exchange reads to
    schedule deferred delivery) as ``deferred_send_datetime`` in the
    response. ``null`` if the property isn't set on the message.

    Pass ``concise=True`` to drop the full ``body`` / ``body_html`` fields
    and surface a single-line ``body_preview`` (first 200 chars, whitespace
    collapsed). Headers (from/to/cc/subject/received/importance/...) are
    preserved. Default False keeps the existing shape.
    """
    message_id = validate_graph_id(message_id)

    deferred_value: str | None = None

    if include_deferred_send:
        # PR_DEFERRED_SEND_TIME extended property fetch.
        #
        # Kiota's typed query-parameter encoder drops the inner
        # `$filter` from `$expand=singleValueExtendedProperties($filter=...)`
        # during URL building, so we build a raw URL via the
        # RAW_URL_KEY path-parameter slot and let the adapter execute it
        # through the normal auth + retry pipeline.
        from urllib.parse import quote

        from kiota_abstractions.method import Method
        from kiota_abstractions.request_information import RequestInformation
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError

        from outlook_mcp.tools.mail_drafts import _PR_DEFERRED_SEND_TIME_ID

        adapter = graph_client.me.messages.by_message_id(
            message_id
        ).request_adapter
        filter_expr = f"id eq '{_PR_DEFERRED_SEND_TIME_ID}'"
        # Keep `=`, space, and single quotes literal — Graph requires
        # them in the $filter clause and they're safe inside a query
        # parameter value. Hex/special chars in the property tag still
        # get percent-encoded.
        encoded_filter = quote(filter_expr, safe="= '")
        raw_url = (
            f"{adapter.base_url}/me/messages/{quote(message_id, safe='')}"
            f"?$expand=singleValueExtendedProperties("
            f"$filter={encoded_filter})"
        )

        info = RequestInformation()
        info.http_method = Method.GET
        info.path_parameters = {RequestInformation.RAW_URL_KEY: raw_url}
        info.url_template = "{+baseurl}"
        info.headers.try_add("Accept", "application/json")
        msg = await adapter.send_async(
            info,
            Message,
            {
                "4XX": ODataError,
                "5XX": ODataError,
            },
        )

        # Graph normalizes the property tag's hex segment to lowercase
        # in the response (`SystemTime 0x3fef`), even when we POST/PATCH
        # it as uppercase. Case-insensitive match.
        target = _PR_DEFERRED_SEND_TIME_ID.lower()
        for prop in msg.single_value_extended_properties or []:
            if (prop.id or "").lower() == target:
                deferred_value = prop.value
                break
    else:
        msg = await graph_client.me.messages.by_message_id(message_id).get()

    return _format_read_message_from_sdk(
        msg, format, concise, deferred_value, include_deferred_send
    )


def _build_read_subrequest_url(message_id: str, include_deferred_send: bool) -> str:
    """Build the per-sub-request URL path for ``read_messages``.

    Matches what ``read_message`` requests:

    - Default: ``/me/messages/{id}`` (all default fields).
    - ``include_deferred_send=True``: adds the singleValueExtendedProperties
      expansion filtered to ``PR_DEFERRED_SEND_TIME``, same as the raw URL
      ``read_message`` builds via Kiota's ``RAW_URL_KEY`` path.
    """
    safe_id = quote(message_id, safe="")
    if not include_deferred_send:
        return f"/me/messages/{safe_id}"

    from outlook_mcp.tools.mail_drafts import _PR_DEFERRED_SEND_TIME_ID

    filter_expr = f"id eq '{_PR_DEFERRED_SEND_TIME_ID}'"
    # Match ``read_message``: keep ``= '`` literal, percent-encode the rest.
    encoded_filter = quote(filter_expr, safe="= '")
    return (
        f"/me/messages/{safe_id}"
        f"?$expand=singleValueExtendedProperties($filter={encoded_filter})"
    )


async def read_messages(
    graph_client: Any,
    message_ids: list[str],
    format: str = "text",
    concise: bool = False,
    include_deferred_send: bool = False,
) -> dict:
    """Read up to 20 messages by ID in a single Graph ``$batch`` round-trip.

    Replaces N sequential ``read_message`` calls with one HTTP request.
    Each per-message entry in ``messages`` matches what ``read_message``
    returns for the same ``(format, concise, include_deferred_send)``
    combo, byte-for-byte. Ordering follows the input ``message_ids``
    list, not Graph's response order.

    Args:
        graph_client: A ``GraphClient`` instance. Needs ``credential`` for
            bearer-token minting (the SDK client isn't used here — raw
            httpx hits the ``$batch`` endpoint directly).
        message_ids: 1-20 Graph message IDs. Each is validated via
            ``validate_graph_id`` before the HTTP call.
        format: ``"text"`` (default), ``"html"``, or ``"full"`` — same
            semantics as ``read_message``.
        concise: When True, drop ``body``/``body_html`` per the v1.8.0
            concise contract; surface a single-line ``body_preview``.
        include_deferred_send: When True, fetch the
            PR_DEFERRED_SEND_TIME extended property on each message and
            surface it as ``deferred_send_datetime``.

    Returns:
        ``{messages, failures, requested, succeeded, failed}`` where
        ``messages`` is the list of read_message-shaped dicts in input
        order (failed IDs skipped), and ``failures`` is a list of
        ``{id, status, code, message}`` for sub-requests that didn't
        return 2xx. Input-validation errors raise ``ValueError``;
        transport-level errors raise ``httpx.HTTPStatusError`` (a 5xx on
        the whole batch is *not* a partial failure).
    """
    if not message_ids:
        raise ValueError("message_ids must not be empty")
    if len(message_ids) > MAX_BATCH_SIZE:
        raise ValueError(
            f"Maximum {MAX_BATCH_SIZE} messages per batch (Graph API limit)"
        )

    # Fail fast on malformed IDs *before* any HTTP work.
    for mid in message_ids:
        validate_graph_id(mid)

    credential = graph_client.credential

    subrequests = []
    for i, mid in enumerate(message_ids):
        subrequests.append({
            "id": str(i),
            "method": "GET",
            "url": _build_read_subrequest_url(mid, include_deferred_send),
        })

    batch_body = {"requests": subrequests}

    tok = credential.get_token(GRAPH_TOKEN_SCOPE)
    headers = {
        "Authorization": f"Bearer {tok.token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            BATCH_URL,
            headers=headers,
            content=json.dumps(batch_body),
        )
        resp.raise_for_status()
        payload = resp.json()

    # Rebuild input ordering. Graph's spec lets responses arrive in any
    # order; we used the input index as the sub-request id so reassembly
    # is trivial.
    responses_by_id: dict[str, dict] = {}
    for sub in payload.get("responses") or []:
        responses_by_id[str(sub.get("id"))] = sub

    messages: list[dict] = []
    failures: list[dict] = []

    for i, mid in enumerate(message_ids):
        sub = responses_by_id.get(str(i))
        if sub is None:
            failures.append({
                "id": mid,
                "status": 0,
                "code": "NoResponse",
                "message": "no sub-response returned by Graph for this id",
            })
            continue

        status = sub.get("status", 0)
        body = sub.get("body") or {}

        if 200 <= status < 300:
            messages.append(
                _format_read_message_from_raw(
                    body, format, concise, include_deferred_send
                )
            )
        else:
            err = body.get("error") or {}
            failures.append({
                "id": mid,
                "status": status,
                "code": err.get("code") or "",
                "message": err.get("message") or "",
            })

    return {
        "messages": messages,
        "failures": failures,
        "requested": len(message_ids),
        "succeeded": len(messages),
        "failed": len(failures),
    }


async def search_mail(
    graph_client: Any,
    query: str,
    count: int = 25,
    folder: str | None = None,
    cursor: str | None = None,
    concise: bool = False,
) -> dict:
    """Search mail using KQL.

    concise: when True, drop ``preview`` and ``categories`` from each message
    (smaller payload for triage scans). Default False preserves the existing shape.
    """
    count = _clamp(count, 1, 100)
    safe_query = sanitize_kql(query)

    query_params = apply_pagination({}, count, cursor)
    query_params["$search"] = safe_query
    query_params["$select"] = (
        "id,subject,from,receivedDateTime,isRead,importance,"
        "bodyPreview,hasAttachments,categories,flag,conversationId"
    )

    if folder:
        folder = await resolve_folder_id(graph_client, folder)
        from msgraph.generated.users.item.mail_folders.item.messages import (
            messages_request_builder as folder_mrb,
        )

        req_config = build_request_config(
            folder_mrb.MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters,
            query_params,
        )
        response = await graph_client.me.mail_folders.by_mail_folder_id(folder).messages.get(
            request_configuration=req_config
        )
    else:
        from msgraph.generated.users.item.messages.messages_request_builder import (
            MessagesRequestBuilder as MeMessagesRequestBuilder,
        )

        req_config = build_request_config(
            MeMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, query_params
        )
        response = await graph_client.me.messages.get(request_configuration=req_config)

    messages = [_format_message_summary(m) for m in (response.value or [])]
    if concise:
        messages = [_concise_summary(m) for m in messages]

    return {
        "messages": messages,
        "count": len(messages),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }


def _folder_to_dict(f: Any) -> dict:
    return {
        "id": f.id,
        "name": sanitize_output(f.display_name or ""),
        "total": f.total_item_count or 0,
        "unread": f.unread_item_count or 0,
        "parent_id": getattr(f, "parent_folder_id", None),
        "child_count": getattr(f, "child_folder_count", 0) or 0,
    }


async def list_folders(
    graph_client: Any,
    cursor: str | None = None,
    recursive: bool = False,
) -> dict:
    """List mail folders.

    Default: top-level folders only, paginated. Set `recursive=True` to return
    the full folder tree (BFS walk of subfolders) — pagination is disabled in
    recursive mode; all folders are returned in one response.
    """
    from msgraph.generated.users.item.mail_folders.mail_folders_request_builder import (
        MailFoldersRequestBuilder,
    )

    if recursive:
        collected: list[dict] = []
        queue: list[Any] = await fetch_all_top_level_folders(graph_client)
        while queue:
            f = queue.pop(0)
            collected.append(_folder_to_dict(f))
            if (getattr(f, "child_folder_count", 0) or 0) > 0:
                children = await fetch_all_child_folders(graph_client, f.id)
                queue.extend(children)

        return {
            "folders": collected,
            "count": len(collected),
            "has_more": False,
            "cursor": None,
        }

    query_params = apply_pagination({}, count=50, cursor=cursor)
    req_config = build_request_config(
        MailFoldersRequestBuilder.MailFoldersRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.mail_folders.get(request_configuration=req_config)

    folders = [_folder_to_dict(f) for f in (response.value or [])]
    return {
        "folders": folders,
        "count": len(folders),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }
