"""Mail read tools: list_inbox, read_message, search_mail, list_folders."""

from __future__ import annotations

from typing import Any

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


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


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
) -> dict:
    """List messages in a folder.

    classification: filter by Focused Inbox classification — "focused" or "other".
    None means no filter (both).
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

    return {
        "messages": messages,
        "count": len(messages),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }


async def read_message(
    graph_client: Any,
    message_id: str,
    format: str = "text",
) -> dict:
    """Read a single message by ID."""
    message_id = validate_graph_id(message_id)

    msg = await graph_client.me.messages.by_message_id(message_id).get()

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

    return {
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


async def search_mail(
    graph_client: Any,
    query: str,
    count: int = 25,
    folder: str | None = None,
    cursor: str | None = None,
) -> dict:
    """Search mail using KQL."""
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
