"""Mail draft tools: list, create, update, send, delete."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_MAIL_DRAFTS, CATEGORY_MAIL_SEND, check_permission
from outlook_mcp.tools.mail_read import _format_message_summary
from outlook_mcp.validation import validate_email, validate_graph_id


async def list_drafts(
    graph_client: Any,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List messages in the Drafts folder.

    Uses cursor-based pagination via apply_pagination / wrap_nextlink.
    Reuses _format_message_summary from mail_read.
    """
    query_params: dict[str, Any] = {
        "$orderby": "lastModifiedDateTime desc",
        "$select": (
            "id,subject,from,receivedDateTime,isRead,importance,"
            "bodyPreview,hasAttachments,categories,flag,conversationId"
        ),
    }
    query_params = apply_pagination(query_params, count, cursor)

    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    req_config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.mail_folders.by_mail_folder_id("drafts").messages.get(
        request_configuration=req_config
    )

    messages = [_format_message_summary(m) for m in (response.value or [])]
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "messages": messages,
        "count": len(messages),
        "next_cursor": next_cursor,
    }


async def create_draft(
    graph_client: Any,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    *,
    config: Config,
) -> dict:
    """Create a draft message in the Drafts folder.

    Validates all email addresses, builds a Message object,
    and POSTs to /me/messages (which creates in Drafts).
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_create_draft")

    # Validate all email addresses
    validated_to = [validate_email(e) for e in to]
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []

    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.message import Message
    from msgraph.generated.models.recipient import Recipient

    def _make_recipient(email: str) -> Recipient:
        r = Recipient()
        r.email_address = EmailAddress()
        r.email_address.address = email
        return r

    msg = Message()
    msg.subject = subject
    msg.body = ItemBody()
    msg.body.content = body
    msg.body.content_type = BodyType.Html if is_html else BodyType.Text
    msg.to_recipients = [_make_recipient(e) for e in validated_to]
    if validated_cc:
        msg.cc_recipients = [_make_recipient(e) for e in validated_cc]
    if validated_bcc:
        msg.bcc_recipients = [_make_recipient(e) for e in validated_bcc]

    importance_map = {
        "low": Importance.Low,
        "normal": Importance.Normal,
        "high": Importance.High,
    }
    msg.importance = importance_map.get(importance, Importance.Normal)

    # POST /me/messages creates a draft
    created = await graph_client.me.messages.post(msg)

    return {
        "status": "created",
        "draft_id": created.id,
    }


async def update_draft(
    graph_client: Any,
    draft_id: str,
    subject: str | None = None,
    body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    *,
    config: Config,
) -> dict:
    """Update an existing draft message.

    Sends a PATCH with only the provided fields.
    Validates draft_id and any email addresses.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_update_draft")
    draft_id = validate_graph_id(draft_id)

    from msgraph.generated.models.message import Message

    msg = Message()

    if subject is not None:
        msg.subject = subject

    if body is not None:
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.item_body import ItemBody

        msg.body = ItemBody()
        msg.body.content = body
        msg.body.content_type = BodyType.Text

    if to is not None:
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.recipient import Recipient

        validated_to = [validate_email(e) for e in to]

        def _make_recipient(email: str) -> Recipient:
            r = Recipient()
            r.email_address = EmailAddress()
            r.email_address.address = email
            return r

        msg.to_recipients = [_make_recipient(e) for e in validated_to]

    if cc is not None:
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.recipient import Recipient

        validated_cc = [validate_email(e) for e in cc]

        def _make_cc_recipient(email: str) -> Recipient:
            r = Recipient()
            r.email_address = EmailAddress()
            r.email_address.address = email
            return r

        msg.cc_recipients = [_make_cc_recipient(e) for e in validated_cc]

    await graph_client.me.messages.by_message_id(draft_id).patch(msg)

    return {
        "status": "updated",
        "draft_id": draft_id,
    }


async def send_draft(
    graph_client: Any,
    draft_id: str,
    *,
    config: Config,
) -> dict:
    """Send an existing draft message.

    POSTs to /me/messages/{id}/send.
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_draft")
    draft_id = validate_graph_id(draft_id)

    await graph_client.me.messages.by_message_id(draft_id).send.post()

    return {
        "status": "sent",
        "draft_id": draft_id,
    }


async def delete_draft(
    graph_client: Any,
    draft_id: str,
    *,
    config: Config,
) -> dict:
    """Delete a draft message.

    DELETEs /me/messages/{id}. Permanent delete since drafts
    don't need soft-delete semantics.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_delete_draft")
    draft_id = validate_graph_id(draft_id)

    await graph_client.me.messages.by_message_id(draft_id).delete()

    return {
        "status": "deleted",
        "draft_id": draft_id,
    }
