"""Mail draft tools: list, create, update, send, delete."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_MAIL_DRAFTS, CATEGORY_MAIL_SEND, check_permission
from outlook_mcp.tools.mail_read import _format_message_summary
from outlook_mcp.validation import validate_datetime, validate_email, validate_graph_id

# MAPI tag PR_DEFERRED_SEND_TIME (0x3FEF, PtypTime) — the legacy extended
# property the Outlook transport reads to schedule deferred sending. Setting
# this on a draft and then sending it instructs Exchange to hold the message
# in the Outbox until the given UTC instant. This is the same mechanism the
# Outlook desktop client uses for "Delay Delivery", and it runs server-side
# so the client doesn't need to be online at the scheduled time.
_PR_DEFERRED_SEND_TIME_ID = "SystemTime 0x3FEF"


def _build_deferred_send_property(deferred_send_datetime: str):
    """Return a SingleValueLegacyExtendedProperty for PR_DEFERRED_SEND_TIME.

    `deferred_send_datetime` must be ISO 8601. validate_datetime normalizes
    to UTC and rejects malformed input. Caller is responsible for passing
    a future timestamp — the server rejects past values with a 400.
    """
    from msgraph.generated.models.single_value_legacy_extended_property import (
        SingleValueLegacyExtendedProperty,
    )

    normalized = validate_datetime(deferred_send_datetime)
    prop = SingleValueLegacyExtendedProperty()
    prop.id = _PR_DEFERRED_SEND_TIME_ID
    prop.value = normalized
    return prop, normalized


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
    reply_to: list[str] | None = None,
    deferred_send_datetime: str | None = None,
    *,
    config: Config,
) -> dict:
    """Create a draft message in the Drafts folder.

    Validates all email addresses, builds a Message object,
    and POSTs to /me/messages (which creates in Drafts).

    Pass ``deferred_send_datetime`` (ISO 8601, UTC preferred) to schedule the
    draft for delayed delivery. The value is set as the
    PR_DEFERRED_SEND_TIME extended property; once the draft is sent
    (e.g. via outlook_send_draft), Exchange holds the message server-side
    until the given instant. This is the standard "Delay Delivery"
    mechanism used by Outlook desktop and survives client offline state.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_create_draft")

    # Validate all email addresses
    validated_to = [validate_email(e) for e in to]
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []
    validated_reply_to = [validate_email(e) for e in reply_to] if reply_to else []

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
    if validated_reply_to:
        msg.reply_to = [_make_recipient(e) for e in validated_reply_to]

    importance_map = {
        "low": Importance.Low,
        "normal": Importance.Normal,
        "high": Importance.High,
    }
    msg.importance = importance_map.get(importance, Importance.Normal)

    deferred_normalized = None
    if deferred_send_datetime is not None:
        prop, deferred_normalized = _build_deferred_send_property(deferred_send_datetime)
        msg.single_value_extended_properties = [prop]

    # POST /me/messages creates a draft
    created = await graph_client.me.messages.post(msg)

    result = {
        "status": "created",
        "draft_id": created.id,
    }
    if deferred_normalized is not None:
        result["deferred_send_datetime"] = deferred_normalized
    return result


async def update_draft(
    graph_client: Any,
    draft_id: str,
    subject: str | None = None,
    body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    reply_to: list[str] | None = None,
    is_html: bool = False,
    deferred_send_datetime: str | None = None,
    *,
    config: Config,
) -> dict:
    """Update an existing draft message.

    Sends a PATCH with only the provided fields.
    Validates draft_id and any email addresses.

    Pass ``is_html=True`` when ``body`` is HTML. Required when overwriting a
    draft that was originally created as HTML (e.g. composed in the Outlook
    web/desktop UI) — PATCHing such a draft with a Text body is rejected by
    the consumer-Outlook MAPI store with ErrorAccessDenied / MapiSetProperties.

    Pass ``deferred_send_datetime`` (ISO 8601) to set or replace the
    PR_DEFERRED_SEND_TIME extended property on the draft, scheduling it
    for delayed delivery once it's sent. Pass an empty string to clear a
    previously-set deferred send time.
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
        msg.body.content_type = BodyType.Html if is_html else BodyType.Text

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

    if reply_to is not None:
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.recipient import Recipient

        validated_reply_to = [validate_email(e) for e in reply_to]

        def _make_reply_to_recipient(email: str) -> Recipient:
            r = Recipient()
            r.email_address = EmailAddress()
            r.email_address.address = email
            return r

        msg.reply_to = [_make_reply_to_recipient(e) for e in validated_reply_to]

    deferred_normalized = None
    if deferred_send_datetime is not None:
        if deferred_send_datetime == "":
            # Clearing the deferred-send property requires PATCHing an
            # empty value via the same extended-property channel — Graph
            # interprets a null/empty value as "remove this property".
            from msgraph.generated.models.single_value_legacy_extended_property import (
                SingleValueLegacyExtendedProperty,
            )
            prop = SingleValueLegacyExtendedProperty()
            prop.id = _PR_DEFERRED_SEND_TIME_ID
            prop.value = ""
            msg.single_value_extended_properties = [prop]
            deferred_normalized = ""
        else:
            prop, deferred_normalized = _build_deferred_send_property(deferred_send_datetime)
            msg.single_value_extended_properties = [prop]

    await graph_client.me.messages.by_message_id(draft_id).patch(msg)

    result = {
        "status": "updated",
        "draft_id": draft_id,
    }
    if deferred_normalized is not None:
        result["deferred_send_datetime"] = deferred_normalized
    return result


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
