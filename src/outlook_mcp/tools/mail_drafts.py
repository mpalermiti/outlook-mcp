"""Mail draft tools: list, create, update, send, delete."""

from __future__ import annotations

import contextlib
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.errors import MESSAGE_MAY_HAVE_BEEN_SENT
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_MAIL_DRAFTS, CATEGORY_MAIL_SEND, check_permission
from outlook_mcp.throttle import call_without_replay
from outlook_mcp.tools._mail_body import above_quote, check_body_format, message_body
from outlook_mcp.tools.mail_read import SUMMARY_SELECT, _format_message_summary
from outlook_mcp.validation import validate_datetime, validate_email, validate_graph_id

# MAPI tag PR_DEFERRED_SEND_TIME (0x3FEF, PtypTime) — the legacy extended
# property the Outlook transport reads to schedule deferred sending. Setting
# this on a draft and then sending it instructs Exchange to hold the message
# in the Outbox until the given UTC instant. This is the same mechanism the
# Outlook desktop client uses for "Delay Delivery", and it runs server-side
# so the client doesn't need to be online at the scheduled time.
_PR_DEFERRED_SEND_TIME_ID = "SystemTime 0x3FEF"


def _build_deferred_send_property(deferred_send_datetime: str, tz: str = "UTC"):
    """Return a SingleValueLegacyExtendedProperty for PR_DEFERRED_SEND_TIME.

    `deferred_send_datetime` must be ISO 8601. validate_datetime normalizes
    to UTC and rejects malformed input. Caller is responsible for passing
    a future timestamp — the server rejects past values with a 400.
    """
    from msgraph.generated.models.single_value_legacy_extended_property import (
        SingleValueLegacyExtendedProperty,
    )

    normalized = validate_datetime(deferred_send_datetime, tz)
    prop = SingleValueLegacyExtendedProperty()
    prop.id = _PR_DEFERRED_SEND_TIME_ID
    prop.value = normalized
    return prop, normalized


async def require_draft(graph_client: Any, message_id: str, tool_name: str) -> None:
    """Refuse a message that is not a draft, before a draft tool changes it.

    A draft is addressed by its message id, and every message in the mailbox
    has one. Left unchecked, the draft tools were general message tools under
    a category the README calls "drafts only": ``outlook_delete_draft`` made
    the same DELETE that ``outlook_delete_message(permanent=True)`` makes, on
    whatever id it was handed — a received message included.

    One GET for ``isDraft`` and nothing else. Only an explicit ``True`` passes:
    a message Graph does not call a draft is not ours to edit or delete here.
    """
    from msgraph.generated.users.item.messages.item.message_item_request_builder import (
        MessageItemRequestBuilder,
    )

    req_config = build_request_config(
        MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters,
        {"$select": "isDraft"},
    )
    message = await graph_client.me.messages.by_message_id(message_id).get(
        request_configuration=req_config
    )
    if getattr(message, "is_draft", None) is not True:
        raise ValueError(
            f"{tool_name} only works on drafts, and this message is not a draft. "
            "Nothing was changed."
        )


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
        "$select": SUMMARY_SELECT,
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


async def _create_in_thread(
    graph_client: Any,
    source_id: str,
    fields: Any,
    action: str,
) -> Any:
    """POST createReply / createReplyAll / createForward and return the draft.

    ``fields`` (recipients, subject, importance) ride on the action's
    ``message``. The body is left out: ``comment`` loses its line breaks on an
    HTML message and ``message.body`` replaces the quoted original
    (``_mail_body``), so the caller patches it in above the quote.
    """
    from msgraph.generated.users.item.messages.item.create_forward import (
        create_forward_post_request_body as fwd,
    )
    from msgraph.generated.users.item.messages.item.create_reply import (
        create_reply_post_request_body as rep,
    )
    from msgraph.generated.users.item.messages.item.create_reply_all import (
        create_reply_all_post_request_body as rep_all,
    )

    source = graph_client.me.messages.by_message_id(source_id)
    if action == "forward":
        request_body: Any = fwd.CreateForwardPostRequestBody()
        builder = source.create_forward
    elif action == "reply_all":
        request_body = rep_all.CreateReplyAllPostRequestBody()
        builder = source.create_reply_all
    else:
        request_body = rep.CreateReplyPostRequestBody()
        builder = source.create_reply
    request_body.message = fields
    return await builder.post(request_body)


async def create_draft(
    graph_client: Any,
    to: list[str] | None,
    subject: str | None,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    reply_to: list[str] | None = None,
    deferred_send_datetime: str | None = None,
    reply_to_message_id: str | None = None,
    reply_all: bool = False,
    forward_message_id: str | None = None,
    is_markdown: bool = False,
    *,
    config: Config,
) -> dict:
    """Create a draft message in the Drafts folder.

    Validates all email addresses, builds a Message object,
    and POSTs to /me/messages (which creates in Drafts).

    Pass ``reply_to_message_id`` (with ``reply_all`` for reply-all) or
    ``forward_message_id`` to create the draft in that message's thread
    instead, through createReply / createReplyAll / createForward: Graph
    quotes the original below the body and, for a reply, picks the
    recipients and subject. Recipients, subject, Reply-To and importance
    passed with it replace Graph's. Nothing is sent.

    Pass ``deferred_send_datetime`` (ISO 8601, UTC preferred) to schedule the
    draft for delayed delivery. The value is set as the
    PR_DEFERRED_SEND_TIME extended property; once the draft is sent
    (e.g. via outlook_send_draft), Exchange holds the message server-side
    until the given instant. This is the standard "Delay Delivery"
    mechanism used by Outlook desktop and survives client offline state.
    """
    check_permission(config, CATEGORY_MAIL_DRAFTS, "outlook_create_draft")
    check_body_format(is_html, is_markdown)

    if reply_to_message_id and forward_message_id:
        raise ValueError("Pass reply_to_message_id or forward_message_id, not both.")
    if reply_all and not reply_to_message_id:
        raise ValueError("reply_all is for a reply: pass reply_to_message_id with it.")
    source_id = reply_to_message_id or forward_message_id
    if source_id is None and (to is None or subject is None):
        raise ValueError(
            "A new draft needs `to` and `subject`. To answer a message, pass "
            "reply_to_message_id or forward_message_id."
        )
    if forward_message_id and not to:
        raise ValueError("A forward needs `to`: Graph has no recipients to pick for it.")
    if source_id is not None:
        source_id = validate_graph_id(source_id)

    # Validate all email addresses
    validated_to = [validate_email(e) for e in to] if to else []
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []
    validated_reply_to = [validate_email(e) for e in reply_to] if reply_to else []

    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.message import Message
    from msgraph.generated.models.recipient import Recipient

    def _make_recipient(email: str) -> Recipient:
        r = Recipient()
        r.email_address = EmailAddress()
        r.email_address.address = email
        return r

    msg = Message()
    if subject is not None:
        msg.subject = subject
    if source_id is None:
        msg.body = message_body(body, is_html, is_markdown)
    # `is not None`, not truthiness: an empty list is a request to clear the
    # recipients, and a reply-all draft arrives with Graph's own to/cc already on
    # it. `if cc:` treated `cc=[]` as "not mentioned", so the draft kept people the
    # caller had just removed — and sending it mailed them.
    if to is not None:
        msg.to_recipients = [_make_recipient(e) for e in validated_to]
    if cc is not None:
        msg.cc_recipients = [_make_recipient(e) for e in validated_cc]
    if bcc is not None:
        msg.bcc_recipients = [_make_recipient(e) for e in validated_bcc]
    if reply_to is not None:
        msg.reply_to = [_make_recipient(e) for e in validated_reply_to]

    importance_map = {
        "low": Importance.Low,
        "normal": Importance.Normal,
        "high": Importance.High,
    }
    msg.importance = importance_map.get(importance, Importance.Normal)

    deferred_prop = None
    deferred_normalized = None
    if deferred_send_datetime is not None:
        deferred_prop, deferred_normalized = _build_deferred_send_property(
            deferred_send_datetime, config.timezone
        )

    if source_id is None:
        if deferred_prop is not None:
            msg.single_value_extended_properties = [deferred_prop]
        # POST /me/messages creates a draft
        created = await graph_client.me.messages.post(msg)
        result = {
            "status": "created",
            "draft_id": created.id,
        }
    else:
        if forward_message_id:
            action = "forward"
        elif reply_all:
            action = "reply_all"
        else:
            action = "reply"
        created = await _create_in_thread(graph_client, source_id, msg, action)

        # What the create call does not carry: the body, above the quote Graph
        # wrote, and the deferred-send property, set the way update_draft sets it.
        patch = Message()
        if body:
            patch.body = above_quote(created.body, body, is_html, is_markdown=is_markdown)
        if deferred_prop is not None:
            patch.single_value_extended_properties = [deferred_prop]
        if body or deferred_prop is not None:
            new_draft = graph_client.me.messages.by_message_id(created.id)
            try:
                await new_draft.patch(patch)
            except Exception:
                # Leave no half-made reply behind (one without its body, or that
                # would send at once): the PATCH's error is the one to report.
                with contextlib.suppress(Exception):
                    await new_draft.delete()
                raise

        result = {
            "status": "created",
            "draft_id": created.id,
            "to": [r.email_address.address for r in created.to_recipients or [] if r.email_address],
            "cc": [r.email_address.address for r in created.cc_recipients or [] if r.email_address],
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
    is_markdown: bool = False,
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
    check_body_format(is_html, is_markdown)
    draft_id = validate_graph_id(draft_id)

    from msgraph.generated.models.message import Message

    msg = Message()

    if subject is not None:
        msg.subject = subject

    if body is not None:
        msg.body = message_body(body, is_html, is_markdown)

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
            prop, deferred_normalized = _build_deferred_send_property(
                deferred_send_datetime, config.timezone
            )
            msg.single_value_extended_properties = [prop]

    await require_draft(graph_client, draft_id, "outlook_update_draft")
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

    send_builder = graph_client.me.messages.by_message_id(draft_id).send
    await call_without_replay(
        lambda once: send_builder.post(request_configuration=once),
        outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT,
    )

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

    await require_draft(graph_client, draft_id, "outlook_delete_draft")
    await graph_client.me.messages.by_message_id(draft_id).delete()

    return {
        "status": "deleted",
        "draft_id": draft_id,
    }
