"""Mail write tools: send_message, reply, forward."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_MAIL_SEND, check_permission
from outlook_mcp.validation import validate_email, validate_graph_id


async def send_message(
    graph_client: Any,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    sensitivity: str = "normal",
    request_read_receipt: bool = False,
    reply_to: list[str] | None = None,
    *,
    config: Config,
) -> dict:
    """Send an email message.

    Validates all recipient addresses, builds a Graph Message object,
    and posts via graph_client.me.send_mail.post().
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_message")

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
    from msgraph.generated.models.sensitivity import Sensitivity
    from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
        SendMailPostRequestBody,
    )

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

    sensitivity_map = {
        "normal": Sensitivity.Normal,
        "personal": Sensitivity.Personal,
        "private": Sensitivity.Private,
        "confidential": Sensitivity.Confidential,
    }
    msg.sensitivity = sensitivity_map.get(sensitivity, Sensitivity.Normal)
    msg.is_read_receipt_requested = request_read_receipt

    request_body = SendMailPostRequestBody()
    request_body.message = msg
    request_body.save_to_sent_items = True

    await graph_client.me.send_mail.post(request_body)

    return {"status": "sent"}


async def reply(
    graph_client: Any,
    message_id: str,
    body: str,
    reply_all: bool = False,
    is_html: bool = False,
    *,
    config: Config,
) -> dict:
    """Reply to a message.

    Uses reply.post() or reply_all.post() based on the reply_all flag.
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_reply")
    message_id = validate_graph_id(message_id)

    from msgraph.generated.users.item.messages.item.reply.reply_post_request_body import (
        ReplyPostRequestBody,
    )
    from msgraph.generated.users.item.messages.item.reply_all.reply_all_post_request_body import (
        ReplyAllPostRequestBody,
    )

    msg_builder = graph_client.me.messages.by_message_id(message_id)

    if reply_all:
        request_body = ReplyAllPostRequestBody()
        request_body.comment = body
        await msg_builder.reply_all.post(request_body)
    else:
        request_body = ReplyPostRequestBody()
        request_body.comment = body
        await msg_builder.reply.post(request_body)

    return {"status": "replied", "reply_all": reply_all}


async def forward(
    graph_client: Any,
    message_id: str,
    to: list[str],
    comment: str | None = None,
    *,
    config: Config,
) -> dict:
    """Forward a message to one or more recipients."""
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_forward")
    message_id = validate_graph_id(message_id)
    validated_to = [validate_email(e) for e in to]

    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.recipient import Recipient
    from msgraph.generated.users.item.messages.item.forward.forward_post_request_body import (
        ForwardPostRequestBody,
    )

    def _make_recipient(email: str) -> Recipient:
        r = Recipient()
        r.email_address = EmailAddress()
        r.email_address.address = email
        return r

    request_body = ForwardPostRequestBody()
    request_body.to_recipients = [_make_recipient(e) for e in validated_to]
    if comment:
        request_body.comment = comment

    await graph_client.me.messages.by_message_id(message_id).forward.post(request_body)

    return {"status": "forwarded"}
