"""Mail write tools: send_message, reply, forward."""

from __future__ import annotations

import contextlib
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.errors import DRAFT_MAY_HAVE_BEEN_SENT, MESSAGE_MAY_HAVE_BEEN_SENT, GraphAPIError
from outlook_mcp.permissions import CATEGORY_MAIL_SEND, check_permission
from outlook_mcp.throttle import call_without_replay
from outlook_mcp.tools._mail_body import above_quote, check_body_format, message_body
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
    request_read_receipt: bool = False,
    reply_to: list[str] | None = None,
    is_markdown: bool = False,
    *,
    config: Config,
) -> dict:
    """Send an email message.

    Validates all recipient addresses, builds a Graph Message object,
    and posts via graph_client.me.send_mail.post(). ``is_markdown`` sends the
    body as HTML converted from Markdown (``_mail_body.markdown_to_html``).
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_send_message")
    check_body_format(is_html, is_markdown)

    # Validate all email addresses
    validated_to = [validate_email(e) for e in to]
    validated_cc = [validate_email(e) for e in cc] if cc else []
    validated_bcc = [validate_email(e) for e in bcc] if bcc else []
    validated_reply_to = [validate_email(e) for e in reply_to] if reply_to else []

    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.message import Message
    from msgraph.generated.models.recipient import Recipient
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
    msg.body = message_body(body, is_html, is_markdown)
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

    # No `sensitivity` here on purpose. msgraph-sdk's Message has no such
    # field (the assignment used to hang a stray attribute kiota never saw),
    # and Graph rejects `"sensitivity"` in a POST body with 400
    # UnableToDeserializePostBody — it is a legacy MAPI property that is not
    # settable through this endpoint. The parameter never did anything.
    msg.is_read_receipt_requested = request_read_receipt

    request_body = SendMailPostRequestBody()
    request_body.message = msg
    request_body.save_to_sent_items = True

    await call_without_replay(
        lambda once: graph_client.me.send_mail.post(request_body, request_configuration=once),
        outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT,
    )

    return {"status": "sent"}


async def _send_in_thread(
    graph_client: Any,
    source_id: str,
    action: str,
    text: str,
    is_html: bool,
    to: list[Any] | None = None,
    is_markdown: bool = False,
) -> None:
    """Reply, reply-all or forward with ``text`` above the quoted original.

    Graph's reply and forward actions take the text as ``comment``, which Graph
    joins with spaces on an HTML message, so paragraphs, line breaks and lists
    arrive as one run-on paragraph, or as ``message.body``, which replaces the
    whole body and sends the reply without the original. So the draft is created
    empty (createReply / createReplyAll / createForward), Graph quotes the original
    into it, one PATCH puts the text above the quote (``_mail_body``), and the
    draft is sent.

    Nothing has gone out before the send. A failed PATCH, or a send Graph refused
    (a 4xx), deletes the draft and raises. After a 5xx, or a send cut off before
    Graph answered (a dropped connection, a timeout), the send may have happened,
    so the draft is left alone and the error says where to look and gives its id.
    """
    from msgraph.generated.models.message import Message
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
        # On `message` only: createForward answers 400 when both are given.
        request_body.message = Message()
        request_body.message.to_recipients = to
        create = source.create_forward
    elif action == "reply_all":
        request_body = rep_all.CreateReplyAllPostRequestBody()
        create = source.create_reply_all
    else:
        request_body = rep.CreateReplyPostRequestBody()
        create = source.create_reply
    draft = await create.post(request_body)

    new_draft = graph_client.me.messages.by_message_id(draft.id)
    try:
        if text:
            patch = Message()
            patch.body = above_quote(draft.body, text, is_html, is_markdown=is_markdown)
            await new_draft.patch(patch)
    except Exception:
        # The PATCH's error is the one to report, even if the delete fails too.
        with contextlib.suppress(Exception):
            await new_draft.delete()
        raise

    in_doubt = f"{DRAFT_MAY_HAVE_BEEN_SENT} The draft's id is {draft.id}."
    try:
        await call_without_replay(
            lambda once: new_draft.send.post(request_configuration=once),
            outcome_unknown=in_doubt,
        )
    except GraphAPIError as exc:
        if 400 <= exc.status_code < 500:
            # Refused, so not sent: leave nothing behind.
            with contextlib.suppress(Exception):
                await new_draft.delete()
        raise
    except Exception as exc:
        # No answer from Graph at all, which leaves the send as much in doubt as a
        # 5xx: the draft stays, and the agent is told where to look instead of
        # getting a bare crash. Only the exception's type is named; its text can
        # carry the request URL.
        raise GraphAPIError(
            0,
            "SendOutcomeUnknown",
            f"The send was cut off before Graph answered ({type(exc).__name__}), so "
            "the message may already have been sent.",
            action=(
                "Check Sent Items before sending it again; if it is not there, it is "
                f"still in Drafts. The draft's id is {draft.id}."
            ),
        ) from exc


async def reply(
    graph_client: Any,
    message_id: str,
    body: str,
    reply_all: bool = False,
    is_html: bool = False,
    is_markdown: bool = False,
    *,
    config: Config,
) -> dict:
    """Reply to a message, with ``body`` above the quoted original.

    Goes through a draft, as ``_send_in_thread`` explains: Graph's reply action
    would flatten plain text (``comment``) or drop the quote (``message.body``,
    the path ``is_html`` took from 1.18.0). On a plain-text message the reply is
    plain text too. ``is_html`` was accepted and ignored until 1.18.0, so every
    "HTML" reply went out as plain text.
    """
    check_permission(config, CATEGORY_MAIL_SEND, "outlook_reply")
    check_body_format(is_html, is_markdown)
    message_id = validate_graph_id(message_id)

    action = "reply_all" if reply_all else "reply"
    await _send_in_thread(graph_client, message_id, action, body, is_html, is_markdown=is_markdown)

    return {"status": "replied", "reply_all": reply_all}


async def forward(
    graph_client: Any,
    message_id: str,
    to: list[str],
    comment: str | None = None,
    is_markdown: bool = False,
    *,
    config: Config,
) -> dict:
    """Forward a message to one or more recipients.

    With a ``comment``, through a draft like ``reply``, so the comment keeps its
    paragraphs and line breaks above the forwarded message (``is_markdown``: the
    comment is Markdown). Without one, one forward call.
    """
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

    recipients = [_make_recipient(e) for e in validated_to]
    if comment:
        await _send_in_thread(
            graph_client, message_id, "forward", comment, False, recipients, is_markdown
        )
        return {"status": "forwarded"}

    request_body = ForwardPostRequestBody()
    request_body.to_recipients = recipients

    forward_builder = graph_client.me.messages.by_message_id(message_id).forward
    await call_without_replay(
        lambda once: forward_builder.post(request_body, request_configuration=once),
        outcome_unknown=MESSAGE_MAY_HAVE_BEEN_SENT,
    )

    return {"status": "forwarded"}
