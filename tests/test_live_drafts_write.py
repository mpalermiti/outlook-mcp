"""Write-tier live guards for drafts in a thread: is the body where Outlook puts a reply?

The mock and wire tests prove what we send. Only Graph can say where a reply draft's
text ends up and what happens to the quoted original: ``comment`` on createReply joins
a multi-line text's lines with spaces on an HTML message, and ``message.body`` replaces
the quote (checked against an Exchange Online mailbox). So these create real reply and
forward drafts of recent Inbox messages, read each back, and delete it.

Nothing is sent: no test here calls a send, and every draft is deleted in a
``finally`` (``tests/test_live_write_never_sends.py`` holds that line). Every draft,
reply or forward, is addressed to ``live.write.guard@example.com`` with
LIVE_WRITE_SUBJECT as its subject, replacing the sender and "RE:" subject Graph picks
for a reply, so no draft names a real correspondent and anything a crash leaks is
greppable in Drafts. Each test checks the draft Graph made carries both.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v tests/test_live_drafts_write.py
"""

from __future__ import annotations

import re
from contextlib import asynccontextmanager

import pytest

from outlook_mcp.tools.mail_drafts import create_draft, delete_draft
from tests.conftest import LIVE_WRITE_SUBJECT

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]

_BODY_OPEN_TAG = re.compile(r"<body\b[^>]*>", re.IGNORECASE)
_GUARD_TO = ["live.write.guard@example.com"]
_TEXT = f"{LIVE_WRITE_SUBJECT}\nsecond line\n\nsecond paragraph\n\n- item one\n- item two"
_HTML = (
    f"<p>{LIVE_WRITE_SUBJECT}<br>second line</p><p>second paragraph</p>"
    "<ul><li>item one</li><li>item two</li></ul>"
)


async def _recent_inbox_message(sdk, content_type: str):
    """One of the 25 newest Inbox messages whose body is ``content_type``."""
    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    from outlook_mcp.pagination import build_request_config

    config = build_request_config(
        MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters,
        {"$top": 25, "$select": "id,body,conversationId"},
    )
    page = await sdk.me.mail_folders.by_mail_folder_id("inbox").messages.get(
        request_configuration=config
    )
    for message in page.value or []:
        if (
            message.body
            and message.body.content_type
            and message.body.content_type.value == content_type
        ):
            return message
    pytest.skip(f"None of the 25 newest Inbox messages has a {content_type} body")


@asynccontextmanager
async def _draft(client, config, **kwargs):
    created = await create_draft(
        client.sdk_client, to=_GUARD_TO, subject=LIVE_WRITE_SUBJECT, config=config, **kwargs
    )
    try:
        yield created["draft_id"]
    finally:
        await delete_draft(client.sdk_client, created["draft_id"], config=config)


async def _read(sdk, draft_id: str):
    from msgraph.generated.users.item.messages.item.message_item_request_builder import (
        MessageItemRequestBuilder,
    )

    from outlook_mcp.pagination import build_request_config

    config = build_request_config(
        MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters,
        {"$select": "body,isDraft,conversationId,subject,toRecipients,ccRecipients"},
    )
    return await sdk.me.messages.by_message_id(draft_id).get(request_configuration=config)


def _assert_addressed_to_the_guard(draft) -> None:
    """Ours, not Graph's: a reply would go to the original's sender, titled "RE: ..."."""
    assert [r.email_address.address for r in draft.to_recipients or []] == _GUARD_TO
    assert not draft.cc_recipients
    assert draft.subject == LIVE_WRITE_SUBJECT


class TestThreadDraftBodies:
    @pytest.mark.parametrize("action", ["reply_to_message_id", "forward_message_id"])
    async def test_plain_text_on_an_html_message_keeps_its_layout_above_the_quote(
        self, real_graph_client, live_write_config, action
    ):
        sdk = real_graph_client.sdk_client
        source = (await _recent_inbox_message(sdk, "html")).id

        # The draft Graph makes with no text: its quote is the reference.
        async with _draft(
            real_graph_client, live_write_config, body="", **{action: source}
        ) as empty_id:
            quote = (await _read(sdk, empty_id)).body.content
        async with _draft(
            real_graph_client, live_write_config, body=_TEXT, **{action: source}
        ) as draft_id:
            draft = await _read(sdk, draft_id)

        assert draft.is_draft is True
        _assert_addressed_to_the_guard(draft)
        assert draft.body.content_type.value == "html"
        content = draft.body.content
        opening = _BODY_OPEN_TAG.search(content)
        assert opening is not None, "Graph's reply draft has no <body> tag"
        # Right after <body>, as written, and the original quoted below, unchanged.
        assert content[opening.end() :].startswith(_HTML), content[opening.end() :][:300]
        assert content[opening.end() + len(_HTML) :] == quote[_BODY_OPEN_TAG.search(quote).end() :]

    async def test_plain_text_on_a_plain_text_message_stays_text_above_the_quote(
        self, real_graph_client, live_write_config
    ):
        sdk = real_graph_client.sdk_client
        source = await _recent_inbox_message(sdk, "text")

        async with _draft(
            real_graph_client, live_write_config, body="", reply_to_message_id=source.id
        ) as empty_id:
            quote = (await _read(sdk, empty_id)).body.content
        async with _draft(
            real_graph_client, live_write_config, body=_TEXT, reply_to_message_id=source.id
        ) as draft_id:
            draft = await _read(sdk, draft_id)

        assert draft.is_draft is True
        _assert_addressed_to_the_guard(draft)
        assert draft.conversation_id == source.conversation_id
        assert draft.body.content_type.value == "text"
        content = draft.body.content.replace("\r\n", "\n")
        assert content.startswith(_TEXT + "\n"), content[:300]
        assert content.endswith(quote.replace("\r\n", "\n").lstrip("\n"))
