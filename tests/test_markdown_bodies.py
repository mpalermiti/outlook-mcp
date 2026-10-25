"""`is_markdown` on the five tools that write a mail body.

Agents write Markdown. Sent as plain text it shows its asterisks; sent with
``is_html=True`` it is not HTML. ``is_markdown=True`` converts a safe subset
(``tools/_mail_body.markdown_to_html``) and escapes the rest.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from msgraph.generated.models.body_type import BodyType

from outlook_mcp import server as server_mod
from outlook_mcp.config import Config
from outlook_mcp.tools import mail_drafts, mail_write

_CFG = Config(client_id="test")
_MD = "**Hi** Sam,\n\n- [notes](https://example.com)\n- `x`"
_HTML = (
    "<p><strong>Hi</strong> Sam,</p>"
    '<ul><li><a href="https://example.com">notes</a></li><li><code>x</code></li></ul>'
)
_QUOTE = "<html><body><div>quoted</div></body></html>"


def _thread_client(content=_QUOTE, html=True):
    created = MagicMock(id="R1")
    created.body = MagicMock(content=content, content_type=BodyType.Html if html else BodyType.Text)
    source = MagicMock()
    for action in ("create_reply", "create_reply_all", "create_forward"):
        getattr(source, action).post = AsyncMock(return_value=created)
    draft = MagicMock()
    draft.patch = AsyncMock()
    draft.delete = AsyncMock()
    draft.send.post = AsyncMock()
    builders = {"AAMkAG123=": source, "R1": draft}
    client = MagicMock()
    client.me.messages.by_message_id = MagicMock(side_effect=lambda i: builders[i])
    return client, source, draft


class TestNewMessages:
    async def test_send_message(self):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        await mail_write.send_message(
            client, ["a@example.com"], "s", _MD, is_markdown=True, config=_CFG
        )

        body = client.me.send_mail.post.call_args[0][0].message.body
        assert (body.content, body.content_type) == (_HTML, BodyType.Html)

    async def test_create_draft(self):
        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=MagicMock(id="D1"))

        await mail_drafts.create_draft(
            client, ["a@example.com"], "s", _MD, is_markdown=True, config=_CFG
        )

        body = client.me.messages.post.call_args[0][0].body
        assert (body.content, body.content_type) == (_HTML, BodyType.Html)

    async def test_update_draft(self):
        builder = MagicMock()
        builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_drafts.update_draft(
            client, "AAMkAG123=", body=_MD, is_markdown=True, config=_CFG
        )

        body = builder.patch.call_args[0][0].body
        assert (body.content, body.content_type) == (_HTML, BodyType.Html)

    async def test_without_it_a_body_is_as_before(self):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        await mail_write.send_message(client, ["a@example.com"], "s", _MD, config=_CFG)

        body = client.me.send_mail.post.call_args[0][0].message.body
        assert (body.content, body.content_type) == (_MD, BodyType.Text)


class TestReplies:
    async def test_reply_puts_the_html_above_the_quote(self):
        client, _, draft = _thread_client()

        await mail_write.reply(client, "AAMkAG123=", _MD, is_markdown=True, config=_CFG)

        body = draft.patch.call_args[0][0].body
        assert body.content == _QUOTE.replace("<body>", f"<body>{_HTML}")
        assert body.content_type == BodyType.Html
        draft.send.post.assert_called_once()

    async def test_on_a_plain_text_thread_the_markdown_goes_as_written(self):
        client, _, draft = _thread_client(content="\r\n> quoted\r\n", html=False)

        await mail_write.reply(client, "AAMkAG123=", "**Hi**\n- a", is_markdown=True, config=_CFG)

        body = draft.patch.call_args[0][0].body
        assert (body.content, body.content_type) == (
            "**Hi**\r\n- a\r\n\r\n> quoted\r\n",
            BodyType.Text,
        )

    async def test_forward_comment(self):
        client, _, draft = _thread_client()

        await mail_write.forward(
            client, "AAMkAG123=", ["a@example.com"], comment=_MD, is_markdown=True, config=_CFG
        )

        assert f"<body>{_HTML}<div>" in draft.patch.call_args[0][0].body.content

    async def test_a_reply_draft_puts_the_html_above_the_quote(self):
        client, source, draft = _thread_client()

        await mail_drafts.create_draft(
            client,
            None,
            None,
            _MD,
            is_markdown=True,
            reply_to_message_id="AAMkAG123=",
            config=_CFG,
        )

        source.create_reply.post.assert_called_once()
        body = draft.patch.call_args[0][0].body
        assert body.content == _QUOTE.replace("<body>", f"<body>{_HTML}")
        assert body.content_type == BodyType.Html
        draft.send.post.assert_not_called()

    async def test_a_forward_draft_too(self):
        client, source, draft = _thread_client()

        await mail_drafts.create_draft(
            client,
            ["a@example.com"],
            None,
            _MD,
            is_markdown=True,
            forward_message_id="AAMkAG123=",
            config=_CFG,
        )

        source.create_forward.post.assert_called_once()
        assert f"<body>{_HTML}<div>" in draft.patch.call_args[0][0].body.content


class TestRefusals:
    @pytest.mark.parametrize(
        "call",
        [
            lambda c: mail_write.send_message(
                c, ["a@example.com"], "s", "b", is_html=True, is_markdown=True, config=_CFG
            ),
            lambda c: mail_write.reply(
                c, "AAMkAG123=", "b", is_html=True, is_markdown=True, config=_CFG
            ),
            lambda c: mail_drafts.create_draft(
                c, ["a@example.com"], "s", "b", is_html=True, is_markdown=True, config=_CFG
            ),
            lambda c: mail_drafts.create_draft(
                c,
                None,
                None,
                "b",
                is_html=True,
                is_markdown=True,
                reply_to_message_id="AAMkAG123=",
                config=_CFG,
            ),
            lambda c: mail_drafts.update_draft(
                c, "AAMkAG123=", body="b", is_html=True, is_markdown=True, config=_CFG
            ),
            lambda c: mail_drafts.update_draft(
                c, "AAMkAG123=", subject="s", is_html=True, is_markdown=True, config=_CFG
            ),
        ],
        ids=[
            "send_message",
            "reply",
            "create_draft",
            "create_draft_reply",
            "update_draft",
            "update_draft_no_body",
        ],
    )
    async def test_html_and_markdown_together_are_refused_before_any_call(self, call):
        client = MagicMock()

        with pytest.raises(ValueError, match="is_html or is_markdown, not both"):
            await call(client)

        assert not client.mock_calls


_TOOLS = {
    "outlook_send_message": ("mail_write", "send_message"),
    "outlook_reply": ("mail_write", "reply"),
    "outlook_forward": ("mail_write", "forward"),
    "outlook_create_draft": ("mail_drafts", "create_draft"),
    "outlook_update_draft": ("mail_drafts", "update_draft"),
}
_ARGS = {
    "outlook_send_message": {"to": ["a@example.com"], "subject": "s", "body": "b"},
    "outlook_reply": {"message_id": "AAMkAG123=", "body": "b"},
    "outlook_forward": {"message_id": "AAMkAG123=", "to": ["a@example.com"], "comment": "c"},
    "outlook_create_draft": {"to": ["a@example.com"], "subject": "s", "body": "b"},
    "outlook_update_draft": {"draft_id": "AAMkAG123=", "body": "b"},
}


class TestTheTools:
    async def test_exactly_the_five_tools_take_it(self):
        from mcp.client import Client

        async with Client(server_mod.mcp) as client:
            tools = (await client.list_tools()).tools

        taking = {t.name for t in tools if "is_markdown" in t.input_schema["properties"]}
        assert taking == set(_TOOLS)
        for tool in tools:
            if tool.name in _TOOLS:
                assert tool.input_schema["properties"]["is_markdown"]["default"] is False

    @pytest.mark.parametrize("tool", sorted(_TOOLS))
    async def test_each_tool_passes_it_through(self, tool):
        module, function = _TOOLS[tool]
        ctx = MagicMock()
        ctx.request_context.lifespan_context = {"auth": MagicMock(), "config": _CFG}
        impl = AsyncMock(return_value={"status": "ok"})

        with (
            patch.object(server_mod, "_get_graph_client", return_value=MagicMock()),
            patch.object(getattr(server_mod, module), function, impl),
        ):
            await getattr(server_mod, tool)(ctx, **_ARGS[tool], is_markdown=True)

        assert impl.call_args.kwargs["is_markdown"] is True
