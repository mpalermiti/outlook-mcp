"""A reply's or forward's text reaches Graph with its layout, above the quoted original.

Graph's reply and forward actions take the new text as ``comment``, which Graph joins
with spaces on an HTML message (two paragraphs and a list arrive as one run-on line), or
as ``message.body``, which replaces the whole body, quote included. Both checked against
an Exchange Online mailbox. So ``outlook_reply`` and ``outlook_forward`` create the
draft empty, patch the text in above the quote, and send the draft.

Like ``test_sends_are_not_replayed.py``, these put a mock transport under the real
middleware, so each assertion is about the request Graph would receive.
"""

from __future__ import annotations

import json

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_abstractions.api_error import APIError
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory

from outlook_mcp.config import Config
from outlook_mcp.errors import GraphAPIError
from outlook_mcp.tools import mail_write

_CFG = Config(client_id="test")

# A createReply draft of an HTML message: Outlook's header block below an <hr>.
_HTML_QUOTE = (
    '<html><head></head><body><hr><div id="divRplyFwdMsg"><b>From:</b> Sender</div>'
    "<div><p>The original message.</p></div></body></html>"
)
# The same for a plain-text message: a line break, a rule, the header block.
_TEXT_QUOTE = "\r\n________________\r\nFrom: Sender\r\n\r\nThe original message.\r\n"
_TEXT = "First line.\nSecond line.\n\nSecond paragraph.\n\n- one\n- two"


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


class _Graph:
    """Answers by method and path suffix, and records every request in order."""

    def __init__(self, answers: dict[tuple[str, str], tuple[int, dict | None]]):
        self._answers = answers
        self.sent: list[httpx.Request] = []

    def _handle(self, request: httpx.Request) -> httpx.Response:
        self.sent.append(request)
        for (method, suffix), (status, body) in self._answers.items():
            if request.method == method and request.url.path.endswith(suffix):
                if body is None:
                    return httpx.Response(status)
                return httpx.Response(status, json=body)
        raise AssertionError(f"unexpected request: {request.method} {request.url.path}")

    def calls(self) -> list[str]:
        return [f"{r.method} {r.url.path.removeprefix('/v1.0')}" for r in self.sent]

    def body_of(self, method: str, suffix: str) -> dict:
        (request,) = [r for r in self.sent if r.method == method and r.url.path.endswith(suffix)]
        return json.loads(request.content) if request.content else {}

    def client(self) -> GraphServiceClient:
        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options,
            client=httpx.AsyncClient(transport=httpx.MockTransport(self._handle)),
        )
        auth = AzureIdentityAuthenticationProvider(
            _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))


def _draft(content: str, content_type: str = "html") -> dict:
    return {"id": "R1", "body": {"contentType": content_type, "content": content}}


def _graph(create: str, draft: dict, *, patch=(200, {}), send=(202, None)) -> _Graph:
    return _Graph(
        {
            ("POST", f"/me/messages/M1/{create}"): (201, draft),
            ("PATCH", "/me/messages/R1"): patch,
            ("POST", "/me/messages/R1/send"): send,
            ("DELETE", "/me/messages/R1"): (204, None),
        }
    )


def _error(status: int) -> tuple[int, dict]:
    return status, {"error": {"code": "ErrorInvalidRequest", "message": "No."}}


class TestReplyText:
    async def test_plain_text_reaches_graph_as_paragraphs_and_line_breaks(self):
        graph = _graph("createReply", _draft(_HTML_QUOTE))

        result = await mail_write.reply(graph.client(), "M1", _TEXT, config=_CFG)

        assert result == {"status": "replied", "reply_all": False}
        assert graph.calls() == [
            "POST /me/messages/M1/createReply",
            "PATCH /me/messages/R1",
            "POST /me/messages/R1/send",
        ]
        # Neither `comment` (flattened) nor `message.body` (drops the quote).
        assert graph.body_of("POST", "/createReply") == {}
        body = graph.body_of("PATCH", "/me/messages/R1")["body"]
        assert body["contentType"] == "html"
        assert body["content"] == _HTML_QUOTE.replace(
            "<body>",
            "<body><p>First line.<br>Second line.</p><p>Second paragraph.</p>"
            "<ul><li>one</li><li>two</li></ul>",
        )

    async def test_reply_all_goes_through_create_reply_all(self):
        graph = _graph("createReplyAll", _draft(_HTML_QUOTE))

        result = await mail_write.reply(graph.client(), "M1", "a\nb", reply_all=True, config=_CFG)

        assert result == {"status": "replied", "reply_all": True}
        assert graph.calls()[0] == "POST /me/messages/M1/createReplyAll"
        assert "<body><p>a<br>b</p><hr>" in graph.body_of("PATCH", "/R1")["body"]["content"]

    async def test_html_goes_above_the_quote_instead_of_replacing_it(self):
        """`message.body` on /reply sent the reply without the original."""
        graph = _graph("createReply", _draft(_HTML_QUOTE))

        await mail_write.reply(
            graph.client(), "M1", "<p>Hi <b>there</b></p>", is_html=True, config=_CFG
        )

        assert graph.body_of("POST", "/createReply") == {}
        assert graph.body_of("PATCH", "/R1")["body"]["content"] == _HTML_QUOTE.replace(
            "<body>", "<body><p>Hi <b>there</b></p>"
        )

    async def test_on_a_plain_text_message_the_reply_stays_plain_text(self):
        """`comment` on a plain-text message turned the reply into HTML as well."""
        graph = _graph("createReply", _draft(_TEXT_QUOTE, "text"))

        await mail_write.reply(graph.client(), "M1", _TEXT, config=_CFG)

        body = graph.body_of("PATCH", "/R1")["body"]
        assert body == {
            "contentType": "text",
            "content": _TEXT.replace("\n", "\r\n") + "\r\n" + _TEXT_QUOTE,
        }


class TestForwardText:
    async def test_a_comment_reaches_graph_as_paragraphs_with_the_recipients(self):
        graph = _graph("createForward", _draft(_HTML_QUOTE))

        result = await mail_write.forward(
            graph.client(), "M1", ["someone@contoso.com"], comment=_TEXT, config=_CFG
        )

        assert result == {"status": "forwarded"}
        assert graph.calls() == [
            "POST /me/messages/M1/createForward",
            "PATCH /me/messages/R1",
            "POST /me/messages/R1/send",
        ]
        # On `message` only: createForward answers 400 when both are given.
        create = graph.body_of("POST", "/createForward")
        assert set(create) == {"Message"}
        assert create["Message"]["toRecipients"] == [
            {"emailAddress": {"address": "someone@contoso.com"}}
        ]
        assert "body" not in create["Message"]
        content = graph.body_of("PATCH", "/R1")["body"]["content"]
        assert content.startswith("<html><head></head><body><p>First line.<br>Second line.</p>")

    async def test_without_a_comment_it_is_one_forward_call(self):
        graph = _Graph({("POST", "/me/messages/M1/forward"): (202, None)})

        result = await mail_write.forward(
            graph.client(), "M1", ["someone@contoso.com"], config=_CFG
        )

        assert result == {"status": "forwarded"}
        assert graph.calls() == ["POST /me/messages/M1/forward"]
        assert "Comment" not in graph.body_of("POST", "/forward")


class TestNothingHalfSent:
    async def test_a_failed_patch_deletes_the_draft_and_sends_nothing(self):
        graph = _graph("createReply", _draft(_HTML_QUOTE), patch=_error(400))

        # The SDK's error; the tool wrapper turns it into a GraphAPIError.
        with pytest.raises(APIError):
            await mail_write.reply(graph.client(), "M1", "x", config=_CFG)

        assert graph.calls() == [
            "POST /me/messages/M1/createReply",
            "PATCH /me/messages/R1",
            "DELETE /me/messages/R1",
        ]

    async def test_a_refused_send_deletes_the_draft(self):
        """A 4xx means the send did not happen: nothing is left behind to retry from."""
        graph = _graph("createReply", _draft(_HTML_QUOTE), send=_error(400))

        with pytest.raises(GraphAPIError):
            await mail_write.reply(graph.client(), "M1", "x", config=_CFG)

        assert graph.calls()[-1] == "DELETE /me/messages/R1"

    @pytest.mark.parametrize("status", [500, 503])
    async def test_a_send_in_doubt_leaves_the_draft_and_says_where_to_look(self, status):
        graph = _graph(
            "createReply",
            _draft(_HTML_QUOTE),
            send=(status, {"error": {"code": "ServiceUnavailable", "message": "Try later."}}),
        )

        with pytest.raises(GraphAPIError) as caught:
            await mail_write.reply(graph.client(), "M1", "x", config=_CFG)

        # It may have gone out: deleting the draft could remove the only trace of it.
        assert "DELETE /me/messages/R1" not in graph.calls()
        assert graph.calls().count("POST /me/messages/R1/send") == 1
        text = str(caught.value)
        assert "may already have been sent" in text, text
        assert "Sent Items" in text and "Drafts" in text, text
        assert "The draft's id is R1." in text, text

    async def test_a_send_cut_off_before_graph_answers_says_the_same(self):
        """A dropped connection is no answer at all: as much in doubt as a 5xx."""

        def drop_the_send(request: httpx.Request) -> httpx.Response:
            if request.method == "POST" and request.url.path.endswith("/R1/send"):
                graph.sent.append(request)
                raise httpx.RemoteProtocolError("Server disconnected", request=request)
            return handle(request)

        graph = _graph("createReply", _draft(_HTML_QUOTE))
        handle, graph._handle = graph._handle, drop_the_send

        with pytest.raises(GraphAPIError) as caught:
            await mail_write.reply(graph.client(), "M1", "x", config=_CFG)

        assert "DELETE /me/messages/R1" not in graph.calls()
        assert graph.calls().count("POST /me/messages/R1/send") == 1
        text = str(caught.value)
        assert "RemoteProtocolError" in text and "/R1/send" not in text, text
        assert "Check Sent Items" in text and "still in Drafts" in text, text
        assert "The draft's id is R1." in text, text
