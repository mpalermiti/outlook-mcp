"""A write that emails someone reaches Graph once, whatever Graph answers.

kiota's ``RetryHandler`` re-sends a request answered 429, 503 or 504, and POST is
in its ``DEFAULT_ALLOWED_METHODS``. A 503 or 504 does not say the first attempt
failed, only that its answer did not arrive: it may have been carried out. For
sendMail, reply, forward, a draft's send, an invitation and an RSVP, every
re-send is then another email to a real recipient, and a re-send of a draft that
did go out finds no draft, so the tool reports a send that happened as failed.

Mocks of the SDK cannot see this: the retry lives in the middleware they replace.
So, like ``test_me_rewrite_reaches_the_wire.py``, these build the client the way
``msgraph`` does, put a mock transport under the real middleware, and count what
reaches it, against whatever kiota the environment resolved.

Which tools are under test is not decided here: ``_tools_that_can_send``, the scan
``test_send_tools_are_marked_destructive.py`` uses, finds every tool gated on
``mail_send``, and each one it finds must have a case below or a named exception.
"""

from __future__ import annotations

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
from outlook_mcp.errors import DRAFT_MAY_HAVE_BEEN_CREATED, GraphAPIError
from outlook_mcp.tools import (
    calendar_attachments,
    calendar_write,
    mail_attachments,
    mail_drafts,
    mail_rules,
    mail_write,
)
from tests.test_send_tools_are_marked_destructive import _tools_that_can_send

_CFG = Config(client_id="test")
_TO = ["someone@contoso.com"]


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


class _Graph:
    """The far side of the transport: answers by path, records every request.

    ``answers`` maps a path suffix to the answers for it, in order; the last one
    repeats, so a retry keeps getting the same failure.
    """

    def __init__(self, answers: dict[str, list[tuple[int, dict, dict | None]]]):
        self._answers = answers
        self.sent: list[httpx.Request] = []

    def _handle(self, request: httpx.Request) -> httpx.Response:
        self.sent.append(request)
        for suffix, queue in self._answers.items():
            if request.url.path.endswith(suffix):
                status, headers, body = queue.pop(0) if len(queue) > 1 else queue[0]
                if body is None:
                    return httpx.Response(status, headers=headers)
                return httpx.Response(status, headers=headers, json=body)
        raise AssertionError(f"unexpected request: {request.method} {request.url.path}")

    def count(self, suffix: str) -> int:
        return sum(1 for r in self.sent if r.url.path.endswith(suffix))

    def client(self) -> GraphServiceClient:
        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options,
            client=httpx.AsyncClient(transport=httpx.MockTransport(self._handle)),
        )
        auth = AzureIdentityAuthenticationProvider(
            _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))


def _server_error(status: int) -> tuple[int, dict, dict]:
    # Retry-After keeps a replay cheap to watch: without it kiota waits 3-5 s a re-send.
    code = "ServiceUnavailable" if status == 503 else "GatewayTimeout"
    return status, {"Retry-After": "1"}, {"error": {"code": code, "message": "Try later."}}


async def _send_with_small_attachment(sdk, tmp_path, monkeypatch):
    attachment = tmp_path / "notes.txt"
    attachment.write_bytes(b"synthetic")
    return await mail_attachments.send_with_attachments(
        sdk,
        _TO,
        "s",
        "b",
        [str(attachment)],
        config=Config(client_id="test", attachments_dir=str(tmp_path)),
    )


async def _send_with_large_attachment(sdk, tmp_path, monkeypatch):
    attachment = tmp_path / "large.bin"
    with open(attachment, "wb") as f:
        f.truncate(mail_attachments._LARGE_FILE_THRESHOLD + 1)

    async def _uploaded(*args, **kwargs):
        return None

    # The chunks go out over a client of their own, not the SDK's; only the
    # draft's send, at the end, is the write under test.
    monkeypatch.setattr(mail_attachments, "_upload_large_file", _uploaded)
    return await mail_attachments.send_with_attachments(
        sdk,
        _TO,
        "s",
        "b",
        [str(attachment)],
        config=Config(client_id="test", attachments_dir=str(tmp_path)),
    )


# What createReply, createReplyAll and createForward answer: the new draft.
_THREAD_DRAFT = (
    201,
    {},
    {"id": "R1", "body": {"contentType": "html", "content": "<body>q</body>"}},
)

_LARGE_SEND_SETUP = {
    "/me/messages": [(201, {}, {"id": "D1"})],
    "/createUploadSession": [(200, {}, {"uploadUrl": "https://upload.example.com/s"})],
}

# Tools the scan finds that have no case below, each with the reason. Empty: every
# tool that can send is exercised.
NOT_EXERCISED: dict[str, str] = {
    "outlook_attach_to_event": (
        "nothing is sent at attach time; the POST itself is covered below, in "
        "test_a_duplicable_write_answered_5xx_is_not_sent_again"
    ),
    # The rule tools are gated on `mail_send` for what a rule can *do* later, not
    # because the call itself sends. Update and delete are idempotent by
    # construction — a repeated PATCH writes the same rule, a repeated DELETE
    # finds nothing — so a replay costs nothing worth a case here. Create is not:
    # it lives in SENDS, because a re-sent POST is a second rule.
    "outlook_update_rule": "a retried PATCH replaces the same rule with the same actions",
    "outlook_delete_rule": "a retried DELETE finds the rule already gone",
}

# (tool, path the write is sent to, the call, answers it needs first, where to check)
SENDS = [
    pytest.param(
        "outlook_send_message",
        "/me/sendMail",
        lambda sdk, *_: mail_write.send_message(sdk, _TO, "s", "b", config=_CFG),
        {},
        "Sent Items",
        id="send_message",
    ),
    # A reply, and a forward with a comment, are sent from the draft their text was
    # patched into (`mail_write._send_in_thread`): the draft's send is the write.
    pytest.param(
        "outlook_reply",
        "/me/messages/R1/send",
        lambda sdk, *_: mail_write.reply(sdk, "M1", "b", config=_CFG),
        {"/createReply": [_THREAD_DRAFT], "/me/messages/R1": [(200, {}, {})]},
        "Sent Items",
        id="reply",
    ),
    pytest.param(
        "outlook_reply",
        "/me/messages/R1/send",
        lambda sdk, *_: mail_write.reply(sdk, "M1", "b", reply_all=True, config=_CFG),
        {"/createReplyAll": [_THREAD_DRAFT], "/me/messages/R1": [(200, {}, {})]},
        "Sent Items",
        id="reply_all",
    ),
    pytest.param(
        "outlook_forward",
        "/me/messages/M1/forward",
        lambda sdk, *_: mail_write.forward(sdk, "M1", _TO, config=_CFG),
        {},
        "Sent Items",
        id="forward",
    ),
    pytest.param(
        "outlook_forward",
        "/me/messages/R1/send",
        lambda sdk, *_: mail_write.forward(sdk, "M1", _TO, comment="c", config=_CFG),
        {"/createForward": [_THREAD_DRAFT], "/me/messages/R1": [(200, {}, {})]},
        "Sent Items",
        id="forward_with_comment",
    ),
    pytest.param(
        "outlook_send_draft",
        "/me/messages/D1/send",
        lambda sdk, *_: mail_drafts.send_draft(sdk, "D1", config=_CFG),
        {},
        "Sent Items",
        id="send_draft",
    ),
    pytest.param(
        "outlook_send_with_attachments",
        "/me/sendMail",
        _send_with_small_attachment,
        {},
        "Sent Items",
        id="send_with_attachments_inline",
    ),
    pytest.param(
        "outlook_send_with_attachments",
        "/me/messages/D1/send",
        _send_with_large_attachment,
        _LARGE_SEND_SETUP,
        "Sent Items",
        id="send_with_attachments_upload_session",
    ),
    pytest.param(
        "outlook_create_event",
        "/me/events",
        lambda sdk, *_: calendar_write.create_event(
            sdk,
            "s",
            "2026-11-02T10:00:00Z",
            "2026-11-02T10:30:00Z",
            attendees=_TO,
            config=_CFG,
        ),
        {},
        "calendar",
        id="create_event",
    ),
    pytest.param(
        "outlook_rsvp",
        "/me/events/E1/accept",
        lambda sdk, *_: calendar_write.rsvp(sdk, "E1", "accept", config=_CFG),
        {},
        "event",
        id="rsvp_accept",
    ),
    pytest.param(
        "outlook_rsvp",
        "/me/events/E1/decline",
        lambda sdk, *_: calendar_write.rsvp(sdk, "E1", "decline", config=_CFG),
        {},
        "event",
        id="rsvp_decline",
    ),
    pytest.param(
        "outlook_rsvp",
        "/me/events/E1/tentativelyAccept",
        lambda sdk, *_: calendar_write.rsvp(sdk, "E1", "tentative", config=_CFG),
        {},
        "event",
        id="rsvp_tentative",
    ),
    # Not a send, and not idempotent either: a re-sent POST is a second rule, and
    # both of them keep running. The gate is `mail_send` because a rule can
    # forward. `where_to_check` is the listing the agent has to read before
    # creating again.
    pytest.param(
        "outlook_create_rule",
        "/messageRules",
        lambda sdk, *_: mail_rules.create_rule(
            sdk,
            config=_CFG,
            name="Rule",
            from_addresses=["someone@contoso.com"],
            mark_as_read=True,
            sequence=100,
        ),
        {},
        "List the rules",
        id="create_rule",
    ),
    pytest.param(
        "outlook_update_event",
        "/me/events/E1",
        lambda sdk, *_: calendar_write.update_event(sdk, "E1", attendees=_TO, config=_CFG),
        {},
        "event",
        id="update_event",
    ),
]


def test_every_tool_that_can_send_is_exercised():
    """A new tool gated on `mail_send` fails here until it has a case or a reason."""
    can_send = _tools_that_can_send()
    exercised = {case.values[0] for case in SENDS}
    assert NOT_EXERCISED.keys() <= can_send, "an exception names a tool that cannot send"
    assert exercised == can_send - NOT_EXERCISED.keys()


@pytest.mark.parametrize("status", [503, 504])
@pytest.mark.parametrize(("tool", "path", "send", "setup", "where_to_check"), SENDS)
async def test_a_send_answered_5xx_is_not_sent_again(
    status, tool, path, send, setup, where_to_check, tmp_path, monkeypatch
):
    graph = _Graph({**setup, path: [_server_error(status)]})

    with pytest.raises((GraphAPIError, APIError)) as caught:
        await send(graph.client(), tmp_path, monkeypatch)

    assert graph.count(path) == 1, (
        f"{path} reached Graph {graph.count(path)} times after a {status}: every "
        "re-send after the first is another email if the first one went out"
    )
    # The outcome is unknown, so the agent must not be told to retry blind.
    assert isinstance(caught.value, GraphAPIError), type(caught.value)
    text = str(caught.value)
    assert "may already" in text, text
    assert where_to_check in text, text
    assert "Retry after a short delay" not in text, text


async def test_an_event_without_attendees_does_not_mention_invitations():
    """Nothing was emailed, so the hint says only that the event may exist."""
    graph = _Graph({"/me/events": [_server_error(503)]})

    with pytest.raises(GraphAPIError) as caught:
        await calendar_write.create_event(
            graph.client(), "s", "2026-11-02T10:00:00Z", "2026-11-02T10:30:00Z", config=_CFG
        )

    assert graph.count("/me/events") == 1
    text = str(caught.value)
    assert "may already exist" in text, text
    assert "invitation" not in text, text


async def test_a_throttled_send_is_sent_again_once_graph_accepts():
    """A 429 is the one answer that says Graph did not process the request, so the
    write goes out again, once, and the tool reports the send that then happened."""
    throttled = (429, {"Retry-After": "0"}, {"error": {"code": "TooManyRequests", "message": "x"}})
    graph = _Graph({"/me/sendMail": [throttled, (202, {}, None)]})

    result = await mail_write.send_message(graph.client(), _TO, "s", "b", config=_CFG)

    assert result == {"status": "sent"}
    assert graph.count("/me/sendMail") == 2


async def test_reads_keep_kiotas_retries():
    """Control: the handler is still on for everything else, or the test above
    would pass for a client that never retries anything."""
    graph = _Graph({"/me": [_server_error(503), (200, {}, {"id": "probe"})]})

    await graph.client().me.get()

    assert graph.count("/me") == 2


# ── Writes that email no one but are not safe to repeat ───────────────────
#
# An attachment POST can leave a duplicate file on the event, and createReply /
# createReplyAll / createForward can leave a second draft. Neither sends mail, so
# they are not in SENDS, but a 503 or 504 that kiota answers with another POST
# is just as much a duplicate.


async def _attach_inline(sdk, tmp_path, monkeypatch):
    (tmp_path / "ticket.txt").write_bytes(b"synthetic")
    return await calendar_attachments.attach_to_event(
        sdk,
        "E1",
        ["ticket.txt"],
        config=Config(client_id="test", attachments_dir=str(tmp_path)),
    )


async def _draft_reply(sdk, tmp_path, monkeypatch):
    return await mail_drafts.create_draft(
        sdk, None, None, "b", reply_to_message_id="M1", config=_CFG
    )


async def _draft_reply_all(sdk, tmp_path, monkeypatch):
    return await mail_drafts.create_draft(
        sdk, None, None, "b", reply_to_message_id="M1", reply_all=True, config=_CFG
    )


async def _draft_forward(sdk, tmp_path, monkeypatch):
    return await mail_drafts.create_draft(sdk, _TO, None, "b", forward_message_id="M1", config=_CFG)


async def _reply_with_text(sdk, tmp_path, monkeypatch):
    return await mail_write.reply(sdk, "M1", "b", config=_CFG)


async def _reply_all_with_text(sdk, tmp_path, monkeypatch):
    return await mail_write.reply(sdk, "M1", "b", reply_all=True, config=_CFG)


async def _forward_with_comment(sdk, tmp_path, monkeypatch):
    return await mail_write.forward(sdk, "M1", _TO, comment="c", config=_CFG)


@pytest.mark.parametrize("status", [503, 504])
@pytest.mark.parametrize(
    ("path", "write", "where_to_check"),
    [
        pytest.param("/me/events/E1/attachments", _attach_inline, "attachments", id="attach_event"),
        pytest.param("/createReply", _draft_reply, "Drafts", id="draft_reply"),
        pytest.param("/createReplyAll", _draft_reply_all, "Drafts", id="draft_reply_all"),
        pytest.param("/createForward", _draft_forward, "Drafts", id="draft_forward"),
        pytest.param("/createReply", _reply_with_text, "Drafts", id="reply_text"),
        pytest.param("/createForward", _forward_with_comment, "Drafts", id="forward_text"),
    ],
)
async def test_a_duplicable_write_answered_5xx_is_not_sent_again(
    status, path, write, where_to_check, tmp_path, monkeypatch
):
    graph = _Graph({path: [_server_error(status)]})

    with pytest.raises(GraphAPIError) as caught:
        await write(graph.client(), tmp_path, monkeypatch)

    assert graph.count(path) == 1, (
        f"{path} reached Graph {graph.count(path)} times after a {status}"
    )
    text = str(caught.value)
    assert "may already" in text, text
    assert where_to_check in text, text
    assert "Retry after a short delay" not in text, text


@pytest.mark.parametrize(
    ("path", "send"),
    [
        pytest.param("/createReply", _reply_with_text, id="reply"),
        pytest.param("/createReplyAll", _reply_all_with_text, id="reply_all"),
        pytest.param("/createForward", _forward_with_comment, id="forward"),
    ],
)
async def test_a_reply_that_failed_at_its_draft_says_it_was_not_sent(
    path, send, tmp_path, monkeypatch
):
    """`outlook_reply` and `outlook_forward` write into a draft and send that. A 5xx on
    the draft's creation comes before any send, so the text says plainly that nothing
    went out, and that the draft may be left over in Drafts."""
    graph = _Graph({path: [_server_error(503)]})

    with pytest.raises(GraphAPIError) as caught:
        await send(graph.client(), tmp_path, monkeypatch)

    assert graph.count("/send") == 0
    text = str(caught.value)
    assert "The reply or forward was not sent" in text, text
    assert "leftover draft" in text and "Drafts" in text, text
    assert caught.value.action != DRAFT_MAY_HAVE_BEEN_CREATED


@pytest.mark.parametrize(
    ("path", "write"),
    [
        pytest.param("/createReply", _draft_reply, id="draft_reply"),
        pytest.param("/createReplyAll", _draft_reply_all, id="draft_reply_all"),
        pytest.param("/createForward", _draft_forward, id="draft_forward"),
    ],
)
async def test_a_draft_tool_keeps_its_check_drafts_hint(path, write, tmp_path, monkeypatch):
    """`outlook_create_draft` sends nothing at all, so its hint is only about the draft."""
    graph = _Graph({path: [_server_error(503)]})

    with pytest.raises(GraphAPIError) as caught:
        await write(graph.client(), tmp_path, monkeypatch)

    assert caught.value.action == DRAFT_MAY_HAVE_BEEN_CREATED
    assert "Check Drafts before creating it again" in str(caught.value)
    assert "not sent" not in str(caught.value)


async def test_an_event_attachment_goes_up_once_and_only_after_the_ownership_lookups(tmp_path):
    """Under a policy that withholds `mail_send`, whose calendar the event is on and
    whether it has attendees are settled before the file is sent: both ownership
    lookups and the attendee read reach Graph first, then the one POST, not repeated
    after its 503."""
    (tmp_path / "ticket.txt").write_bytes(b"synthetic")
    me = {"address": "me@contoso.com"}
    graph = _Graph(
        {
            # Not the default calendar, so its owner is compared with the default's:
            # the same address, so the calendar is the user's own.
            "/me/events/E1/calendar": [
                (200, {}, {"id": "C2", "isDefaultCalendar": False, "owner": me})
            ],
            "/me/calendar": [(200, {}, {"owner": me})],
            # The event has no attendees, so the policy lets the file through.
            "/me/events/E1": [(200, {}, {"id": "E1"})],
            "/me/events/E1/attachments": [_server_error(503)],
        }
    )
    policy = Config(
        client_id="test", attachments_dir=str(tmp_path), allow_categories=["calendar_write"]
    )

    with pytest.raises(GraphAPIError) as caught:
        await calendar_attachments.attach_to_event(
            graph.client(), "E1", ["ticket.txt"], config=policy
        )

    on_the_wire = [
        (r.method, r.url.path.removeprefix("/v1.0"), r.url.params.get("$select"))
        for r in graph.sent
    ]
    assert on_the_wire == [
        ("GET", "/me/events/E1/calendar", "id,isDefaultCalendar,owner"),
        ("GET", "/me/calendar", "owner"),
        ("GET", "/me/events/E1", None),
        ("POST", "/me/events/E1/attachments", None),
    ]
    assert "Read the event's attachments" in str(caught.value)
