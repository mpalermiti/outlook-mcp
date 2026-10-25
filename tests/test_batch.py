"""Tests for batch triage tool (Graph /$batch endpoint)."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.batch import batch_triage

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _make_batch_response(statuses: list[int], errors: list[str | None] | None = None) -> bytes:
    """Build a Graph /$batch reply body with the given per-sub-request statuses."""
    errors = errors or [None] * len(statuses)
    responses = []
    for i, status in enumerate(statuses):
        entry = {"id": str(i), "status": status}
        if errors[i] is not None:
            entry["body"] = {"error": {"code": "Err", "message": errors[i]}}
        else:
            entry["body"] = {}
        responses.append(entry)
    return _payload_response(responses)


def _payload_response(responses: list[dict]) -> bytes:
    """The bytes `send_primitive_async(..., "bytes", ...)` returns for a 2xx batch."""
    return json.dumps({"responses": responses}).encode("utf-8")


def _mock_client_with_batch_response(content: bytes) -> MagicMock:
    mock_client = MagicMock()
    mock_client.request_adapter.send_primitive_async = AsyncMock(return_value=content)
    return mock_client


class TestBatchTriageValidation:
    async def test_read_only_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="move",
                value="inbox",
                config=_CFG_RO,
            )

    async def test_empty_message_ids_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="must not be empty"):
            await batch_triage(
                mock_client,
                message_ids=[],
                action="move",
                value="inbox",
                config=_CFG,
            )

    async def test_max_20_messages(self):
        mock_client = MagicMock()
        ids = [f"AAMkAG{i}=" for i in range(21)]
        with pytest.raises(ValueError, match="Maximum 20"):
            await batch_triage(
                mock_client,
                message_ids=ids,
                action="move",
                value="inbox",
                config=_CFG,
            )

    async def test_exactly_20_messages_allowed(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200] * 20))
        ids = [f"AAMkAG{i}=" for i in range(20)]
        result = await batch_triage(
            mock_client,
            message_ids=ids,
            action="move",
            value="inbox",
            config=_CFG,
        )
        assert result["success_count"] == 20
        assert result["failure_count"] == 0

    async def test_invalid_action_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="action must be one of"):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="delete",
                value="inbox",
                config=_CFG,
            )

    async def test_invalid_message_id_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await batch_triage(
                mock_client,
                message_ids=["<script>alert(1)</script>"],
                action="move",
                value="inbox",
                config=_CFG,
            )

    async def test_invalid_flag_status_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="flag value must be one of"):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="flag",
                value="bogus",
                config=_CFG,
            )

    async def test_move_validates_folder_name(self):
        mock_client = MagicMock()
        empty_response = MagicMock()
        empty_response.value = []
        empty_response.odata_next_link = None
        mock_client.me.mail_folders.get = AsyncMock(return_value=empty_response)
        with pytest.raises(ValueError, match="not found"):
            await batch_triage(
                mock_client,
                message_ids=["AAMkAG123="],
                action="move",
                value="<bad folder>",
                config=_CFG,
            )


class TestBatchTriageSingleRequest:
    """One /$batch call per batch_triage invocation — regardless of batch size."""

    async def test_single_http_call_for_move(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([201, 201, 201]))
        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2=", "AAMkAG3="],
            action="move",
            value="archive",
            config=_CFG,
        )
        assert result["success_count"] == 3
        mock_client.request_adapter.send_primitive_async.assert_called_once()

    async def test_batch_body_shape(self):
        """Verify the batch request body matches Graph's expected schema."""
        mock_client = _mock_client_with_batch_response(_make_batch_response([201, 201]))
        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="move",
            value="inbox",
            config=_CFG,
        )
        call = mock_client.request_adapter.send_primitive_async.call_args
        req = call.args[0]
        assert req.url == "https://graph.microsoft.com/v1.0/$batch"
        body = json.loads(req.content.decode("utf-8"))
        assert len(body["requests"]) == 2
        assert body["requests"][0]["method"] == "POST"
        assert body["requests"][0]["url"] == "/me/messages/AAMkAG1%3D/move"
        assert body["requests"][0]["body"] == {"destinationId": "inbox"}
        assert body["requests"][1]["url"] == "/me/messages/AAMkAG2%3D/move"

    @pytest.mark.parametrize(
        "action,value",
        [("move", "inbox"), ("flag", "flagged"), ("categorize", "Red"), ("mark_read", True)],
    )
    async def test_an_id_stays_one_path_segment(self, action, value):
        """IDs may contain `/` and `+`. Encoded the way the typed SDK encodes them
        on every other call, so `AA/BB` cannot become two segments of the URL."""
        mock_client = _mock_client_with_batch_response(_make_batch_response([200]))
        await batch_triage(
            mock_client, message_ids=["AA/BB+CC="], action=action, value=value, config=_CFG
        )
        req = mock_client.request_adapter.send_primitive_async.call_args.args[0]
        url = json.loads(req.content.decode("utf-8"))["requests"][0]["url"]
        assert url.startswith("/me/messages/AA%2FBB%2BCC%3D")


class TestBatchTriageActions:
    async def test_flag_builds_patch_with_flag_status(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200]))
        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="flag",
            value="flagged",
            config=_CFG,
        )
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        sub = body["requests"][0]
        assert sub["method"] == "PATCH"
        assert sub["body"] == {"flag": {"flagStatus": "flagged"}}

    async def test_categorize_splits_comma_separated(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200]))
        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="categorize",
            value="Red, Blue, Green",
            config=_CFG,
        )
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        sub = body["requests"][0]
        assert sub["method"] == "PATCH"
        assert sub["body"] == {"categories": ["Red", "Blue", "Green"]}

    async def test_mark_read_true(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200]))
        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="mark_read",
            value="true",
            config=_CFG,
        )
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"isRead": True}

    async def test_mark_read_false(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200]))
        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="mark_read",
            value="false",
            config=_CFG,
        )
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"isRead": False}

    async def test_move_resolves_folder_display_name(self):
        """Display-name folder is resolved once before batch assembly."""
        folder_entity = MagicMock()
        folder_entity.id = "AAMkAG_tldr_folder="
        folder_entity.display_name = "TLDR"
        folders_response = MagicMock()
        folders_response.value = [folder_entity]
        folders_response.odata_next_link = None

        mock_client = _mock_client_with_batch_response(_make_batch_response([201]))
        mock_client.me.mail_folders.get = AsyncMock(return_value=folders_response)

        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="move",
            value="TLDR",
            config=_CFG,
        )
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"destinationId": "AAMkAG_tldr_folder="}

    async def test_move_resolves_a_folder_path(self):
        travel = MagicMock(id="AAMkTravelFolder=", display_name="Travel", child_folder_count=0)
        mock_client = _mock_client_with_batch_response(_make_batch_response([201]))
        inbox = MagicMock()
        inbox.child_folders.get = AsyncMock(
            return_value=MagicMock(value=[travel], odata_next_link=None)
        )
        mock_client.me.mail_folders.by_mail_folder_id = MagicMock(return_value=inbox)

        await batch_triage(
            mock_client,
            message_ids=["AAMkAG1="],
            action="move",
            value="Inbox/Travel",
            config=_CFG,
        )

        mock_client.me.mail_folders.by_mail_folder_id.assert_called_once_with("inbox")
        body = json.loads(
            mock_client.request_adapter.send_primitive_async.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"destinationId": "AAMkTravelFolder="}


class TestBatchTriageErrorHandling:
    async def test_per_item_failure_captured(self):
        http_resp = _make_batch_response([201, 404], errors=[None, "Not found"])
        mock_client = _mock_client_with_batch_response(http_resp)
        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="move",
            value="inbox",
            config=_CFG,
        )
        assert result["success_count"] == 1
        assert result["failure_count"] == 1
        assert result["results"][0]["status"] == "success"
        assert result["results"][1]["status"] == "error"
        assert "Not found" in result["results"][1]["error"]

    async def test_all_failures_counted(self):
        http_resp = _make_batch_response(
            [500, 500, 500], errors=["Graph error", "Graph error", "Graph error"]
        )
        mock_client = _mock_client_with_batch_response(http_resp)
        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2=", "AAMkAG3="],
            action="move",
            value="inbox",
            config=_CFG,
        )
        assert result["success_count"] == 0
        assert result["failure_count"] == 3

    async def test_results_preserve_message_order(self):
        """Results ordered by input message_ids even if /$batch returns them shuffled."""
        http_resp = _payload_response(
            [
                {"id": "2", "status": 201, "body": {}},
                {"id": "0", "status": 201, "body": {}},
                {"id": "1", "status": 404, "body": {"error": {"message": "Not found"}}},
            ]
        )
        mock_client = _mock_client_with_batch_response(http_resp)
        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2=", "AAMkAG3="],
            action="move",
            value="inbox",
            config=_CFG,
        )
        assert result["results"][0]["id"] == "AAMkAG1="
        assert result["results"][0]["status"] == "success"
        assert result["results"][1]["id"] == "AAMkAG2="
        assert result["results"][1]["status"] == "error"
        assert result["results"][2]["id"] == "AAMkAG3="
        assert result["results"][2]["status"] == "success"

    async def test_missing_subresponse_marked_error(self):
        """If Graph omits a sub-response, that message shows as error."""
        http_resp = _payload_response(
            [{"id": "0", "status": 201, "body": {}}]
            # index 1 missing
        )
        mock_client = _mock_client_with_batch_response(http_resp)
        result = await batch_triage(
            mock_client,
            message_ids=["AAMkAG1=", "AAMkAG2="],
            action="move",
            value="inbox",
            config=_CFG,
        )
        assert result["results"][0]["status"] == "success"
        assert result["results"][1]["status"] == "error"
        assert "no response" in result["results"][1]["error"]


# ── Throttled sub-requests, envelope errors, mark_read values ──────────────
#
# These run the batch through a real GraphServiceClient whose transport is a
# function, built the way `test_me_rewrite_reaches_the_wire.py` builds it: the
# envelope is a real `httpx.Response` and the real adapter and middleware
# handle it, so nothing here depends on which adapter method the tool calls.


def _graph_client_over(handler):
    """A real GraphServiceClient that sends every request to ``handler``."""
    import httpx
    from azure.core.credentials import AccessToken
    from kiota_authentication_azure.azure_identity_authentication_provider import (
        AzureIdentityAuthenticationProvider,
    )
    from msgraph import GraphServiceClient
    from msgraph import graph_request_adapter as gra
    from msgraph_core import GraphClientFactory

    class _Credential:
        def get_token(self, *args, **kwargs):
            return AccessToken("not-a-real-token", 9_999_999_999)

    http_client = GraphClientFactory.create_with_default_middleware(
        client=httpx.AsyncClient(transport=httpx.MockTransport(handler)),
        options=getattr(gra, "options", None),
    )
    auth = AzureIdentityAuthenticationProvider(
        _Credential(), scopes=["https://graph.microsoft.com/.default"]
    )
    return GraphServiceClient(request_adapter=gra.GraphRequestAdapter(auth, http_client))


class _Graph:
    """Answers each `$batch` with the next reply, and keeps what was sent."""

    def __init__(self, *replies):
        self.replies = list(replies)
        self.sent: list[list[dict]] = []

    def __call__(self, request):
        import httpx

        self.sent.append(json.loads(request.content)["requests"])
        reply = self.replies[min(len(self.sent), len(self.replies)) - 1]
        if isinstance(reply, httpx.Response):
            return reply
        return httpx.Response(200, json={"responses": reply})

    def client(self):
        return _graph_client_over(self)


def _throttled(sub_id: str, message: str = "Please retry again later.") -> dict:
    return {
        "id": sub_id,
        "status": 429,
        # "0" keeps these instant; retry_throttled_subrequests honours it.
        "headers": {"Retry-After": "0"},
        "body": {"error": {"code": "TooManyRequests", "message": message}},
    }


async def _flag(graph: _Graph, ids: list[str]) -> dict:
    return await batch_triage(
        graph.client(), message_ids=ids, action="flag", value="flagged", config=_CFG
    )


class TestBatchTriageThrottling:
    """Graph answers a `$batch` with 200 even when sub-requests were throttled.

    Each sub-request is checked against the mailbox's limits on its own and can
    come back 429 with its own Retry-After, and Graph documents that the SDK's
    retry handler does not retry those. `outlook_read_messages` already re-sends
    them through `throttle.retry_throttled_subrequests`; the write batch has to
    as well, or a throttled flag is reported as a failed one.
    """

    async def test_throttled_subrequest_is_retried_then_succeeds(self):
        graph = _Graph(
            [{"id": "0", "status": 200, "body": {}}, _throttled("1")],
            [{"id": "1", "status": 200, "body": {}}],
        )

        result = await _flag(graph, ["AAMkAG1=", "AAMkAG2="])

        assert [r["status"] for r in result["results"]] == ["success", "success"]
        assert (result["success_count"], result["failure_count"]) == (2, 0)
        # The retry carries only the throttled sub-request, unchanged.
        assert len(graph.sent) == 2
        assert graph.sent[1] == [graph.sent[0][1]]

    async def test_persistent_throttle_is_bounded_and_reported_with_graphs_text(self):
        graph = _Graph([_throttled("0")])

        result = await _flag(graph, ["AAMkAG1="])

        assert result["failure_count"] == 1
        assert result["results"][0]["error"] == "Please retry again later."
        # The first send plus the three retries retry_throttled_subrequests allows.
        assert len(graph.sent) == 4

    async def test_a_permanent_failure_is_not_retried(self):
        graph = _Graph([{"id": "0", "status": 404, "body": {"error": {"message": "Not found"}}}])

        result = await _flag(graph, ["AAMkAG1="])

        assert result["results"][0]["error"] == "Not found"
        assert len(graph.sent) == 1

    async def test_a_failed_retry_round_keeps_what_the_first_round_changed(self):
        """By the retry, the first round has already changed a message.

        Raising there would tell the agent nothing happened while one message was
        flagged; only the re-sent one is reported as failed, in Graph's words.
        """
        import httpx

        graph = _Graph(
            [{"id": "0", "status": 200, "body": {}}, _throttled("1")],
            httpx.Response(
                400, json={"error": {"code": "BadRequest", "message": "Invalid batch payload."}}
            ),
        )

        result = await _flag(graph, ["AAMkAG1=", "AAMkAG2="])

        assert result["results"][0] == {"id": "AAMkAG1=", "status": "success"}
        assert result["results"][1]["status"] == "error"
        assert "Invalid batch payload." in result["results"][1]["error"]
        assert result["success_count"] == 1

    async def test_an_unreadable_retry_reply_keeps_what_the_first_round_changed(self):
        """The same holds when the retry's reply cannot be read: its outcome is unknown."""
        import httpx

        graph = _Graph(
            [{"id": "0", "status": 200, "body": {}}, _throttled("1")],
            httpx.Response(200, text="<html>gateway</html>"),
        )

        result = await _flag(graph, ["AAMkAG1=", "AAMkAG2="])

        assert result["results"][0] == {"id": "AAMkAG1=", "status": "success"}
        assert result["results"][1]["status"] == "error"
        assert "not known whether this message was changed" in result["results"][1]["error"]
        assert result["success_count"] == 1


class TestBatchTriageEnvelopeErrors:
    async def test_an_envelope_error_reaches_the_model_as_graphs_message(self):
        """A refused `$batch` is one error, in Graph's words.

        It used to come back as a successful call with every message marked
        "no response from batch", and Graph's reason dropped.
        """
        import httpx
        from mcp.client import Client
        from mcp.server.mcpserver import MCPServer

        from outlook_mcp.server import _wrap_tool_errors

        graph = _Graph(
            httpx.Response(
                400,
                json={"error": {"code": "BadRequest", "message": "Invalid batch payload format."}},
            )
        )
        server = MCPServer("batch-envelope-guard", version="0")

        @server.tool()
        @_wrap_tool_errors
        async def triage() -> dict:
            return await _flag(graph, ["AAMkAG1="])

        async with Client(server) as client:
            result = await client.call_tool("triage", {})

        assert result.is_error is True
        text = result.content[0].text
        assert "Invalid batch payload format." in text
        assert "no response from batch" not in text


class TestBatchTriageMarkReadValues:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("true", True),
            ("True", True),
            (" true ", True),
            ("read", True),
            (True, True),
            ("false", False),
            ("FALSE", False),
            ("unread", False),
            (False, False),
        ],
    )
    async def test_accepted_values(self, value, expected):
        graph = _Graph([{"id": "0", "status": 200, "body": {}}])
        await batch_triage(
            graph.client(), message_ids=["AAMkAG1="], action="mark_read", value=value, config=_CFG
        )
        assert graph.sent[0][0]["body"] == {"isRead": expected}

    @pytest.mark.parametrize("value", ["yes", "1", "flagged", "", "trueish"])
    async def test_any_other_value_is_refused_before_anything_is_sent(self, value):
        """Anything but "true" used to mean unread, for every message in the batch."""
        graph = _Graph([{"id": "0", "status": 200, "body": {}}])
        with pytest.raises(ValueError, match="mark_read value must be"):
            await batch_triage(
                graph.client(),
                message_ids=["AAMkAG1="],
                action="mark_read",
                value=value,
                config=_CFG,
            )
        assert graph.sent == []
