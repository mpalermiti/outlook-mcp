"""Tests for batch triage tool (Graph /$batch endpoint)."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.batch import batch_triage

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _make_batch_response(statuses: list[int], errors: list[str | None] | None = None):
    """Build a mock Graph /$batch response with the given per-sub-request statuses."""
    errors = errors or [None] * len(statuses)
    responses = []
    for i, status in enumerate(statuses):
        entry = {"id": str(i), "status": status}
        if errors[i] is not None:
            entry["body"] = {"error": {"code": "Err", "message": errors[i]}}
        else:
            entry["body"] = {}
        responses.append(entry)
    http_resp = MagicMock()
    http_resp.json.return_value = {"responses": responses}
    return http_resp


def _mock_client_with_batch_response(http_resp: MagicMock) -> MagicMock:
    mock_client = MagicMock()
    mock_client.request_adapter.get_http_response_message = AsyncMock(
        return_value=http_resp
    )
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
                mock_client, message_ids=[], action="move", value="inbox", config=_CFG,
            )

    async def test_max_20_messages(self):
        mock_client = MagicMock()
        ids = [f"AAMkAG{i}=" for i in range(21)]
        with pytest.raises(ValueError, match="Maximum 20"):
            await batch_triage(
                mock_client, message_ids=ids, action="move", value="inbox", config=_CFG,
            )

    async def test_exactly_20_messages_allowed(self):
        mock_client = _mock_client_with_batch_response(_make_batch_response([200] * 20))
        ids = [f"AAMkAG{i}=" for i in range(20)]
        result = await batch_triage(
            mock_client, message_ids=ids, action="move", value="inbox", config=_CFG,
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
        mock_client.request_adapter.get_http_response_message.assert_called_once()

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
        call = mock_client.request_adapter.get_http_response_message.call_args
        req = call.args[0]
        assert req.url == "https://graph.microsoft.com/v1.0/$batch"
        body = json.loads(req.content.decode("utf-8"))
        assert len(body["requests"]) == 2
        assert body["requests"][0]["method"] == "POST"
        assert body["requests"][0]["url"] == "/me/messages/AAMkAG1=/move"
        assert body["requests"][0]["body"] == {"destinationId": "inbox"}
        assert body["requests"][1]["url"] == "/me/messages/AAMkAG2=/move"


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
            mock_client.request_adapter.get_http_response_message.call_args.args[0].content
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
            mock_client.request_adapter.get_http_response_message.call_args.args[0].content
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
            mock_client.request_adapter.get_http_response_message.call_args.args[0].content
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
            mock_client.request_adapter.get_http_response_message.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"isRead": False}

    async def test_move_resolves_folder_display_name(self):
        """Display-name folder is resolved once before batch assembly."""
        folder_entity = MagicMock()
        folder_entity.id = "AAMkAG_tldr_folder="
        folder_entity.display_name = "TLDR"
        folders_response = MagicMock()
        folders_response.value = [folder_entity]

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
            mock_client.request_adapter.get_http_response_message.call_args.args[0].content
        )
        assert body["requests"][0]["body"] == {"destinationId": "AAMkAG_tldr_folder="}


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
        http_resp = MagicMock()
        http_resp.json.return_value = {
            "responses": [
                {"id": "2", "status": 201, "body": {}},
                {"id": "0", "status": 201, "body": {}},
                {"id": "1", "status": 404, "body": {"error": {"message": "Not found"}}},
            ]
        }
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
        http_resp = MagicMock()
        http_resp.json.return_value = {
            "responses": [{"id": "0", "status": 201, "body": {}}]
            # index 1 missing
        }
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
