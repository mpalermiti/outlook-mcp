"""Large-file upload sessions: every chunk's reply is checked (#125).

A refused chunk used to be ignored, so the draft was sent, or the attachment reported as
added, without the file. These drive the real ``_upload_large_file`` against an
``httpx.MockTransport``, so what is checked is the reply Graph would give, not a patched-out
helper.
"""

from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import GraphAPIError, wrap_http_status_error
from outlook_mcp.tools import mail_attachments
from outlook_mcp.tools.mail_attachments import (
    _upload_large_file,
    attach_to_draft,
    send_with_attachments,
)

_UPLOAD_URL = "https://graph.microsoft.com/upload/session"


def _graph_replies(monkeypatch, replies, chunk):
    """Answer each chunk PUT with the next entry in ``replies`` (then 202); return the
    requests seen. An entry may be an exception, raised as the transport would. An empty
    chunk is refused, as Graph refuses a backwards range, so a loop that sends one ends."""
    seen: list[httpx.Request] = []
    queue = iter(replies)

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        if not request.content:
            error = {"error": {"code": "InvalidRange", "message": "Empty chunk."}}
            return httpx.Response(416, json=error)
        reply = next(queue, 202)
        if isinstance(reply, Exception):
            raise reply
        if reply >= 400:
            error = {"error": {"code": "InvalidRequest", "message": "The chunk was refused."}}
            return httpx.Response(reply, json=error)
        return httpx.Response(reply, json={})

    real_client = httpx.AsyncClient
    monkeypatch.setattr(
        httpx, "AsyncClient", lambda *a, **k: real_client(transport=httpx.MockTransport(handler))
    )
    monkeypatch.setattr(mail_attachments, "_UPLOAD_CHUNK_SIZE", chunk)
    return seen


def _file(tmp_path, size, name="big.bin"):
    path = tmp_path / name
    path.write_bytes(b"x" * size)
    return path


class TestUploadLargeFile:
    async def test_a_completed_upload_returns(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 2 * 1024 + 100)
        seen = _graph_replies(monkeypatch, [202, 202, 201], chunk=1024)

        await _upload_large_file(_UPLOAD_URL, str(big), 2148)

        assert [r.headers["Content-Range"] for r in seen] == [
            "bytes 0-1023/2148",
            "bytes 1024-2047/2148",
            "bytes 2048-2147/2148",
        ]

    async def test_a_refused_chunk_raises_and_stops(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 2 * 1024 + 100)
        seen = _graph_replies(monkeypatch, [202, 500, 202], chunk=1024)

        with pytest.raises(httpx.HTTPStatusError):
            await _upload_large_file(_UPLOAD_URL, str(big), 2148)

        assert len(seen) == 2  # nothing after the refused chunk

    async def test_an_upload_graph_never_reports_complete_raises(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 2 * 1024 + 100)
        _graph_replies(monkeypatch, [202, 202, 202], chunk=1024)

        with pytest.raises(GraphAPIError) as info:
            await _upload_large_file(_UPLOAD_URL, str(big), 2148)

        assert info.value.error_code == "IncompleteUpload"

    async def test_a_file_that_comes_up_short_stops_before_an_empty_put(
        self, tmp_path, monkeypatch
    ):
        """The size Graph was given is larger than what the file now holds."""
        short = _file(tmp_path, 1500)
        seen = _graph_replies(monkeypatch, [202, 202], chunk=1024)

        with pytest.raises(GraphAPIError) as info:
            await _upload_large_file(_UPLOAD_URL, str(short), 2148)

        assert info.value.error_code == "IncompleteUpload"
        assert "changed while it was being uploaded" in str(info.value)
        assert [len(r.content) for r in seen] == [1024, 476]  # no empty PUT

    async def test_the_model_sees_graphs_error_not_the_upload_url(self, tmp_path, monkeypatch):
        """``_wrap_tool_errors`` turns the refusal into Graph's own error; the upload URL,
        which carries its own authorization, is not in it."""
        big = _file(tmp_path, 2 * 1024 + 100)
        _graph_replies(monkeypatch, [403], chunk=1024)

        with pytest.raises(httpx.HTTPStatusError) as info:
            await _upload_large_file(_UPLOAD_URL, str(big), 2148)
        error = wrap_http_status_error(info.value)

        assert error.error_code == "InvalidRequest"
        assert "upload/session" not in str(error)


def _send_client():
    """A Graph client whose draft is ``draft123`` and whose send is a recorded no-op."""
    mock_client = MagicMock()
    mock_client.me.messages.post = AsyncMock(return_value=MagicMock(id="draft123"))
    builder = mock_client.me.messages.by_message_id.return_value
    builder.attachments.create_upload_session.post = AsyncMock(
        return_value=MagicMock(upload_url=_UPLOAD_URL)
    )
    builder.send.post = AsyncMock()
    return mock_client, builder


async def _send(mock_client, big, tmp_path):
    await send_with_attachments(
        mock_client,
        to=["a@b.com"],
        subject="Large file",
        body="See attached",
        attachment_paths=[str(big)],
        config=Config(client_id="test", attachments_dir=str(tmp_path)),
    )


class TestNothingGoesOutAfterAFailedUpload:
    async def test_send_names_the_draft_after_a_refused_chunk(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        _graph_replies(monkeypatch, [202, 500], chunk=1024 * 1024)
        mock_client, builder = _send_client()

        with pytest.raises(GraphAPIError) as info:
            await _send(mock_client, big, tmp_path)

        builder.send.post.assert_not_called()
        assert info.value.error_code == "InvalidRequest"  # still Graph's own error
        assert "saved in Drafts (id draft123)" in str(info.value)
        assert "upload/session" not in str(info.value)

    async def test_send_names_the_draft_when_the_upload_is_not_complete(
        self, tmp_path, monkeypatch
    ):
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        _graph_replies(monkeypatch, [], chunk=1024 * 1024)  # every chunk 202, never done
        mock_client, builder = _send_client()

        with pytest.raises(GraphAPIError) as info:
            await _send(mock_client, big, tmp_path)

        builder.send.post.assert_not_called()
        assert info.value.error_code == "IncompleteUpload"
        assert "saved in Drafts (id draft123)" in str(info.value)
        assert "Try again" not in str(info.value)  # a plain retry would make a second draft

    async def test_send_names_the_draft_when_the_connection_drops(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        _graph_replies(monkeypatch, [202, httpx.WriteTimeout("timed out")], chunk=1024 * 1024)
        mock_client, builder = _send_client()

        with pytest.raises(GraphAPIError) as info:
            await _send(mock_client, big, tmp_path)

        builder.send.post.assert_not_called()
        assert info.value.error_code == "UploadInterrupted"
        assert "saved in Drafts (id draft123)" in str(info.value)

    async def test_an_unexpected_failure_is_not_dressed_up(self, tmp_path):
        """A crash stays a crash (its text stays on the server), and still nothing is sent."""
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        mock_client, builder = _send_client()

        with (
            patch.object(
                mail_attachments, "_upload_large_file", AsyncMock(side_effect=RuntimeError("x"))
            ),
            pytest.raises(RuntimeError),
        ):
            await _send(mock_client, big, tmp_path)

        builder.send.post.assert_not_called()

    async def test_attach_to_draft_does_not_report_it_attached(self, tmp_path, monkeypatch):
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        _graph_replies(monkeypatch, [202, 500], chunk=1024 * 1024)

        mock_client = MagicMock()
        builder = mock_client.me.messages.by_message_id.return_value
        builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        builder.attachments.create_upload_session.post = AsyncMock(
            return_value=MagicMock(upload_url=_UPLOAD_URL)
        )

        with pytest.raises(httpx.HTTPStatusError):
            await attach_to_draft(
                mock_client,
                draft_id="AAMkAG123=",
                attachment_paths=[str(big)],
                config=Config(client_id="test", attachments_dir=str(tmp_path)),
            )

    async def test_attach_to_event_does_not_report_it_attached(self, tmp_path, monkeypatch):
        """The third caller (#160). Skipped on a tree that predates it."""
        calendar_attachments = pytest.importorskip("outlook_mcp.tools.calendar_attachments")
        attach_to_event = calendar_attachments.attach_to_event
        big = _file(tmp_path, 3 * 1024 * 1024 + 1)
        _graph_replies(monkeypatch, [202, 500], chunk=1024 * 1024)

        mock_client = MagicMock()
        builder = mock_client.me.events.by_event_id.return_value
        builder.attachments.create_upload_session.post = AsyncMock(
            return_value=MagicMock(upload_url=_UPLOAD_URL)
        )

        with pytest.raises(httpx.HTTPStatusError):
            await attach_to_event(
                mock_client,
                event_id="AAMkAGevent=",
                attachment_paths=[str(big)],
                config=Config(client_id="test", attachments_dir=str(tmp_path)),
            )
