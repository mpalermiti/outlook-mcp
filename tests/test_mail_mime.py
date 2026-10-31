"""MIME export uses bounded raw streams, not a MIME parser or text decoder."""

import base64
import hashlib
import json
from types import SimpleNamespace
from unittest.mock import MagicMock

import httpx
import pytest
from mcp.client import Client
from mcp.types import BlobResourceContents, EmbeddedResource, TextContent

from outlook_mcp.tools import mail_mime


class TrackedStream(httpx.AsyncByteStream):
    def __init__(self, chunks):
        self.chunks = chunks
        self.read = 0
        self.closed = False

    async def __aiter__(self):
        for chunk in self.chunks:
            self.read += 1
            yield chunk

    async def aclose(self):
        self.closed = True


@pytest.fixture
def graph():
    credential = MagicMock()
    credential.get_token.return_value = SimpleNamespace(token="synthetic-graph-token")
    return SimpleNamespace(credential=credential)


def mock_graph(monkeypatch, stream, *, status=200, headers=None, check=None):
    original_client = httpx.AsyncClient

    def respond(request):
        if check:
            check(request)
        return httpx.Response(status, headers=headers, stream=stream)

    monkeypatch.setattr(
        mail_mime.httpx,
        "AsyncClient",
        lambda **kwargs: original_client(transport=httpx.MockTransport(respond), **kwargs),
    )


@pytest.mark.asyncio
async def test_original_multipart_bytes_and_id_are_preserved(monkeypatch, graph):
    data = (
        b"From: sender@example.com\r\nContent-Type: multipart/mixed; boundary=x\r\n\r\n"
        b"--x\r\nContent-Type: application/octet-stream\r\n\r\n\x00\xff\x80\r\n--x--\r\n"
    )
    stream = TrackedStream([data[:37], data[37:]])

    def check(request):
        assert request.method == "GET"
        assert request.url.host == "graph.microsoft.com"
        assert request.url.raw_path == b"/v1.0/me/messages/A%2FB%2BC%3D/$value"
        assert request.headers["authorization"] == "Bearer synthetic-graph-token"
        assert request.headers["accept-encoding"] == "identity"

    mock_graph(monkeypatch, stream, check=check)
    result = await mail_mime.get_mime(graph, "A/B+C=")
    assert len(result) == 2
    assert isinstance(result[0], TextContent)
    assert json.loads(result[0].text) == {
        "message_id": "A/B+C=",
        "size_bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "content_type": "message/rfc822",
    }
    assert isinstance(result[1], EmbeddedResource)
    resource = result[1].resource
    assert isinstance(resource, BlobResourceContents)
    assert str(resource.uri) == "message:///A%2FB%2BC%3D.eml"
    assert resource.mime_type == "message/rfc822"
    assert base64.b64decode(resource.blob) == data
    graph.credential.get_token.assert_called_once_with("https://graph.microsoft.com/.default")
    assert stream.closed


@pytest.mark.asyncio
@pytest.mark.parametrize("headers", [{}, {"content-length": "1"}, {"content-length": "invalid"}])
async def test_stream_limit_does_not_trust_metadata(monkeypatch, graph, headers):
    monkeypatch.setattr(mail_mime, "MAX_MIME_BYTES", 8)
    monkeypatch.setattr(mail_mime, "_CHUNK_BYTES", 4)
    stream = TrackedStream([b"1234", b"5678", b"9abc", b"unread"])
    mock_graph(monkeypatch, stream, headers=headers)
    with pytest.raises(ValueError, match="export limit"):
        await mail_mime.get_mime(graph, "message")
    assert stream.read == 3
    assert stream.closed


@pytest.mark.asyncio
async def test_declared_oversize_is_refused_before_body(monkeypatch, graph):
    stream = TrackedStream([b"not read"])
    mock_graph(monkeypatch, stream, headers={"content-length": str(mail_mime.MAX_MIME_BYTES + 1)})
    with pytest.raises(ValueError, match="export limit"):
        await mail_mime.get_mime(graph, "message")
    assert stream.read == 0
    assert stream.closed


@pytest.mark.asyncio
async def test_exact_limit_is_accepted(monkeypatch, graph):
    monkeypatch.setattr(mail_mime, "MAX_MIME_BYTES", 8)
    stream = TrackedStream([b"1234", b"5678"])
    mock_graph(monkeypatch, stream)
    assert json.loads((await mail_mime.get_mime(graph, "message"))[0].text)["size_bytes"] == 8
    assert stream.closed


@pytest.mark.asyncio
@pytest.mark.parametrize("extra", [0, 1])
async def test_real_ten_mib_boundary(monkeypatch, graph, extra):
    data = b"x" * (10 * 1024 * 1024 + extra)
    stream = TrackedStream([data])
    mock_graph(monkeypatch, stream)
    if extra:
        with pytest.raises(ValueError, match="10 MiB export limit"):
            await mail_mime.get_mime(graph, "message")
    else:
        result = await mail_mime.get_mime(graph, "message")
        assert json.loads(result[0].text)["size_bytes"] == len(data)
        assert base64.b64decode(result[1].resource.blob) == data
    assert stream.closed


@pytest.mark.asyncio
@pytest.mark.parametrize("data", [b"Subject: text\r\n\r\nhello\r\n", b"\x00\xff\x80\r\n"])
async def test_export_reaches_mcp_client_as_exact_blob_without_duplicate_json(
    monkeypatch, graph, data
):
    from outlook_mcp import server

    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: graph)
    # The suite normally loads default groups; expose this opt-in for this call only.
    monkeypatch.setattr(server.mcp._tool_manager, "_tools", dict(server.mcp._tool_manager._tools))
    if "outlook_export_message" not in server.mcp._tool_manager._tools:
        server.mcp.add_tool(server.outlook_export_message)
    stream = TrackedStream([data])
    mock_graph(monkeypatch, stream)
    async with Client(server.mcp) as client:
        result = await client.call_tool("outlook_export_message", {"message_id": "message"})
    assert result.is_error is False
    assert result.structured_content is None
    assert [block.type for block in result.content] == ["text", "resource"]
    meta = json.loads(result.content[0].text)
    assert meta["size_bytes"] == len(data)
    assert meta["sha256"] == hashlib.sha256(data).hexdigest()
    assert "content_base64" not in meta
    resource = result.content[1].resource
    assert isinstance(resource, BlobResourceContents)
    assert resource.mime_type == "message/rfc822"
    assert base64.b64decode(resource.blob) == data
    assert stream.closed


@pytest.mark.asyncio
async def test_empty_export_is_not_a_success(monkeypatch, graph):
    stream = TrackedStream([])
    mock_graph(monkeypatch, stream)
    with pytest.raises(ValueError, match="empty MIME"):
        await mail_mime.get_mime(graph, "message")
    assert stream.closed


@pytest.mark.asyncio
async def test_compressed_body_is_refused_without_decompressing(monkeypatch, graph):
    stream = TrackedStream([b"not read"])
    mock_graph(monkeypatch, stream, headers={"content-encoding": "gzip"})
    with pytest.raises(ValueError, match="unsupported encoding"):
        await mail_mime.get_mime(graph, "message")
    assert stream.read == 0
    assert stream.closed


@pytest.mark.asyncio
@pytest.mark.parametrize("message_id", ["", "..", "../other", "id?x=y", "id\r\n", "x" * 1025])
async def test_invalid_id_is_refused_before_token_or_network(graph, message_id):
    with pytest.raises(ValueError):
        await mail_mime.get_mime(graph, message_id)
    graph.credential.get_token.assert_not_called()


@pytest.mark.asyncio
async def test_graph_error_reaches_existing_structured_translation(monkeypatch, graph):
    from outlook_mcp.errors import GraphAPIError
    from outlook_mcp.server import _wrap_tool_errors

    stream = TrackedStream([b'{"error":{"code":"ErrorItemNotFound","message":"Not found"}}'])
    mock_graph(monkeypatch, stream, status=404)
    with pytest.raises(GraphAPIError) as error:
        await _wrap_tool_errors(mail_mime.get_mime)(graph, "message")
    assert error.value.status_code == 404
    assert error.value.error_code == "ErrorItemNotFound"
    assert stream.closed


@pytest.mark.asyncio
async def test_redirect_is_not_followed(monkeypatch, graph):
    stream = TrackedStream([])
    requests = []
    mock_graph(
        monkeypatch,
        stream,
        status=302,
        headers={"location": "https://example.com/elsewhere"},
        check=lambda request: requests.append(request),
    )
    with pytest.raises(httpx.HTTPStatusError):
        await mail_mime.get_mime(graph, "message")
    assert len(requests) == 1
    assert stream.closed


@pytest.mark.asyncio
async def test_error_body_is_bounded_and_closed(monkeypatch, graph):
    monkeypatch.setattr(mail_mime, "_ERROR_BYTES", 8)
    monkeypatch.setattr(mail_mime, "_CHUNK_BYTES", 4)
    stream = TrackedStream([b"1234", b"5678", b"unread"])
    mock_graph(monkeypatch, stream, status=503)
    with pytest.raises(httpx.HTTPStatusError) as error:
        await mail_mime.get_mime(graph, "message")
    assert error.value.response.content == b"12345678"
    assert stream.read == 2
    assert stream.closed


@pytest.mark.asyncio
async def test_stream_failure_closes_response(monkeypatch, graph):
    class BrokenStream(TrackedStream):
        async def __aiter__(self):
            yield b"partial"
            raise httpx.ReadError("interrupted")

    stream = BrokenStream([])
    mock_graph(monkeypatch, stream)
    with pytest.raises(httpx.ReadError):
        await mail_mime.get_mime(graph, "message")
    assert stream.closed


@pytest.mark.asyncio
async def test_compressed_error_is_not_decompressed(monkeypatch, graph):
    from outlook_mcp.errors import GraphAPIError
    from outlook_mcp.server import _wrap_tool_errors

    stream = TrackedStream([b"not a gzip stream"])
    mock_graph(monkeypatch, stream, status=503, headers={"content-encoding": "gzip"})
    with pytest.raises(GraphAPIError) as error:
        await _wrap_tool_errors(mail_mime.get_mime)(graph, "message")
    assert error.value.status_code == 503
    assert stream.closed
