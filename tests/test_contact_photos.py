"""``outlook_set_contact_photo`` (#121): what reaches Graph, and what is refused first.

Every refusal must happen before the PUT, so each one asserts the PUT was never
awaited and that its *own* text came back. A test that only checks "it raised"
would pass for a neighbouring guard, and removing the guard it is named for would
leave it green.

The size guard is proven with a file that is a valid PNG in every other respect,
so with the size check removed the call reaches the PUT and the test fails for
that reason, not because the bytes stopped looking like an image.
"""

from __future__ import annotations

import builtins
import inspect
import json
import os
import struct
import zlib
from unittest.mock import AsyncMock, MagicMock

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from mcp.client import Client
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError, ToolInputError
from outlook_mcp.tools import contact_photos
from outlook_mcp.tools.contact_photos import _MAX_PHOTO_BYTES, set_contact_photo

CONTACT_ID = "AAMkAGContact123="


# ── image fixtures ─────────────────────────────────────────────────────


def _chunk(kind: bytes, data: bytes) -> bytes:
    crc = zlib.crc32(kind + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", crc)


def png(rgb: tuple[int, int, int] = (128, 128, 128), side: int = 8, *, size: int | None = None):
    """A real PNG; with ``size``, padded by a ``tEXt`` chunk to exactly that many bytes."""
    raw = b"".join(b"\x00" + bytes(rgb) * side for _ in range(side))
    head = b"\x89PNG\r\n\x1a\n" + _chunk(
        b"IHDR", struct.pack(">IIBBBBB", side, side, 8, 2, 0, 0, 0)
    )
    body = _chunk(b"IDAT", zlib.compress(raw))
    tail = _chunk(b"IEND", b"")
    if size is not None:
        pad = size - len(head) - len(body) - len(tail) - 12  # 12: the tEXt chunk's own framing
        assert pad >= 0
        body += _chunk(b"tEXt", b"k\x00" + b"x" * (pad - 2))
    out = head + body + tail
    assert size is None or len(out) == size
    return out


GIF = bytes.fromhex(
    "47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b"
)

# SOI, an EXIF APP1 holding a thumbnail with its own EOI, a scan, EOI, then a
# Samsung-style trailer after it: a real phone photo ends like this (#183).
JPEG = (
    b"\xff\xd8"
    + b"\xff\xe1"
    + struct.pack(">H", 2 + 12)
    + b"Exif\x00\x00\xff\xd8\xff\xd9\x00\x00"
    + b"\xff\xda"
    + struct.pack(">H", 8)
    + b"\x01\x01\x00\x00\x3f\x00"
    + b"\x12\x34\x56"
    + b"\xff\xd9"
    + b"SEFH\x00\x00SEFT"
)

WEBP = b"RIFF" + struct.pack("<I", 12) + b"WEBPVP8L" + b"\x00" * 4


# ── harness ────────────────────────────────────────────────────────────


@pytest.fixture
def cfg(tmp_path):
    folder = tmp_path / "attachments"
    folder.mkdir()
    return Config(client_id="test", attachments_dir=str(folder))


def _write(cfg: Config, name: str, data: bytes) -> str:
    path = os.path.join(cfg.attachments_dir, name)
    with open(path, "wb") as f:
        f.write(data)
    return name


def _client():
    client = MagicMock()
    put = AsyncMock(return_value=None)
    client.me.contacts.by_contact_id.return_value.photo.content.put = put
    return client, put


class _ReadSpy:
    """Records every size passed to ``read`` on files the module opens."""

    def __init__(self):
        self.reads: list[int] = []
        self.opened = 0

    def __call__(self, *args, **kwargs):
        self.opened += 1
        handle = builtins.open(*args, **kwargs)
        spy = self
        real_read = handle.read

        def read(n=-1):
            spy.reads.append(n)
            return real_read(n)

        handle.read = read
        return handle


@pytest.fixture
def read_spy(monkeypatch):
    spy = _ReadSpy()
    monkeypatch.setattr(contact_photos, "open", spy, raising=False)
    return spy


# ── accepted ───────────────────────────────────────────────────────────


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("name", "data"),
    [("card.png", png()), ("card.jpg", JPEG), ("card.gif", GIF)],
    ids=["png", "jpeg-with-exif-and-trailer", "gif"],
)
async def test_each_allowed_type_is_put_byte_for_byte(cfg, name, data):
    client, put = _client()

    result = await set_contact_photo(client, CONTACT_ID, _write(cfg, name, data), config=cfg)

    client.me.contacts.by_contact_id.assert_called_once_with(CONTACT_ID)
    put.assert_awaited_once_with(data)
    assert result == {"status": "updated", "contact_id": CONTACT_ID, "uploaded_bytes": len(data)}


@pytest.mark.asyncio
async def test_the_type_comes_from_the_bytes_not_the_name(cfg):
    client, put = _client()
    data = png()

    await set_contact_photo(client, CONTACT_ID, _write(cfg, "photo.webp", data), config=cfg)

    put.assert_awaited_once_with(data)


@pytest.mark.asyncio
async def test_a_file_of_exactly_the_limit_is_sent(cfg, read_spy):
    client, put = _client()
    data = png(size=_MAX_PHOTO_BYTES)

    await set_contact_photo(client, CONTACT_ID, _write(cfg, "big.png", data), config=cfg)

    assert put.await_args.args[0] == data
    assert read_spy.reads == [_MAX_PHOTO_BYTES + 1]


# ── refused before the PUT, each with its own text ─────────────────────


@pytest.mark.asyncio
async def test_one_byte_over_the_limit_is_refused_with_a_bounded_read(cfg, read_spy):
    client, put = _client()
    name = _write(cfg, "huge.png", png(size=_MAX_PHOTO_BYTES + 1))

    with pytest.raises(ToolInputError, match="over this server's 20 MiB limit") as caught:
        await set_contact_photo(client, CONTACT_ID, name, config=cfg)

    put.assert_not_awaited()
    assert "Ask the user for a smaller copy" in str(caught.value)
    # The read is the gate: never more than one byte past the limit is requested.
    assert read_spy.reads == [_MAX_PHOTO_BYTES + 1]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("name", "data"),
    [
        ("anim.webp", WEBP),
        ("cut.png", png()[:-12]),
        ("notes.jpg", b"not an image at all\n"),
        ("cut.jpg", JPEG[:-14]),
        ("cut.gif", GIF[:-1]),
    ],
    ids=[
        "webp",
        "png-truncated-before-iend",
        "text-named-jpg",
        "jpeg-without-eoi",
        "gif-without-trailer",
    ],
)
async def test_anything_but_jpeg_png_or_gif_is_refused(cfg, name, data):
    client, put = _client()

    with pytest.raises(ToolInputError, match="is not a JPEG, PNG or GIF file"):
        await set_contact_photo(client, CONTACT_ID, _write(cfg, name, data), config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_an_empty_file_is_refused(cfg):
    client, put = _client()

    with pytest.raises(ToolInputError, match="is empty, so no photo was set"):
        await set_contact_photo(client, CONTACT_ID, _write(cfg, "empty.png", b""), config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_missing_file_is_refused(cfg):
    client, put = _client()

    with pytest.raises(ToolInputError, match="Photo file not found"):
        await set_contact_photo(client, CONTACT_ID, "nowhere.png", config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_directory_is_refused(cfg):
    client, put = _client()
    os.mkdir(os.path.join(cfg.attachments_dir, "folder.png"))

    with pytest.raises(ToolInputError, match="Photo file not found"):
        await set_contact_photo(client, CONTACT_ID, "folder.png", config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
@pytest.mark.parametrize("where", ["absolute", "dotdot"])
async def test_a_path_outside_attachments_dir_is_refused(cfg, tmp_path, where):
    client, put = _client()
    outside = tmp_path / "secret.png"
    outside.write_bytes(png())
    path = str(outside) if where == "absolute" else os.path.join("..", "secret.png")

    with pytest.raises(ValueError, match="attachments") as caught:
        await set_contact_photo(client, CONTACT_ID, path, config=cfg)

    assert not isinstance(caught.value, OSError)
    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_symlink_out_of_attachments_dir_is_refused(cfg, tmp_path):
    client, put = _client()
    outside = tmp_path / "secret.png"
    outside.write_bytes(png())
    link = os.path.join(cfg.attachments_dir, "link.png")
    try:
        os.symlink(outside, link)
    except (OSError, NotImplementedError):
        pytest.skip("this platform or account cannot create symlinks")

    with pytest.raises(ValueError, match="attachments"):
        await set_contact_photo(client, CONTACT_ID, "link.png", config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_an_unreadable_file_is_reported_not_crashed(cfg, monkeypatch):
    client, put = _client()
    name = _write(cfg, "locked.png", png())

    def locked(*args, **kwargs):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(contact_photos, "open", locked, raising=False)

    with pytest.raises(ToolInputError, match="Could not open locked.png in attachments_dir"):
        await set_contact_photo(client, CONTACT_ID, name, config=cfg)

    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_failure_preparing_attachments_dir_is_reported_not_crashed(cfg, monkeypatch):
    client, put = _client()

    def broken(*args, **kwargs):
        raise OSError(28, "No space left on device")

    monkeypatch.setattr(contact_photos, "resolve_attachment_path", broken)

    with pytest.raises(ToolInputError, match="Could not open card.png in attachments_dir"):
        await set_contact_photo(client, CONTACT_ID, "card.png", config=cfg)

    put.assert_not_awaited()


# ── ordering: policy and the id before the disk ────────────────────────


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("overrides", "error"),
    [
        ({"read_only": True}, ReadOnlyError),
        ({"allow_categories": ["mail_drafts"]}, PermissionDeniedError),
    ],
    ids=["read_only", "contacts_write-not-allowed"],
)
async def test_policy_refuses_before_the_file_is_read(cfg, read_spy, overrides, error):
    client, put = _client()
    name = _write(cfg, "card.png", png())
    gated = cfg.model_copy(update=overrides)

    with pytest.raises(error):
        await set_contact_photo(client, CONTACT_ID, name, config=gated)

    assert read_spy.opened == 0
    put.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_malformed_contact_id_is_refused_before_the_file_is_read(cfg, read_spy):
    client, put = _client()
    name = _write(cfg, "card.png", png())

    with pytest.raises(ValueError, match="Graph ID contains invalid characters"):
        await set_contact_photo(client, "bad id/../x", name, config=cfg)

    assert read_spy.opened == 0
    put.assert_not_awaited()


# ── the wire, through the real SDK pipeline ────────────────────────────


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def _real_graph() -> tuple[GraphServiceClient, list[httpx.Request]]:
    """A real SDK client over a mock transport, built as in test_body_type_reaches_the_wire."""
    sent: list[httpx.Request] = []

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(request)
        return httpx.Response(200)

    factory = GraphClientFactory.create_with_default_middleware
    kwargs = {}
    params = inspect.signature(factory).parameters
    if "options" in params:
        kwargs["options"] = getattr(_gra, "options", None)
    if "client" in params:
        kwargs["client"] = httpx.AsyncClient(transport=httpx.MockTransport(transport))
    auth = AzureIdentityAuthenticationProvider(
        _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
    )
    client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, factory(**kwargs)))
    return client, sent


@pytest.mark.asyncio
async def test_the_photo_goes_out_as_one_put_of_the_file_bytes(cfg):
    client, sent = _real_graph()
    data = png((200, 10, 10))

    await set_contact_photo(client, CONTACT_ID, _write(cfg, "card.png", data), config=cfg)

    assert len(sent) == 1
    request = sent[0]
    assert request.method == "PUT"
    assert request.url.path == f"/v1.0/me/contacts/{CONTACT_ID}/photo/$value"
    assert request.content == data
    # The type Graph was probed accepting (2026-11-14). If the SDK ever changes
    # the label it sends, this says so before a user finds out.
    assert request.headers["Content-Type"] == "application/octet-stream"


# ── through an MCP client, with the opt-in tool loaded ─────────────────


@pytest.fixture
def mcp_with_photo_tool(monkeypatch, cfg):
    from outlook_mcp import server

    client, put = _client()
    monkeypatch.setattr(server, "_get_graph_client", lambda ctx: MagicMock(sdk_client=client))
    monkeypatch.setattr(server, "_get_config", lambda ctx: cfg)
    # The suite loads the default groups; expose this opt-in tool for this test only.
    monkeypatch.setattr(server.mcp._tool_manager, "_tools", dict(server.mcp._tool_manager._tools))
    if "outlook_set_contact_photo" not in server.mcp._tool_manager._tools:
        server.mcp.add_tool(server.outlook_set_contact_photo)
    return server.mcp, client, put


async def _call(mcp, arguments):
    async with Client(mcp) as client:
        return await client.call_tool("outlook_set_contact_photo", arguments)


@pytest.mark.asyncio
async def test_a_call_through_the_server_puts_the_file_for_that_contact(mcp_with_photo_tool, cfg):
    mcp, client, put = mcp_with_photo_tool
    data = png((10, 10, 200))
    name = _write(cfg, "card.png", data)

    result = await _call(mcp, {"contact_id": CONTACT_ID, "file_path": name})

    assert result.is_error is False
    client.me.contacts.by_contact_id.assert_called_once_with(CONTACT_ID)
    put.assert_awaited_once_with(data)
    assert json.loads(result.content[0].text) == {
        "status": "updated",
        "contact_id": CONTACT_ID,
        "uploaded_bytes": len(data),
    }


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("setup", "expected"),
    [
        ("not-an-image", "is not a JPEG, PNG or GIF file"),
        ("locked", "Could not open card.png in attachments_dir"),
        ("resolver-oserror", "Could not open card.png in attachments_dir"),
        ("outside", "attachments"),
    ],
)
async def test_each_refusal_reaches_the_client_as_text(
    mcp_with_photo_tool, cfg, monkeypatch, tmp_path, setup, expected
):
    mcp, _, put = mcp_with_photo_tool
    path = "card.png"
    if setup == "not-an-image":
        _write(cfg, path, b"hello")
    elif setup == "locked":
        _write(cfg, path, png())

        def locked(*args, **kwargs):
            raise PermissionError(13, "Permission denied")

        monkeypatch.setattr(contact_photos, "open", locked, raising=False)
    elif setup == "resolver-oserror":

        def broken(*args, **kwargs):
            raise OSError(28, "No space left on device")

        monkeypatch.setattr(contact_photos, "resolve_attachment_path", broken)
    else:
        (tmp_path / "secret.png").write_bytes(png())
        path = str(tmp_path / "secret.png")

    result = await _call(mcp, {"contact_id": CONTACT_ID, "file_path": path})

    assert result.is_error is True
    text = result.content[0].text
    # The SDK prefixes forwarded text with "Error executing tool <name>: "; a
    # withheld message is that prefix alone, so the refusal's own words are the check.
    assert expected in text
    put.assert_not_awaited()
