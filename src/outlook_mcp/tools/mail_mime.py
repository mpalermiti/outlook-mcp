"""Bounded, byte-preserving MIME export on the existing Graph sign-in."""

from __future__ import annotations

import base64
import hashlib
import json
from typing import Any
from urllib.parse import quote

import httpx
from mcp.types import BlobResourceContents, ContentBlock, EmbeddedResource, TextContent

from outlook_mcp.tools._delta import GRAPH_BASE, _bearer_token
from outlook_mcp.validation import validate_graph_id

MAX_MIME_BYTES = 10 * 1024 * 1024
_CHUNK_BYTES = 64 * 1024
_ERROR_BYTES = 64 * 1024


async def get_mime(client: Any, message_id: str) -> list[ContentBlock]:
    """Return the original message bytes without parsing or reserializing MIME."""
    validate_graph_id(message_id)
    url = GRAPH_BASE + "me/messages/" + quote(message_id, safe="") + "/$value"
    headers = {
        "Authorization": "Bearer " + _bearer_token(client.credential),
        "Accept": "message/rfc822",
        "Accept-Encoding": "identity",
    }
    async with httpx.AsyncClient(timeout=30.0, follow_redirects=False) as http:
        async with http.stream("GET", url, headers=headers) as response:
            if not response.is_success:
                # Graph's error JSON gets the same structured error translation
                # as other tools, without reading an unbounded error body.
                error_body = bytearray()
                async for chunk in response.aiter_raw(chunk_size=_CHUNK_BYTES):
                    error_body.extend(chunk[: _ERROR_BYTES - len(error_body)])
                    if len(error_body) >= _ERROR_BYTES:
                        break
                httpx.Response(
                    response.status_code,
                    content=bytes(error_body),
                    request=response.request,
                ).raise_for_status()

            # Request identity encoding and refuse anything else: raw chunks
            # both preserve the MIME bytes and bound memory before decompression.
            if response.headers.get("content-encoding", "identity").lower() != "identity":
                raise ValueError("Graph returned an unsupported encoding for MIME export")
            try:
                declared_size = int(response.headers.get("content-length", "0"))
            except ValueError:
                declared_size = 0
            if declared_size > MAX_MIME_BYTES:
                raise ValueError("Message MIME exceeds the 10 MiB export limit")
            content = bytearray()
            async for chunk in response.aiter_raw(chunk_size=_CHUNK_BYTES):
                if len(content) + len(chunk) > MAX_MIME_BYTES:
                    raise ValueError("Message MIME exceeds the 10 MiB export limit")
                content.extend(chunk)
    if not content:
        raise ValueError("Graph returned empty MIME content")
    metadata = {
        "message_id": message_id,
        "size_bytes": len(content),
        "sha256": hashlib.sha256(content).hexdigest(),
        "content_type": "message/rfc822",
    }
    resource = BlobResourceContents(
        uri=f"message:///{quote(message_id, safe='')}.eml",
        mime_type="message/rfc822",
        blob=base64.b64encode(content).decode("ascii"),
    )
    return [
        TextContent(type="text", text=json.dumps(metadata)),
        EmbeddedResource(type="resource", resource=resource),
    ]
