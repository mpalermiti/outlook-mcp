"""Read-only MIME qualification on a personal Microsoft mailbox."""

import base64
import hashlib
import json

import pytest

from outlook_mcp.tools.mail_mime import get_mime
from outlook_mcp.tools.mail_read import list_inbox

pytestmark = [pytest.mark.live, pytest.mark.asyncio]


async def test_mime_matches_native_graph_bytes(real_graph_client, consumer_mailbox_only):
    inbox = await list_inbox(real_graph_client.sdk_client, count=1)
    if not inbox["messages"]:
        pytest.skip("Inbox is empty; no message available for MIME qualification")
    message_id = inbox["messages"][0]["id"]
    exported = await get_mime(real_graph_client, message_id)
    metadata = json.loads(exported[0].text)
    native = await real_graph_client.sdk_client.me.messages.by_message_id(message_id).content.get()
    assert native, "Graph SDK returned no MIME bytes"
    # Compare digests and sizes so test failures do not print message content.
    assert metadata["size_bytes"] == len(native)
    assert metadata["sha256"] == hashlib.sha256(native).hexdigest()
    assert (
        hashlib.sha256(base64.b64decode(exported[1].resource.blob)).hexdigest()
        == metadata["sha256"]
    )
