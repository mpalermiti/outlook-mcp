"""Write-tier live guards for contacts: does every field survive a round trip?

Tier 3 of the silent-no-op audit. The mock suite and the wire-level payload
tests prove what we *send*. Only Graph can say what it *kept* — `is_online`
was accepted and dropped on personal accounts with a 201 and no error. Each
test here writes every parameter a tool accepts, reads the object back
through the tool's own read path, and asserts each value persisted.

Contacts have no outward side effect: nothing is emailed, nothing is shared.
Everything created is deleted in a `finally` and named with LIVE_WRITE_NAME.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
"""

from __future__ import annotations

from contextlib import asynccontextmanager

import pytest

from outlook_mcp.tools.contacts import create_contact, delete_contact, get_contact, update_contact
from tests.conftest import LIVE_WRITE_NAME

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]


@asynccontextmanager
async def _temporary_contact(client, config, **kwargs):
    created = await create_contact(client.sdk_client, config=config, **kwargs)
    contact_id = created["id"]
    try:
        yield contact_id
    finally:
        await delete_contact(client.sdk_client, contact_id, config=config)


class TestCreateContactRoundTrip:
    async def test_every_create_field_persists(self, real_graph_client, live_write_config):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Roundtrip",
            email="live.write.guard@example.com",
            phone="+15555550199",
            company="LiveWriteGuard Co",
            title="Probe",
        ) as contact_id:
            got = await get_contact(real_graph_client.sdk_client, contact_id)

            assert got["first_name"] == LIVE_WRITE_NAME
            assert got["last_name"] == "Roundtrip"
            assert "live.write.guard@example.com" in [
                e.get("address") if isinstance(e, dict) else e for e in got["email_addresses"]
            ]
            assert got["mobile_phone"] == "+15555550199"
            assert got["company"] == "LiveWriteGuard Co"
            assert got["title"] == "Probe"


class TestUpdateContactRoundTrip:
    async def test_every_update_field_persists(self, real_graph_client, live_write_config):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Before",
            email="before@example.com",
            phone="+15555550198",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                first_name=f"{LIVE_WRITE_NAME}-After",
                last_name="After",
                email="after@example.com",
                phone="+15555550197",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["first_name"] == f"{LIVE_WRITE_NAME}-After"
            assert got["last_name"] == "After"
            assert got["mobile_phone"] == "+15555550197"
            assert "after@example.com" in [
                e.get("address") if isinstance(e, dict) else e for e in got["email_addresses"]
            ]

    async def test_partial_update_leaves_other_fields_alone(
        self, real_graph_client, live_write_config
    ):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Stable",
            phone="+15555550196",
            company="Keep Me",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                last_name="Changed",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["last_name"] == "Changed"
            assert got["first_name"] == LIVE_WRITE_NAME
            assert got["mobile_phone"] == "+15555550196"
            assert got["company"] == "Keep Me"
