"""Write-tier live guards for To Do: does every field survive a round trip?

Tier 3 of the silent-no-op audit — see test_live_contacts_write.py for the
rationale. Tasks have no outward side effect. Everything created is deleted
in a `finally` and titled with LIVE_WRITE_NAME.

Read-back goes through get_task (single task, $expand=checklistItems) or
list_tasks (overview shape). The default list is scanned up to 100 entries.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
"""

from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import date, timedelta
from pathlib import Path

import pytest

from outlook_mcp.tools.todo import (
    add_checklist_item,
    create_task,
    delete_checklist_item,
    delete_task,
    get_task,
    list_tasks,
    update_checklist_item,
    update_task,
)
from outlook_mcp.tools.todo_attachments import (
    delete_task_attachment,
    download_task_attachment,
    list_task_attachments,
    upload_task_attachment,
)
from tests.conftest import LIVE_WRITE_NAME

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]


async def _find_task(client, task_id: str) -> dict:
    page = await list_tasks(client.sdk_client, count=100)
    for task in page["tasks"]:
        if task["id"] == task_id:
            return task
    raise AssertionError(f"task {task_id} not found in the first 100 of the default list")


@asynccontextmanager
async def _temporary_task(client, config, **kwargs):
    created = await create_task(client.sdk_client, config=config, **kwargs)
    task_id = created["task_id"]
    try:
        yield task_id
    finally:
        await delete_task(client.sdk_client, task_id, config=config)


class TestCreateTaskRoundTrip:
    async def test_every_create_field_persists(self, real_graph_client, live_write_config):
        due = (date.today() + timedelta(days=30)).isoformat()

        async with _temporary_task(
            real_graph_client,
            live_write_config,
            title=f"{LIVE_WRITE_NAME} create",
            due=f"{due}T17:00:00Z",
            importance="high",
            body="live write guard body",
            reminder=True,
            recurrence={
                "pattern": {"type": "daily", "interval": 2},
                "range": {"type": "noEnd", "startDate": due},
            },
        ) as task_id:
            got = await _find_task(real_graph_client, task_id)

            assert got["title"] == f"{LIVE_WRITE_NAME} create"
            assert got["importance"] == "high"
            assert due in (got["due"] or "")
            assert "live write guard body" in (got["body"] or "")
            assert got["is_reminder_on"] is True
            assert got["has_recurrence"] is True


class TestUpdateTaskRoundTrip:
    async def test_every_update_field_persists(self, real_graph_client, live_write_config):
        due = (date.today() + timedelta(days=31)).isoformat()

        async with _temporary_task(
            real_graph_client,
            live_write_config,
            title=f"{LIVE_WRITE_NAME} before",
            importance="low",
        ) as task_id:
            await update_task(
                real_graph_client.sdk_client,
                task_id=task_id,
                title=f"{LIVE_WRITE_NAME} after",
                due=f"{due}T09:00:00Z",
                body="updated body",
                importance="high",
                config=live_write_config,
            )

            got = await _find_task(real_graph_client, task_id)
            assert got["title"] == f"{LIVE_WRITE_NAME} after"
            assert got["importance"] == "high"
            assert due in (got["due"] or "")
            assert "updated body" in (got["body"] or "")

    async def test_partial_update_leaves_other_fields_alone(
        self, real_graph_client, live_write_config
    ):
        async with _temporary_task(
            real_graph_client,
            live_write_config,
            title=f"{LIVE_WRITE_NAME} stable",
            importance="high",
            body="keep me",
        ) as task_id:
            await update_task(
                real_graph_client.sdk_client,
                task_id=task_id,
                title=f"{LIVE_WRITE_NAME} renamed",
                config=live_write_config,
            )

            got = await _find_task(real_graph_client, task_id)
            assert got["title"] == f"{LIVE_WRITE_NAME} renamed"
            assert got["importance"] == "high"
            assert "keep me" in (got["body"] or "")


class TestChecklistRoundTrip:
    async def test_checklist_items_survive_a_round_trip(self, real_graph_client, live_write_config):
        async with _temporary_task(
            real_graph_client,
            live_write_config,
            title=f"{LIVE_WRITE_NAME} checklist",
        ) as task_id:
            first = await add_checklist_item(
                real_graph_client.sdk_client,
                task_id=task_id,
                display_name="step one",
                config=live_write_config,
            )
            await add_checklist_item(
                real_graph_client.sdk_client,
                task_id=task_id,
                display_name="step two",
                config=live_write_config,
            )

            got = await get_task(real_graph_client.sdk_client, task_id)
            # Graph does not promise an expansion order; compare as sets.
            assert {i["display_name"] for i in got["checklist_items"]} == {
                "step one",
                "step two",
            }

            await update_checklist_item(
                real_graph_client.sdk_client,
                task_id=task_id,
                checklist_item_id=first["checklist_item_id"],
                is_checked=True,
                config=live_write_config,
            )

            got = await get_task(real_graph_client.sdk_client, task_id)
            # The order an agent reads as "the next step": we added "step one"
            # then "step two" and checked "step one" — expect that derived
            # order by name, not the list compared against its own sort.
            assert [i["display_name"] for i in got["checklist_items"]] == [
                "step two",
                "step one",
            ]
            checked = next(i for i in got["checklist_items"] if i["is_checked"])
            assert checked["id"] == first["checklist_item_id"]
            # checkedDateTime is server-maintained from isChecked, and comes
            # back as a real datetime — must be emitted as ISO 8601.
            assert checked["checked_at"] is not None
            assert "T" in checked["checked_at"], checked["checked_at"]

            await delete_checklist_item(
                real_graph_client.sdk_client,
                task_id=task_id,
                checklist_item_id=first["checklist_item_id"],
                config=live_write_config,
            )
            got = await get_task(real_graph_client.sdk_client, task_id)
            assert got["checklist_count"] == 1
            assert got["checklist_items"][0]["display_name"] == "step two"


class TestAttachmentRoundTrip:
    """The inline base64 POST is the one thing mocks cannot vouch for."""

    async def test_attachment_survives_upload_download_delete(
        self, real_graph_client, live_write_config
    ):
        # 512 KiB, non-trivially binary — well inside the 20 MiB inline
        # ceiling, far past the point where encoding bugs could hide.
        payload = bytes(range(256)) * 2048
        base = Path(live_write_config.attachments_dir)
        src = base / f"{LIVE_WRITE_NAME}-src.bin"
        dst = base / f"{LIVE_WRITE_NAME}-dst.bin"
        base.mkdir(parents=True, exist_ok=True)
        src.write_bytes(payload)
        try:
            async with _temporary_task(
                real_graph_client,
                live_write_config,
                title=f"{LIVE_WRITE_NAME} attachment",
            ) as task_id:
                up = await upload_task_attachment(
                    real_graph_client.sdk_client,
                    task_id=task_id,
                    file_path=str(src),
                    config=live_write_config,
                )
                assert up["size"] == len(payload)
                assert up["name"] == src.name
                assert up["attachment_id"]

                listing = await list_task_attachments(real_graph_client.sdk_client, task_id)
                assert listing["count"] == 1
                att = listing["attachments"][0]
                assert att["name"] == src.name
                # Graph's `size` may carry service-side overhead above the raw
                # byte count (the Exchange-backed store), so >= rather than ==;
                # the download's byte-for-byte comparison below is the real
                # fidelity check.
                assert att["size"] >= len(payload)

                await download_task_attachment(
                    real_graph_client.sdk_client,
                    task_id=task_id,
                    attachment_id=att["id"],
                    save_path=str(dst),
                    config=live_write_config,
                )
                assert dst.read_bytes() == payload

                await delete_task_attachment(
                    real_graph_client.sdk_client,
                    task_id=task_id,
                    attachment_id=att["id"],
                    config=live_write_config,
                )
                listing = await list_task_attachments(real_graph_client.sdk_client, task_id)
                assert listing["count"] == 0
        finally:
            src.unlink(missing_ok=True)
            dst.unlink(missing_ok=True)
