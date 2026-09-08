"""Write-tier live guards for To Do: does every field survive a round trip?

Tier 3 of the silent-no-op audit — see test_live_contacts_write.py for the
rationale. Tasks have no outward side effect. Everything created is deleted
in a `finally` and titled with LIVE_WRITE_NAME.

There is no get_task tool; read-back goes through list_tasks and finds the
task by id. The default list is scanned up to 100 entries.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
"""

from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import date, timedelta

import pytest

from outlook_mcp.tools.todo import create_task, delete_task, list_tasks, update_task
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
