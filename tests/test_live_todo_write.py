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
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4
from zoneinfo import ZoneInfo

import pytest

from outlook_mcp.tools.todo import (
    _resolve_list_id,
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
        # Pinned to UTC: the written date is asserted on a repeating task, and
        # under another zone Graph can move it (see TestDueIsAnchoredInTheConfiguredZone).
        config = live_write_config.model_copy(update={"timezone": "UTC"})
        due = (date.today() + timedelta(days=30)).isoformat()

        async with _temporary_task(
            real_graph_client,
            config,
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


_LA = "America/Los_Angeles"
# The #72 probe sent a Windows name and got it back; which spelling Graph
# returns for an IANA name has not been seen, so either one counts.
_LA_SPELLINGS = (_LA, "Pacific Standard Time")


def _zone_kept_or_coerced(due_zone, rule_zone, sent) -> bool:
    """Whether a repeating task's stored zones are one of the two seen live.

    The zone sent (in any spelling in ``sent``) or UTC, on the due and the rule
    alike. Michael's account (2026-11-14) stored UTC on both whatever was sent;
    the account in the #72 probe (2026-09-25) kept the zone. Per account, like
    `is_online` (#70), so either passes; a third zone, or the two disagreeing, fails.
    """
    return due_zone == rule_zone == "UTC" or (due_zone in sent and rule_zone in sent)


async def _raw_task(client, task_id: str):
    """The task as Graph holds it: `recurrenceTimeZone` is not in the tool's output."""
    list_id = await _resolve_list_id(client.sdk_client, None)
    tasks = client.sdk_client.me.todo.lists.by_todo_task_list_id(list_id).tasks
    return await tasks.by_todo_task_id(task_id).get()


def _stored_zones(task) -> tuple:
    return task.due_date_time.time_zone, task.recurrence.range.recurrence_time_zone


async def _ids_titled(client, title: str) -> list[str]:
    page = await list_tasks(client.sdk_client, count=100)
    return [t["id"] for t in page["tasks"] if t["title"] == title]


async def _one_task_titled(client, config, title: str, task_id: str) -> None:
    """Fail if moving a due left a second task with ``title``; delete any copy first."""
    copies = [i for i in await _ids_titled(client, title) if i != task_id]
    for copy in copies:
        await delete_task(client.sdk_client, copy, config=config)
    assert not copies, (
        f"changing a repeating task's due left {len(copies)} more task(s) titled "
        f"{title!r} (deleted): a bug in update_task, separate from #72"
    )


def _unique(label: str) -> str:
    return f"{LIVE_WRITE_NAME} {label} {uuid4().hex[:8]}"


class TestDueIsAnchoredInTheConfiguredZone:
    """#72: what Graph does with the zone, on a plain task and on a repeating one.

    The probe on #72 sent 09:00 on 10-15 as China Standard Time. A plain task
    came back as 16:00 UTC on 10-14, midnight on the 15th in that zone: the
    date survived, the hour did not. Michael's run (2026-11-14) confirmed that
    west and east of UTC.

    A repeating task's zone is per account: the #72 probe got its zone back,
    Michael's account stored UTC on the due and the rule whatever was sent.
    These tests accept either (see `_zone_kept_or_coerced`) and assert no date
    on a repeating task, whose due the #72 probe saw move a day.
    """

    @staticmethod
    def _recurrence(start: date) -> dict:
        return {
            "pattern": {"type": "daily", "interval": 2},
            "range": {"type": "noEnd", "startDate": start.isoformat()},
        }

    # Each case sets its own zone, so the machine running it makes no difference.
    # West of UTC midnight there is the written day in UTC; east of it, the day before.
    @pytest.mark.parametrize("zone", [_LA, "Asia/Tokyo"])
    async def test_a_plain_tasks_due_is_midnight_on_the_written_day_in_the_zone(
        self, real_graph_client, live_write_config, zone
    ):
        """A zone Graph ignored would come back at midnight UTC: same day, other hour."""
        config = live_write_config.model_copy(update={"timezone": zone})
        written = date.today() + timedelta(days=30)
        midnight = datetime(written.year, written.month, written.day, tzinfo=ZoneInfo(zone))

        async with _temporary_task(
            real_graph_client,
            config,
            title=f"{LIVE_WRITE_NAME} zone plain",
            due=f"{written.isoformat()}T23:30:00",
        ) as task_id:
            got = await _find_task(real_graph_client, task_id)

            assert got["due_time_zone"] == "UTC"
            stored = datetime.fromisoformat((got["due"] or "").split(".")[0])
            assert stored.replace(tzinfo=timezone.utc) == midnight.astimezone(timezone.utc)

    async def test_a_repeating_task_keeps_the_zone_sent_or_utc(
        self, real_graph_client, live_write_config
    ):
        """Graph accepts `recurrenceTimeZone` on a task; the account decides what it stores."""
        config = live_write_config.model_copy(update={"timezone": _LA})
        written = date.today() + timedelta(days=30)

        async with _temporary_task(
            real_graph_client,
            config,
            title=f"{LIVE_WRITE_NAME} zone repeating",
            due=f"{written.isoformat()}T09:00:00",
            recurrence=self._recurrence(written),
        ) as task_id:
            zones = _stored_zones(await _raw_task(real_graph_client, task_id))

            assert _zone_kept_or_coerced(*zones, _LA_SPELLINGS), (
                f"stored (due, rule) zones {zones}: neither the zone sent nor UTC on both"
            )

    async def test_moving_a_repeating_tasks_due_keeps_its_stored_zone(
        self, real_graph_client, live_write_config
    ):
        """The update sends the task's own zone, not this server's: a 400 fails here.

        Also counts the tasks with this title afterwards: moving a repeating
        task's due must not leave a second one behind.
        """
        created_in = live_write_config.model_copy(update={"timezone": _LA})
        updated_in = live_write_config.model_copy(update={"timezone": "Asia/Tokyo"})
        written = date.today() + timedelta(days=30)
        title = _unique("zone update")

        async with _temporary_task(
            real_graph_client,
            created_in,
            title=title,
            due=f"{written.isoformat()}T09:00:00",
            recurrence=self._recurrence(written),
        ) as task_id:
            before = _stored_zones(await _raw_task(real_graph_client, task_id))
            assert _zone_kept_or_coerced(*before, _LA_SPELLINGS), before

            await update_task(
                real_graph_client.sdk_client,
                task_id=task_id,
                due=f"{(written + timedelta(days=2)).isoformat()}T09:00:00",
                config=updated_in,
            )

            after = _stored_zones(await _raw_task(real_graph_client, task_id))
            await _one_task_titled(real_graph_client, created_in, title, task_id)
            assert after == before, f"stored zones moved from {before} to {after}"

    async def test_mains_due_patch_on_a_repeating_task_leaves_one_task(
        self, real_graph_client, live_write_config
    ):
        """The payload `main` sends, bypassing the tools: is a duplicate a bug there too?

        Before #72, `create_task` sent a UTC-labelled due and no rule zone (as it
        still does on a UTC config) and `update_task` patched the due labelled
        UTC. This sends that patch through the SDK directly, so the answer does
        not depend on this branch's `update_task`.
        """
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.todo_task import TodoTask

        config = live_write_config.model_copy(update={"timezone": "UTC"})
        written = date.today() + timedelta(days=30)
        title = _unique("main due patch")

        async with _temporary_task(
            real_graph_client,
            config,
            title=title,
            due=f"{written.isoformat()}T09:00:00",
            recurrence=self._recurrence(written),
        ) as task_id:
            patch = TodoTask()
            patch.due_date_time = DateTimeTimeZone(
                date_time=f"{(written + timedelta(days=2)).isoformat()}T09:00:00",
                time_zone="UTC",
            )
            list_id = await _resolve_list_id(real_graph_client.sdk_client, None)
            tasks = real_graph_client.sdk_client.me.todo.lists.by_todo_task_list_id(list_id)
            await tasks.tasks.by_todo_task_id(task_id).patch(patch)

            await _one_task_titled(real_graph_client, config, title, task_id)


class TestUpdateTaskRoundTrip:
    async def test_every_update_field_persists(self, real_graph_client, live_write_config):
        # Pinned to UTC: a plain task keeps only the date, picked in the
        # configured zone, so east of UTC 09:00Z would be stored the day before.
        config = live_write_config.model_copy(update={"timezone": "UTC"})
        due = (date.today() + timedelta(days=31)).isoformat()

        async with _temporary_task(
            real_graph_client,
            config,
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
                config=config,
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
