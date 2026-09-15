"""Tests for the aggregate multi-account read tools.

The contract under test: fan out to every account concurrently, tag each item
with its account, sort globally, truncate — and never let one account's
failure take the merged listing down.
"""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.aggregation import (
    list_events_all,
    list_inbox_all,
    list_tasks_all,
    require_aggregate,
)
from outlook_mcp.config import Config


class TestRequireAggregate:
    def test_off_refuses_with_the_remedy(self):
        with pytest.raises(ValueError, match="allow_aggregate=true"):
            require_aggregate(Config())

    def test_on_passes(self):
        require_aggregate(Config(allow_aggregate=True))


class TestInboxAll:
    async def test_merges_tags_and_sorts_newest_first(self, monkeypatch):
        from outlook_mcp.tools import mail_read

        async def fake(graph_client, **kwargs):
            if graph_client is NET.sdk_client:
                return {"messages": [{"id": "n1", "received": "2026-09-14T10:00:00"}]}
            return {"messages": [{"id": "k1", "received": "2026-09-15T09:00:00"}]}

        monkeypatch.setattr(mail_read, "list_inbox", fake)
        result = await list_inbox_all({"net": NET, "neko": NEKO}, [], "UTC")

        assert [m["id"] for m in result["messages"]] == ["k1", "n1"]  # newest first
        assert result["messages"][0]["account"] == "neko"
        assert result["messages"][1]["account"] == "net"
        assert result["accounts_queried"] == ["net", "neko"]
        assert result["errors"] == []

    async def test_one_account_failing_does_not_sink_the_rest(self, monkeypatch):
        from outlook_mcp.tools import mail_read

        async def fake(graph_client, **kwargs):
            if graph_client is NET.sdk_client:
                raise RuntimeError("token exploded")
            return {"messages": [{"id": "k1", "received": "2026-09-15T09:00:00"}]}

        monkeypatch.setattr(mail_read, "list_inbox", fake)
        result = await list_inbox_all({"net": NET, "neko": NEKO}, [], "UTC")

        assert [m["id"] for m in result["messages"]] == ["k1"]
        assert len(result["errors"]) == 1
        assert result["errors"][0]["account"] == "net"
        assert "token exploded" in result["errors"][0]["error"]

    async def test_truncates_after_the_global_sort(self, monkeypatch):
        from outlook_mcp.tools import mail_read

        async def fake(graph_client, **kwargs):
            if graph_client is NET.sdk_client:
                return {"messages": [{"id": "n1", "received": "2026-09-14T10:00:00"}]}
            return {"messages": [{"id": "k1", "received": "2026-09-15T09:00:00"}]}

        monkeypatch.setattr(mail_read, "list_inbox", fake)
        result = await list_inbox_all({"net": NET, "neko": NEKO}, [], "UTC", count=1)

        assert [m["id"] for m in result["messages"]] == ["k1"]  # the newest survives
        assert result["count"] == 1
        assert result["has_more"] is True

    async def test_skipped_accounts_are_reported_not_silently_dropped(self, monkeypatch):
        from outlook_mcp.tools import mail_read

        async def fake(graph_client, **kwargs):
            return {"messages": []}

        monkeypatch.setattr(mail_read, "list_inbox", fake)
        result = await list_inbox_all({"net": NET}, ["neko"], "UTC")

        assert result["skipped_unauthenticated"] == ["neko"]


class TestEventsAll:
    async def test_sorts_soonest_first_across_accounts(self, monkeypatch):
        from outlook_mcp.tools import calendar_read

        async def fake(graph_client, **kwargs):
            if graph_client is NET.sdk_client:
                return {"events": [{"id": "e1", "start": "2026-09-16T09:00:00 (UTC)"}]}
            return {"events": [{"id": "e2", "start": "2026-09-15T09:00:00 (UTC)"}]}

        monkeypatch.setattr(calendar_read, "list_events", fake)
        result = await list_events_all({"net": NET, "neko": NEKO}, [], "UTC")

        assert [e["id"] for e in result["events"]] == ["e2", "e1"]  # chronological
        assert result["events"][0]["account"] == "neko"


class TestTasksAll:
    async def test_flattens_lists_and_tags_both_account_and_list(self, monkeypatch):
        from outlook_mcp.tools import todo

        async def fake_lists(graph_client):
            if graph_client is NET.sdk_client:
                return {"task_lists": [{"id": "L1", "display_name": "任务"}]}
            return {"task_lists": [{"id": "L2", "display_name": "作业"}]}

        async def fake_tasks(graph_client, list_id=None, status=None, count=25):
            if list_id == "L1":
                return {"tasks": [{"id": "t1", "created": "2026-09-13T08:00:00"}]}
            return {"tasks": [{"id": "t2", "created": "2026-09-14T08:00:00"}]}

        monkeypatch.setattr(todo, "list_task_lists", fake_lists)
        monkeypatch.setattr(todo, "list_tasks", fake_tasks)
        result = await list_tasks_all({"net": NET, "neko": NEKO}, [])

        assert [t["id"] for t in result["tasks"]] == ["t2", "t1"]  # newest first
        assert result["tasks"][0]["account"] == "neko"
        assert result["tasks"][0]["list"] == "作业"
        assert result["tasks"][1]["list"] == "任务"


class TestSingleAccountShape:
    async def test_legacy_account_key_renders_as_default(self, monkeypatch):
        from outlook_mcp.tools import mail_read

        async def fake(graph_client, **kwargs):
            return {"messages": [{"id": "s1", "received": "2026-09-15T00:00:00"}]}

        monkeypatch.setattr(mail_read, "list_inbox", fake)
        result = await list_inbox_all({None: SINGLE}, [], "UTC")

        assert result["messages"][0]["account"] == "default"
        assert result["accounts_queried"] == ["default"]


# Module-level fakes: identity matters — the runners branch on `is`, which is
# exactly what the real code does with per-account clients.
NET = MagicMock()
NEKO = MagicMock()
SINGLE = MagicMock()
