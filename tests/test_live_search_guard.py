"""Exercise the live search guard's folder choice without mailbox credentials."""

import importlib.util
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest


@pytest.fixture
def guard():
    spec = importlib.util.spec_from_file_location(
        "search_query_guard", Path(__file__).with_name("test_live_query_shape.py")
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.mark.asyncio
async def test_a_small_inbox_does_not_hide_a_populated_archive(guard, monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_LIVE_SEARCH_FOLDER", raising=False)
    listing = AsyncMock(
        return_value={
            "folders": [
                {"id": "inbox-id", "total": 13},
                {"id": "archive-id", "total": 400},
                {"id": "sent-id", "total": 100},
            ]
        }
    )
    search = AsyncMock(
        return_value={
            "count": 275,
            "messages": [{"id": str(i)} for i in range(275)],
            "cursor": None,
            "incomplete": "Older matches may be missing",
        }
    )
    monkeypatch.setattr(guard, "list_folders", listing)
    monkeypatch.setattr(guard, "search_mail", search)
    client = MagicMock()
    try:
        await guard.test_a_search_graph_cut_short_says_so(SimpleNamespace(sdk_client=client))
    except pytest.skip.Exception as exc:
        pytest.fail(f"the populated archive must exercise the guard, not skip: {exc}")
    assert search.call_args.kwargs["folder"] == "archive-id"
    listing.assert_awaited_once_with(client)
    client.me.mail_folders.by_mail_folder_id.assert_not_called()


@pytest.mark.asyncio
async def test_an_explicit_folder_can_reach_beyond_the_listing(guard, monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_LIVE_SEARCH_FOLDER", "Chosen")
    resolve = AsyncMock(return_value="chosen-id")
    listing = AsyncMock()
    search = AsyncMock(
        return_value={
            "count": 275,
            "messages": [{"id": str(i)} for i in range(275)],
            "cursor": None,
            "incomplete": "Older matches may be missing",
        }
    )
    monkeypatch.setattr(guard, "resolve_folder_id", resolve)
    monkeypatch.setattr(guard, "list_folders", listing)
    monkeypatch.setattr(guard, "search_mail", search)
    client = MagicMock()
    client.me.mail_folders.by_mail_folder_id.return_value.get = AsyncMock(
        return_value=SimpleNamespace(total_item_count=400)
    )
    await guard.test_a_search_graph_cut_short_says_so(SimpleNamespace(sdk_client=client))
    resolve.assert_awaited_once_with(client, "Chosen")
    assert search.call_args.kwargs["folder"] == "chosen-id"
    listing.assert_not_awaited()
