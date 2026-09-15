"""Aggregated multi-account reads: fan out concurrently, merge, tag, return.

The per-capability routing (``capability_accounts``) deliberately shows the
agent *one* account per capability. The aggregate tools are the other shape:
one listing across **every** authenticated account, each item tagged with the
account it came from — "what's new anywhere", not "what's new on the routed
account". ``allow_aggregate`` gates them independently of
``allow_cross_account`` (cross gates deliberately switching routing; aggregate
gates bulk cross-account reads).

Design stances worth knowing before editing:

- **Per-account failure isolation.** One account's expired token or Graph
  error lands in ``errors``; the other accounts' results still return. This is
  the opposite trade-off from ``digest.changes_since``, where a genuine error
  cancels siblings — there the three resources belong to one mailbox and a
  partial digest misleads; here the accounts are independent and a partial
  aggregate is exactly what the operator asked for.
- **No cross-account pagination.** Merging per-account cursors into one is a
  token-passing protocol the caller would have to round-trip; instead the
  merged listing is sorted and truncated to ``count``, and ``has_more`` says
  whether any account had more. Deep pagination is what the per-account tools
  are for.
- **Session routing overrides are ignored.** switch_account(capability=...)
  re-routes the *single-account* tools; the aggregate tools' definition is
  "all authenticated accounts", so overrides don't apply.
"""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.tools import calendar_read, mail_read, todo

_ERROR_TEXT_CAP = 300


def require_aggregate(config: Config) -> None:
    """Refuse unless the aggregate switch is on. ValueError, not an
    OutlookMCPError, so _wrap_tool_errors routes it to the model as an
    anticipated failure the caller can read — same pattern as the
    allow_cross_account gate in AuthManager.switch_account."""
    if not config.allow_aggregate:
        raise ValueError(
            "Aggregated multi-account queries are disabled (allow_aggregate=false). "
            "Set allow_aggregate=true in ~/.outlook-mcp/config.json to let the "
            "outlook_list_*_all tools fan out across every authenticated account."
        )


async def _fan_out(
    clients: dict[str | None, Any],
    runner: Callable[[Any], Awaitable[list[dict]]],
) -> tuple[dict[str | None, list[dict]], list[dict]]:
    """Run ``runner`` against every client concurrently, isolating failures.

    Returns (account -> items, errors) where errors is a list of
    ``{"account", "error"}`` dicts — a failed account never cancels its
    siblings, and the error text keeps any recovery hint the exception
    carried (OutlookMCPError embeds one in str()).
    """

    async def _one(
        account: str | None, client: Any
    ) -> tuple[str | None, list[dict] | None, str | None]:
        try:
            return account, await runner(client), None
        except Exception as exc:  # noqa: BLE001 — the whole point is isolation
            return account, None, str(exc)[:_ERROR_TEXT_CAP]

    results = await asyncio.gather(*(_one(a, c) for a, c in clients.items()))
    items: dict[str | None, list[dict]] = {}
    errors: list[dict] = []
    for account, ok_items, error in results:
        if error is not None:
            # gather completion order is nondeterministic; the wire output is not.
            errors.append({"account": account or "default", "error": error})
        else:
            items[account] = ok_items or []
    return items, errors


def _merged(
    per_account: dict[str | None, list[dict]],
    sort_key: Callable[[dict], str],
    reverse: bool,
    count: int,
    items_key: str,
    accounts: list[str],
    skipped: list[str],
    errors: list[dict],
) -> dict:
    """Tag, globally sort, truncate, and wrap into the shared response shape.

    ``has_more`` is the truncation signal (more merged items than ``count``) —
    per-account next-pages are deliberately not surfaced; see the module
    docstring's no-cross-account-pagination stance.
    """
    tagged: list[dict] = []
    for account, entries in per_account.items():
        for entry in entries:
            entry["account"] = account or "default"
            tagged.append(entry)
    tagged.sort(key=sort_key, reverse=reverse)
    return {
        items_key: tagged[:count],
        "count": min(len(tagged), count),
        "accounts_queried": accounts,
        "skipped_unauthenticated": skipped,
        "errors": errors,
        "has_more": len(tagged) > count,
    }


async def list_inbox_all(
    clients: dict[str | None, Any],
    skipped: list[str],
    timezone: str,
    folder: str = "inbox",
    count: int = 25,
    unread_only: bool = False,
    concise: bool = False,
) -> dict:
    """One inbox listing across every account, newest first."""

    async def runner(client: Any) -> list[dict]:
        result = await mail_read.list_inbox(
            client.sdk_client,
            folder=folder,
            count=count,
            unread_only=unread_only,
            concise=concise,
            timezone=timezone,
        )
        return result["messages"]

    per_account, errors = await _fan_out(clients, runner)
    return _merged(
        per_account,
        sort_key=lambda m: m.get("received") or "",
        reverse=True,
        count=count,
        items_key="messages",
        accounts=[a or "default" for a in clients],
        skipped=skipped,
        errors=errors,
    )


async def list_events_all(
    clients: dict[str | None, Any],
    skipped: list[str],
    timezone: str,
    days: int = 7,
    count: int = 50,
    concise: bool = False,
) -> dict:
    """One event listing across every account, soonest first."""

    async def runner(client: Any) -> list[dict]:
        result = await calendar_read.list_events(
            client.sdk_client,
            days=days,
            count=count,
            timezone=timezone,
            concise=concise,
        )
        return result["events"]

    per_account, errors = await _fan_out(clients, runner)
    return _merged(
        per_account,
        # start is "2026-09-15T10:00:00.0000000 (UTC)" — the string sorts
        # chronologically within one timezone label, which is what a merged
        # day-view needs; exact cross-timezone ordering isn't worth parsing.
        sort_key=lambda e: e.get("start") or "",
        reverse=False,
        count=count,
        items_key="events",
        accounts=[a or "default" for a in clients],
        skipped=skipped,
        errors=errors,
    )


async def list_tasks_all(
    clients: dict[str | None, Any],
    skipped: list[str],
    status: str | None = None,
    count: int = 25,
) -> dict:
    """One task listing across every account and every task list, newest first.

    Each account's lists are queried concurrently too: N accounts x M lists is
    the latency of the slowest single request, not the sum.
    """

    async def runner(client: Any) -> list[dict]:
        lists_result = await todo.list_task_lists(client.sdk_client)
        lists = lists_result["task_lists"]

        async def _one_list(task_list: dict) -> list[dict]:
            tasks_result = await todo.list_tasks(
                client.sdk_client,
                list_id=task_list["id"],
                status=status,
                count=count,
            )
            for task in tasks_result["tasks"]:
                task["list"] = task_list["display_name"]
            return tasks_result["tasks"]

        per_list = await asyncio.gather(*(_one_list(tl) for tl in lists))
        return [task for tasks in per_list for task in tasks]

    per_account, errors = await _fan_out(clients, runner)
    return _merged(
        per_account,
        sort_key=lambda t: t.get("created") or "",
        reverse=True,
        count=count,
        items_key="tasks",
        accounts=[a or "default" for a in clients],
        skipped=skipped,
        errors=errors,
    )
