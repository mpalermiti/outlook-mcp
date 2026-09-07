"""Probe: what does Graph actually do with the datetimes create_event sends?

One-off evidence gathering, not a guard. `create_event` sets
``event.start.date_time`` to the caller's raw string and hardcodes
``time_zone = "UTC"``, discarding what ``validate_datetime`` normalized. For a
`Z`-suffixed input that is self-consistent. For an offset-bearing input
(``12:30:00+02:00``) or a naive one (``12:30:00``) the payload is ambiguous —
an offset inside a field labelled UTC — and Graph's behavior there is not
documented.

The fix for that is not obvious in the safe direction either: the discarded
normalization calls ``.astimezone()`` on naive input, which resolves against
the *server's* local clock, so the same string means different instants on a
UTC container and on a laptop in California. Before changing any of it, this
script establishes what the three input shapes currently produce end to end.

Creates three throwaway events, reads back the instant Graph resolved for each
(via ``Prefer: outlook.timezone="UTC"``), prints a table, and deletes them.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run python scripts/probe_datetime_semantics.py
"""

from __future__ import annotations

import asyncio
import os
import sys
from datetime import date, timedelta

import httpx

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config
from outlook_mcp.graph import GraphClient
from outlook_mcp.tools.calendar_write import create_event, delete_event

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
SUBJECT = "[outlook-mcp probe] datetime semantics — safe to delete"


def _token() -> str:
    config = load_config()
    am = AuthManager(config)
    am.try_cached_token(am.get_token_scopes())
    return am.get_credential().get_token("https://graph.microsoft.com/.default").token


def _read_back(token: str, event_id: str) -> tuple[str, str]:
    """The instant Graph resolved, asked for in UTC, plus what it stored verbatim."""
    r = httpx.get(
        urljoin_event(event_id),
        headers={
            "Authorization": f"Bearer {token}",
            # Ask Graph to project start/end into UTC — this is the resolved instant.
            "Prefer": 'outlook.timezone="UTC"',
        },
        timeout=20,
    )
    r.raise_for_status()
    body = r.json()
    start = body.get("start") or {}
    return start.get("dateTime", "?"), start.get("timeZone", "?")


def urljoin_event(event_id: str) -> str:
    from urllib.parse import quote

    return f"{GRAPH_BASE}me/events/{quote(event_id, safe='')}"


# Three spellings of the same wall-clock intent, at 12:30 on the anchor day.
# `Z` is the shape the tool's own docstring example uses; the other two are what
# real clients send.
CASES = [
    ("Z-suffixed  ", "T12:30:00Z", "T13:30:00Z"),
    ("+02:00 offset", "T12:30:00+02:00", "T13:30:00+02:00"),
    ("naive       ", "T12:30:00", "T13:30:00"),
]


async def run() -> int:
    if os.environ.get("OUTLOOK_MCP_LIVE_WRITE") != "1":
        print("Refusing to run: set OUTLOOK_MCP_LIVE_WRITE=1 (this creates real events).")
        return 2

    config = load_config()
    if config.read_only:
        print("Config is read_only — nothing to probe.")
        return 2

    am = AuthManager(config)
    if not am.try_cached_token(am.get_token_scopes()):
        print("Not authenticated — run `outlook-mcp auth` first.")
        return 2

    token = _token()
    client = GraphClient(am.get_credential())
    anchor = (date.today() + timedelta(days=45)).isoformat()

    print(f"anchor day: {anchor}   host TZ: {os.environ.get('TZ', '(system default)')}")
    print(f"config.timezone: {config.timezone}\n")
    print(f"{'input':14s} {'sent as':26s} {'Graph resolved (UTC)':28s}")
    print("-" * 70)

    rows: list[tuple[str, str, str]] = []
    for label, start_suffix, end_suffix in CASES:
        start = f"{anchor}{start_suffix}"
        event_id = None
        try:
            created = await create_event(
                client.sdk_client,
                subject=SUBJECT,
                start=start,
                end=f"{anchor}{end_suffix}",
                config=config,
            )
            event_id = created["event_id"]
            resolved, tz = _read_back(token, event_id)
            rows.append((label, start, f"{resolved} ({tz})"))
        except Exception as exc:  # noqa: BLE001 — a rejection is itself the finding
            rows.append((label, start, f"REJECTED: {type(exc).__name__}: {exc}"))
        finally:
            if event_id:
                await delete_event(client.sdk_client, event_id, config=config)

    for label, sent, resolved in rows:
        print(f"{label:14s} {sent:26s} {resolved}")

    print(
        "\nRead this as: identical wall-clock intent should resolve to the same instant "
        "only where the caller said so. Any row that lands somewhere the caller did not "
        "ask for is the PR 2 bug, and the naive row is the one to re-run under a "
        "different TZ= to check for host dependence."
    )
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(run()))
