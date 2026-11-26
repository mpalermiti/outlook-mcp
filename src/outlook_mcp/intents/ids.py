"""Immutable ids: one id family per process (contracts §6.3, decision D2).

Outlook ids change when an item moves between folders, unless the request asks
for immutable ids with ``Prefer: IdType="ImmutableId"``. Messages, attachments,
events, meeting messages and contacts support them; folders and calendars keep
constant ids anyway, and To Do tasks and inbox rules are not covered.

A process that loads intent tools sends the preference on every Outlook request,
so one process hands out one id family and an agent never compares two ids for
the same message. Whether Graph honours it on a personal account is not
documented, so the family is learned at run time from ``Preference-Applied``:

- the first id-bearing reply sets it, ``"immutable"`` or ``"rest"``;
- a later reply that disagrees makes it ``"mixed"`` for the rest of the process,
  with a warning logged.

The family flag is short-lived action state (the "no local caching" carve-out in
contracts §0): one word, never mailbox data, gone at exit.

**Which requests.** ``IMMUTABLE_IDS`` is set once, at import, by the server: true
when the process loads an intent tool (``wanted_for``), false otherwise, so a
classic install sends nothing new. When it is on:

- every SDK request carries the preference, through hooks on the Graph client's
  HTTP client (``sdk_client``), and each id-bearing reply is heard;
- the classic tools' raw paths (delta pages, ``$batch`` envelopes and each
  sub-request) add it through ``with_preference``;
- the intent tools' own path (``transport``) always adds it.
"""

from __future__ import annotations

import logging
from collections.abc import Collection, Mapping
from typing import Any
from urllib.parse import urlsplit

logger = logging.getLogger(__name__)

IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"'

# Set once by server.py at import: does this process load an intent tool?
IMMUTABLE_IDS = False

# Path segments whose replies carry Outlook ids that the preference changes. Folders,
# calendars, To Do and rules keep their ids either way, so they don't vote.
_ID_SEGMENTS = frozenset({"messages", "events", "calendarview", "contacts", "attachments"})


def _preferences(header: str | None) -> list[str]:
    return [part.strip() for part in (header or "").split(",") if part.strip()]


def _is_immutable_preference(part: str) -> bool:
    name, _, value = part.partition("=")
    return name.strip().lower() == "idtype" and value.strip().strip('"').lower() == "immutableid"


def combine_prefer(existing: str | None) -> str:
    """``existing``'s preferences with ``IdType="ImmutableId"`` added once, comma-separated.

    Never assigned over another preference: ``odata.maxpagesize`` sizes delta
    pages and ``outlook.timezone`` sets the zone of event times, and replacing
    either changes the answer without an error.
    """
    parts = _preferences(existing)
    if not any(_is_immutable_preference(p) for p in parts):
        parts.append(IMMUTABLE_ID_PREFERENCE)
    return ", ".join(parts)


def immutable_applied(preference_applied: str | None) -> bool:
    """Whether a ``Preference-Applied`` value says Graph returned immutable ids."""
    return any(_is_immutable_preference(p) for p in _preferences(preference_applied))


class IdFamily:
    """The id family this process has been handed: ``None`` until the first id-bearing reply."""

    def __init__(self) -> None:
        self.family: str | None = None

    def observe(self, preference_applied: str | None) -> str:
        """Record one successful id-bearing reply's ``Preference-Applied`` header."""
        seen = "immutable" if immutable_applied(preference_applied) else "rest"
        if self.family is None:
            self.family = seen
        elif self.family not in (seen, "mixed"):
            self.family = "mixed"
            logger.warning(
                "Graph returned both immutable and regular ids in this process; ids from "
                "earlier and later answers may not compare equal (id family: mixed)."
            )
        return self.family


# Process-wide: every intent request reports to it.
ID_FAMILY = IdFamily()


# ── The process switch ────────────────────────────────────────────────────


def wanted_for(loaded_tools: Collection[str]) -> bool:
    """Whether a process loading ``loaded_tools`` asks for immutable ids (decision D2)."""
    from outlook_mcp import toolsets

    return any(toolsets.TOOL_GROUPS.get(name) == toolsets.INTENTS for name in loaded_tools)


def with_preference(headers: Mapping[str, str] | None) -> Any:
    """``headers`` with the preference combined into ``Prefer``, when the switch is on.

    Off, ``headers`` comes back as it was (the same object), so a classic install's
    requests are byte for byte what they were. On, a copy comes back.
    """
    if not IMMUTABLE_IDS:
        return headers if headers is not None else {}
    combined = dict(headers or {})
    key = next((k for k in combined if k.lower() == "prefer"), "Prefer")
    combined[key] = combine_prefer(combined.get(key))
    return combined


def carries_outlook_ids(url: object) -> bool:
    """Whether a reply to ``url`` carries ids the preference applies to."""
    segments = urlsplit(str(url)).path.lower().split("/")
    return any(segment in _ID_SEGMENTS for segment in segments)


async def _ask(request: Any) -> None:
    request.headers["Prefer"] = combine_prefer(request.headers.get("Prefer"))


async def _hear(response: Any) -> None:
    if 200 <= response.status_code < 300 and carries_outlook_ids(response.request.url):
        ID_FAMILY.observe(response.headers.get("Preference-Applied"))


def base_client() -> Any:
    """kiota's default HTTP client, the one the SDK builds when given none."""
    from kiota_http.kiota_client_factory import KiotaClientFactory

    client = KiotaClientFactory.get_default_client()
    client.base_url = "https://graph.microsoft.com/v1.0"
    return client


def sdk_client() -> Any:
    """The Graph SDK's HTTP client with its default middleware, plus the preference hooks.

    The request hook runs before the middleware, so kiota's own retries carry the
    header too; the response hook hears each id-bearing reply.
    """
    from msgraph import graph_request_adapter
    from msgraph_core import GraphClientFactory

    client = base_client()
    client.event_hooks = {"request": [_ask], "response": [_hear]}
    return GraphClientFactory.create_with_default_middleware(
        client=client, options=getattr(graph_request_adapter, "options", None)
    )
