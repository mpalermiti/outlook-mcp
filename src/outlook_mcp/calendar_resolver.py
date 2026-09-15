"""Resolve calendar references (display names or IDs) to Graph calendar IDs.

The sibling of ``folder_resolver``, with one deliberate difference. A mail
folder can be addressed by a well-known name that never needs a lookup, but a
calendar has no such names — only the default, which has its own URL — so
every non-default reference costs one ``/me/calendars`` listing. That listing
is what the ID check runs against too: an ID that is not one of the user's
calendars is refused here, with the real list in the message, instead of
going to Graph and coming back as a bare 400 or 404.
"""

from __future__ import annotations

import unicodedata
from typing import Any

from outlook_mcp.pagination import build_request_config
from outlook_mcp.validation import sanitize_output

# The one alias for the default calendar; blank means the same thing.
_PRIMARY = "primary"


def _fold(name: str) -> str:
    """Case- and normalization-insensitive key for a display name.

    ``casefold`` rather than ``lower`` so "STRASSE" finds "Straße"; NFC so a
    name typed with a combining accent (macOS, some models) matches the
    precomposed form Graph stores.
    """
    return unicodedata.normalize("NFC", name).casefold()


async def fetch_all_calendars(graph_client: Any, top: int = 100) -> list[Any]:
    """Every calendar in ``/me/calendars``, following ``@odata.nextLink``.

    Graph documents no page size for this collection, only that any collection
    may page; asking for ``$top`` and following the link is the contract that
    holds either way (``folder_resolver.fetch_all_top_level_folders`` does the
    same for mail folders).
    """
    from msgraph.generated.users.item.calendars.calendars_request_builder import (
        CalendarsRequestBuilder,
    )

    config = build_request_config(
        CalendarsRequestBuilder.CalendarsRequestBuilderGetQueryParameters, {"$top": top}
    )
    response = await graph_client.me.calendars.get(request_configuration=config)
    collected = list(response.value) if response and response.value else []
    while True:
        next_link = getattr(response, "odata_next_link", None)
        if not isinstance(next_link, str) or not next_link:
            break
        response = await graph_client.me.calendars.with_url(next_link).get()
        collected.extend(list(response.value) if response and response.value else [])
    return collected


async def resolve_calendar_id(graph_client: Any, calendar_ref: str | None) -> str | None:
    """Resolve a caller-supplied calendar reference to a Graph calendar ID.

    Returns ``None`` for the default calendar — ``None``, blank, or "primary"
    in any case — so the caller keeps the single-round-trip ``/me/calendarView``
    path. Anything else is matched against the user's calendars: first as an
    exact ID, then as a display name (case-insensitive). Names are matched
    without first guessing whether the value "looks like" an ID, because
    calendar names routinely carry the characters and lengths such a guess
    trips on ("Kids + School", "Calendar - Jane Smith (jane@…)").

    Raises ValueError, naming the calendars that do exist, when nothing matches
    or more than one name does.
    """
    if calendar_ref is None:
        return None
    trimmed = calendar_ref.strip()
    if not trimmed or trimmed.casefold() == _PRIMARY:
        return None

    calendars = await fetch_all_calendars(graph_client)

    for cal in calendars:
        if cal.id and cal.id == trimmed:
            return cal.id

    wanted = _fold(trimmed)
    matches = [cal for cal in calendars if cal.name and _fold(cal.name) == wanted]
    shown = sanitize_output(trimmed[:50])

    if len(matches) == 1:
        if not matches[0].id:
            # None means "the default calendar" to the caller, so a calendar
            # Graph returned without an ID must not fall through as one.
            raise ValueError(f"Calendar '{shown}' was returned with no ID; cannot query it.")
        return matches[0].id
    if len(matches) > 1:
        candidates = ", ".join(f"{sanitize_output(cal.name)} = {cal.id}" for cal in matches)
        raise ValueError(
            f"Calendar name '{shown}' is ambiguous ({len(matches)} matches): {candidates}. "
            "Pass the calendar ID instead."
        )
    available = ", ".join(sanitize_output(cal.name or "(unnamed)") for cal in calendars)
    raise ValueError(
        f"Calendar '{shown}' not found. Available calendars: {available}. "
        "Pass a display name or an ID from outlook_list_calendars."
    )
