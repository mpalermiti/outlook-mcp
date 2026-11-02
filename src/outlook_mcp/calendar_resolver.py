"""Resolve calendar references (display names or IDs) to Graph calendar IDs.

The sibling of ``folder_resolver``, with one deliberate difference. A mail
folder can be addressed by a well-known name that never needs a lookup, but a
calendar has no such names — only the default, which has its own URL — so
every non-default reference costs one ``/me/calendars`` listing. That listing
is what the ID check runs against too: an ID that is not one of the user's
calendars is refused here, with the real list in the message, instead of
going to Graph and coming back as a bare 400 or 404.

It also says whose a calendar is, for the writes that need ``mail_send`` when
it is not the user's own: whatever is written to a calendar someone shared is
shown to them, as an invitation would be.
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


def _named(calendar_ref: str | None) -> str | None:
    """The reference to look up, or ``None`` when it means the default calendar."""
    if calendar_ref is None:
        return None
    trimmed = calendar_ref.strip()
    if not trimmed or trimmed.casefold() == _PRIMARY:
        return None
    return trimmed


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
    or more than one name does. A write resolves through
    ``resolve_calendar_for_write`` instead.
    """
    trimmed = _named(calendar_ref)
    if trimmed is None:
        return None
    return _match(await fetch_all_calendars(graph_client), trimmed).id


async def resolve_calendar_for_write(
    graph_client: Any, calendar_ref: str | None
) -> tuple[str | None, bool]:
    """The calendar a new event goes to, and whether it belongs to someone else.

    Resolved as ``resolve_calendar_id`` does, with two things more, both read
    off the same ``/me/calendars`` listing, so neither costs a request:

    - A calendar Graph lists with ``canEdit: false`` (a birthdays or holiday
      calendar, a view-only share) is refused here, naming the calendars that
      can be written, because Graph's own answer is a 403 whose hint blames the
      account type. Only an explicit ``false`` refuses: a calendar listed
      without ``canEdit`` goes through, and Graph has the last word.
    - Whose it is, by ``_belongs_to_someone_else``. The listing carries
      ``owner`` and ``isDefaultCalendar`` because it asks for no ``$select``.

    Returns ``(None, False)`` for the default calendar, with no listing: it is
    the user's own.
    """
    trimmed = _named(calendar_ref)
    if trimmed is None:
        return None, False

    calendars = await fetch_all_calendars(graph_client)
    found = _match(calendars, trimmed)
    if getattr(found, "can_edit", None) is False:
        name = sanitize_output((found.name or trimmed)[:50])
        others = ", ".join(
            sanitize_output(cal.name or "(unnamed)")
            for cal in calendars
            if cal is not found and getattr(cal, "can_edit", None) is not False
        )
        raise ValueError(
            f"Calendar '{name}' is read-only to this account (Graph lists it with "
            f"canEdit: false), so nothing can be written to it. Calendars you can "
            f"write to: {others or '(none listed)'}."
        )
    default = next((cal for cal in calendars if _is_default(cal)), None)
    return found.id, _belongs_to_someone_else(found, default)


def _is_default(calendar: Any) -> bool:
    """Only an explicit ``isDefaultCalendar: true`` counts."""
    return getattr(calendar, "is_default_calendar", None) is True


def _owner_address(calendar: Any) -> str | None:
    """The calendar owner's address, lower-cased, or ``None`` when Graph gave none.

    ``lower`` rather than ``casefold``: the comparison only has to ignore case,
    and casefold also maps distinct characters ("ß" and "ss") onto each other,
    which would let two different addresses compare equal.
    """
    address = getattr(getattr(calendar, "owner", None), "address", None)
    if not isinstance(address, str) or not address:
        return None
    return address.lower()


def _belongs_to_someone_else(calendar: Any, default_calendar: Any | None) -> bool:
    """Whether ``calendar`` is someone else's, judged against the default calendar.

    The default calendar is always the user's own. Any other calendar is the
    user's when its ``owner.address`` matches the default calendar's, ignoring
    case: verified live, a consumer account's own calendars all carry the same
    owner address as its default one. Anything that cannot be compared — no
    default calendar, or an owner without an address on either side — counts as
    someone else's, so the write needs ``mail_send`` rather than going through
    on a guess.
    """
    if _is_default(calendar):
        return False
    theirs = _owner_address(calendar)
    mine = _owner_address(default_calendar) if default_calendar is not None else None
    return theirs is None or mine is None or theirs != mine


async def event_calendar_belongs_to_someone_else(graph_client: Any, event_id: str) -> bool:
    """Whether the calendar holding event ``event_id`` is someone else's.

    ``GET /me/events/{id}/calendar?$select=id,isDefaultCalendar,owner`` — the
    default calendar ends it there. Any other is compared, by owner address,
    with ``GET /me/calendar?$select=owner``. Both verified live on a consumer
    account to answer 200 with exactly those fields.

    Fails closed: a lookup that raises counts as someone else's, the same as an
    owner Graph does not report, because the write it guards then needs
    ``mail_send`` rather than going ahead unchecked.
    """
    from kiota_abstractions.base_request_configuration import RequestConfiguration
    from msgraph.generated.users.item.calendar.calendar_request_builder import (
        CalendarRequestBuilder as DefaultCalendarRequestBuilder,
    )
    from msgraph.generated.users.item.events.item.calendar.calendar_request_builder import (
        CalendarRequestBuilder as EventCalendarRequestBuilder,
    )

    event_calendar = EventCalendarRequestBuilder.CalendarRequestBuilderGetQueryParameters(
        select=["id", "isDefaultCalendar", "owner"]
    )
    default_calendar = DefaultCalendarRequestBuilder.CalendarRequestBuilderGetQueryParameters(
        select=["owner"]
    )
    try:
        calendar = await graph_client.me.events.by_event_id(event_id).calendar.get(
            request_configuration=RequestConfiguration(query_parameters=event_calendar)
        )
        if _is_default(calendar):
            return False
        if _owner_address(calendar) is None:
            return True
        default = await graph_client.me.calendar.get(
            request_configuration=RequestConfiguration(query_parameters=default_calendar)
        )
    except Exception:
        return True
    return _belongs_to_someone_else(calendar, default)


def _match(calendars: list[Any], trimmed: str) -> Any:
    """The one calendar ``trimmed`` names: an exact ID first, then a display name."""
    for cal in calendars:
        if cal.id and cal.id == trimmed:
            return cal

    wanted = _fold(trimmed)
    matches = [cal for cal in calendars if cal.name and _fold(cal.name) == wanted]
    shown = sanitize_output(trimmed[:50])

    if len(matches) == 1:
        if not matches[0].id:
            # None means "the default calendar" to the caller, so a calendar
            # Graph returned without an ID must not fall through as one.
            raise ValueError(f"Calendar '{shown}' was returned with no ID; cannot query it.")
        return matches[0]
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
