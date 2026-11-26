"""``outlook_agenda``: your days in a window (contracts §12.4).

Meetings, overlaps, free blocks, invitations and tasks due, from the owner's own
calendars (at most five, the default first) and To Do lists. Free blocks are the
gaps between ``day_start`` and ``day_end`` that no busy event covers: they say
nothing about anyone else's time.

Sections, in order: ``days`` (one item per local date), ``overlaps``, ``invites``
and ``tasks_due``. agenda reads no sent mail, so an invitation's organizer has no
``known`` flag here; ``catch_up`` and ``brief`` carry it.

Requests (§12.4): the calendar list (A1), a calendar view per calendar of at most
three pages (A2), and the task lists with the open tasks of at most eight (A3):
25 at most, under a ceiling of 30.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone, tzinfo
from typing import Any
from urllib.parse import quote

from outlook_mcp.calendar_resolver import fetch_all_calendars
from outlook_mcp.intents import budget
from outlook_mcp.intents.budget import RequestBudget, run_sections
from outlook_mcp.intents.envelope import (
    SIZE_CAP,
    build_envelope,
    build_next,
    has_blocking_warning,
    others_wrote,
    section,
    size_of,
    suggestion,
    trim_to_cap,
    utc,
    why,
)
from outlook_mcp.intents.reading import (
    event_item,
    event_span,
    failed_like,
    is_busy,
    is_open_invite,
    local_dates,
    local_iso,
    local_midnight,
    open_tasks,
    outcome_of,
    read_view,
    response_of,
    show_as_of,
    task_item,
    task_order,
    window_echo,
    zone_info,
)
from outlook_mcp.intents.resolver import check_words, open_lookup
from outlook_mcp.validation import CONFIG_REMEDY_TEXT, resolve_window, validate_clock

TOOL = "outlook_agenda"
MAX_DAYS = 31
CALENDARS_MAX = 5
EVENT_PAGES = 3
EVENT_TOP = 100
TASK_LISTS = 8
DAY_EVENTS = 30
OVERLAP_PARTNERS = 5
MIN_FREE = (5, 480)
CAPS = {"days": MAX_DAYS, "overlaps": 20, "invites": 20, "tasks_due": 30}
# A2's `$select` (§12.4).
EVENT_SELECT = (
    "id,iCalUId,subject,start,end,isAllDay,isOrganizer,organizer,responseStatus,"
    "isCancelled,showAs,location,seriesMasterId,type"
)
# The view is read a day wider on each side: an all-day event is a floating date,
# which Graph places by UTC midnight here, so one on the window's first or last
# local day could fall just outside a view of the exact window. What is shown is
# decided locally.
_VIEW_MARGIN = timedelta(days=1)
_ZONE_REMEDY = "Pass an IANA name such as Europe/Paris, or leave zone out."


def _clock(value: str, name: str, *, end_of_day: bool = False) -> int:
    try:
        return validate_clock(value, end_of_day=end_of_day)
    except ValueError as exc:
        raise ValueError(f"{name}: {exc}") from None


def _check_calendars(calendars: list[str] | None) -> list[str] | None:
    if calendars is None:
        return None
    if not isinstance(calendars, list) or not 1 <= len(calendars) <= CALENDARS_MAX:
        raise ValueError(f"calendars takes 1-{CALENDARS_MAX} calendar names or ids.")
    return [check_words(c, "a calendar name").strip() for c in calendars]


def _choose_calendars(found: list[Any], wanted: list[str] | None) -> tuple[list[Any], int]:
    """The calendars to read, the default first, and how many more were left unread."""
    ordered = sorted(found, key=lambda c: 0 if getattr(c, "is_default_calendar", False) else 1)
    if wanted is None:
        return ordered[:CALENDARS_MAX], max(0, len(ordered) - CALENDARS_MAX)
    chosen: list[Any] = []
    for ref in wanted:
        by_id = [c for c in ordered if c.id == ref]
        by_name = [c for c in ordered if (c.name or "").casefold() == ref.casefold()]
        matches = by_id or by_name
        if not matches:
            raise ValueError(
                f"No calendar called '{ref[:60]}'. Pass a calendar's name or id, as "
                "outlook_list_calendars shows them."
            )
        if len(matches) > 1:
            raise ValueError(
                f"Several calendars are called '{ref[:60]}'. Pass the id of the one you mean."
            )
        if matches[0] not in chosen:
            chosen.append(matches[0])
    chosen.sort(key=lambda c: 0 if getattr(c, "is_default_calendar", False) else 1)
    return chosen, 0


def _intervals_union(spans: list[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    merged: list[tuple[datetime, datetime]] = []
    for start, end in sorted(spans):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _clip(
    spans: list[tuple[datetime, datetime]], low: datetime, high: datetime
) -> list[tuple[datetime, datetime]]:
    return [(max(s, low), min(e, high)) for s, e in spans if s < high and e > low]


def _minutes(span: tuple[datetime, datetime]) -> int:
    return int((span[1] - span[0]).total_seconds() // 60)


def _wall(day: date, minutes: int, tz: tzinfo) -> datetime:
    """``minutes`` after local midnight of ``day`` in wall-clock terms, as UTC."""
    if minutes >= 1440:
        return local_midnight(day + timedelta(days=1), tz)
    if minutes == 0:
        return local_midnight(day, tz)
    wall = datetime.combine(day, datetime.min.time()) + timedelta(minutes=minutes)
    return wall.replace(tzinfo=tz).astimezone(timezone.utc)


def _concise(row: dict, calendar: int, tz: tzinfo, now: datetime) -> dict:
    """A day's event: when, which calendar (its place in ``understood.calendars``), and
    how it stands, without the organizer."""
    start, end = event_span(row, tz)
    full = event_item(row, calendar=calendar, tz=tz, zone_name="", now=now)
    return {
        "kind": "event",
        "id": row.get("id"),
        "start": local_iso(start, tz) if start else None,
        "end": local_iso(end, tz) if end else None,
        "all_day": row.get("isAllDay") is True,
        "calendar": calendar,
        "show_as": show_as_of(row),
        "response": response_of(row),
        "is_organizer": row.get("isOrganizer") is True,
        "is_cancelled": row.get("isCancelled") is True,
        "why": full["why"],
        "warnings": full["warnings"],
        "others_wrote": others_wrote(
            subject=full["others_wrote"].get("subject"),
            location=full["others_wrote"].get("location"),
        ),
    }


async def read_agenda(
    graph_client: Any,
    *,
    window: str = "today",
    zone: str | None = None,
    day_start: str = "09:00",
    day_end: str = "18:00",
    calendars: list[str] | None = None,
    min_free: int = 30,
    include_tasks: bool = True,
    config_zone: str = "UTC",
    signed_in: str | None = None,
    loaded_tools: frozenset[str] = frozenset(),
    now: datetime | None = None,
) -> dict:
    """The agenda for ``window`` (§12.4), in the envelope.

    ``zone`` (IANA) both reads the window and shows the times; ``config_zone`` is
    used when it is not given. ``loaded_tools`` is the tool set this process loads,
    the only tools ``next`` may name. ``now`` is fixed by tests.
    """
    moment = now or datetime.now(timezone.utc)

    # Every argument is checked before any request goes out.
    zone_name = zone.strip() if isinstance(zone, str) and zone.strip() else config_zone
    tz = zone_info(zone_name, remedy=_ZONE_REMEDY if zone else CONFIG_REMEDY_TEXT)
    start, end = resolve_window(
        window,
        zone_name,
        now=moment,
        max_days=MAX_DAYS,
        zone_remedy=_ZONE_REMEDY if zone else CONFIG_REMEDY_TEXT,
    )
    first_minute = _clock(day_start, "day_start")
    last_minute = _clock(day_end, "day_end", end_of_day=True)
    if last_minute <= first_minute:
        raise ValueError("day_end must be later than day_start.")
    if (
        isinstance(min_free, bool)
        or not isinstance(min_free, int)
        or not MIN_FREE[0] <= min_free <= MIN_FREE[1]
    ):
        raise ValueError(
            f"min_free is minutes, a whole number from {MIN_FREE[0]} to {MIN_FREE[1]}."
        )
    if not isinstance(include_tasks, bool):
        raise ValueError("include_tasks is true or false.")
    wanted = _check_calendars(calendars)
    days = local_dates(start, end, tz)
    today = moment.astimezone(tz).date()

    spent = RequestBudget(budget.CEILINGS[TOOL])
    chosen_calendars: list[Any] = []
    not_covered: list[str] = []
    beyond = 0
    # The start of the last event read on a calendar whose read was cut at its cap:
    # past it, nothing is known, so nothing is called free (review A1).
    read_until: datetime | None = None

    async with open_lookup(
        graph_client, spent, now=moment, zone=zone_name, signed_in=signed_in
    ) as lookup:

        async def read_calendars() -> list[tuple[dict, str]]:
            nonlocal beyond, read_until
            found = await fetch_all_calendars(lookup.sdk)
            chosen, beyond = _choose_calendars(found, wanted)
            chosen_calendars.extend(chosen)
            if beyond:
                not_covered.append("calendars_beyond_5")
            seen: dict[str, tuple[dict, str]] = {}
            for calendar in chosen:
                rows, more = await read_view(
                    lookup,
                    f"/me/calendars/{quote(calendar.id, safe='')}/calendarView",
                    start - _VIEW_MARGIN,
                    end + _VIEW_MARGIN,
                    EVENT_SELECT,
                    pages=EVENT_PAGES,
                    top=EVENT_TOP,
                )
                if more:
                    if "events_beyond_300" not in not_covered:
                        not_covered.append("events_beyond_300")
                    starts = [s for s in (event_span(r, tz)[0] for r in rows) if s]
                    if starts:
                        last = max(starts)
                        read_until = last if read_until is None else min(read_until, last)
                for row in rows:
                    if not isinstance(row.get("id"), str):
                        continue
                    # The same meeting on two calendars is kept once, from the first
                    # calendar read: the default one, when it has it.
                    key = row.get("iCalUId") if isinstance(row.get("iCalUId"), str) else row["id"]
                    seen.setdefault(key, (row, calendar.id))
            return [
                (row, cal)
                for row, cal in seen.values()
                if _overlaps(event_span(row, tz), start, end)
            ]

        async def read_tasks() -> list[tuple[dict, str]]:
            pairs, missed = await open_tasks(lookup, lists_max=TASK_LISTS)
            not_covered.extend(missed)
            return pairs

        runs = [("days", read_calendars)]
        if include_tasks:
            runs.append(("tasks_due", read_tasks))
        results = {r.name: r for r in await run_sections(runs)}

    sections: list[dict] = []
    failed: list[dict] = []
    suggestions: list[dict] = []
    calendar_result = results["days"]
    if calendar_result.status == "ok":
        events = sorted(
            calendar_result.value,
            key=lambda pair: (event_span(pair[0], tz)[0] or start, str(pair[0].get("id"))),
        )
        place = {c.id: n for n, c in enumerate(chosen_calendars)}
        day_items = [
            _day_item(
                d,
                events,
                tz,
                moment,
                start,
                end,
                first_minute,
                last_minute,
                min_free,
                place=place,
                read_until=read_until,
            )
            for d in days
        ]
        overlap_items = _overlap_items(events, tz, zone_name, moment)
        invites = [
            event_item(row, calendar=cal, tz=tz, zone_name=zone_name, now=moment)
            for row, cal in events
            if is_open_invite(row) and _starts_in(row, tz, start, end)
        ]
        exact = read_until is None
        sections += [
            section("days", day_items, cap=CAPS["days"], total_exact=exact),
            section("overlaps", overlap_items, cap=CAPS["overlaps"], total_exact=exact),
            section("invites", invites, cap=CAPS["invites"], total_exact=exact),
        ]
        suggestions += _suggestions(invites, events, tz, moment)
    else:
        for name in ("days", "overlaps", "invites"):
            sections.append(section(name, [], cap=CAPS[name], status=calendar_result.status))
            failed.append(failed_like(name, calendar_result.failure))

    if include_tasks:
        task_result = results["tasks_due"]
        if task_result.status == "ok":
            due_items = _tasks_due(task_result.value, days, today)
            tasks_exact = not any(
                c.startswith(("task_lists_beyond", "tasks_beyond")) for c in not_covered
            )
            sections.append(
                section("tasks_due", due_items, cap=CAPS["tasks_due"], total_exact=tasks_exact)
            )
        else:
            sections.append(
                section("tasks_due", [], cap=CAPS["tasks_due"], status=task_result.status)
            )
            failed.append(task_result.failure)

    understood: dict = {
        "window": window_echo(window, start, end, zone_name),
        "zone": zone_name,
        "day_start": day_start,
        "day_end": day_end,
        "min_free": min_free,
        "include_tasks": include_tasks,
        "calendars": [
            {
                "id": c.id,
                "is_default": bool(getattr(c, "is_default_calendar", False)),
                # A shared calendar is named by its owner.
                "others_wrote": others_wrote(name=c.name),
            }
            for c in chosen_calendars
        ],
        "not_covered": not_covered,
    }
    if beyond:
        understood["calendars_beyond_5"] = beyond
    outcome = outcome_of(sections)
    if read_until is not None:
        # A calendar's read was cut: the days past it are not known (review A1).
        understood["events_read_until"] = utc(read_until)
        outcome = "incomplete"
    env = build_envelope(
        tool=TOOL,
        outcome=outcome,
        as_of=moment,
        understood=understood,
        sections=sections,
        failed_sections=failed,
        graph_requests=spent.spent,
        next=build_next(suggestions, loaded_tools),
        capped=bool(not_covered),
        trim=False,
    )
    return _fit(env)


def _fit(env: dict) -> dict:
    """The envelope within the size cap, keeping every day (review A3).

    Before the envelope's own trim (which takes items from the last section first),
    the days give up what they can spare: first their events' locations, then their
    events, from the last day back. Each day keeps its date, free blocks, busy minutes
    and ``events_total``, so a long window still answers when you are free.
    """
    days = next((s for s in env["sections"] if s["name"] == "days"), None)
    reduced = False
    if days is not None and size_of(env) > SIZE_CAP:
        for day in days["items"]:
            for shown in day["events"]:
                if shown["others_wrote"].pop("location", None) is not None:
                    reduced = True
        for day in reversed(days["items"]):
            while day["events"] and size_of(env) > SIZE_CAP:
                day["events"].pop()
                reduced = True
            if size_of(env) <= SIZE_CAP:
                break
    env = trim_to_cap(env)
    if reduced:
        env["trimmed"] = True
        env["partial"] = True
    return env


def _overlaps(span: tuple[datetime | None, datetime | None], low: datetime, high: datetime) -> bool:
    start, end = span
    return start is not None and end is not None and start < high and end > low


def _starts_in(row: dict, tz: tzinfo, low: datetime, high: datetime) -> bool:
    start = event_span(row, tz)[0]
    return start is not None and low <= start < high


def _day_item(
    day: date,
    events: list[tuple[dict, str]],
    tz: tzinfo,
    now: datetime,
    window_start: datetime,
    window_end: datetime,
    first_minute: int,
    last_minute: int,
    min_free: int,
    *,
    place: dict[str, int],
    read_until: datetime | None,
) -> dict:
    """One local date: its events, the free blocks between day_start and day_end, and
    the minutes busy events take inside those bounds. No time at or after
    ``read_until`` is called free: past a capped read, nothing is known."""
    midnight, next_midnight = local_midnight(day, tz), local_midnight(day + timedelta(days=1), tz)
    on_day = [
        (row, cal) for row, cal in events if _overlaps(event_span(row, tz), midnight, next_midnight)
    ]
    low = max(_wall(day, first_minute, tz), window_start)
    high = min(_wall(day, last_minute, tz), window_end)
    busy = _intervals_union(
        [
            span  # type: ignore[misc]
            for span in (event_span(row, tz) for row, _ in on_day if is_busy(row))
            if span[0] is not None and span[1] is not None
        ]
    )
    busy_in_bounds = _clip(busy, low, high) if low < high else []
    # Free time never starts before now: today's begins at now, a past day has none.
    free_from = max(low, now)
    if read_until is not None:
        high = min(high, read_until)
    free: list[dict] = []
    cursor = free_from
    for span_start, span_end in _clip(busy, free_from, high) if free_from < high else []:
        if span_start > cursor:
            free.append((cursor, span_start))  # type: ignore[arg-type]
        cursor = max(cursor, span_end)
    if free_from < high and cursor < high:
        free.append((cursor, high))  # type: ignore[arg-type]
    blocks = [
        {"start": local_iso(s, tz), "end": local_iso(e, tz), "minutes": _minutes((s, e))}
        for s, e in free  # type: ignore[misc]
        if _minutes((s, e)) >= min_free
    ]
    shown = [_concise(row, place.get(cal, 0), tz, now) for row, cal in on_day]
    return {
        "kind": "day",
        "id": day.isoformat(),
        "date": day.isoformat(),
        "events": shown[:DAY_EVENTS],
        "events_total": len(shown),
        "free": blocks,
        "busy_minutes": sum(_minutes(s) for s in busy_in_bounds),
        "why": [why("today")] if day == now.astimezone(tz).date() else [],
        "warnings": [],
        "others_wrote": {},
    }


def _overlap_items(
    events: list[tuple[dict, str]], tz: tzinfo, zone_name: str, now: datetime
) -> list[dict]:
    """Busy events that overlap another busy event, each with whom it overlaps."""
    busy = []
    for row, cal in events:
        span = event_span(row, tz)
        if is_busy(row) and span[0] is not None and span[1] is not None:
            busy.append((row, cal, span))
    partners: dict[str, list[str]] = {}
    for n, (a, _, (a_start, a_end)) in enumerate(busy):
        for b, _, (b_start, b_end) in busy[n + 1 :]:
            if a_start < b_end and b_start < a_end:
                partners.setdefault(a["id"], []).append(b["id"])
                partners.setdefault(b["id"], []).append(a["id"])
    items = []
    for row, cal, _ in busy:
        if row["id"] not in partners:
            continue
        item = event_item(row, calendar=cal, tz=tz, zone_name=zone_name, now=now)
        # At most OVERLAP_PARTNERS reasons: a day of back-to-back meetings would otherwise
        # repeat every other id on every item.
        found = partners[row["id"]]
        item["why"] = [
            *item["why"],
            *(why("overlaps", **{"with": p}) for p in found[:OVERLAP_PARTNERS]),
        ]
        items.append(item)
    return items


def _tasks_due(pairs: list[tuple[dict, str]], days: list[date], today: date) -> list[dict]:
    """Open tasks due on one of the window's local dates, and overdue ones when the
    window includes today (§12.4)."""
    from outlook_mcp.intents.reading import due_date

    first, last = days[0], days[-1]
    items = []
    for row, list_id in pairs:
        due = due_date(row)
        if due is None:
            continue
        in_window = first <= due <= last
        overdue = due < today and first <= today <= last
        if in_window or overdue:
            items.append(task_item(row, list_id, today=today, in_window=in_window))
    return sorted(items, key=task_order)


def _suggestions(
    invites: list[dict], events: list[tuple[dict, str]], tz: tzinfo, now: datetime
) -> list[dict]:
    """``answer`` for the first invite; ``brief_meeting`` for the next meeting."""
    offers = []
    unblocked = [i for i in invites if not has_blocking_warning(i)]
    if unblocked:
        offers.append(
            suggestion(
                "answer", "outlook_answer_invites", {"event_ids": [unblocked[0]["id"]]}, ["answer"]
            )
        )
    for row, _ in events:
        start = event_span(row, tz)[0]
        if (
            start is not None
            and start >= now
            and row.get("isAllDay") is not True
            and row.get("isCancelled") is not True
            # The owner's own, or one they accepted: agenda reads no sent mail, so it
            # can't tell a stranger's invitation, whose subject and attendees a brief
            # would then search for (review A4).
            and (row.get("isOrganizer") is True or response_of(row) == "accepted")
        ):
            offers.append(suggestion("brief_meeting", "outlook_brief", {"meeting": row["id"]}))
            break
    return offers
