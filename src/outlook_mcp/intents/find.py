"""``outlook_find``: words or an exact reference resolved to items, with reasons (contracts §12.1).

The resolver (``resolver.resolve``) behind a tool's arguments. ``find_items``
checks every argument before any request goes out, refuses one that does not
apply to the kind (nothing is silently ignored), and answers in the envelope
with one section, ``matches``. Its outcome is ``exact``, ``single``,
``candidates`` or ``nothing_found``: a name that matches nothing is an answer
here, not an error.

Registered by the server as ``outlook_find`` in the opt-in ``intents`` group.

``incomplete`` is a fifth outcome, beyond the contracts' four (§1.2): the search
did not finish, or a list it read was cut at its cap, and it matched nothing in
what it did read. ``understood.not_searched`` says what was left out.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from outlook_mcp.intents import budget
from outlook_mcp.intents.budget import RequestBudget, run_sections
from outlook_mcp.intents.envelope import (
    build_envelope,
    build_next,
    has_blocking_warning,
    section,
    suggestion,
    utc,
)
from outlook_mcp.intents.resolver import (
    KINDS,
    LIMIT_MAX,
    Filters,
    check_words,
    open_lookup,
    resolve,
)
from outlook_mcp.validation import (
    CONFIG_REMEDY_TEXT,
    resolve_timezone,
    resolve_window,
    validate_datetime,
)

TOOL = "outlook_find"

# Which kinds take each optional argument (§12.1). Anything else is refused.
APPLIES_TO = {
    "sender": ("message", "thread", "file"),
    "after": ("message", "thread", "file", "draft"),
    "before": ("message", "thread", "file", "draft"),
    "window": ("event", "task"),
    "has_file": ("message", "thread"),
    "folder": ("message", "thread", "file"),
    "task_list": ("task",),
}
# The longest event window find reads (§3.2).
EVENT_WINDOW_DAYS = 400
# Suggested next calls, by the kind of a match (§12.1): for at most this many.
_NEXT_FOR = 3
_NEXT = {
    "message": ("open", "outlook_read_message", "message_id", "id"),
    "thread": ("open_thread", "outlook_list_thread", "conversation_id", "conversation_id"),
    "person": ("brief_person", "outlook_brief", "person", "address"),
}


def _instant(value: str, name: str, zone: str, now: datetime) -> datetime:
    try:
        moment = validate_datetime(value, zone, now=now)
    except ValueError as exc:
        raise ValueError(f"{name}: {exc}") from None
    return datetime.strptime(moment, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)


def _window_echo(value: str | None, start: datetime, end: datetime, zone: str) -> dict:
    """``understood.window`` (§3.2): the value as given, and both ends in UTC and local."""
    tz = resolve_timezone(zone, remedy=CONFIG_REMEDY_TEXT)
    return {
        "value": value,
        "start": utc(start),
        "end": utc(end),
        "zone": zone,
        "start_local": start.astimezone(tz).isoformat(),
        "end_local": end.astimezone(tz).isoformat(),
    }


def _suggestions(items: list[dict]) -> list[dict]:
    """``open``, ``open_thread`` or ``brief_person`` for up to three unblocked matches."""
    offers = []
    for item in items:
        if len(offers) == _NEXT_FOR:
            break
        spec = _NEXT.get(item["kind"])
        if spec is None or has_blocking_warning(item) or not item.get(spec[3]):
            continue
        template, tool, argument, field = spec
        offers.append(suggestion(template, tool, {argument: item[field]}))
    return offers


async def find_items(
    graph_client: Any,
    *,
    kind: str,
    words: str,
    sender: str | None = None,
    after: str | None = None,
    before: str | None = None,
    window: str | None = None,
    has_file: bool | None = None,
    folder: str | None = None,
    task_list: str | None = None,
    limit: int = 5,
    zone: str = "UTC",
    signed_in: str | None = None,
    loaded_tools: frozenset[str] = frozenset(),
    now: datetime | None = None,
) -> dict:
    """Resolve ``words`` to ``kind`` items (§12.1) and answer in the envelope.

    ``zone`` (the config's time zone) reads zone-less times and windows;
    ``signed_in`` is the signed-in username, one of the owner's own addresses;
    ``loaded_tools`` is the tool set this process loads, the only tools ``next``
    may name. ``now`` is fixed by tests.
    """
    moment = now or datetime.now(timezone.utc)

    # Every argument is checked before any request goes out.
    if kind not in KINDS:
        raise ValueError(f"kind must be one of: {', '.join(KINDS)}.")
    words = check_words(words)
    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= LIMIT_MAX:
        raise ValueError(f"limit must be a whole number from 1 to {LIMIT_MAX}.")
    resolve_timezone(zone, remedy=CONFIG_REMEDY_TEXT)
    given = {
        "sender": sender,
        "after": after,
        "before": before,
        "window": window,
        "has_file": has_file,
        "folder": folder,
        "task_list": task_list,
    }
    for name, value in given.items():
        if value is not None and kind not in APPLIES_TO[name]:
            raise ValueError(
                f"`{name}` applies to kind {', '.join(APPLIES_TO[name])}; this call's kind "
                f"is '{kind}'. Leave it out."
            )
    if sender is not None:
        sender = check_words(sender, "sender")
    for name, value in (("folder", folder), ("task_list", task_list)):
        if value is not None:
            check_words(value, name)
    start = _instant(after, "after", zone, moment) if after is not None else None
    end = _instant(before, "before", zone, moment) if before is not None else None
    if start is not None and end is not None and end <= start:
        raise ValueError("`before` must be later than `after`.")
    span = None
    if window is not None:
        span = resolve_window(
            window,
            zone,
            now=moment,
            max_days=EVENT_WINDOW_DAYS if kind == "event" else None,
        )
    filters = Filters(
        sender=sender,
        after=start,
        before=end,
        window=span,
        has_file=has_file,
        folder=folder.strip() if folder else None,
        task_list=task_list.strip() if task_list else None,
    )

    spent = RequestBudget(budget.CEILINGS[TOOL])
    async with open_lookup(
        graph_client, spent, now=moment, zone=zone, signed_in=signed_in
    ) as lookup:

        async def matches():
            return await resolve(lookup, kind, words, filters, "read", limit)

        (result,) = await run_sections([("matches", matches)])
        queries = list(lookup.graph.queries)

    capped = False
    if result.status == "ok":
        resolution = result.value
        items = [c.item for c in resolution.candidates]
        capped = bool(resolution.not_searched)
        matched = section(
            "matches",
            items,
            cap=limit,
            total=resolution.total,
            # A capped search may have missed matches: its count is a floor (review S4).
            total_exact=not capped,
            total_display=f"{resolution.total}+" if capped else None,
        )
        outcome = resolution.outcome
        understood = dict(resolution.understood)
        failed: list[dict] = []
    else:
        # The search did not finish (budget, deadline, a refused feature): that is not
        # "nothing matched", so it is never answered as `nothing_found` (review S4).
        items = []
        matched = section("matches", [], cap=limit, status=result.status)
        outcome = "incomplete"
        understood = {"kind": kind, "words": words, "not_searched": [result.failure["message"]]}
        failed = [result.failure]

    understood["queries"] = queries
    understood["limit"] = limit
    understood["zone"] = zone
    if kind == "event" or span is not None:
        notes = understood.get("window") or {}
        if span is None and notes:
            low = datetime.strptime(notes["start"], "%Y-%m-%dT%H:%M:%SZ")
            high = datetime.strptime(notes["end"], "%Y-%m-%dT%H:%M:%SZ")
            span = (low.replace(tzinfo=timezone.utc), high.replace(tzinfo=timezone.utc))
        if span is not None:
            understood["window"] = _window_echo(window, span[0], span[1], zone)

    return build_envelope(
        tool=TOOL,
        outcome=outcome,
        as_of=moment,
        understood=understood,
        sections=[matched],
        failed_sections=failed,
        graph_requests=spent.spent,
        next=build_next(_suggestions(items), loaded_tools),
        capped=capped,
    )
