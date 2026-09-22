"""Shared recurrence conversion for calendar events and To Do tasks.

Graph models recurrence identically on `event` and `todoTask` — a
`patternedRecurrence` of `{pattern, range}` — so the conversion lives here
rather than in either tool module.

Three layers, deliberately separate:

``build_patterned_recurrence``
    Strict dict → typed SDK model. This is To Do's converter from 1.5.0, moved
    here unchanged; ``outlook_create_task`` still gets exactly its old
    behavior, including rejecting a bare string.

``build_event_recurrence``
    The calendar entry point. Additionally accepts a JSON-encoded object (some
    client bridges stringify nested arguments) and the shorthands the tool
    docstring has always advertised, and reconciles ``range.startDate`` with
    the event's own start before delegating to the strict converter.

``serialize_recurrence``
    Typed SDK model → the same JSON shape, for read paths. What create accepts,
    get returns.
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from typing import Any

# Graph's camelCase day names, Monday-first to match date.weekday().
_WEEKDAYS = (
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday",
)

# The shorthands `outlook_create_event`'s docstring has advertised since 1.0.
# They were accepted and silently dropped until #41; agents in the wild send
# them, so they expand to real patterns rather than becoming a hard error.
_SHORTHANDS = ("daily", "weekdays", "weekly", "monthly", "yearly")

# Graph serializes datetimes with seven fractional digits ("...T09:00:00.0000000").
# datetime.fromisoformat accepts at most six before 3.11, and the project floor is
# 3.10 — so trim to six when reading an event's own start back off the wire.
_OVERLONG_FRACTION = re.compile(r"(\.\d{6})\d+")


def build_patterned_recurrence(recurrence: dict) -> Any:
    """Convert a Graph JSON-shape recurrence dict into a typed PatternedRecurrence.

    Expects the documented Microsoft Graph JSON shape with camelCase keys:
      {"pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"], ...},
       "range":   {"type": "endDate", "startDate": "2026-04-22", "endDate": "2026-12-31", ...}}

    String enums are mapped to the SDK enum members; ISO dates are parsed.
    """
    from msgraph.generated.models.day_of_week import DayOfWeek
    from msgraph.generated.models.patterned_recurrence import PatternedRecurrence
    from msgraph.generated.models.recurrence_pattern import RecurrencePattern
    from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType
    from msgraph.generated.models.recurrence_range import RecurrenceRange
    from msgraph.generated.models.recurrence_range_type import RecurrenceRangeType
    from msgraph.generated.models.week_index import WeekIndex

    if not isinstance(recurrence, dict):
        raise ValueError("recurrence must be a dict with 'pattern' and 'range' keys")

    pattern_in = recurrence.get("pattern") or {}
    range_in = recurrence.get("range") or {}
    if not pattern_in or not range_in:
        raise ValueError("recurrence must include both 'pattern' and 'range'")

    def _enum_lookup(enum_cls: Any, value: str, label: str) -> Any:
        # SDK enum members are PascalCase; Graph JSON uses camelCase
        try:
            return enum_cls(value)
        except ValueError:
            target = value[:1].upper() + value[1:]
            try:
                return enum_cls[target]
            except KeyError as e:
                valid = [m.value for m in enum_cls]
                raise ValueError(f"Invalid {label} '{value}'. Must be one of: {valid}") from e

    pattern = RecurrencePattern()
    if "type" in pattern_in:
        pattern.type = _enum_lookup(RecurrencePatternType, pattern_in["type"], "pattern.type")
    if "interval" in pattern_in:
        pattern.interval = int(pattern_in["interval"])
    if "month" in pattern_in:
        pattern.month = int(pattern_in["month"])
    if "dayOfMonth" in pattern_in:
        pattern.day_of_month = int(pattern_in["dayOfMonth"])
    if "daysOfWeek" in pattern_in:
        pattern.days_of_week = [
            _enum_lookup(DayOfWeek, d, "pattern.daysOfWeek") for d in pattern_in["daysOfWeek"]
        ]
    if "firstDayOfWeek" in pattern_in:
        pattern.first_day_of_week = _enum_lookup(
            DayOfWeek, pattern_in["firstDayOfWeek"], "pattern.firstDayOfWeek"
        )
    if "index" in pattern_in:
        pattern.index = _enum_lookup(WeekIndex, pattern_in["index"], "pattern.index")

    rng = RecurrenceRange()
    if "type" in range_in:
        rng.type = _enum_lookup(RecurrenceRangeType, range_in["type"], "range.type")
    if "startDate" in range_in:
        rng.start_date = date.fromisoformat(range_in["startDate"])
    if "endDate" in range_in:
        rng.end_date = date.fromisoformat(range_in["endDate"])
    if "numberOfOccurrences" in range_in:
        rng.number_of_occurrences = int(range_in["numberOfOccurrences"])
    if "recurrenceTimeZone" in range_in:
        rng.recurrence_time_zone = range_in["recurrenceTimeZone"]

    pr = PatternedRecurrence()
    pr.pattern = pattern
    pr.range = rng
    return pr


def maybe_zone(name: str | None) -> Any:
    """``ZoneInfo(name)`` if it resolves, else ``None``.

    Graph hands back Windows zone names ("Pacific Standard Time") as readily as
    IANA ones, and no mapping between the two ships with Python. An
    unresolvable name means the date below falls back to the text — the
    behaviour before anchoring existed, which is wrong only for a start whose
    offset disagrees with its zone, and never worse than not trying.
    """
    if not name:
        return None
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError, OSError):
        # OSError because ZoneInfo resolves the name against the filesystem: a
        # value that is not a usable path component fails there rather than as
        # a missing key. Every one of the three means the same thing here —
        # no zone, so use the date as written.
        return None


def event_start_date(start: str, zone: str | None = None) -> date:
    """The calendar date the event's first occurrence falls on, in its own zone.

    Graph requires ``range.startDate`` to be the date of the first occurrence,
    and expands a series against the zone the master is anchored in — so the
    date that matters is the one the event has *there*.

    Deliberately *not* UTC-normalized: a ``00:30+02:00`` start is the 7th to
    the person scheduling it, and converting to UTC first would move an
    early-morning or late-evening series a day off.

    ``zone`` resolves the remaining case, and it is not hypothetical. While
    every event was anchored in UTC the text date and the event's own local
    date were the same day by construction; once the anchor is real they can
    differ. ``2026-10-29T01:00:00Z`` anchored in ``America/Los_Angeles`` is
    Wednesday the 28th at 18:00 there — and taking the text date built a
    *Thursday* pattern starting the 29th, which Graph accepted and scheduled a
    full day late, reported as ``status: created``. Verified live 2026-09-21.

    A naive start needs no conversion: it is already wall-clock time in
    ``zone``. Only an offset-bearing or ``Z`` start names an instant whose
    local date has to be worked out.
    """
    text = _OVERLONG_FRACTION.sub(r"\1", start.strip())
    if text.endswith(("Z", "z")):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError as e:
        raise ValueError(f"Invalid event start for recurrence: {start[:50]}") from e

    if parsed.tzinfo is not None:
        tz = maybe_zone(zone)
        if tz is not None:
            # astimezone on an aware datetime, never timedelta arithmetic:
            # PEP 495 resets `fold` on the latter, which is how a refactor that
            # read as cleanup moved a window an hour in 1.21.0.
            return parsed.astimezone(tz).date()
    return parsed.date()


def _expand_shorthand(name: str, start: date) -> dict:
    """Turn "weekly" and friends into a real Graph pattern anchored on the start."""
    key = name.strip().lower()
    if key not in _SHORTHANDS:
        raise ValueError(
            f"Invalid recurrence '{name[:50]}'. Pass a Graph recurrence object "
            f"({{'pattern': ..., 'range': ...}}) or one of: {list(_SHORTHANDS)}"
        )

    if key == "daily":
        pattern: dict[str, Any] = {"type": "daily", "interval": 1}
    elif key == "weekdays":
        pattern = {"type": "weekly", "interval": 1, "daysOfWeek": list(_WEEKDAYS[:5])}
    elif key == "weekly":
        pattern = {"type": "weekly", "interval": 1, "daysOfWeek": [_WEEKDAYS[start.weekday()]]}
    elif key == "monthly":
        pattern = {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": start.day}
    else:  # yearly
        pattern = {
            "type": "absoluteYearly",
            "interval": 1,
            "month": start.month,
            "dayOfMonth": start.day,
        }

    # Shorthands carry no end — the caller who wants one sends a full object.
    return {"pattern": pattern, "range": {"type": "noEnd"}}


def _reconcile_range(payload: dict, start: date) -> dict:
    """Fill in or verify ``range.startDate`` against the event's first occurrence.

    Graph requires the range to begin on the same day as the series master and
    returns ErrorInvalidRecurrenceRange otherwise. Callers usually omit it, so
    default it; when they do send one, a mismatch is a mistake worth naming
    here rather than surfacing as an opaque 400.
    """
    rng = dict(payload.get("range") or {})
    rng.setdefault("type", "noEnd")

    given = rng.get("startDate")
    if given is None:
        rng["startDate"] = start.isoformat()
        return {**payload, "range": rng}

    try:
        parsed = date.fromisoformat(str(given))
    except ValueError as e:
        raise ValueError(
            f"recurrence range.startDate must be YYYY-MM-DD; got {str(given)[:50]!r}"
        ) from e

    if parsed != start:
        raise ValueError(
            f"recurrence range.startDate ({parsed.isoformat()}) must match the event's start "
            f"date ({start.isoformat()}); Graph rejects a series whose range begins on a "
            f"different day than its first occurrence"
        )

    rng["startDate"] = parsed.isoformat()
    return {**payload, "range": rng}


def build_event_recurrence(recurrence: dict | str, *, start: str, zone: str | None = None) -> Any:
    """Build a typed PatternedRecurrence for a calendar event.

    Accepts the Graph recurrence object, a JSON-encoded string of one, or a
    shorthand ("daily", "weekdays", "weekly", "monthly", "yearly") anchored on
    ``start``. ``range.startDate`` is defaulted from ``start`` when omitted.

    ``zone`` is the zone the event is anchored in. It decides which calendar
    date an offset-bearing ``start`` falls on — both for that default and for
    the weekday a shorthand expands to. Omitted, the date comes from the text,
    which is right only while the two agree; see ``event_start_date`` for the
    case where they do not.

    ``range.recurrenceTimeZone`` is passed through rather than stripped, and
    that is a decision, not an oversight. Graph derives the range's zone from
    the event's own when the field is absent, so supplying one can only agree
    (redundant) or disagree — and a disagreement is refused outright with
    ``400 ErrorPropertyValidationFailure`` rather than silently honoured, so it
    is not the silent class of bug. Verified live 2026-09-21, along with the
    reason not to police it here: Graph *accepts* the two vocabularies mixed
    (event ``America/Los_Angeles`` with range ``Pacific Standard Time``, and
    the reverse), and nothing in the standard library maps between them, so a
    string comparison would refuse the read-modify-write round trip that works
    today — ``serialize_recurrence`` hands back whichever spelling Graph
    chose. The opaque Graph error carries a hint naming this cause instead
    (``errors._HINT_TABLE``).
    """
    start_date = event_start_date(start, zone)

    if isinstance(recurrence, str):
        text = recurrence.strip()
        if text.startswith(("{", "[")):
            try:
                parsed = json.loads(text)
            except json.JSONDecodeError as e:
                raise ValueError(f"recurrence is not valid JSON: {e}") from e
            if not isinstance(parsed, dict):
                raise ValueError(
                    "recurrence JSON must be an object with 'pattern' and 'range' keys"
                )
            payload = parsed
        else:
            payload = _expand_shorthand(text, start_date)
    elif isinstance(recurrence, dict):
        payload = dict(recurrence)
    else:
        raise ValueError(
            "recurrence must be a Graph recurrence object, a JSON string of one, "
            f"or one of: {list(_SHORTHANDS)}"
        )

    return build_patterned_recurrence(_reconcile_range(payload, start_date))


# SDK attribute → Graph JSON key, in the order Graph documents them.
_PATTERN_FIELDS = (
    ("type", "type"),
    ("interval", "interval"),
    ("month", "month"),
    ("day_of_month", "dayOfMonth"),
    ("days_of_week", "daysOfWeek"),
    ("first_day_of_week", "firstDayOfWeek"),
    ("index", "index"),
)

_RANGE_ATTRS = (
    ("type", "type"),
    ("start_date", "startDate"),
    ("end_date", "endDate"),
    ("number_of_occurrences", "numberOfOccurrences"),
    ("recurrence_time_zone", "recurrenceTimeZone"),
)


def _plain(value: Any) -> Any:
    """Unwrap SDK enums and dates into JSON-safe primitives."""
    if isinstance(value, list):
        return [_plain(v) for v in value]
    if hasattr(value, "value"):  # SDK enum member
        return value.value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def _section(obj: Any, fields: tuple[tuple[str, str], ...]) -> dict:
    out: dict[str, Any] = {}
    for attr, key in fields:
        value = getattr(obj, attr, None)
        if value is not None:
            out[key] = _plain(value)
    return out


def serialize_recurrence(recurrence: Any) -> dict | None:
    """Convert a typed PatternedRecurrence back to the documented JSON shape.

    Read paths previously did ``str(event.recurrence)``, which leaked a
    multi-hundred-character Python repr into the response. Unset fields are
    omitted so a recurring event doesn't pad every detail payload with nulls.
    """
    if recurrence is None:
        return None

    out: dict[str, Any] = {}
    pattern = _section(getattr(recurrence, "pattern", None), _PATTERN_FIELDS)
    if pattern:
        out["pattern"] = pattern
    rng = _section(getattr(recurrence, "range", None), _RANGE_ATTRS)
    if rng:
        out["range"] = rng
    return out or None
