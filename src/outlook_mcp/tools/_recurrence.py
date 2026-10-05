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


def _object(value: Any, label: str) -> dict:
    """``value`` as a JSON object, or a refusal naming ``label``.

    ``None`` reads as absent. Anything else that is not a dict is refused rather
    than probed: ``"type" in "daily"`` is a substring test, so a string where an
    object belongs would otherwise build an empty pattern without complaint.
    """
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError(f"recurrence {label} must be an object; got {type(value).__name__}")
    return value


def _integer(value: Any, label: str) -> int:
    """``value`` as a whole number, or a refusal naming ``label``.

    A digit string is accepted, as it always was. A bool or a fractional float
    is refused rather than coerced: ``int(True)`` is 1 and ``int(2.5)`` is 2, and
    silently reinterpreting what the caller asked for is worse than saying so.
    """
    refusal = ValueError(f"recurrence {label} must be a whole number; got {value!r:.50}")
    if isinstance(value, bool) or (isinstance(value, float) and not value.is_integer()):
        raise refusal
    try:
        return int(value)
    except (TypeError, ValueError) as e:
        raise refusal from e


def _iso_date(value: Any, label: str) -> date:
    if not isinstance(value, str):
        raise ValueError(f"recurrence {label} must be a YYYY-MM-DD string; got {value!r:.50}")
    try:
        return date.fromisoformat(value)
    except ValueError as e:
        raise ValueError(f"recurrence {label} must be YYYY-MM-DD; got {value[:50]!r}") from e


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

    pattern_in = _object(recurrence.get("pattern"), "pattern")
    range_in = _object(recurrence.get("range"), "range")
    if not pattern_in or not range_in:
        raise ValueError("recurrence must include both 'pattern' and 'range'")

    def _enum_lookup(enum_cls: Any, value: str, label: str) -> Any:
        valid = [m.value for m in enum_cls]
        if not isinstance(value, str):
            raise ValueError(f"Invalid {label} {value!r:.50}. Must be one of: {valid}")
        # SDK enum members are PascalCase; Graph JSON uses camelCase
        try:
            return enum_cls(value)
        except ValueError:
            target = value[:1].upper() + value[1:]
            try:
                return enum_cls[target]
            except KeyError as e:
                raise ValueError(f"Invalid {label} '{value}'. Must be one of: {valid}") from e

    pattern = RecurrencePattern()
    if "type" in pattern_in:
        pattern.type = _enum_lookup(RecurrencePatternType, pattern_in["type"], "pattern.type")
    if "interval" in pattern_in:
        pattern.interval = _integer(pattern_in["interval"], "pattern.interval")
    if "month" in pattern_in:
        pattern.month = _integer(pattern_in["month"], "pattern.month")
    if "dayOfMonth" in pattern_in:
        pattern.day_of_month = _integer(pattern_in["dayOfMonth"], "pattern.dayOfMonth")
    if "daysOfWeek" in pattern_in:
        if not isinstance(pattern_in["daysOfWeek"], list):
            raise ValueError(
                'recurrence pattern.daysOfWeek must be a list of day names, e.g. ["monday"]'
            )
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
        rng.start_date = _iso_date(range_in["startDate"], "range.startDate")
    if "endDate" in range_in:
        rng.end_date = _iso_date(range_in["endDate"], "range.endDate")
    if "numberOfOccurrences" in range_in:
        rng.number_of_occurrences = _integer(
            range_in["numberOfOccurrences"], "range.numberOfOccurrences"
        )
    if "recurrenceTimeZone" in range_in:
        if not isinstance(range_in["recurrenceTimeZone"], str):
            raise ValueError("recurrence range.recurrenceTimeZone must be a zone name string")
        rng.recurrence_time_zone = range_in["recurrenceTimeZone"]

    pr = PatternedRecurrence()
    pr.pattern = pattern
    pr.range = rng
    return pr


def maybe_zone(name: str | None) -> Any:
    """The ``ZoneInfo`` for ``name`` if this host has one, else ``None``.

    Delegates to ``validation.resolve_timezone`` rather than calling
    ``ZoneInfo`` again, so there is one place that knows how zone resolution
    fails — which platform-specific exception each failure raises, and that
    membership in ``available_timezones()`` is the portable test. This copy
    caught ``OSError`` when the shared one did not, and the divergence was the
    bug: an over-long name reached the model as a message-free crash through
    the path that had *not* been updated.

    The difference that remains is deliberate. ``resolve_timezone`` raises a
    ``ValueError`` the model can act on, because a bad zone there is the
    caller's mistake. Here a name that will not resolve is usually Graph's own
    Windows spelling, which is nobody's mistake, and the answer is to fall back
    to the date as written.
    """
    if not name:
        return None
    from outlook_mcp.validation import resolve_timezone

    try:
        return resolve_timezone(name)
    except ValueError:
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
    rng = dict(_object(payload.get("range"), "range"))
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


def event_recurrence_payload(recurrence: dict | str, start_date: date) -> dict:
    """Any accepted recurrence shape, as the Graph JSON object it stands for.

    A dict is copied, a JSON string decoded, and a shorthand expanded against
    ``start_date``. Nothing is reconciled yet — that is ``build_event_recurrence``'s
    job — so a caller can adjust the payload in between.
    """
    if isinstance(recurrence, dict):
        return dict(recurrence)
    if not isinstance(recurrence, str):
        raise ValueError(
            "recurrence must be a Graph recurrence object, a JSON string of one, "
            f"or one of: {list(_SHORTHANDS)}"
        )
    text = recurrence.strip()
    if not text.startswith(("{", "[")):
        return _expand_shorthand(text, start_date)
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"recurrence is not valid JSON: {e}") from e
    if not isinstance(parsed, dict):
        raise ValueError("recurrence JSON must be an object with 'pattern' and 'range' keys")
    return parsed


def check_recurrence_shape(recurrence: dict | str) -> None:
    """Refuse a malformed recurrence without needing the event's start.

    Everything except ``range.startDate`` can be checked before a single read:
    the JSON, the shorthand name, both sections present, every enum and date.
    ``startDate`` is left to ``build_event_recurrence``, which is the only place
    that knows the date it has to match. A shorthand is expanded against an
    arbitrary date here, since only the shape is kept.
    """
    payload = event_recurrence_payload(recurrence, date(2000, 1, 3))
    rng = {k: v for k, v in _object(payload.get("range"), "range").items() if k != "startDate"}
    rng.setdefault("type", "noEnd")
    build_patterned_recurrence({**payload, "range": rng})


def _pattern_type(pattern: dict) -> str:
    """The pattern's type in Graph's own spelling, however the caller cased it.

    The converter accepts both `weekly` and `Weekly` (the SDK's member name), so
    anything that branches on the type has to see one spelling, or a PascalCase
    pattern skips every type-specific comparison and move.
    """
    from msgraph.generated.models.recurrence_pattern_type import RecurrencePatternType

    raw = str(pattern.get("type") or "")
    for member in RecurrencePatternType:
        if member.value.lower() == raw.lower():
            return member.value
    return raw


def _whole(pattern: dict, field: str) -> int | None:
    """``pattern[field]`` read the way the converter reads it, or ``None`` if absent.

    The converter accepts ``"15"`` for 15, so a comparison or a move that took the
    raw value would call an echoed ``"15"`` a different day from the stored 15.
    """
    value = pattern.get(field)
    return None if value is None else _integer(value, f"pattern.{field}")


def _week_boundary(pattern: dict) -> str | None:
    """The day a weekly pattern's weeks start on, or ``None`` where it decides nothing.

    Only a multi-week interval reads the week boundary: it decides which days
    share a week, so moving it reshapes a fortnightly Sunday-and-Monday series.
    Every week, it changes no occurrence, only the week start Outlook displays.
    A pattern that omits it gets Graph's default, Sunday — which only a
    hand-written pattern does, since Graph's read-back always carries it.
    """
    if (_whole(pattern, "interval") or 1) <= 1:
        return None
    return str(pattern.get("firstDayOfWeek") or "sunday").lower()


def _pattern_key(pattern: Any) -> tuple | None:
    """The fields that decide which days a pattern lands on, for comparison.

    Graph fills in defaults on read — a weekly pattern comes back carrying
    ``month: 0``, ``dayOfMonth: 0`` and ``index: "first"`` — so comparing whole
    dicts would call a hand-written pattern different from the stored one it
    describes. Only the fields the pattern's type actually reads take part.
    """
    if not isinstance(pattern, dict) or not pattern.get("type"):
        return None
    kind = _pattern_type(pattern)
    key: tuple = (kind, _whole(pattern, "interval") or 1)
    if kind in ("weekly", "relativeMonthly", "relativeYearly"):
        key += (frozenset(str(d).lower() for d in pattern.get("daysOfWeek") or ()),)
    if kind == "weekly" and (boundary := _week_boundary(pattern)):
        # Counting the boundary every week would call an echo that dropped the
        # field an edit.
        key += (boundary,)
    if kind in ("absoluteMonthly", "absoluteYearly"):
        key += (_whole(pattern, "dayOfMonth"),)
    if kind in ("absoluteYearly", "relativeYearly"):
        key += (_whole(pattern, "month"),)
    if kind in ("relativeMonthly", "relativeYearly"):
        key += (str(pattern.get("index") or "first").lower(),)
    return key


def same_pattern(a: Any, b: Any) -> bool:
    """Whether two Graph patterns schedule the same days."""
    key = _pattern_key(a)
    return key is not None and key == _pattern_key(b)


def _refuse_move(kind: str, old: date, new: date, why: str) -> ValueError:
    return ValueError(
        f"This series' {kind} pattern was set up for a series starting "
        f"{old.isoformat()}, and the new start falls on {new.isoformat()} in the "
        f"event's zone. {why} Pass `recurrence` with the pattern you want alongside "
        f"`start`, and it is sent as given."
    )


def move_pattern(payload: dict, *, old: date, new: date) -> dict:
    """``payload``'s pattern moved from a series starting ``old`` to one starting ``new``.

    A series' pattern names days — ``daysOfWeek``, ``dayOfMonth`` — and those
    days belong to its start date. Re-deriving ``range.startDate`` alone moves
    the range and leaves the pattern behind: a Thursday 02:00Z series
    re-anchored to Los Angeles starts on Wednesday at 18:00 but still says
    ``daysOfWeek: ["thursday"]``, and Graph schedules every occurrence on
    Thursday. Verified live — reported as ``updated``, every instance a day late.

    So the pattern moves by the same number of days the start did. Weekly days
    shift together, and a multi-week pattern's ``firstDayOfWeek`` with them —
    Sunday when the pattern omits it — so a fortnightly Sunday-and-Monday series
    stays one block rather than being split across the week boundary. Every
    week, the boundary is left as it was: it schedules nothing there, and
    shifting it would only move the week start Outlook displays. Absolute
    patterns take the new date's day (and month,
    for yearly), provided the stored pattern was anchored on the old one.

    Refused, rather than approximated, where the moved series has no exact
    expression: a relative pattern ("the first Thursday") whose day moves, and a
    monthly one pushed into a neighbouring month — the day before "the 1st" is
    not any one ``dayOfMonth``.
    """
    delta = (new - old).days
    pattern = dict(payload.get("pattern") or {})
    kind = _pattern_type(pattern)
    if delta == 0 or kind == "daily":
        return payload

    if kind == "weekly":

        def shifted(day: Any) -> str:
            return _WEEKDAYS[(_WEEKDAYS.index(str(day).lower()) + delta) % 7]

        pattern["daysOfWeek"] = [shifted(d) for d in pattern.get("daysOfWeek") or ()]
        if boundary := _week_boundary(pattern):
            pattern["firstDayOfWeek"] = shifted(boundary)
    elif kind in ("absoluteMonthly", "absoluteYearly"):
        anchored = _whole(pattern, "dayOfMonth") == old.day and (
            kind == "absoluteMonthly" or _whole(pattern, "month") == old.month
        )
        if not anchored:
            raise _refuse_move(
                kind,
                old,
                new,
                "Its day does not match that start date, so there is no telling what "
                "moving it should mean.",
            )
        if kind == "absoluteMonthly" and (new.year, new.month) != (old.year, old.month):
            raise _refuse_move(
                kind,
                old,
                new,
                "That crosses into a different month, which no single dayOfMonth expresses.",
            )
        pattern["dayOfMonth"] = new.day
        if kind == "absoluteYearly":
            pattern["month"] = new.month
    else:
        raise _refuse_move(
            kind or "unknown",
            old,
            new,
            "A relative pattern cannot be moved by a day and stay exact: the day "
            "before the first Thursday is not always the first Wednesday.",
        )
    return {**payload, "pattern": pattern}


def build_event_recurrence(
    recurrence: dict | str,
    *,
    start: str,
    zone: str | None = None,
    drop_range_timezone: bool = False,
) -> Any:
    """Build a typed PatternedRecurrence for a calendar event.

    Accepts the Graph recurrence object, a JSON-encoded string of one, or a
    shorthand ("daily", "weekdays", "weekly", "monthly", "yearly") anchored on
    ``start``. ``range.startDate`` is defaulted from ``start`` when omitted.

    ``zone`` is the zone the event is anchored in. It decides which calendar
    date an offset-bearing ``start`` falls on — both for that default and for
    the weekday a shorthand expands to. Omitted, the date comes from the text,
    which is right only while the two agree; see ``event_start_date`` for the
    case where they do not.

    ``drop_range_timezone`` removes ``range.recurrenceTimeZone`` from the
    payload. Set it when the caller has named the event's zone by another route
    — an explicit ``timezone`` on ``update_event`` — because the two then
    describe the same thing and Graph refuses the pair when they disagree. The
    read-modify-write round trip makes that collision ordinary rather than
    exotic: ``outlook_get_event`` hands back ``recurrenceTimeZone``, so a caller
    echoing a recurrence while asking for a new zone sends the old zone with it.
    Graph re-derives the range's zone from the event's own when the field is
    absent, so dropping it is lossless.

    Off by default, because the general passthrough below is deliberate; this
    resolves one specific conflict rather than reversing that decision.

    ``range.recurrenceTimeZone`` is otherwise passed through rather than stripped, and
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
    payload = event_recurrence_payload(recurrence, start_date)

    if drop_range_timezone and payload.get("range"):
        payload = {
            **payload,
            "range": {
                k: v
                for k, v in _object(payload["range"], "range").items()
                if k != "recurrenceTimeZone"
            },
        }
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
