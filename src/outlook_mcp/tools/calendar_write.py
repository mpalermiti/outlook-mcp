"""Calendar write tools: create_event, update_event, delete_event, rsvp."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from datetime import timezone as dt_timezone
from typing import Any

from outlook_mcp._windows_zones import WINDOWS_TO_IANA
from outlook_mcp.calendar_resolver import (
    event_calendar_belongs_to_someone_else,
    resolve_calendar_for_write,
)
from outlook_mcp.config import Config
from outlook_mcp.errors import (
    EVENT_MAY_HAVE_BEEN_CREATED,
    INVITATIONS_MAY_HAVE_BEEN_SENT,
    RESPONSE_MAY_HAVE_BEEN_SENT,
    UPDATE_MAY_HAVE_BEEN_SENT,
    OutlookMCPError,
)
from outlook_mcp.permissions import (
    CATEGORY_CALENDAR_WRITE,
    check_permission,
    check_sends_mail,
    mail_send_withheld,
)
from outlook_mcp.throttle import call_without_replay
from outlook_mcp.tools._recurrence import (
    _OVERLONG_FRACTION,
    build_event_recurrence,
    check_recurrence_shape,
    describe_recurrence,
    event_recurrence_payload,
    event_start_date,
    lands_on,
    maybe_zone,
    move_pattern,
    move_series,
    same_pattern,
    serialize_recurrence,
)
from outlook_mcp.tools.calendar_read import EVENT_KEY_PROPERTY_ID
from outlook_mcp.validation import (
    resolve_config_event_timezone,
    resolve_relative_datetime,
    validate_email,
    validate_event_key,
    validate_event_timezone,
    validate_graph_id,
    validate_header_zone,
)


def _now() -> datetime:
    """The clock a call resolves its relative ``start``/``end`` against (#74).

    Read once per call and shared by both ends, so ``+2h``/``+3h`` are exactly
    an hour apart: two reads can straddle a second boundary, and the
    seconds-precision result would then make the event a second longer.
    """
    return datetime.now(dt_timezone.utc)


def _in_zone(instant: str, zone: str, given: str) -> str:
    """A resolved relative time as wall-clock time in ``zone``, the event's anchor (#74).

    Same moment, written without a zone or offset, because that is the only form
    Graph keeps the zone for: handed a ``Z`` or an offset, it anchors a single
    event in UTC whatever ``timeZone`` says (#236). Measured live on a consumer
    mailbox, 2026-11-10; a naive time kept both an IANA and a Windows label.

    Refused, rather than sent some other way, when that form cannot say which
    moment is meant: in a zone ``zoneinfo`` cannot convert into (Graph's Windows
    names — each series this server created in the 2026-11-10 measurement came
    back named that way, though it was sent an IANA zone), and inside
    the interval a zone repeats when daylight saving ends, where one wall-clock
    time names two moments. ``given`` is the caller's own text, for the message.
    """
    tz = maybe_zone(zone)
    if tz is None:
        raise ValueError(
            f"This event is anchored in {zone!r}, which is not an IANA zone name, so a "
            "relative time is not placed in it. Nothing was "
            "modified. Pass `start` and `end` as local time without a zone or offset "
            "(2026-11-10T09:00:00); they are read in the event's own zone."
        )
    moment = datetime.fromisoformat(instant.replace("Z", "+00:00"))
    local = _wall_clock(moment, tz)
    if local is None:
        # No offset is offered as the remedy: a typed one anchors a single event
        # in UTC (#236), so only a time outside the repeat is safe to suggest.
        raise ValueError(
            f"{given!r} falls in the interval {zone} repeats when daylight saving ends: "
            f"{moment.astimezone(tz):%Y-%m-%dT%H:%M:%S} happens twice there, so it names "
            "two different times. Nothing was modified. Pass a time outside that interval."
        )
    return local


def _wall_clock(moment: datetime, tz: Any) -> str | None:
    """``moment`` as naive wall-clock time in ``tz``, or ``None`` if that names two moments.

    The one conversion behind both ``_in_zone`` (relative times, #74) and
    ``_typed_in_zone`` (a typed ``Z`` or offset, #236). ``None`` inside the
    interval a zone repeats when daylight saving ends, where the wall-clock
    text alone cannot say which of the two moments is meant.
    """
    # `astimezone`, never timedelta arithmetic: it keeps `fold`, and cannot land
    # in a spring-forward gap, so the one case left to catch is the repeat.
    local = moment.astimezone(tz)
    if local.replace(fold=0).astimezone(dt_timezone.utc) != local.replace(fold=1).astimezone(
        dt_timezone.utc
    ):
        return None
    # `isoformat` writes no fraction for a whole second and six digits otherwise.
    return local.replace(tzinfo=None, fold=0).isoformat()


def _typed_in_zone(value: str, zone: str | None) -> str | None:
    """A typed ``Z``/offset ``value`` as wall-clock time in ``zone`` (#236).

    Graph keeps the instant of a ``Z`` or offset but anchors a single event in
    UTC whatever ``timeZone`` says; the same moment written as local time keeps
    the zone. Measured live on a consumer mailbox, 2026-11-10. A zone-less value
    is already local and comes back as is. A Windows name Graph gave
    ("Pacific Standard Time") converts through CLDR's default IANA zone for it
    (``_windows_zones``); the name sent stays Graph's own. ``None`` when there is
    no safe way to convert: a zone neither knows, a value this Python cannot
    parse, a moment past the years it can hold (``9999-12-31`` east of UTC), or
    one in the repeated hour (``_wall_clock``).
    """
    moment = _parse_iso(value)
    if moment is None:
        return None
    if moment.tzinfo is None:
        return value
    tz = maybe_zone(zone) or maybe_zone(WINDOWS_TO_IANA.get(zone or ""))
    try:
        return None if tz is None else _wall_clock(moment, tz)
    except OverflowError:
        return None


def _parse_iso(value: str) -> datetime | None:
    """``value`` parsed, a ``Z`` read as UTC and Graph's seven fractional digits cut to six."""
    text = _OVERLONG_FRACTION.sub(r"\1", value.strip())
    try:
        return datetime.fromisoformat(text[:-1] + "+00:00" if text.endswith("Z") else text)
    except ValueError:
        return None


def _typed_in_zones(ends: dict[str, tuple[str, str | None]]) -> dict[str, str]:
    """Each typed end, ``{name: (value, zone)}``, through ``_typed_in_zone``: all or none.

    If any end cannot be converted, every end is sent as written, which is the
    behaviour before #236: both ends filed in UTC. Converting only the other
    would file the two ends in different zones, a state nothing made before.
    """
    converted = {name: _typed_in_zone(value, zone) for name, (value, zone) in ends.items()}
    if any(v is None for v in converted.values()):
        return {name: value for name, (value, _) in ends.items()}
    return {name: v for name, v in converted.items() if v is not None}


# The check reads an event's times back with a GET of only ``start`` and ``end``,
# asking for them in UTC. Graph honours the header on a GET:
# ``_start_in_its_own_zone`` relies on it for a named zone, verified live, and
# UTC is what a GET returns anyway, which the read tools and their live tests
# rely on. It is asked for regardless, so the check rests on a request rather
# than on a default. Not on a write's reply: measured live 2026-11-25, a POST
# sent this header answered ``2026-12-27T18:00:00`` in ``America/Los_Angeles``,
# the event's own zone, so a check that read the reply never ran.
_IN_UTC = 'outlook.timezone="UTC"'
_TIMES_SELECT = ["start", "end"]

# The furthest a converted time is put back. Two sets of rules for one zone
# disagree by a whole step of its offset: an hour for daylight saving, half an
# hour at Lord Howe, two at Troll, or a change to the standard offset (Magadan
# moved two hours in 2014). Three hours covers any of these with room to spare.
# A larger difference is something else (a series' first occurrence days from its
# typed start, a date-line jump like Samoa's in 2011): it is reported, not changed.
_MOST_A_ZONE_RULE_MOVES = timedelta(hours=3)


def _converted(ends: dict[str, tuple[str, str | None]], sent: dict[str, str]) -> dict:
    """``{name: (typed value, zone)}`` for each end ``_typed_in_zones`` actually changed."""
    return {name: ends[name] for name, value in sent.items() if value != ends[name][0]}


def _is_single_instance(event: Any) -> bool:
    """Whether Graph says this event stands alone: no series' master, occurrence or exception.

    ``False`` when the stored event's ``type`` cannot be read: not knowing is not single.
    """
    event_type = getattr(event, "type", None)
    if event_type is None:
        return False
    return (getattr(event_type, "value", None) or str(event_type)) == "singleInstance"


async def _times_in_utc(target: Any) -> Any:
    """The event's ``start`` and ``end`` as Graph reads them back in UTC, or ``None``.

    One GET of only those two fields, through ``target``, the request builder
    the write went to, so an event in another calendar is read there. Sent once,
    with kiota's retry off: the write has succeeded, and a throttled read
    retried after ``Retry-After`` would hold back its answer for minutes, long
    enough for a client to give up and create the event again. ``None`` when
    Graph or the connection fails: the check then fails open, as when the times
    cannot be read, rather than turning a write that succeeded into an error.
    Anything else is a bug, and is not caught.
    """
    import httpx
    from kiota_abstractions.api_error import APIError
    from kiota_abstractions.base_request_configuration import RequestConfiguration
    from kiota_http.middleware.options.retry_handler_option import RetryHandlerOption
    from msgraph.generated.users.item.events.item.event_item_request_builder import (
        EventItemRequestBuilder,
    )

    query = EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters(select=_TIMES_SELECT)
    request_config = RequestConfiguration(
        query_parameters=query, options=[RetryHandlerOption(should_retry=False)]
    )
    request_config.headers.add("Prefer", _IN_UTC)
    try:
        return await target().get(request_configuration=request_config)
    except (APIError, httpx.HTTPError):
        return None


def _utc_instant(side: Any) -> datetime | None:
    """The moment a ``start``/``end`` names, or ``None`` unless it is text in ``UTC``."""
    text = getattr(side, "date_time", None)
    if not isinstance(text, str) or getattr(side, "time_zone", None) != "UTC":
        return None
    moment = _parse_iso(text)
    if moment is None:
        return None
    return moment.replace(tzinfo=dt_timezone.utc) if moment.tzinfo is None else moment


def _read_back_instant(event: Any, which: str) -> datetime | None:
    """The UTC moment ``_times_in_utc`` read for ``start`` or ``end``, or ``None``.

    ``None`` unless the event carries that end as text in ``UTC``, the zone
    ``_IN_UTC`` asked for: anything else would be a guess.
    """
    side = getattr(event, "start", None) if which == "start" else getattr(event, "end", None)
    return _utc_instant(side)


def _read_back_instants(event: Any, converted: dict) -> dict[str, datetime] | None:
    """Each converted end's moment in the read-back ``event``, or ``None`` if any cannot be read.

    Not knowing one end is not knowing where the event is.
    """
    got = {}
    for which in converted:
        moment = _read_back_instant(event, which)
        if moment is None:
            return None
        got[which] = moment
    return got


def _stored_instants(event: Any, ends: dict) -> dict[str, datetime]:
    """Where each of ``ends`` stood in ``event`` as a plain GET returned it, before a write.

    A plain GET gives ``start`` and ``end`` in UTC; an end it did not give that
    way is left out, and then nothing is compared with it.
    """
    sides = {"start": getattr(event, "start", None), "end": getattr(event, "end", None)}
    found = {which: _utc_instant(sides[which]) for which in ends}
    return {which: moment for which, moment in found.items() if moment is not None}


def _as_it_was(got: dict[str, datetime], earlier: dict[str, datetime], converted: dict) -> bool:
    """Whether a read found an end exactly where it stood before the latest write, not as typed.

    That is most likely the event as it was, served by a read that came too
    soon after the write (not seen for events; Graph has lagged straight after a
    write for contacts, #259), not where Graph put the time sent. Taken at its
    word, the gap is the move just asked for, which passes for a zone rule's and
    would be "corrected", filing the event in UTC and mailing any attendees
    again. So such a read says nothing. A real disagreement over a zone's rules
    lands on the old time only when the move happens to equal the rules' gap,
    and that event is left as Graph put it.
    """
    for which, (typed, _) in converted.items():
        was, wanted = earlier.get(which), _parse_iso(typed)
        if was is None or wanted is None:
            continue
        at_old = abs(got[which] - was) < timedelta(seconds=1)
        if at_old and abs(got[which] - wanted) >= timedelta(seconds=1):
            return True
    return False


def _a_zone_rule_could_move(got: datetime, wanted: datetime) -> bool:
    """Whether two sets of rules for one zone could put the time ``wanted`` at ``got``.

    Offsets are whole minutes, so the gap is an exact number of minutes, at most
    ``_MOST_A_ZONE_RULE_MOVES``: measured from the time typed, or from it without
    its fraction of a second, which Graph may drop.
    """
    for meant in (wanted, wanted.replace(microsecond=0)):
        gap = abs(got - meant)
        if gap <= _MOST_A_ZONE_RULE_MOVES and gap % timedelta(minutes=1) == timedelta(0):
            return True
    return False


def _moved_ends(read: dict[str, datetime], converted: dict) -> tuple[list[str], bool]:
    """Each converted end ``read`` puts at another moment than typed, described, and
    whether a zone rule could explain every one of those gaps.
    """
    moved, explained = [], True
    for which, (typed, _) in converted.items():
        got, wanted = read[which], _parse_iso(typed)
        if wanted is None:
            return [], True
        if abs(got - wanted) < timedelta(seconds=1):
            continue
        explained = explained and _a_zone_rule_could_move(got, wanted)
        moved.append(
            f"the {which} at {got.astimezone(dt_timezone.utc):%Y-%m-%dT%H:%M:%S} UTC "
            f"instead of {wanted.astimezone(dt_timezone.utc):%Y-%m-%dT%H:%M:%S} UTC"
        )
    return moved, explained


def _zone_not_kept(moved: list[str], converted: dict, still: list[str] | None) -> str:
    """The ``note`` for a corrected write: what moved, how far, and what is known now.

    ``still`` is what the read after the correction found away from the moment
    typed: empty when it found every end there, ``None`` when it could not say.
    Only the first says the time is right.
    """
    ends = " and ".join(converted)
    both = len(converted) > 1
    placed = (
        f"Graph placed {' and '.join(moved)}: its rules for the event's time zone differ "
        f"from this server's. The {ends} {'were' if both else 'was'} saved again as typed"
    )
    filed = (
        f"{'they are' if both else 'it is'} filed in UTC rather than in the event's own time zone"
    )
    if still is None:
        return (
            f"{placed}, so {filed}. Reading the event back afterwards did not confirm the "
            "time; outlook_get_event shows where it is."
        )
    if still:
        return (
            f"{placed}, but reading the event back puts {' and '.join(still)}. Read the "
            "event with outlook_get_event and check the time with the user."
        )
    return f"{placed}, so the time is right, but {filed}."


async def _keep_typed_instants(
    target: Any, converted: dict, done: str, before: dict[str, datetime] | None = None
) -> dict:
    """Check a write that converted a typed time by reading the event back (#262 review).

    The conversion uses this host's zone rules (IANA, through ``_windows_zones``
    for a Windows name) and Graph reads the local time back with its own. Where
    they disagree for that zone and date, or this host's tz data is out of date,
    the event lands at another moment. Then each converted end is sent again as
    typed: the old behaviour, the right moment, the zone lost. ``{"note": ...}``
    says so; ``{}`` when nothing moved or the read cannot say. A gap no zone
    rule makes (``_a_zone_rule_could_move``) is not the conversion's doing, so
    nothing is written and the ``note`` reports it for the caller to check.

    Where the event is comes from a GET (``_times_in_utc``), not the write's
    reply, which Graph gives in the event's own zone whatever ``Prefer`` asks.
    One GET when ``converted`` is not empty, none otherwise. A read that fails,
    cannot be parsed, or finds an end exactly where ``before`` says it stood
    before the write (``_as_it_was``; an update passes it, a create has nothing
    before) means no correction and no ``note``, as before the check.

    After a correction the event is read once more, so the ``note`` says the
    time is right only when that read finds it so: a GET, on a path taken only
    when Graph and this host already disagreed. If it finds the event elsewhere,
    nothing more is sent and the ``note`` says where. If it cannot say, or finds
    the event where the first read did, the ``note`` says the time was not
    confirmed.

    Callers pass ``converted`` only for a single event. A series' start can
    read back as its first occurrence rather than the start typed, days away,
    and a patch without its recurrence would file the series in UTC or be refused.

    ``target`` makes the request builder the write went to. The correction goes
    through ``call_without_replay``: an idempotent patch, but on an event with
    attendees each one is mailed as an update. Not pinned with ``If-Match``: it
    follows the write it corrects at once, and sends only the two times.
    """
    if not converted:
        return {}
    read = _read_back_instants(await _times_in_utc(target), converted)
    if read is None or _as_it_was(read, before or {}, converted):
        return {}
    moved, explained = _moved_ends(read, converted)
    if not moved:
        return {}
    if not explained:
        return {
            "note": f"Graph placed {' and '.join(moved)}. No time zone rule explains a "
            "difference like that, so it was not changed. Read the event with "
            "outlook_get_event and check the time with the user."
        }

    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.event import Event

    fix = Event()
    if "start" in converted:
        fix.start = DateTimeTimeZone()
        fix.start.date_time, fix.start.time_zone = converted["start"]
    if "end" in converted:
        fix.end = DateTimeTimeZone()
        fix.end.date_time, fix.end.time_zone = converted["end"]
    builder = target()
    try:
        await call_without_replay(
            lambda once: builder.patch(fix, request_configuration=once),
            outcome_unknown=UPDATE_MAY_HAVE_BEEN_SENT,
        )
    except Exception as exc:
        # Graph's own text without its hint: the action below is the one that
        # fits, and on a 5xx it already covers "it may have been saved".
        detail = exc.message if isinstance(exc, OutlookMCPError) else type(exc).__name__
        raise OutlookMCPError(
            "event_time_not_restored",
            f"The event was {done}, but Graph placed {' and '.join(moved)}. Sending "
            f"the time again as typed ended in an error: {detail}",
            "Read the event with outlook_get_event to see where it is now. If it is "
            "still at the wrong time, move it with outlook_update_event, not by "
            "creating it again: pass `start` and `end` as local time without a zone or "
            "offset, which needs no conversion.",
        ) from exc
    after = _read_back_instants(await _times_in_utc(target), converted)
    if after is None or _as_it_was(after, read, converted):
        return {"note": _zone_not_kept(moved, converted, None)}
    still, _ = _moved_ends(after, converted)
    return {"note": _zone_not_kept(moved, converted, still)}


def _usable_zone(value: Any) -> str | None:
    """A zone name we can send back, or ``None`` when there isn't one.

    An event created against a custom time zone reports the sentinel
    ``tzone://Microsoft/Custom``. It cannot be echoed — Graph answers it with a
    400 — and every substitute moves the event, so the caller is asked for a
    zone rather than given a guess. Documented behaviour; not something I have
    reproduced live, which is why this refuses rather than translates.
    """
    if not value or str(value).startswith("tzone://"):
        return None
    return str(value)


def _anchor_zones(event: Any) -> tuple[str | None, str | None]:
    """The zones ``event``'s start and end are anchored in, as a plain GET reports them.

    Deliberately **not** ``event.start.time_zone``. Graph projects ``start`` and
    ``end`` into UTC unless the request carries a ``Prefer: outlook.timezone``
    header, and this server never sends one — so ``start.time_zone`` reads
    ``"UTC"`` for every event ever fetched, whatever it is anchored in. The
    ``original*TimeZone`` fields are the ones that survive that projection.

    Reading the wrong one is not a near-miss: it makes "keep the zone the event
    is already stored in" resolve to UTC every time, which is precisely the
    behaviour this module was changed to stop — a New York meeting patched to a
    new time is relocated to UTC and its series re-anchored, reported as
    ``updated``.

    Start and end are returned **separately** because Graph stores them
    separately, and a transcontinental flight is the ordinary case: verified
    live 2026-09-21, an event sent as 08:00 ``America/New_York`` to 11:00
    ``America/Los_Angeles`` comes back as 13:00Z to 19:00Z with both anchors
    intact. Collapsing them onto the start's zone would relabel that landing
    time and move it three hours.

    Read one attribute at a time rather than through a loop: ``getattr`` with a
    variable attribute is invisible to the guards built for exactly this class
    of field-name drift.
    """
    return (
        _usable_zone(getattr(event, "original_start_time_zone", None)),
        _usable_zone(getattr(event, "original_end_time_zone", None)),
    )


# Why `timezone` and an all-day event cannot be combined. One constant rather
# than the same sentence in two branches: the check fires either before any
# network call (the caller said `is_all_day=True`) or after one read (only the
# stored event knew), and a refusal that drifts between those two paths would
# read as two different rules.
_ALL_DAY_ZONE_REFUSAL = (
    "timezone cannot be applied to an all-day event. Graph stores one anchored in "
    "UTC whatever zone it is sent — verified live — so there is no zone to "
    "re-anchor it to, and accepting this would report success for nothing. "
    "To turn it into a timed event instead, pass is_all_day=False with real "
    "start and end times."
)

# Graph accepts `isOnlineMeeting: false` on a PATCH and keeps the meeting: the
# reply and a fresh GET both still report it, unchanged 90 seconds later, on a
# personal account (probed 2026-11-09, #70). Answering `updated` would report
# success for nothing, so the argument is refused before anything is sent.
_ONLINE_REMOVAL_REFUSAL = (
    "is_online=False cannot remove an online meeting: Graph accepts the request "
    "and keeps the meeting (verified on a personal account). To drop it, delete "
    "the event and create it again without is_online. On an event with attendees, "
    "that sends them a cancellation and a new invitation."
)


def _online_reply(response: Any) -> dict:
    """``is_online`` as Graph's write reply states it, or nothing.

    Graph honours ``isOnlineMeeting`` on some personal accounts and ignores it
    on others (#70), and its reply says which — even for a request it ignored.
    A reply that does not carry the field as a boolean yields no key rather
    than ``False``: absent is not "no meeting".
    """
    value = getattr(response, "is_online_meeting", None)
    return {"is_online": value} if isinstance(value, bool) else {}


# Graph needs an all-day event's ends at midnight, which a relative time almost
# never is; its 400 says so, but only after the round trip. One constant for
# both tools and both paths through `update_event` (the caller said
# `is_all_day=True`, or only the stored event knew).
_ALL_DAY_RELATIVE_REFUSAL = (
    "A relative time (+2h, now, 7d) cannot place an all-day event: Graph needs "
    "both ends at midnight. Nothing was modified. Pass dates at midnight, e.g. "
    "start=2026-10-22T00:00:00, end=2026-10-23T00:00:00."
)


def _placed(local: str, zone: str) -> str:
    """A relative time as it was written, in the form the read tools show (#74).

    Reported on the write result, so a caller who passed ``start="7d"`` can see
    it made an event a week in the past.
    """
    return f"{local} ({zone})"


def _key_property(key: str) -> Any:
    """The extended property that stamps an event with ``key``.

    Sent in the event's own POST or PATCH; Graph upserts it by id, so an
    update replaces the key in place. This server offers no way to clear one.
    """
    from msgraph.generated.models.single_value_legacy_extended_property import (
        SingleValueLegacyExtendedProperty,
    )

    prop = SingleValueLegacyExtendedProperty()
    prop.id = EVENT_KEY_PROPERTY_ID
    prop.value = key
    return prop


def _has_attendees(event: Any) -> bool:
    """True when this event has other people on it, whoever organizes it.

    The organizer's edit is mailed to the attendees as an update. An attendee's
    edit stays in their own calendar — but the response Exchange sends when
    they later accept or decline is built from that copy, so its text reaches
    the organizer all the same. Neither side is exempt.
    """
    return bool(getattr(event, "attendees", None))


# What a refusal says when the calendar is not the user's own. Content written
# to a calendar someone shared is shown to them, as an invitation would be, so
# it needs `mail_send` under a policy that withholds it.
SOMEONE_ELSES_CALENDAR = "write to a calendar not confirmed as the user's own"


async def refuse_someone_elses_calendar(
    graph_client: Any, event_id: str, config: Config, tool_name: str
) -> None:
    """Refuse adding content to an event on someone else's calendar without ``mail_send``.

    Looked up only under a policy that withholds ``mail_send``; otherwise no
    request is made. See ``event_calendar_belongs_to_someone_else`` for the
    lookup and for what counts as unconfirmed.
    """
    if not mail_send_withheld(config):
        return
    if await event_calendar_belongs_to_someone_else(graph_client, event_id):
        check_sends_mail(config, tool_name, SOMEONE_ELSES_CALENDAR, emails=False)


def _is_series_master(event: Any) -> bool:
    """Whether Graph considers this event the master of a recurring series.

    Only a master's zone or date change needs its recurrence re-sent; a single
    instance and an occurrence both patch cleanly without one.
    """
    event_type = getattr(event, "type", None)
    if event_type is None:
        return False
    return (getattr(event_type, "value", None) or str(event_type)) == "seriesMaster"


async def _start_in_its_own_zone(graph_client: Any, event_id: str, zone: str) -> str:
    """The event's start as local wall clock in ``zone``.

    Graph returns ``start.dateTime`` projected into UTC unless the request asks
    otherwise, and it names the event's own zone in Windows terms
    ("Pacific Standard Time") for anything it was not given an IANA name for —
    which Python cannot map. So the date an 18:00 Pacific event falls on is
    underivable locally: `2026-11-05T02:00:00Z` is the 5th, the event is on the
    4th, and a weekly series built from the former lands on the wrong weekday.

    Rather than carry a Windows-to-IANA table, ask the side that has the
    mapping. ``Prefer: outlook.timezone`` makes Graph do the projection, and it
    accepts the same Windows name it just handed us. Verified live: the header
    is honoured on the SDK's own request builder, so this needs no raw-httpx
    path and inherits kiota's retry handling rather than owing its own.

    Every way this read can fail to deliver that is refused rather than papered
    over, because the only fallback is the date as written — the UTC date, which
    is the wrong-day bug this exists to fix. A failed read is not caught; a reply
    without a start is refused; and so is a start projected into any zone but
    the one asked for. Graph echoes the requested name back verbatim in ``start.timeZone``
    — verified live for both a Windows and an IANA name — so a mismatch means the
    header was not honoured, and the text beside it is UTC or worse. Read as
    local time it builds a wrong-day series that looks like success.
    """
    from kiota_abstractions.base_request_configuration import RequestConfiguration

    # `zone` is written into the header as is, so it is checked first: Graph
    # supplied it, but a name that ends the quoted value or the line, or that a
    # header cannot encode, must not become part of a request.
    validate_header_zone(
        zone,
        remedy="It is the zone Graph reports for this event. Nothing was modified. Pass "
        "`start` alongside `recurrence` to anchor the series explicitly.",
    )
    request_config = RequestConfiguration()
    request_config.headers.add("Prefer", f'outlook.timezone="{zone}"')
    localized = await graph_client.me.events.by_event_id(event_id).get(
        request_configuration=request_config
    )
    start = getattr(localized, "start", None)
    date_time = getattr(start, "date_time", None)
    projected = getattr(start, "time_zone", None)
    if not date_time or projected != zone:
        got = f"it in {projected!r}" if date_time else "no start at all"
        raise ValueError(
            f"Asked Graph for this event's start in {zone!r} and got {got}, so the "
            f"day its series falls on cannot be trusted. Nothing was modified. Pass "
            f"`start` alongside `recurrence` to anchor the series explicitly."
        )
    return str(date_time)


# Graph discards a series' edited and deleted occurrences whenever its master's
# times or recurrence change, and says nothing. Measured live on a consumer
# mailbox, one edited and one deleted occurrence each time: moving the start an
# hour in the same zone, moving only the end, re-anchoring into another zone,
# extending the range and adding a weekday all restored both, while a subject
# patch and a recurrence re-sent unchanged kept them. So the loss follows the
# change to the series' shape, not the recurrence a zone change carries along.
#
# A plain GET leaves both collections out, so they are asked for by name; the
# nested exceptions take the same projection, which is why `subject` and `start`
# are here. `_occurrences_a_series_change_discards` is the half that reads them.
_OCCURRENCE_SELECT = "subject,start,cancelledOccurrences,exceptionOccurrences"


def _cancelled_date(occurrence_id: Any) -> str:
    """The date Graph encodes at the end of a cancelled occurrence's id.

    Graph lists a deleted occurrence as ``OID.<master id>.2026-11-16`` rather
    than as an event, so the date is the only thing there is to name.
    """
    return str(occurrence_id).rsplit(".", 1)[-1]


def _occurrences_a_series_change_discards(master: Any) -> list[str]:
    """Each edited or deleted occurrence of ``master``, described for a refusal.

    Fails closed: a read that omits either collection cannot say the series is
    clean, and a guess here is how data goes missing without an error.
    """
    cancelled = getattr(master, "cancelled_occurrences", None)
    edited = getattr(master, "exception_occurrences", None)
    if cancelled is None or edited is None:
        raise ValueError(
            "Could not read which occurrences of this series have been edited or "
            "deleted, and changing a series' times or recurrence discards them. "
            "Nothing was modified."
        )
    return [
        f"edited {getattr(getattr(e, 'start', None), 'date_time', '?')} UTC "
        f"({getattr(e, 'subject', None) or 'no subject'})"
        for e in edited
    ] + [f"deleted {_cancelled_date(c)}" for c in cancelled]


async def _refuse_to_discard_occurrences(graph_client: Any, event_id: str) -> None:
    """Refuse a times or recurrence change to a master with edited or deleted occurrences."""
    from kiota_abstractions.base_request_configuration import RequestConfiguration
    from msgraph.generated.users.item.events.item.event_item_request_builder import (
        EventItemRequestBuilder,
    )

    query = EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters(
        select=_OCCURRENCE_SELECT.split(",")
    )
    master = await graph_client.me.events.by_event_id(event_id).get(
        request_configuration=RequestConfiguration(query_parameters=query)
    )
    lost = _occurrences_a_series_change_discards(master)
    if not lost:
        return
    raise ValueError(
        f"Changing this series' times or recurrence would discard {len(lost)} "
        f"occurrence change(s) — Graph restores every edited and deleted occurrence "
        f"when a series is reshaped, without saying so: {'; '.join(lost)}. Nothing was "
        f"modified. To keep them, change individual occurrences instead, and omit "
        f"`recurrence` if it is unchanged. If the user accepts losing them, the change "
        f"has to be made outside this tool."
    )


def _follow_new_start(
    payload: dict, existing: Any, *, start: str, zone: str | None, given: bool = True
) -> dict:
    """A recurrence payload made to agree with a new ``start``.

    ``range.startDate`` is dropped so the builder re-derives it from ``start``:
    a series whose start moves has to move its range with it, and Graph refuses
    a range beginning on a different day than the first occurrence.

    The pattern moves too when it is the series' own — the one Graph holds,
    whether read back here for a zone change or handed back by the caller from
    ``outlook_get_event``. A zone change reads the same instants in a new zone,
    so its days shift as far as the stored start's local day does, all of them
    (``move_pattern``). A start the caller also moved to another day then
    follows the rules of a same-zone move (#98 review): the series' own,
    re-sent, goes through ``move_series``; one the caller handed back keeps a
    single day following the new start, as agents moving a weekly meeting
    rely on, and several days exactly as written. A pattern that differs from
    the stored one is the caller's new instruction and is sent as given.

    Except where the stored pattern skips its own ``startDate`` (a "weekdays"
    series begun on a Saturday): Graph may have put its first occurrence later,
    so no move can be measured from that date (#98). The caller's pattern
    (``given``) is then sent as written; the series' own, read back for a zone
    change, is refused with the remedy of passing one.
    """
    rng = {k: v for k, v in (payload.get("range") or {}).items() if k != "startDate"}
    payload = {**payload, "range": rng}

    stored = serialize_recurrence(getattr(existing, "recurrence", None)) or {}
    if not _is_series_master(existing) or not same_pattern(
        payload.get("pattern"), stored.get("pattern")
    ):
        return payload

    stored_start = (stored.get("range") or {}).get("startDate")
    if not stored_start:
        raise ValueError(
            "Could not read the date this series currently starts on, so there is no "
            "telling whether its days have to move with the new start. Pass "
            "`recurrence` with the pattern you want alongside `start`."
        )
    old = date.fromisoformat(str(stored_start))
    if lands_on(stored.get("pattern"), old) is False:
        if given:
            return payload
        raise ValueError(
            f"This series' pattern does not land on {old.isoformat()}, the day its range "
            "starts, so there is no telling how far its days should move with the new "
            "start. Nothing was modified. Pass `recurrence` with the days you want "
            "alongside `start`; it is sent as given."
        )
    new = event_start_date(start, zone)
    by_zone = _day_in_new_zone(existing, zone) or old
    shifted = move_pattern(payload, old=old, new=by_zone)
    if new == by_zone:
        return shifted
    if not given:
        return move_series(shifted, old=by_zone, new=new)
    if len({str(d).lower() for d in (payload.get("pattern") or {}).get("daysOfWeek") or ()}) > 1:
        return payload
    return move_pattern(payload, old=old, new=new)


def _day_in_new_zone(existing: Any, zone: str | None) -> date | None:
    """The day the stored start falls on in ``zone``, if that is a zone change.

    Where a zone change alone puts the series: Thursday 02:00Z is Wednesday in
    Los Angeles. ``None`` with the zone unchanged, or when it cannot be worked
    out here, and the stored ``startDate`` then stands for it.
    """
    stored_zone, _ = _anchor_zones(existing)
    stored = getattr(existing, "start", None)
    at = getattr(stored, "date_time", None)
    if stored_zone == zone or not at or maybe_zone(zone) is None:
        return None
    return event_start_date(_as_instant(str(at), getattr(stored, "time_zone", None)), zone)


def _resend_recurrence(event: Any, *, anchor: str | None, zone: str | None) -> Any:
    """The event's own recurrence, rebuilt as a payload safe to send back.

    Not ``event.recurrence`` handed straight over: that model came off the wire
    with its unset fields explicitly ``None``, and the request adapter
    serializes through the backing store, which writes those nulls out — onto
    the *parent*, under their Python names (#63/#64).
    ``serialize_recurrence`` omits what is unset, and the strict builder then
    produces a fresh model carrying only the fields Graph actually sent.

    It never returns ``None``. A master that comes back without a readable
    recurrence cannot have its zone changed at all — Graph refuses the patch
    unless one travels with it — and assigning ``None`` to ``event.recurrence``
    is the unsendable-null shape from #63/#64 besides.

    ``startDate`` is dropped and the pattern moved with the start; see
    ``_follow_new_start``. ``recurrenceTimeZone`` is dropped because it is the
    *old* zone. Sending it back beside
    a new ``start.timeZone`` is a self-contradictory patch, and Graph answers
    the contradiction with the same opaque ``400
    ErrorPropertyValidationFailure`` this function exists to avoid. Omitted,
    Graph derives the range's zone from the event's own — verified live, a
    series created with no ``recurrenceTimeZone`` at all came back carrying the
    event's.
    """
    payload = serialize_recurrence(getattr(event, "recurrence", None))
    if not payload:
        raise ValueError(
            "This series came back without a readable recurrence, and Graph refuses "
            "to change a series' zone unless its recurrence travels with the patch. "
            "Nothing was modified. Pass `recurrence` alongside `timezone`."
        )

    start = anchor or getattr(getattr(event, "start", None), "date_time", None)
    if not start:
        raise ValueError(
            "Could not read the event's current start to re-anchor its recurrence "
            "for a timezone change; pass `start` and `end` alongside `timezone`."
        )

    # `recurrenceTimeZone` is dropped by the builder itself, which is the one
    # place that knows why (see `drop_range_timezone`).
    payload = _follow_new_start(payload, event, start=start, zone=zone, given=False)
    return build_event_recurrence(payload, start=start, zone=zone, drop_range_timezone=True)


def _unconverted(start: str, zone: str) -> ValueError:
    """Why a ``Z``/offset ``start`` left as typed cannot move a series (#98)."""
    return ValueError(
        f"{start!r} could not be converted into local time in this series' zone, "
        f"{zone!r}. A series' start can move to another day only as local time, which "
        "takes its repeat rule along. Nothing was modified. Pass `start` and `end` as "
        "local time without a zone or offset (2026-11-10T09:00:00); they are read in "
        "the event's own zone."
    )


async def _day_move(
    graph_client: Any, event_id: str, master: Any, start: str | None, zone: str
) -> tuple[date, date] | None:
    """The days series ``master`` moves between when ``start`` is sent, or ``None`` (#98).

    Graph answers a master's start moved to another day ``updated`` and snaps it
    back onto its pattern unless the recurrence travels with it: every
    occurrence stays where it was. Measured live, a timed and an all-day weekly
    series. ``None`` for a change within the day, or no ``start`` at all.

    The day it is on now is its ``range.startDate`` when the pattern lands
    there. Otherwise (a "weekdays" series begun on a Saturday) it is read off
    the master's own start, since Graph may have put that later, and a move is
    refused: no shift can be measured from ``startDate``. Refused too where the
    move cannot be sent as local time: no ``startDate``, or a ``Z``/offset
    ``start`` that ``_typed_in_zones`` left as typed (an unknown zone, the hour
    that repeats when clocks go back).
    """
    if start is None:
        return None
    stored = serialize_recurrence(getattr(master, "recurrence", None)) or {}
    first = (stored.get("range") or {}).get("startDate")
    if not first:
        raise ValueError(
            "Could not read the date this series currently starts on, so there is no "
            "telling whether the new start moves it to another day. Nothing was "
            "modified. Pass `recurrence` with the pattern you want alongside `start`."
        )
    typed = _parse_iso(start)
    left_typed = typed is not None and typed.tzinfo is not None
    if left_typed and maybe_zone(zone) is None:
        raise _unconverted(start, zone)
    first = date.fromisoformat(str(first))
    current = first
    if lands_on(stored.get("pattern"), first) is not True:
        now = getattr(getattr(master, "start", None), "date_time", None)
        if not now:
            raise ValueError(
                "Could not read the day this series currently starts on. Nothing was "
                "modified. Pass `recurrence` with the pattern you want alongside `start`."
            )
        if maybe_zone(zone) is None:
            now = await _start_in_its_own_zone(graph_client, event_id, zone)
        else:
            now = _as_instant(str(now), getattr(master.start, "time_zone", None))
        current = event_start_date(now, zone)
    new = event_start_date(start, zone)
    if new == current:
        return None
    if left_typed:
        raise _unconverted(start, zone)
    if current != first:
        raise ValueError(
            f"This series' pattern does not land on {first.isoformat()}, the day its "
            f"range starts, and its first occurrence is {current.isoformat()}, so there "
            f"is no telling how far its days should move to start on {new.isoformat()}. "
            "Nothing was modified. Pass `recurrence` with the days you want alongside "
            "`start`; it is sent as given."
        )
    return first, new


def _moved_recurrence(master: Any, *, start: str, zone: str, old: date, new: date) -> Any:
    """``master``'s own recurrence for its start moved from ``old`` to ``new`` (#98).

    ``move_series`` decides the days; rebuilt as ``_resend_recurrence`` does,
    for the same reasons, ``recurrenceTimeZone`` dropped because Graph derives
    the same one from the event's zone.
    """
    payload = move_series(serialize_recurrence(master.recurrence) or {}, old=old, new=new)
    return build_event_recurrence(payload, start=start, zone=zone, drop_range_timezone=True)


def _series_note(master: Any, sent: Any) -> str | None:
    """What a re-sent recurrence now says, when it schedules other days than before.

    The caller passed no ``recurrence``, so the result is the only place the
    agent learns its series' rule was rewritten.
    """
    before = describe_recurrence(serialize_recurrence(getattr(master, "recurrence", None)) or {})
    after = describe_recurrence(serialize_recurrence(sent) or {})
    if before == after:
        return None
    return f"The series' repeat rule moved with its start and now reads: {after}."


def _as_instant(date_time: str, projected_zone: str | None) -> str:
    """Mark a stored start as the instant it is, so its local date can be found.

    Graph returns ``start.dateTime`` as a **naive** string already projected
    into ``start.timeZone`` — "2026-10-29T01:00:00.0000000" with
    ``timeZone: "UTC"`` beside it. Handed on as-is it reads as wall-clock time,
    so ``event_start_date`` takes the UTC date and never converts. Appending
    the offset turns it back into the instant Graph meant, which is the thing
    a zone can be applied to.

    Only UTC is handled, because that is the only projection this server ever
    asks for: no request here sends ``Prefer: outlook.timezone``. Anything else
    is returned untouched rather than guessed at.
    """
    text = date_time.strip()
    if not projected_zone or projected_zone.strip().upper() != "UTC":
        return text
    if text.endswith(("Z", "z")) or "+" in text[10:] or "-" in text[10:]:
        return text
    return text + "Z"


# Spellings an agent plausibly emits for the two Graph values whose names don't
# match Outlook's own menu labels ("Out of office", "Working elsewhere"), mapped
# to the normalised form of the real value. Widening what we accept cannot break
# the advertised contract; the refusal path is where a sanitizer does damage
# (#30), so that message names every valid value instead of guessing.
_FREE_BUSY_ALIASES = {
    "out_of_office": "oof",
    "outofoffice": "oof",
    "working_elsewhere": "workingelsewhere",
}


def _free_busy(value: str) -> Any:
    """Resolve a ``show_as`` string to the SDK's ``FreeBusyStatus`` member.

    Accepts the six Graph values case-insensitively, plus the aliases above;
    spaces and hyphens normalise to underscores, so "Out of office" and
    "working-elsewhere" both land.

    ``unknown`` is accepted even though Outlook's menu does not offer it. It is
    Graph's sentinel for a status it cannot determine, and Graph stores it on
    both POST and PATCH — verified live on a consumer mailbox, sent and read
    back on a fresh GET. Refusing it would break the round trip this project
    keeps elsewhere: an event read back as ``show_as: "unknown"`` has to be
    writable back unchanged, so the value ``get`` returns is a value ``create``
    takes.
    """
    from msgraph.generated.models.free_busy_status import FreeBusyStatus

    valid = [member.value for member in FreeBusyStatus]
    if not isinstance(value, str) or not value.strip():
        raise ValueError(
            f"show_as must be one of: {valid}. Omit it to leave the event's "
            f"busy status alone — this tool cannot clear it."
        )

    normalised = value.strip().lower().replace(" ", "_").replace("-", "_")
    normalised = _FREE_BUSY_ALIASES.get(normalised, normalised)

    for member in FreeBusyStatus:
        if member.value.lower() == normalised:
            return member

    # The valid set is derived from the enum so it cannot drift. The sentence
    # after it names only the two values whose Graph spelling differs from the
    # label Outlook shows — the same two `_FREE_BUSY_ALIASES` exists for — and
    # no longer claims to translate the whole list. Naming all six would mean
    # inventing an Outlook label for `unknown`, which the UI does not offer,
    # so the enumeration is the part that had to go rather than the hint.
    raise ValueError(
        f"Invalid show_as '{value[:50]}'. Must be one of: {valid}. Outlook labels "
        f"'oof' as Out of office and 'workingElsewhere' as Working elsewhere."
    )


async def create_event(
    graph_client: Any,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    is_all_day: bool = False,
    is_online: bool = False,
    recurrence: dict | str | None = None,
    timezone: str | None = None,
    show_as: str | None = None,
    calendar: str | None = None,
    key: str | None = None,
    *,
    config: Config,
) -> dict:
    """Create a calendar event.

    Validates inputs, builds a Graph Event object, and posts via
    graph_client.me.events.post().

    ``recurrence`` accepts a Graph recurrence object, a JSON string of one, or
    a shorthand ("daily", "weekdays", "weekly", "monthly", "yearly"). Setting
    it makes the created event a series master rather than a single occurrence.

    ``timezone`` is the IANA zone the event is *anchored* in, defaulting to
    ``config.timezone``. An offset or ``Z`` in the datetime pins the instant;
    the zone decides what the second occurrence of a weekly series does when
    daylight saving ends. Graph anchors a single event sent a ``Z`` or offset in
    UTC whatever ``timeZone`` says, so such a value is sent as the same moment
    in wall-clock time in the zone (``_typed_in_zones``, #236), then, for a
    single event, read back and checked (``_keep_typed_instants``). Anchored in
    ``UTC`` — which is what this tool sent unconditionally before 1.23.0 — a
    09:00 series silently becomes 08:00 the week the clocks go back.

    A zone-less ``start``/``end`` ("2026-10-28T09:00:00") is read by Graph as
    wall-clock time *in that zone*, which is what ``config.timezone`` has
    always meant for input interpretation everywhere else in this server.

    Otherwise the datetime string is passed through as the caller wrote it
    rather than normalized to UTC first. That is load-bearing twice over:
    normalizing a naive value needs a zone to resolve it against, and resolving
    it against the host's clock is the bug 1.15.0 removed; and Graph requires an
    all-day event to start on a midnight boundary, which a UTC conversion of
    local midnight is not.

    A relative ``start``/``end`` (``+2h``, ``now``) is the exception (#74). It
    names a moment, so it is resolved before anything is sent, and then written
    as wall-clock time in the anchor zone (``_in_zone``), the form Graph keeps
    the zone for. Sent as typed, Graph refused it with a 400. The result
    carries each end that was relative as written (``start``/``end``, in the
    read tools' ``"<local> (<zone>)"`` form), so ``7d`` shows it is in the past.
    A relative time on an all-day event is refused, since Graph needs midnight.

    ``show_as`` is Graph's ``showAs`` — the free/busy status Outlook labels
    "Show as". Omitted, Graph applies its own default (``busy``); we do not
    send one, so the default stays Graph's to change.

    New parameters are appended to this signature rather than inserted:
    ``create_event`` is importable from the published package and nothing after
    ``graph_client`` is keyword-only, so a mid-signature insert would rebind a
    positional argument for an external caller. ``show_as`` therefore sits
    after ``timezone``, and ``server.py``'s positional call has to agree — a
    disagreement there only misfires when a value is actually supplied, so a
    green suite proves nothing about it. ``calendar`` is appended after
    ``show_as`` for the same reason, and ``key`` after ``calendar``.

    ``calendar`` is the one ``outlook_list_events`` reads, resolved the same
    way: ``None``, blank or "primary" keeps ``POST /me/events`` with no extra
    round trip; anything else is a display name or ID matched against
    ``/me/calendars`` by ``calendar_resolver``, and the event is posted to that
    calendar's own collection. A calendar Graph lists as read-only is refused
    there rather than sent for a 403. Nothing else needs to know which calendar
    an event is on: ``/me/events/{id}`` reaches events on any of the user's own
    calendars — checked read-only on four calendars of one mailbox, and for an
    event created here by ``TestCreatingInAnotherCalendar`` — so the tools that
    take an event ID are unchanged. A calendar someone else shared is not
    claimed: Graph documents its events under ``/me/calendars/{id}/events``.

    A calendar that is not the user's own — judged from the same listing, so
    at no extra cost — needs ``mail_send`` under a policy that withholds it:
    its owner sees the event (``SOMEONE_ELSES_CALENDAR``).

    ``key`` is the caller's own reference for the event — a booking number,
    say — stored in a hidden extended property (``EVENT_KEY_PROPERTY_ID``) so
    ``find_events_by_key`` can find the event again when a change to the same
    booking arrives. Nothing here refuses a key another event already has:
    looking first is the caller's step, and a check here could not stop two
    concurrent creates anyway.

    ``is_online`` asks Graph for an online meeting. Some personal accounts
    honour it and some ignore it (#70), so when it is ``True`` the result
    carries ``is_online`` as Graph's reply states it — omitted if the reply
    lacks the field. Only ``True`` is reported: the default ``False`` cannot be
    told apart from an argument nobody passed.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")
    if attendees:
        # Graph emails the subject and body to every attendee.
        check_sends_mail(config, "outlook_create_event", "invite attendees")

    # Validated, and only a relative form resolved; see the docstring.
    moment = _now()
    start_resolved = resolve_relative_datetime(start, now=moment)
    end_resolved = resolve_relative_datetime(end, now=moment)
    # `is None`, not `or`: an explicitly empty timezone is a mistake to refuse,
    # not a synonym for "use the configured one". The two sources are held to
    # different standards on purpose — a stored config value can be legacy, a
    # freshly written argument cannot.
    zone = (
        resolve_config_event_timezone(config.timezone)
        if timezone is None
        else validate_event_timezone(timezone)
    )
    if is_all_day:
        # Graph stores an all-day event anchored in UTC whatever zone it is
        # sent — verified live 2026-09-21, both a midnight `Z` and a naive
        # midnight labelled America/Los_Angeles came back
        # `originalStartTimeZone: UTC`. Sending the zone therefore buys nothing
        # there and costs something here: `00:00Z` labelled Los Angeles is
        # 17:00 the previous day, so the recurrence would be built for the
        # wrong date and weekday. Label it as Graph will store it.
        zone = "UTC"
    # The resolver hands ISO input back as the same object, so identity says
    # which ends were relative.
    placed: dict[str, str] = {}
    if is_all_day and (start_resolved is not start or end_resolved is not end):
        raise ValueError(_ALL_DAY_RELATIVE_REFUSAL)
    if start_resolved is not start:
        start = _in_zone(start_resolved, zone, start)
        placed["start"] = _placed(start, zone)
    if end_resolved is not end:
        end = _in_zone(end_resolved, zone, end)
        placed["end"] = _placed(end, zone)
    # A typed `Z` or offset goes out as local time in the zone (#236). An all-day
    # event is stored in UTC whatever it is sent, so it has no zone to keep.
    converted: dict = {}
    if not is_all_day:
        typed = {"start": (start, start_resolved), "end": (end, end_resolved)}
        ends = {k: (v, zone) for k, (v, r) in typed.items() if r is v}
        sent = _typed_in_zones(ends)
        converted = _converted(ends, sent)
        start, end = sent.get("start", start), sent.get("end", end)

    # Validate attendee emails if provided
    validated_attendees = []
    if attendees:
        validated_attendees = [validate_email(e) for e in attendees]

    # Resolved here rather than at the assignment, so every argument is checked
    # before anything is sent. `update_event` has the same ordering for a
    # sharper reason (see there); these two stay the same shape deliberately.
    resolved_show_as = _free_busy(show_as) if show_as is not None else None
    if key is not None:
        key = validate_event_key(key)

    from msgraph.generated.models.attendee import Attendee
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.event import Event
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.location import Location

    event = Event()
    event.subject = subject

    event.start = DateTimeTimeZone()
    event.start.date_time = start
    event.start.time_zone = zone

    event.end = DateTimeTimeZone()
    event.end.date_time = end
    event.end.time_zone = zone

    event.is_all_day = is_all_day
    event.is_online_meeting = is_online

    if location:
        event.location = Location()
        event.location.display_name = location

    if body:
        event.body = ItemBody()
        event.body.content = body
        event.body.content_type = BodyType.Text

    if validated_attendees:
        event.attendees = []
        for email in validated_attendees:
            att = Attendee()
            att.email_address = EmailAddress()
            att.email_address.address = email
            event.attendees.append(att)

    if recurrence:
        event.recurrence = build_event_recurrence(recurrence, start=start, zone=zone)

    if resolved_show_as is not None:
        event.show_as = resolved_show_as

    if key is not None:
        event.single_value_extended_properties = [_key_property(key)]

    # Last, because it is the only check that asks Graph: every argument above
    # is refused for free, and the calendar listing is paid for only by a call
    # that is going to be sent. Whose calendar it is comes off the same listing.
    calendar_id, someone_elses = await resolve_calendar_for_write(graph_client, calendar)
    if someone_elses:
        check_sends_mail(config, "outlook_create_event", SOMEONE_ELSES_CALENDAR, emails=False)
    events = (
        graph_client.me.events
        if calendar_id is None
        else graph_client.me.calendars.by_calendar_id(calendar_id).events
    )

    response = await call_without_replay(
        lambda once: events.post(event, request_configuration=once),
        outcome_unknown=(
            INVITATIONS_MAY_HAVE_BEEN_SENT if validated_attendees else EVENT_MAY_HAVE_BEEN_CREATED
        ),
    )
    # Only a single event is read back and checked; see `_keep_typed_instants`.
    # A reply without an id leaves nothing to read: the SDK will not build a
    # request for one, and the event exists either way.
    note = await _keep_typed_instants(
        lambda: events.by_event_id(response.id),
        converted if event.recurrence is None and isinstance(response.id, str) else {},
        f"created (id {response.id})",
    )

    return {
        "status": "created",
        "event_id": response.id,
        "subject": response.subject,
        **placed,
        **(_online_reply(response) if is_online else {}),
        **note,
    }


async def update_event(
    graph_client: Any,
    event_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    recurrence: dict | str | None = None,
    remove_recurrence: bool = False,
    attendees: list[str] | None = None,
    is_all_day: bool | None = None,
    show_as: str | None = None,
    timezone: str | None = None,
    key: str | None = None,
    is_online: bool | None = None,
    *,
    config: Config,
) -> dict:
    """Update an existing calendar event.

    Only patches changed fields. ``None`` means "leave alone" for every
    argument, so ``False`` and ``[]`` are real instructions, not absences.

    ``attendees`` **replaces the whole collection** — Graph has no add-one
    operation, so pass the full intended list, and expect Outlook to email
    invitations to everyone on it and cancellations to anyone dropped. ``[]``
    removes them all. This is the one argument here with an outward-facing
    side effect.

    ``is_all_day`` requires ``start`` and ``end`` in the *same* call — Graph
    returns ``ErrorInvalidRequest: Missing parameters: Event.Start`` for a
    lone isAllDay patch — and both must fall on midnight boundaries. We reject
    the incomplete call locally rather than pass through that error.

    A ``start``/``end`` patch keeps the zone the event is already anchored in.
    Graph rejects a ``start`` patch carrying no ``timeZone`` at all
    (``TimeZoneNotSupportedException`` on the empty string), so one has to be
    sent, and the event's own is the only one that does not move it —
    patching a colleague's 09:00 New York meeting must not relocate it to this
    server's zone. That costs one GET, issued only when ``start``/``end`` is
    in play, and the zone comes off ``originalStartTimeZone`` rather than
    ``start.timeZone``: see ``_anchor_zones`` for why the obvious field is the
    wrong one, and why start and end are read separately.

    A consequence worth stating, because nothing else says it: a zone-less
    ``start``/``end`` here is read in the **event's** zone — not UTC as before
    this change, and not ``config.timezone`` as everywhere else in this server.
    Patching a colleague's New York meeting to ``"2026-11-03T09:00:00"`` means
    09:00 in New York, not 09:00 where this server is configured.

    A ``Z`` or offset ``start``/``end`` is sent as wall-clock time in the zone
    that end is anchored in, as in ``create_event`` (#236), when ``zoneinfo``
    or the Windows-name table can convert into it. If either end cannot be
    converted, both are sent as written, as before. On an event that stands
    alone, with no ``recurrence`` in the call, it is read back and checked as on
    create.

    A relative ``start``/``end`` (``+2h``) is written as wall-clock time in the
    zone that end is anchored in, as in ``create_event``, so it is placed only
    once that zone has been read. It is refused on an event Graph names in
    Windows terms ("Pacific Standard Time", as the series this server created
    came back when measured), and inside the interval a zone repeats when
    daylight saving ends; see ``_in_zone``. As on create, the result carries
    each end that was relative as written, and an all-day event refuses one.

    ``timezone`` re-anchors the event into a *different* zone, and requires
    ``start`` and ``end`` in the same call for the reason above: Graph rejects a
    ``start`` patch carrying no ``timeZone``, so the zone is never an
    independent edit. Passing it without them is refused rather than answered
    ``updated``. One zone re-anchors both ends, which is what "move this to
    Eastern" means; an event whose ends genuinely differ keeps them by omitting
    the argument.

    Changing a **series master's** zone needs its recurrence re-sent in the
    same patch. Without it Graph answers ``400 ErrorPropertyValidationFailure``,
    which names neither the zone nor the property; with it the identical patch
    succeeds. Both verified live, with the zone unchanged and a single instance
    as controls. When the caller has not supplied a new recurrence the event's
    existing one is read and sent back, so a series created before events
    carried a zone can be repaired in place rather than deleted and rebuilt.
    Its days move as far as the zone moves the stored start's local date
    (``_follow_new_start``): a Thursday 02:00Z series re-anchored to Los Angeles
    becomes a Wednesday one. A start also moved to another day by the caller
    then follows the same-zone rules below, as a separate step.

    Moving a series master's start to another **day** in the same zone needs
    the same (#98). Without it Graph answers ``updated`` and snaps the master
    back onto its pattern, every occurrence where it was — measured live, timed
    and all-day. So the existing recurrence is re-sent (``move_series``): a
    pattern that already lands on the new day is kept and only the range moves;
    a one-day pattern moves to the new day; several weekdays moved to a day
    they skip are refused, as is a pattern with no exact move. An ``endDate``
    moves as far as the start. Whenever a re-sent rule schedules other days, the
    reply's ``note`` names it (``_series_note``). The day is the start's own in
    the event's zone, so a ``Z``/offset start left as typed cannot move a
    series, nor can a series whose pattern skips its range's first day
    (``_day_move``). A time moved within the day sends no recurrence, and a
    patch carrying one gets no read-back correction (``_keep_typed_instants``).

    Any ``start``, ``end`` or ``recurrence`` patch to a series master that has
    edited or deleted occurrences is refused, naming them. Graph restores every
    one of them when a master's times or pattern change — zone or not — and
    reports success. Measured live; see ``_OCCURRENCE_SELECT``.

    It cannot be combined with an all-day event, and that is refused rather than
    ignored: Graph stores an all-day event anchored in UTC whatever zone it is
    sent, so there is no zone to move it to and accepting the argument would
    report success for nothing.

    ``is_online=True`` asks Graph to add an online meeting; it is idempotent (a
    second request keeps the same join link). Some personal accounts honour it
    and some ignore it (#70), so the result carries ``is_online`` as Graph's
    reply states it, as ``create_event`` does. ``is_online=False`` is refused
    (``_ONLINE_REMOVAL_REFUSAL``): Graph accepts it and keeps the meeting.
    Appended after ``key`` for the reason ``create_event`` gives.

    Setting ``recurrence`` turns a single event into a series, or replaces the
    pattern of an existing one. It takes the same shapes as ``create_event``.
    Graph anchors the range on the series master's start, so when ``start``
    isn't part of the same patch the event's current start is read first. When
    it is, a supplied ``range.startDate`` is dropped and re-derived from it. A
    pattern equal to the stored one moves with a zone change, and a one-day one
    follows a start moved to another day, while several days are sent exactly
    as written — so ``outlook_get_event``'s recurrence can be handed back beside
    a new start and zone.
    Passing ``None`` leaves any existing recurrence untouched — this is a
    partial patch — except that it is carried along when a series' start moves
    to another day, as above.

    ``remove_recurrence=True`` turns a series master back into a single event,
    keeping the first occurrence's time. It is a separate flag rather than a
    sentinel value on ``recurrence`` because ``None`` there already means
    "leave alone", and the two are mutually exclusive.

    It has to go through ``additional_data``: Graph clears a series with an
    explicit ``"recurrence": null``, and ``event.recurrence = None`` does not
    produce one on every kiota this package accepts. The adapter serializes
    through the backing store. Before kiota-serialization-json 1.15 that writes
    a top-level null *beside* the object body and then refuses the mixed
    document with ``ValueError("Invalid Json output")``, an unsendable request;
    from 1.15 it writes the null inside the body under the Python attribute
    name, which for ``recurrence`` happens to be Graph's (#233).
    ``test_setting_event_recurrence_none_is_not_sendable_on_every_supported_kiota``
    pins both. This can be simplified only once the kiota floor is 1.15.

    ``show_as`` patches Graph's ``showAs`` and needs nothing else resent — it is
    not ``is_all_day``. Graph honours all six values on PATCH for a consumer
    mailbox, verified live. Omitting it leaves the event's current status alone;
    there is no value that clears one, because Graph has no such state.

    ``key`` sets or replaces the event's key (see ``create_event``) and needs no
    read: it is neither a time nor a pattern, so the occurrence check and the
    pin for a series reshape do not apply. Whether Graph agrees that keying a
    master leaves its changed occurrences alone is a live-tier question, asked
    by ``test_keying_a_series_keeps_them``.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_update_event")
    if attendees:
        check_sends_mail(config, "outlook_update_event", "invite attendees")
    event_id = validate_graph_id(event_id)

    # Everything that can reject this call without asking Graph anything goes
    # above the first `await`. Before the anchor-zone read existed these all
    # rejected with zero network calls, and hoisting the read above them had
    # two costs: a stale id plus a malformed time surfaced as
    # `404 ErrorItemNotFound` instead of the input error, and every rejected
    # patch paid for a full-event round trip.
    # Validated, and only a relative form resolved, as in `create_event`.
    # Each resolved instant is placed in its zone once that zone has been read,
    # below; the resolver hands ISO input back as the same object.
    moment = _now()
    start_resolved = start if start is None else resolve_relative_datetime(start, now=moment)
    end_resolved = end if end is None else resolve_relative_datetime(end, now=moment)
    if is_all_day is not None and (start is None or end is None):
        raise ValueError(
            "is_all_day requires start and end in the same call; Graph rejects a lone "
            "isAllDay patch with 'Missing parameters: Event.Start'. Both must be "
            "midnight boundaries, e.g. start=2026-10-22T00:00:00, end=2026-10-23T00:00:00"
        )
    if remove_recurrence and recurrence is not None:
        raise ValueError(
            "Pass either recurrence or remove_recurrence, not both — they ask for opposite things"
        )
    if timezone is not None and (start is None or end is None):
        raise ValueError(
            "timezone requires start and end in the same call. It is the zone those "
            "datetimes are read in and the series is anchored to, not a conversion "
            "applied to the times already stored — Graph rejects a start patch that "
            "carries no timeZone, so the two travel together. To re-anchor an event "
            "without moving it, pass its current start and end with the new timezone."
        )
    # Resolved here, not at the assignment below, for the reason the block above
    # exists: the assignment is past `current_event()`, so a bad zone refused
    # there would have cost a full-event GET and surfaced a stale id's 404
    # ahead of the actionable input error.
    if timezone is not None and is_all_day:
        raise ValueError(_ALL_DAY_ZONE_REFUSAL)
    relative = start_resolved is not start or end_resolved is not end
    if relative and is_all_day:
        raise ValueError(_ALL_DAY_RELATIVE_REFUSAL)
    requested_zone = validate_event_timezone(timezone) if timezone is not None else None

    validated_attendees = None
    if attendees is not None:
        validated_attendees = [validate_email(e) for e in attendees]

    # Resolved here, with the other argument checks, rather than at the
    # assignment far below: every read path in this function is lazy, so a
    # `show_as` refused late would still have cost the `current_event()` GET
    # that a recurrence or a time patch triggers. Validating before any await
    # is the rule this function already follows for datetimes and attendees.
    resolved_show_as = _free_busy(show_as) if show_as is not None else None
    # The same for a malformed recurrence, and sharper since the occurrence
    # check: on a series with changed occurrences, a bad shape would otherwise
    # surface as that refusal, two reads later, instead of as the input error.
    if recurrence is not None:
        check_recurrence_shape(recurrence)
    if key is not None:
        key = validate_event_key(key)
    if is_online is False:
        raise ValueError(_ONLINE_REMOVAL_REFUSAL)

    # The event as Graph currently holds it, fetched at most once and only when
    # something actually needs it: the zone a start/end patch must carry, the
    # anchor date for a recurrence sent without a start, or the existing
    # recurrence a zone change has to bring along.
    _current: Any = None

    async def current_event() -> Any:
        nonlocal _current
        if _current is None:
            _current = await graph_client.me.events.by_event_id(event_id).get()
        return _current

    # Any change to an event on someone else's calendar is shown to its owner.
    await refuse_someone_elses_calendar(graph_client, event_id, config, "outlook_update_event")

    # A reworded meeting reaches the other people on it — as the organizer's
    # update, or in an attendee's next response — so rewording one needs
    # `mail_send` as inviting them did. Only a restricted policy pays for the
    # read that finds out. Adding an online meeting is counted with them: it
    # writes the join details into the meeting, which presumably goes out the
    # same way — inferred, not observed, since the live tier has no attendees.
    rewords = not (subject is None and body is None and location is None and is_online is None)
    if mail_send_withheld(config) and rewords:
        if _has_attendees(await current_event()):
            check_sends_mail(
                config,
                "outlook_update_event",
                "change the subject, body, location or online meeting of an event "
                "that has attendees",
            )

    from msgraph.generated.models.attendee import Attendee
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.event import Event
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.location import Location

    event = Event()

    if subject is not None:
        event.subject = subject

    start_zone: str | None = None
    end_zone: str | None = None
    # Set only when the patch reshapes a series master; see the occurrence check
    # below for why that patch gets pinned.
    if_match: str | None = None
    all_day = False
    if start is not None or end is not None:
        existing_event = await current_event()
        if requested_zone is not None:
            # One explicit zone re-anchors both ends, which is what "move this
            # meeting to Eastern" means. An event whose ends genuinely differ
            # keeps them by omitting the argument.
            start_zone = end_zone = requested_zone
        else:
            # Echoed back verbatim, not validated: Graph stores Windows zone
            # names ("Pacific Standard Time") as readily as IANA ones, and a
            # name it gave us is by definition one it accepts.
            start_zone, end_zone = _anchor_zones(existing_event)
        # An all-day event is stored anchored in UTC whatever zone it is sent
        # (see `create_event`), so label it the way Graph will hold it. The
        # stored flag decides when the caller did not, because a patch that
        # leaves `is_all_day` alone still has to agree with what the event is.
        all_day = (
            is_all_day
            if is_all_day is not None
            else bool(getattr(existing_event, "is_all_day", False))
        )
        if all_day:
            if requested_zone is not None:
                # Reached only when the caller left `is_all_day` alone and the
                # stored event turns out to be all-day; the same refusal as
                # above, one round trip later because nothing knew sooner.
                raise ValueError(_ALL_DAY_ZONE_REFUSAL)
            if relative:
                # The stored event is all-day and the caller did not say so;
                # the same refusal as above, one read later.
                raise ValueError(_ALL_DAY_RELATIVE_REFUSAL)
            start_zone = end_zone = "UTC"

    # A relative end is placed in the zone that end is anchored in, which is
    # only known now. Before the occurrence read below, since nothing it finds
    # could rescue a time `_in_zone` refuses. A zone of None (Graph's custom-zone
    # sentinel) is left for `_anchored` to refuse with its own message.
    placed: dict[str, str] = {}
    if start_resolved is not start and start_zone is not None:
        start = _in_zone(start_resolved, start_zone, start)
        placed["start"] = _placed(start, start_zone)
    if end_resolved is not end and end_zone is not None:
        end = _in_zone(end_resolved, end_zone, end)
        placed["end"] = _placed(end, end_zone)
    # A typed `Z` or offset goes out as local time in its zone, as on create
    # (#236). An all-day event is anchored in UTC, so it has nothing to keep.
    converted: dict = {}
    if not all_day:
        typed = {"start": (start, start_resolved, start_zone), "end": (end, end_resolved, end_zone)}
        ends = {k: (v, z) for k, (v, r, z) in typed.items() if v is not None and r is v}
        sent = _typed_in_zones(ends)
        converted = _converted(ends, sent)
        start, end = sent.get("start", start), sent.get("end", end)

    # Last of the checks, because it is the only one that costs a second read,
    # and only a master can have occurrences to lose. `remove_recurrence` is left
    # out on purpose: collapsing the series is what it asks for.
    if (start is not None or end is not None or recurrence is not None) and (
        _is_series_master(await current_event())
    ):
        await _refuse_to_discard_occurrences(graph_client, event_id)
        # Both halves of this patch rest on state read before it: the check just
        # made, and — for a zone change — the recurrence re-sent below, echoed
        # from the first read. Between those reads and the PATCH another client
        # can edit an occurrence (which this patch would then discard) or
        # re-pattern the series (which the echo would revert). Editing or
        # deleting an occurrence moves the master's change key, as does any
        # master edit — verified live — so pinning to the *first* read's key
        # covers everything after it; Graph refuses a stale one with
        # `412 ErrorIrresolvableConflict`, which `_HINT_TABLE` answers. On the
        # SDK's own request builder, so this owes no raw-HTTP path. A missing key
        # degrades to the last-writer-wins every other write has (#87).
        if_match = (getattr(await current_event(), "additional_data", None) or {}).get(
            "@odata.etag"
        )

    def _anchored(zone: str | None, which: str) -> str:
        if zone is None:
            raise ValueError(
                f"Could not read the zone this event's {which} is anchored in, so "
                f"patching its time would move the event. Graph reports no usable "
                f"zone in two cases: an event created against a custom time zone, "
                f"where it returns a sentinel it will not accept back, and one where "
                f"the field is absent entirely. Read the event back with "
                f"outlook_get_event to see which."
            )
        return zone

    if start is not None:
        event.start = DateTimeTimeZone()
        event.start.date_time = start
        event.start.time_zone = _anchored(start_zone, "start")

    if end is not None:
        event.end = DateTimeTimeZone()
        event.end.date_time = end
        event.end.time_zone = _anchored(end_zone, "end")

    if location is not None:
        event.location = Location()
        event.location.display_name = location

    if body is not None:
        event.body = ItemBody()
        event.body.content = body
        event.body.content_type = BodyType.Text

    if validated_attendees is not None:
        event.attendees = []
        for email in validated_attendees:
            att = Attendee()
            att.email_address = EmailAddress()
            att.email_address.address = email
            event.attendees.append(att)

    if is_all_day is not None:
        event.is_all_day = is_all_day

    if remove_recurrence:
        # `event.recurrence = None` does not serialize before kiota 1.15 (see
        # the docstring); additional_data is the explicit JSON null Graph needs
        # on every version.
        event.additional_data = {**(event.additional_data or {}), "recurrence": None}

    # Set when a recurrence the caller did not pass goes out with new days.
    series_note: str | None = None
    if recurrence is not None:
        anchor = start
        anchor_zone = start_zone
        if anchor is None:
            existing = await current_event()
            anchor = getattr(getattr(existing, "start", None), "date_time", None)
            if not anchor:
                raise ValueError(
                    "Could not read the event's current start to anchor the recurrence "
                    "range; pass `start` alongside `recurrence`."
                )
            if anchor_zone is None:
                anchor_zone, _ = _anchor_zones(existing)
            stored_start = getattr(existing, "start", None)
            anchor = _as_instant(anchor, getattr(stored_start, "time_zone", None))
            if anchor_zone and maybe_zone(anchor_zone) is None:
                # Graph named the zone in Windows terms, so the instant above
                # cannot be converted here. Ask Graph for the local wall clock
                # instead of guessing; a naive value needs no conversion, so
                # `event_start_date` then reads the date straight off it.
                anchor = await _start_in_its_own_zone(graph_client, event_id, anchor_zone)
        else:
            # A new start re-derives the range, so a `startDate` handed back from
            # `outlook_get_event` is stale by definition rather than a mistake to
            # refuse — and when the pattern is the series' own, its days move
            # with the start. One place, shared with the zone-change re-send.
            recurrence = _follow_new_start(
                event_recurrence_payload(recurrence, event_start_date(anchor, anchor_zone)),
                await current_event(),
                start=anchor,
                zone=anchor_zone,
            )
        event.recurrence = build_event_recurrence(
            recurrence,
            start=anchor,
            zone=anchor_zone,
            # An explicit `timezone` in the same call names the event's zone, so
            # a `recurrenceTimeZone` echoed back from `outlook_get_event` is the
            # *old* one. Graph refuses the pair when they disagree, which makes
            # the ordinary read-modify-write round trip a 400 rather than an
            # edit. Dropped only here: the general passthrough stays deliberate.
            drop_range_timezone=requested_zone is not None,
        )
    elif start_zone is not None and not remove_recurrence:
        # A start/end patch that changes a series master's zone is refused with
        # `400 ErrorPropertyValidationFailure` unless the recurrence travels
        # with it; the same patch succeeds when it does. Verified live against a
        # consumer mailbox — both directions, with the zone unchanged as the
        # control, and a single instance as the other control. Without this,
        # every series created before zones existed is unrepairable except by
        # deleting it.
        existing = await current_event()
        # The recurrence follows the *start* anchor — that is the one Graph
        # derives `range.recurrenceTimeZone` from — so an end-only patch never
        # needs the recurrence to travel with it.
        stored_zone, _ = _anchor_zones(existing)
        # `stored_zone` is None when Graph names a zone we cannot send back — a
        # custom-zone event reports `tzone://Microsoft/Custom`. That is still a
        # zone *change*, and skipping the re-send for it hands the caller the
        # opaque 400 this branch exists to prevent, so None compares as different
        # rather than being filtered out by a truthiness check.
        #
        # A start moved to another day in the same zone needs it too (#98):
        # without it Graph answers `updated` and leaves every occurrence where
        # it was. Its days follow `move_series`, not a zone change's shift.
        #
        # Either way the patch echoes state read above, so it is pinned; see the
        # occurrence check, which set `if_match` for every master reshape.
        if _is_series_master(existing):
            if stored_zone != start_zone:
                event.recurrence = _resend_recurrence(existing, anchor=start, zone=start_zone)
            elif moved := await _day_move(graph_client, event_id, existing, start, start_zone):
                event.recurrence = _moved_recurrence(
                    existing, start=start, zone=start_zone, old=moved[0], new=moved[1]
                )
            if event.recurrence is not None:
                series_note = _series_note(existing, event.recurrence)

    if resolved_show_as is not None:
        event.show_as = resolved_show_as

    if key is not None:
        event.single_value_extended_properties = [_key_property(key)]

    if is_online is not None:
        event.is_online_meeting = is_online

    # Only a single event is read back and checked; see `_keep_typed_instants`.
    # A converted time means the event was already read for its zone, so telling
    # a single event from a series costs no request.
    checked = (
        converted
        if converted
        and event.recurrence is None
        and "recurrence" not in (event.additional_data or {})
        and _is_single_instance(await current_event())
        else {}
    )

    request_builder = graph_client.me.events.by_event_id(event_id)

    def _patch(once: Any) -> Any:
        if if_match:
            once.headers.try_add("If-Match", if_match)
        return request_builder.patch(event, request_configuration=once)

    # Not replayed after a 5xx either: on an event with attendees the patch is
    # mailed to them as an update, and a pinned replay would only be refused
    # with a 412 for the change key this very patch moved (#131).
    # Where the converted ends stood before this patch, off the read already
    # made, so a read-back that still shows them there is not taken for a move.
    before = _stored_instants(await current_event(), checked) if checked else {}
    response = await call_without_replay(_patch, outcome_unknown=UPDATE_MAY_HAVE_BEEN_SENT)
    note = await _keep_typed_instants(lambda: request_builder, checked, "updated", before)
    notes = [n for n in (series_note, note.get("note")) if n]

    return {
        "status": "updated",
        "event_id": response.id,
        **placed,
        **(_online_reply(response) if is_online is not None else {}),
        **({"note": " ".join(notes)} if notes else {}),
    }


async def delete_event(
    graph_client: Any,
    event_id: str,
    *,
    config: Config,
) -> dict:
    """Delete a calendar event."""
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_delete_event")
    event_id = validate_graph_id(event_id)

    await graph_client.me.events.by_event_id(event_id).delete()

    return {"status": "deleted", "event_id": event_id}


async def rsvp(
    graph_client: Any,
    event_id: str,
    response: str,
    message: str | None = None,
    *,
    config: Config,
) -> dict:
    """RSVP to a calendar event.

    response must be one of: accept, decline, tentative.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_rsvp")
    if message:
        # The comment is mailed to the organizer; the bare response is not text.
        check_sends_mail(config, "outlook_rsvp", "send a message with the response")
    event_id = validate_graph_id(event_id)

    event_builder = graph_client.me.events.by_event_id(event_id)

    if response == "accept":
        from msgraph.generated.users.item.events.item.accept.accept_post_request_body import (
            AcceptPostRequestBody,
        )

        request_body = AcceptPostRequestBody()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await call_without_replay(
            lambda once: event_builder.accept.post(request_body, request_configuration=once),
            outcome_unknown=RESPONSE_MAY_HAVE_BEEN_SENT,
        )
        return {"status": "accepted", "event_id": event_id}

    elif response == "decline":
        from msgraph.generated.users.item.events.item.decline.decline_post_request_body import (
            DeclinePostRequestBody,
        )

        request_body = DeclinePostRequestBody()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await call_without_replay(
            lambda once: event_builder.decline.post(request_body, request_configuration=once),
            outcome_unknown=RESPONSE_MAY_HAVE_BEEN_SENT,
        )
        return {"status": "declined", "event_id": event_id}

    elif response == "tentative":
        from msgraph.generated.users.item.events.item.tentatively_accept import (  # noqa: E501
            tentatively_accept_post_request_body,
        )

        cls = tentatively_accept_post_request_body.TentativelyAcceptPostRequestBody
        request_body = cls()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await call_without_replay(
            lambda once: event_builder.tentatively_accept.post(
                request_body, request_configuration=once
            ),
            outcome_unknown=RESPONSE_MAY_HAVE_BEEN_SENT,
        )
        return {"status": "tentativelyAccepted", "event_id": event_id}

    else:
        raise ValueError(f"response must be accept/decline/tentative; got {response}")
