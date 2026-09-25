"""Calendar write tools: create_event, update_event, delete_event, rsvp."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_CALENDAR_WRITE, check_permission
from outlook_mcp.tools._recurrence import (
    build_event_recurrence,
    maybe_zone,
    serialize_recurrence,
)
from outlook_mcp.validation import (
    resolve_config_event_timezone,
    validate_datetime,
    validate_email,
    validate_event_timezone,
    validate_graph_id,
)


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


def _is_series_master(event: Any) -> bool:
    """Whether Graph considers this event the master of a recurring series.

    Only a master's zone change needs its recurrence re-sent; a single instance
    and an occurrence both patch cleanly without one.
    """
    event_type = getattr(event, "type", None)
    if event_type is None:
        return False
    return (getattr(event_type, "value", None) or str(event_type)) == "seriesMaster"


async def _start_in_its_own_zone(graph_client: Any, event_id: str, zone: str) -> str | None:
    """The event's start as local wall clock in ``zone``, or ``None``.

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

    Returns ``None`` rather than raising if anything about the second read is
    unusable — the caller then falls back to the date as written, which is what
    it did before this existed.
    """
    from kiota_abstractions.base_request_configuration import RequestConfiguration

    request_config = RequestConfiguration()
    request_config.headers.add("Prefer", f'outlook.timezone="{zone}"')
    localized = await graph_client.me.events.by_event_id(event_id).get(
        request_configuration=request_config
    )
    date_time = getattr(getattr(localized, "start", None), "date_time", None)
    return str(date_time) if date_time else None


def _resend_recurrence(event: Any, *, anchor: str | None, zone: str | None) -> Any:
    """The event's own recurrence, rebuilt as a payload safe to send back.

    Not ``event.recurrence`` handed straight over: that model came off the wire
    with its unset fields explicitly ``None``, and the request adapter
    serializes through the backing store, which writes those nulls out — onto
    the *parent*, under their Python names (#63/#64).
    ``serialize_recurrence`` omits what is unset, and the strict builder then
    produces a fresh model carrying only the fields Graph actually sent.

    Two range fields are dropped rather than echoed:

    ``startDate``, so the builder re-derives it from ``anchor`` — a patch that
    moves the series to a different day has to move the range with it, and Graph
    refuses a range beginning on a different day than the first occurrence.

    ``recurrenceTimeZone``, because it is the *old* zone. Sending it back beside
    a new ``start.timeZone`` is a self-contradictory patch, and Graph answers
    the contradiction with the same opaque ``400
    ErrorPropertyValidationFailure`` this function exists to avoid. Omitted,
    Graph derives the range's zone from the event's own — verified live, a
    series created with no ``recurrenceTimeZone`` at all came back carrying the
    event's.
    """
    payload = serialize_recurrence(getattr(event, "recurrence", None))
    if not payload:
        return None

    start = anchor or getattr(getattr(event, "start", None), "date_time", None)
    if not start:
        raise ValueError(
            "Could not read the event's current start to re-anchor its recurrence "
            "for a timezone change; pass `start` and `end` alongside `timezone`."
        )

    stale = {"startDate", "recurrenceTimeZone"}
    payload["range"] = {k: v for k, v in (payload.get("range") or {}).items() if k not in stale}
    return build_event_recurrence(payload, start=start, zone=zone)


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
    ``config.timezone``. It is the zone that governs the series, not a
    conversion applied to ``start``: an offset or ``Z`` in the datetime still
    pins the instant (Graph honours it), while the zone decides what the second
    occurrence of a weekly series does when daylight saving ends. Anchored in
    ``UTC`` — which is what this tool sent unconditionally before 1.23.0 — a
    09:00 series silently becomes 08:00 the week the clocks go back.

    A zone-less ``start``/``end`` ("2026-10-28T09:00:00") is read by Graph as
    wall-clock time *in that zone*, which is what ``config.timezone`` has
    always meant for input interpretation everywhere else in this server.

    The datetime string is passed through as the caller wrote it rather than
    normalized to UTC first. That is load-bearing twice over: normalizing a
    naive value needs a zone to resolve it against, and resolving it against
    the host's clock is the bug 1.15.0 removed; and Graph requires an all-day
    event to start on a midnight boundary, which a UTC conversion of local
    midnight is not.

    ``show_as`` is Graph's ``showAs`` — the free/busy status Outlook labels
    "Show as". Omitted, Graph applies its own default (``busy``); we do not
    send one, so the default stays Graph's to change.

    New parameters are appended to this signature rather than inserted:
    ``create_event`` is importable from the published package and nothing after
    ``graph_client`` is keyword-only, so a mid-signature insert would rebind a
    positional argument for an external caller. ``show_as`` therefore sits
    after ``timezone``, and ``server.py``'s positional call has to agree — a
    disagreement there only misfires when a value is actually supplied, so a
    green suite proves nothing about it.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_create_event")

    # Validate datetime inputs
    validate_datetime(start)
    validate_datetime(end)
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

    # Validate attendee emails if provided
    validated_attendees = []
    if attendees:
        validated_attendees = [validate_email(e) for e in attendees]

    # Resolved here rather than at the assignment, so every argument is checked
    # before anything is sent. `update_event` has the same ordering for a
    # sharper reason (see there); these two stay the same shape deliberately.
    resolved_show_as = _free_busy(show_as) if show_as is not None else None

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

    response = await graph_client.me.events.post(event)

    return {
        "status": "created",
        "event_id": response.id,
        "subject": response.subject,
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

    It cannot be combined with an all-day event, and that is refused rather than
    ignored: Graph stores an all-day event anchored in UTC whatever zone it is
    sent, so there is no zone to move it to and accepting the argument would
    report success for nothing.

    There is deliberately no ``is_online`` here. Graph accepts ``isOnlineMeeting``
    on a personal (consumer) account and silently ignores it — a created or
    patched event comes back ``isOnlineMeeting: False, onlineMeetingProvider:
    'unknown'``. Teams meetings need a work/school account, which this server
    does not target. Adding the parameter would only ship a convincing no-op.

    Setting ``recurrence`` turns a single event into a series, or replaces the
    pattern of an existing one. It takes the same shapes as ``create_event``.
    Graph anchors the range on the series master's start, so when ``start``
    isn't part of the same patch the event's current start is read first.
    Passing ``None`` leaves any existing recurrence untouched — this is a
    partial patch.

    ``remove_recurrence=True`` turns a series master back into a single event,
    keeping the first occurrence's time. It is a separate flag rather than a
    sentinel value on ``recurrence`` because ``None`` there already means
    "leave alone", and the two are mutually exclusive.

    It has to go through ``additional_data``: Graph clears a series with an
    explicit ``"recurrence": null``, and ``event.recurrence = None`` does not
    produce one. The adapter serializes through the backing store, which writes
    a top-level null *beside* the object body rather than inside it, and then
    refuses the mixed document with ``ValueError("Invalid Json output")`` — so
    the plain assignment is not a silent no-op, it is an unsendable request.
    ``test_setting_event_recurrence_none_is_unsendable`` pins that; if kiota
    ever starts emitting a usable top-level null, it fails and this can be
    simplified.

    ``show_as`` patches Graph's ``showAs`` and needs nothing else resent — it is
    not ``is_all_day``. Graph honours all six values on PATCH for a consumer
    mailbox, verified live. Omitting it leaves the event's current status alone;
    there is no value that clears one, because Graph has no such state.
    """
    check_permission(config, CATEGORY_CALENDAR_WRITE, "outlook_update_event")
    event_id = validate_graph_id(event_id)

    # Everything that can reject this call without asking Graph anything goes
    # above the first `await`. Before the anchor-zone read existed these all
    # rejected with zero network calls, and hoisting the read above them had
    # two costs: a stale id plus a malformed time surfaced as
    # `404 ErrorItemNotFound` instead of the input error, and every rejected
    # patch paid for a full-event round trip.
    if start is not None:
        validate_datetime(start)
    if end is not None:
        validate_datetime(end)
    if is_all_day is not None and (start is None or end is None):
        raise ValueError(
            "is_all_day requires start and end in the same call; Graph rejects a lone "
            "isAllDay patch with 'Missing parameters: Event.Start'. Both must be "
            "midnight boundaries, e.g. start=2026-10-22T00:00:00, end=2026-10-23T00:00:00"
        )
    if remove_recurrence and recurrence is not None:
        raise ValueError(
            "Pass either recurrence or remove_recurrence, not both — they ask for "
            "opposite things"
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
        raise ValueError(
            "timezone cannot be applied to an all-day event. Graph stores one anchored in "
            "UTC whatever zone it is sent — verified live — so there is no zone to "
            "re-anchor it to, and accepting this would report success for nothing. "
            "To turn it into a timed event instead, pass is_all_day=False with real "
            "start and end times."
        )
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
                raise ValueError(
                    "timezone cannot be applied to an all-day event. Graph stores one anchored in "
            "UTC whatever zone it is sent — verified live — so there is no zone to "
            "re-anchor it to, and accepting this would report success for nothing. "
            "To turn it into a timed event instead, pass is_all_day=False with real "
            "start and end times."
                )
            start_zone = end_zone = "UTC"

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
        # `event.recurrence = None` does not serialize (see the docstring);
        # additional_data survives as the explicit JSON null Graph needs.
        event.additional_data = {**(event.additional_data or {}), "recurrence": None}

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
                localized = await _start_in_its_own_zone(graph_client, event_id, anchor_zone)
                if localized:
                    anchor = localized
        event.recurrence = build_event_recurrence(recurrence, start=anchor, zone=anchor_zone)
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
        if stored_zone != start_zone and _is_series_master(existing):
            event.recurrence = _resend_recurrence(existing, anchor=start, zone=start_zone)

    if resolved_show_as is not None:
        event.show_as = resolved_show_as

    response = await graph_client.me.events.by_event_id(event_id).patch(event)

    return {
        "status": "updated",
        "event_id": response.id,
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
        await event_builder.accept.post(request_body)
        return {"status": "accepted", "event_id": event_id}

    elif response == "decline":
        from msgraph.generated.users.item.events.item.decline.decline_post_request_body import (
            DeclinePostRequestBody,
        )

        request_body = DeclinePostRequestBody()
        if message:
            request_body.comment = message
        request_body.send_response = True
        await event_builder.decline.post(request_body)
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
        await event_builder.tentatively_accept.post(request_body)
        return {"status": "tentativelyAccepted", "event_id": event_id}

    else:
        raise ValueError(f"response must be accept/decline/tentative; got {response}")
