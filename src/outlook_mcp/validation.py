"""Input validation — patterns ported from olkcli (MIT)."""

import logging
import re
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError, available_timezones

logger = logging.getLogger(__name__)

# Where a bad `config.timezone` is fixed. Kept next to the logger so the
# read path and the write path quote the same sentence.
CONFIG_REMEDY_TEXT = "Set `timezone` in ~/.outlook-mcp/config.json."


@lru_cache(maxsize=1)
def _known_zone_keys() -> frozenset[str]:
    """Every zone key this host's database holds, built once.

    Membership here, rather than "does ``ZoneInfo`` raise", because the two
    disagree by platform: macOS resolves names against a case-insensitive
    ``/usr/share/zoneinfo``, so ``ZoneInfo("america/los_angeles")`` succeeds
    there and the name goes to Graph exactly as written, while a Linux
    tzdata-only install refuses the same string. A validator built on the
    exception therefore passes on a contributor's laptop and rejects in
    production. ``available_timezones()`` is the same set everywhere.

    Cached because it walks the database — ~50 ms per call, and
    ``config.timezone`` is resolved on every calendar read.
    """
    return frozenset(available_timezones())


# Process-local latch so the legacy-zone substitution warns at most once per
# run rather than on every event created.
_warned_legacy_config_zone: set[str] = set()

# Graph API entity ID pattern: alphanumeric, =, +, /, -
_GRAPH_ID_RE = re.compile(r"^[a-zA-Z0-9_=+/\-]{1,1024}$")

# Email pattern (simplified but sufficient for validation)
_EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

# Phone pattern
_PHONE_RE = re.compile(r"^[0-9 ()+.\-]{1,30}$")

# KQL characters to strip. The value is embedded as `$search="<value>"`, so the
# only real boundary is the string literal itself. Verified live against Graph:
#   "  MUST strip — an embedded quote makes Graph SILENTLY DISCARD $search and
#      return the whole mailbox (200, no error). This is the actual vector.
#   \  MUST strip — a real escape metachar. `\s` is a 400 (unrecognized escape)
#      and a trailing `\` escapes our closing quote (400, unterminated literal).
#   *  stripped — buys nothing (Graph already prefix-matches: `subject:Amaz` and
#      `subject:Amaz*` both return 50), and a bare `*` would become a silent
#      whole-mailbox read where it is currently a loud 400.
#   &|! stripped — not operators. Graph's operators are the uppercase words
#      AND/OR/NOT; the symbol forms silently zero out an otherwise-matching query.
# `:` `(` `)` are deliberately NOT stripped — they are the property-restriction and
# grouping syntax the tools document, and stripping `:` silently turned every
# documented query into a zero-result free-text phrase.
_KQL_DANGEROUS = re.compile(r'["&|!*\\]')

# Relative datetime offsets: `7d`, `-7d`, `+2h`, `30m`, `2w`, `now`.
# Bare and `-` mean before now; `+` means after. Bounded to 5 digits so a
# nonsense magnitude is a validation error rather than a datetime overflow.
_RELATIVE_RE = re.compile(r"^([+-]?)(\d{1,5})([mhdw])$", re.IGNORECASE)
_RELATIVE_UNITS = {"m": "minutes", "h": "hours", "d": "days", "w": "weeks"}

# Control characters (C0 + C1 + DEL), excluding \n and \t
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
# ANSI escape sequences
_ANSI_ESCAPE = re.compile(r"\x1b\[[0-9;]*[a-zA-Z]")

# ISO 8601 datetime: strict pattern to reject injection attempts
_ISO_DATETIME_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}"           # date
    r"(?:T\d{2}:\d{2}:\d{2}"        # optional time
    r"(?:\.\d+)?"                    # optional fractional seconds
    r"(?:Z|[+-]\d{2}:\d{2})?"       # optional timezone
    r")?$"
)

# Time zone abbreviations an agent reaches for when a user says "3pm Pacific".
# None of these resolve as IANA keys, so they land in `resolve_timezone`'s error
# path; naming them there turns "not a zone name the database contains" into the
# sentence that actually fixes the call.
#
# The invariant is "only abbreviations zoneinfo *cannot* resolve", and it is
# guarded in both directions by `test_the_abbreviation_table_holds_only_
# unresolvable_names`. It was false when this shipped: `CET`, `EET` and `WET`
# all resolve, so the refusal branch was dead for them and they were handed
# to Graph unchecked. They live in the table below now.
_TIME_ZONE_ABBREVIATIONS = frozenset(
    {
        "PT", "PST", "PDT", "MT", "MDT", "CT", "CST", "CDT", "ET", "EDT",
        "AKST", "AKDT", "ADT", "NST", "NDT", "BST", "CEST",
        "EEST", "WEST", "AEST", "AEDT", "ACST", "ACDT", "AWST",
        "NZST", "NZDT", "SGT", "PHT", "KST", "JST", "IST", "ICT",
    }
)

# Zone names the IANA database *does* resolve and Graph rejects with
# `400 TimeZoneNotSupportedException` — verified live 2026-09-21 (EST/MST/HST)
# and 2026-09-23 (CET/EET/WET), alongside `America/Los_Angeles`, `Pacific
# Standard Time`, `GMT`, `US/Pacific`, `Etc/UTC` and `UTC`, which are all
# accepted. Each maps to the zone a caller reaching for it meant.
_ZONES_GRAPH_REFUSES = {
    "EST": "America/New_York",
    "MST": "America/Denver",
    "HST": "Pacific/Honolulu",
    "CET": "Europe/Paris",
    "EET": "Europe/Athens",
    "WET": "Europe/Lisbon",
}

# Of those, the ones that are *also* fixed offsets which never observe daylight
# saving — so they would be the wrong anchor for a DST-crossing series even if
# Graph took them. `CET`/`EET`/`WET` are deliberately not here: tzdata gives
# them CEST/EEST/WEST, verified 2026-09-23 (Jan +1:00, Jul +2:00 for CET), so
# saying that of them would be false and the reason has to stay true per entry.
_FIXED_OFFSET_ZONES = frozenset({"EST", "MST", "HST"})

WELL_KNOWN_FOLDERS = {
    "inbox", "drafts", "sentitems", "deleteditems",
    "junkemail", "archive", "outbox",
}


def validate_graph_id(value: str) -> str:
    """Validate a Microsoft Graph entity ID."""
    if not value:
        raise ValueError("Graph ID must not be empty")
    if len(value) > 1024:
        raise ValueError("Graph ID too long (max 1024 chars)")
    if not _GRAPH_ID_RE.match(value):
        raise ValueError(f"Graph ID contains invalid characters: {value[:50]}")
    return value


def validate_email(value: str) -> str:
    """Validate an email address."""
    if not _EMAIL_RE.match(value):
        raise ValueError(f"Invalid email address: {value[:50]}")
    return value


def has_time_zone_database() -> bool:
    """Whether *any* IANA database is reachable, asked by resolving a key that
    every database has.

    Importability of ``tzdata`` is a proxy for this, not the thing itself: a
    POSIX host with ``/usr/share/zoneinfo`` and no ``tzdata`` installed has a
    database, and would have been told its install was broken.
    """
    try:
        ZoneInfo("UTC")
    except (ZoneInfoNotFoundError, ValueError, OSError):
        return False
    return True


def resolve_timezone(name: str, remedy: str | None = None) -> ZoneInfo:
    """Return the ``ZoneInfo`` for ``name``, or say why it would not load.

    Shared by the calendar read path (where ``name`` is ``config.timezone``)
    and the calendar write path (where it is that or a caller-supplied
    ``timezone`` argument). An unhandled ``ZoneInfoNotFoundError`` reaches the
    model as a message-free ``Error executing tool …`` — ``_wrap_tool_errors``
    keeps the text of an *unexpected* exception on the server, which is right
    for a crash and wrong for a bad zone name. Raising ``ValueError`` routes it
    through ``ToolInputError`` instead, so the message survives the trip.

    The three ways the lookup fails need different fixes, so they get different
    messages: an abbreviation is the mistake an agent actually makes and the
    one Graph answers with an opaque ``TimeZoneNotSupportedException``; a name
    no database contains is a typo; no database at all is a broken install —
    ``tzdata`` is a dependency exactly so that Windows and slim Linux images
    have one.

    ``ValueError`` is caught alongside ``ZoneInfoNotFoundError`` because
    zoneinfo raises it, not the subclass, for a path-shaped key:
    ``/etc/localtime`` is a plausible thing to put in a config file and would
    otherwise escape both the truncation and the hint. ``OSError`` joins them
    because ``ZoneInfo`` resolves a name against the filesystem where one
    exists, so a name too long to be a path component fails there instead —
    ``'x' * 256`` raises ``OSError [Errno 63] File name too long`` on macOS
    while raising ``ZoneInfoNotFoundError`` on a Windows tzdata-only install.
    That platform split is why it went unnoticed: the escape is invisible on
    the machine this was written on, and reaches the model as a message-free
    crash on the maintainer's.

    ``remedy`` is appended to the not-a-zone message, and the caller supplies
    it because only the caller knows where the value came from. Naming
    ``config.json`` unconditionally told a *model* to go and edit the
    operator's config file to fix its own tool argument.
    """
    try:
        if name not in _known_zone_keys():
            raise ZoneInfoNotFoundError(f"No time zone found with key {name}")
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError, OSError) as exc:
        if not has_time_zone_database():
            raise ValueError(
                f"Invalid timezone: {name[:50]} — this host has no IANA time zone "
                "database, so no zone name would resolve. Reinstall "
                "outlook-graph-mcp to pick up its `tzdata` dependency."
            ) from exc
        if name.strip().upper() in _TIME_ZONE_ABBREVIATIONS:
            raise ValueError(
                f"Invalid timezone: {name[:50]} — that is a time zone "
                "abbreviation, not a zone name. Use the IANA name for the "
                "place, which carries the daylight-saving rules with it: "
                "America/Los_Angeles, America/New_York, Europe/London."
            ) from exc
        raise ValueError(
            f"Invalid timezone: {name[:50]} — not a zone name the IANA database "
            f"contains. Use an IANA name like America/Los_Angeles or UTC."
            + (f" {remedy}" if remedy else "")
        ) from exc


def validate_event_timezone(name: str) -> str:
    """Validate a zone name for the ``timeZone`` field of a Graph event.

    Stricter than :func:`resolve_timezone`, and deliberately a separate
    function rather than a stricter mode of it. ``resolve_timezone`` also
    answers for ``config.timezone``, which is loaded at startup: tightening it
    would turn a working install into one that dies on the next upgrade, and
    the rule there is accept-and-warn for legacy config, hard-error only where
    no stored config can carry the mistake. A ``timezone`` tool argument is
    written fresh on every call, so it can be held to the higher standard.

    What the extra standard buys: ``EST`` resolves perfectly well in Python and
    Graph answers it with ``400 TimeZoneNotSupportedException`` — a network
    round trip to learn what we already knew, reported in Graph's words rather
    than ours. Returns the name unchanged so callers rebind rather than
    validate-and-discard.
    """
    key = name.strip() if isinstance(name, str) else name
    if not key:
        # An empty string is a caller who meant something and sent nothing.
        # Falling back to the configured zone would anchor the event somewhere
        # they never named and report success, while a whitespace-only value
        # one keystroke away is refused — so refuse both, the same way.
        raise ValueError(
            "timezone is empty. Omit it to anchor the event in the server's "
            "configured zone, or pass an IANA zone name like America/Los_Angeles."
        )
    replacement = _ZONES_GRAPH_REFUSES.get(key.upper())
    if replacement:
        # The second clause is true of EST/MST/HST and false of CET/EET/WET,
        # which do observe daylight saving — so it is stated per entry rather
        # than of the table, and the table's own comment says why.
        also_fixed = (
            " and it is a fixed-offset zone that would not follow daylight saving "
            "even if it did"
            if key.upper() in _FIXED_OFFSET_ZONES
            else ""
        )
        raise ValueError(
            f"Invalid timezone: {key[:50]} — Microsoft Graph rejects it{also_fixed}. "
            f"Use {replacement}."
        )
    resolve_timezone(key, remedy="Pass it as the `timezone` argument.")
    return key


def resolve_config_event_timezone(name: str) -> str:
    """The zone to anchor an event in when the caller named none.

    Same standard as :func:`validate_event_timezone` with one exception, and
    the exception is the point: a *stored* ``config.timezone`` of ``EST`` came
    from an install that worked. Before events carried a real zone at all, the
    value was never sent — every event went out labelled ``UTC`` — so nothing
    ever rejected it, and a config file on disk today can perfectly well hold
    one. Refusing it here would leave that server running, reading calendars
    happily, and failing every ``outlook_create_event`` that does not pass an
    explicit ``timezone``. An upgrade must not do that.

    So such a value falls back to ``UTC`` — which is *exactly* what every event
    this server wrote was anchored in before 1.23.0 — and says so, once per run
    per name, naming the config key and the one-line change that earns the fix.
    Nothing about that install gets worse; it simply does not get better until
    someone edits the config.

    Translating ``EST`` to ``America/New_York`` was the obvious alternative and
    is wrong: they are not the same zone. ``EST`` is a fixed UTC−05:00 with no
    daylight saving, which is how ``resolve_timezone`` — and therefore every
    calendar *read* — already interprets that config value. Anchoring writes in
    a DST-observing zone would make the two halves of the server disagree about
    the same string every summer, and would be this function inventing a
    meaning the operator never asked for. The same argument is why a *typo* is
    refused rather than guessed at; it applies here too.

    An explicit ``timezone`` argument gets no tolerance at all — it is written
    fresh on every call and cannot be a legacy anything.
    """
    key = name.strip() if isinstance(name, str) else name
    if key and str(key).upper() in _ZONES_GRAPH_REFUSES:
        better = _ZONES_GRAPH_REFUSES[str(key).upper()]
        if key not in _warned_legacy_config_zone:
            _warned_legacy_config_zone.add(key)
            # The parenthetical is true of EST/MST/HST and false of CET/EET/WET,
            # so it is chosen per entry — the same split `validate_event_timezone`
            # makes. Telling an operator that CET "never observes daylight saving"
            # would be false diagnostic information in the one message they get.
            why = (
                f" (not {key!r}, which is a fixed offset that never observes "
                f"daylight saving)"
                if str(key).upper() in _FIXED_OFFSET_ZONES
                else ""
            )
            logger.warning(
                "config.timezone is %r, which Microsoft Graph rejects for calendar "
                "events. Anchoring them in UTC, as this server did before it sent a "
                "zone at all — so recurring events will still shift an hour across a "
                "daylight-saving change. Set `timezone` in ~/.outlook-mcp/config.json "
                "to %s%s to fix that.",
                key,
                better,
                why,
            )
        return "UTC"
    # Not `validate_event_timezone`: its remedy names the `timezone` argument,
    # and this value came from the config file.
    resolve_timezone(name, remedy=CONFIG_REMEDY_TEXT)
    return name.strip()


def validate_datetime(value: str, tz: str = "UTC", *, now: datetime | None = None) -> str:
    """Validate and re-serialize a datetime string to UTC.

    Accepts ISO 8601, and relative offsets against the current instant:
    ``7d`` / ``-7d`` are seven days *ago*, ``+7d`` is seven days from now, and
    ``now`` is this moment. Units are ``m`` / ``h`` / ``d`` / ``w``, case
    insensitive. Bare-means-ago follows the convention every other CLI uses and
    the one agents reach for unprompted — but it is an asymmetry, so it is stated
    in every docstring that takes a date.

    Rejects injection attempts. Returns a safe ISO 8601 string suitable for
    OData filters.

    `tz` is the IANA zone that *zone-less* input is interpreted in — a naive
    datetime ("2026-10-22T12:30:00") or a bare date ("2026-10-22"), which the
    caller has told us nothing about. Input that already carries a `Z` or an
    offset pins its own instant and is unaffected.

    Callers with a Config pass `config.timezone`; the default keeps zone-less
    input meaning UTC. It must never be resolved against the host clock: that
    made an identical query mean different instants on a UTC container and on
    a laptop in California, silently skewing every date filter and the send
    time of every scheduled message.
    """
    reference = now or datetime.now(timezone.utc)

    if value.strip().lower() == "now":
        return reference.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    relative = _RELATIVE_RE.match(value.strip())
    if relative:
        sign, magnitude, unit = relative.groups()
        offset = timedelta(**{_RELATIVE_UNITS[unit.lower()]: int(magnitude)})
        # Bare and `-` look backwards; only `+` looks forward.
        resolved = reference + offset if sign == "+" else reference - offset
        return resolved.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    # First pass: reject anything that doesn't look like a date
    if not _ISO_DATETIME_RE.match(value):
        raise ValueError(
            f"Invalid datetime format: {value[:50]}. Use ISO 8601 "
            "(2026-10-22 or 2026-10-22T14:30:00Z), or a relative offset — "
            "7d for seven days ago, +7d for seven days from now, `now` for "
            "this moment (units: m, h, d, w)."
        )

    # `tz` is `config.timezone` on every caller, so the remedy is the config file.
    zone = resolve_timezone(tz, remedy=CONFIG_REMEDY_TEXT)

    # Second pass: actually parse it to ensure validity
    try:
        if "T" in value:
            # Full datetime
            if value.endswith("Z"):
                dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
            else:
                dt = datetime.fromisoformat(value)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=zone)
            # Re-serialize to UTC ISO 8601
            utc_dt = dt.astimezone(timezone.utc)
            return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            # Date only — midnight in `tz`
            dt = datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=zone)
            return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    except (ValueError, OverflowError) as e:
        raise ValueError(f"Invalid datetime: {value[:50]}") from e


def sanitize_kql(query: str) -> str:
    """Sanitize a KQL search query to prevent injection.

    The value is embedded as ``$search="<value>"``; the surrounding quotes are
    load-bearing (an unquoted ``:`` is a Graph 400), so the strip only has to
    protect the string-literal boundary. See ``_KQL_DANGEROUS``.
    """
    sanitized = _KQL_DANGEROUS.sub("", query)
    if not sanitized.strip():
        # Would send $search="" — Graph answers with an opaque BadRequest.
        raise ValueError(
            f"Search query is empty after sanitization: {query[:50]!r}. "
            'Quotes, backslashes and the characters & | ! * are removed; '
            "use KQL syntax like subject:budget or from:sarah@acme.com."
        )
    return f'"{sanitized}"'


def validate_folder_name(name: str) -> str:
    """Validate a folder name — well-known names or Graph IDs."""
    lower = name.lower()
    if lower in WELL_KNOWN_FOLDERS:
        return lower
    return validate_graph_id(name)


def validate_phone(value: str) -> str:
    """Validate a phone number."""
    if not _PHONE_RE.match(value):
        raise ValueError(f"Invalid phone number: {value[:30]}")
    return value


def sanitize_output(text: str, multiline: bool = False) -> str:
    """Strip control characters and ANSI escapes from output text."""
    text = _ANSI_ESCAPE.sub("", text)
    text = _CONTROL_CHARS.sub("", text)
    if not multiline:
        text = text.replace("\n", " ").replace("\t", " ")
    else:
        text = text.replace("\t", " ")
    return text
