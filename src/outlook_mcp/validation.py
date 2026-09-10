"""Input validation — patterns ported from olkcli (MIT)."""

import re
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

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

    try:
        zone = ZoneInfo(tz)
    except (ZoneInfoNotFoundError, ValueError) as e:
        raise ValueError(f"Invalid timezone: {tz[:50]}") from e

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
