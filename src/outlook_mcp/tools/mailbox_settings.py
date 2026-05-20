"""Mailbox settings tools: timezone get/set, auto-reply (OOF) get/set."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_MAILBOX_SETTINGS, check_permission
from outlook_mcp.validation import sanitize_output, validate_datetime

_CONTROL_CHARS_BODY = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


# ── Timezone ──────────────────────────────────────────


async def get_timezone(graph_client: Any) -> dict:
    """Get the server-side mailbox timezone from /me/mailboxSettings.

    Returns the `timeZone` field of the user's MailboxSettings — this is the
    timezone Exchange uses for calendar items and OOF/auto-reply scheduling,
    and is distinct from the local `config.timezone` used by this server for
    relative-date math on user input.

    Returns an empty string when the mailbox has no timezone configured.
    """
    settings = await graph_client.me.mailbox_settings.get()
    time_zone = getattr(settings, "time_zone", None) if settings else None
    return {"timezone": time_zone or ""}


async def set_timezone(
    graph_client: Any,
    timezone: str,
    *,
    config: Config,
) -> dict:
    """Set the server-side mailbox timezone via PATCH /me/mailboxSettings.

    Accepts both IANA names (e.g. "America/Los_Angeles") and Windows display
    names (e.g. "Pacific Standard Time"); unknown values are rejected by
    Microsoft Graph, not by a local allowlist.
    """
    check_permission(config, CATEGORY_MAILBOX_SETTINGS, "outlook_set_timezone")

    if not timezone or not timezone.strip():
        raise ValueError("timezone must not be empty")
    if re.search(r"[\x00-\x1f\x7f]", timezone):
        raise ValueError("timezone contains control characters")
    timezone = timezone.strip()
    # (do NOT call sanitize_output — Graph will reject unknown zones)

    from msgraph.generated.models.mailbox_settings import MailboxSettings

    body = MailboxSettings()
    body.time_zone = timezone

    response = await graph_client.me.mailbox_settings.patch(body)
    echoed = getattr(response, "time_zone", None) if response else None
    return {"status": "updated", "timezone": echoed or timezone}


# ── Auto-reply (OOF) ─────────────────────────────────


def _status_pairs():
    """Single source of truth for (string, enum) status pairs.

    Lazy import keeps `msgraph` out of the module-import path so the rest of
    the file (and its tests) stay fast to import.
    """
    from msgraph.generated.models.automatic_replies_status import AutomaticRepliesStatus

    return (
        ("disabled", AutomaticRepliesStatus.Disabled),
        ("always", AutomaticRepliesStatus.AlwaysEnabled),
        ("scheduled", AutomaticRepliesStatus.Scheduled),
    )


def _audience_pairs():
    """Single source of truth for (string, enum) audience pairs."""
    from msgraph.generated.models.external_audience_scope import ExternalAudienceScope

    return (
        ("none", ExternalAudienceScope.None_),
        ("contacts_only", ExternalAudienceScope.ContactsOnly),
        ("all", ExternalAudienceScope.All),
    )


# CLDR windowsZones canonical mapping (subset). Outlook.com mailboxes
# default to Windows display names; zoneinfo only ships IANA, so we
# translate the common ones before falling back.
_WINDOWS_TO_IANA = {
    "Dateline Standard Time": "Etc/GMT+12",
    "UTC-11": "Etc/GMT+11",
    "Aleutian Standard Time": "America/Adak",
    "Hawaiian Standard Time": "Pacific/Honolulu",
    "Marquesas Standard Time": "Pacific/Marquesas",
    "Alaskan Standard Time": "America/Anchorage",
    "UTC-09": "Etc/GMT+9",
    "Pacific Standard Time (Mexico)": "America/Tijuana",
    "UTC-08": "Etc/GMT+8",
    "Pacific Standard Time": "America/Los_Angeles",
    "US Mountain Standard Time": "America/Phoenix",
    "Mountain Standard Time (Mexico)": "America/Chihuahua",
    "Mountain Standard Time": "America/Denver",
    "Central America Standard Time": "America/Guatemala",
    "Central Standard Time": "America/Chicago",
    "Easter Island Standard Time": "Pacific/Easter",
    "Central Standard Time (Mexico)": "America/Mexico_City",
    "Canada Central Standard Time": "America/Regina",
    "SA Pacific Standard Time": "America/Bogota",
    "Eastern Standard Time (Mexico)": "America/Cancun",
    "Eastern Standard Time": "America/New_York",
    "Haiti Standard Time": "America/Port-au-Prince",
    "Cuba Standard Time": "America/Havana",
    "US Eastern Standard Time": "America/Indianapolis",
    "Turks And Caicos Standard Time": "America/Grand_Turk",
    "Paraguay Standard Time": "America/Asuncion",
    "Atlantic Standard Time": "America/Halifax",
    "Venezuela Standard Time": "America/Caracas",
    "Central Brazilian Standard Time": "America/Cuiaba",
    "SA Western Standard Time": "America/La_Paz",
    "Pacific SA Standard Time": "America/Santiago",
    "Newfoundland Standard Time": "America/St_Johns",
    "Tocantins Standard Time": "America/Araguaina",
    "E. South America Standard Time": "America/Sao_Paulo",
    "SA Eastern Standard Time": "America/Cayenne",
    "Argentina Standard Time": "America/Buenos_Aires",
    "Greenland Standard Time": "America/Godthab",
    "Montevideo Standard Time": "America/Montevideo",
    "Magallanes Standard Time": "America/Punta_Arenas",
    "Saint Pierre Standard Time": "America/Miquelon",
    "Bahia Standard Time": "America/Bahia",
    "UTC-02": "Etc/GMT+2",
    "Azores Standard Time": "Atlantic/Azores",
    "Cape Verde Standard Time": "Atlantic/Cape_Verde",
    "UTC": "Etc/UTC",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central Europe Standard Time": "Europe/Budapest",
    "Romance Standard Time": "Europe/Paris",
    "Morocco Standard Time": "Africa/Casablanca",
    "Sao Tome Standard Time": "Africa/Sao_Tome",
    "Central European Standard Time": "Europe/Warsaw",
    "W. Central Africa Standard Time": "Africa/Lagos",
    "Jordan Standard Time": "Asia/Amman",
    "GTB Standard Time": "Europe/Bucharest",
    "Middle East Standard Time": "Asia/Beirut",
    "Egypt Standard Time": "Africa/Cairo",
    "E. Europe Standard Time": "Europe/Chisinau",
    "Syria Standard Time": "Asia/Damascus",
    "West Bank Standard Time": "Asia/Hebron",
    "South Africa Standard Time": "Africa/Johannesburg",
    "FLE Standard Time": "Europe/Kiev",
    "Israel Standard Time": "Asia/Jerusalem",
    "South Sudan Standard Time": "Africa/Juba",
    "Kaliningrad Standard Time": "Europe/Kaliningrad",
    "Sudan Standard Time": "Africa/Khartoum",
    "Libya Standard Time": "Africa/Tripoli",
    "Namibia Standard Time": "Africa/Windhoek",
    "Arabic Standard Time": "Asia/Baghdad",
    "Turkey Standard Time": "Europe/Istanbul",
    "Arab Standard Time": "Asia/Riyadh",
    "Belarus Standard Time": "Europe/Minsk",
    "Russian Standard Time": "Europe/Moscow",
    "E. Africa Standard Time": "Africa/Nairobi",
    "Iran Standard Time": "Asia/Tehran",
    "Arabian Standard Time": "Asia/Dubai",
    "Astrakhan Standard Time": "Europe/Astrakhan",
    "Azerbaijan Standard Time": "Asia/Baku",
    "Russia Time Zone 3": "Europe/Samara",
    "Mauritius Standard Time": "Indian/Mauritius",
    "Saratov Standard Time": "Europe/Saratov",
    "Georgian Standard Time": "Asia/Tbilisi",
    "Volgograd Standard Time": "Europe/Volgograd",
    "Caucasus Standard Time": "Asia/Yerevan",
    "Afghanistan Standard Time": "Asia/Kabul",
    "West Asia Standard Time": "Asia/Tashkent",
    "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
    "Pakistan Standard Time": "Asia/Karachi",
    "Qyzylorda Standard Time": "Asia/Qyzylorda",
    "India Standard Time": "Asia/Calcutta",
    "Sri Lanka Standard Time": "Asia/Colombo",
    "Nepal Standard Time": "Asia/Katmandu",
    "Central Asia Standard Time": "Asia/Almaty",
    "Bangladesh Standard Time": "Asia/Dhaka",
    "Omsk Standard Time": "Asia/Omsk",
    "Myanmar Standard Time": "Asia/Rangoon",
    "SE Asia Standard Time": "Asia/Bangkok",
    "Altai Standard Time": "Asia/Barnaul",
    "W. Mongolia Standard Time": "Asia/Hovd",
    "North Asia Standard Time": "Asia/Krasnoyarsk",
    "N. Central Asia Standard Time": "Asia/Novosibirsk",
    "Tomsk Standard Time": "Asia/Tomsk",
    "China Standard Time": "Asia/Shanghai",
    "North Asia East Standard Time": "Asia/Irkutsk",
    "Singapore Standard Time": "Asia/Singapore",
    "W. Australia Standard Time": "Australia/Perth",
    "Taipei Standard Time": "Asia/Taipei",
    "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
    "Aus Central W. Standard Time": "Australia/Eucla",
    "Transbaikal Standard Time": "Asia/Chita",
    "Tokyo Standard Time": "Asia/Tokyo",
    "North Korea Standard Time": "Asia/Pyongyang",
    "Korea Standard Time": "Asia/Seoul",
    "Yakutsk Standard Time": "Asia/Yakutsk",
    "Cen. Australia Standard Time": "Australia/Adelaide",
    "AUS Central Standard Time": "Australia/Darwin",
    "E. Australia Standard Time": "Australia/Brisbane",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "West Pacific Standard Time": "Pacific/Port_Moresby",
    "Tasmania Standard Time": "Australia/Hobart",
    "Vladivostok Standard Time": "Asia/Vladivostok",
    "Lord Howe Standard Time": "Australia/Lord_Howe",
    "Bougainville Standard Time": "Pacific/Bougainville",
    "Russia Time Zone 10": "Asia/Srednekolymsk",
    "Magadan Standard Time": "Asia/Magadan",
    "Norfolk Standard Time": "Pacific/Norfolk",
    "Sakhalin Standard Time": "Asia/Sakhalin",
    "Central Pacific Standard Time": "Pacific/Guadalcanal",
    "Russia Time Zone 11": "Asia/Kamchatka",
    "New Zealand Standard Time": "Pacific/Auckland",
    "UTC+12": "Etc/GMT-12",
    "Fiji Standard Time": "Pacific/Fiji",
    "Kamchatka Standard Time": "Asia/Kamchatka",
    "Chatham Islands Standard Time": "Pacific/Chatham",
    "UTC+13": "Etc/GMT-13",
    "Tonga Standard Time": "Pacific/Tongatapu",
    "Samoa Standard Time": "Pacific/Apia",
    "Line Islands Standard Time": "Pacific/Kiritimati",
}


def _dttz_to_utc_iso(dttz: Any) -> str:
    """Convert a Graph DateTimeTimeZone object to a UTC ISO 8601 string.

    Translates Windows display names (e.g. "Pacific Standard Time") to IANA
    via `_WINDOWS_TO_IANA` before consulting `zoneinfo.ZoneInfo`. Raises
    `ValueError` if neither the raw name nor its translation is recognized —
    we never silently emit a string that looks like UTC but isn't.
    """
    if dttz is None:
        return ""
    date_time = getattr(dttz, "date_time", None)
    time_zone = getattr(dttz, "time_zone", None) or "UTC"
    if not date_time:
        return ""

    # Python 3.10's fromisoformat rejects >6 fractional digits; Graph emits 7.
    s = re.sub(r"\.(\d{1,6})\d*", r".\1", date_time)

    try:
        naive = datetime.fromisoformat(s)
    except ValueError as exc:
        raise ValueError(
            f"Cannot parse datetime {date_time!r} from Microsoft Graph response."
        ) from exc

    iana = _WINDOWS_TO_IANA.get(time_zone, time_zone)
    try:
        tz = ZoneInfo(iana)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        raise ValueError(
            f"Cannot parse timezone {time_zone!r} from Microsoft Graph response. "
            "Neither the raw name nor its Windows-to-IANA translation is recognized "
            "by the local zoneinfo database."
        ) from exc

    aware = naive.replace(tzinfo=tz)
    utc_zone = ZoneInfo("UTC")
    return aware.astimezone(utc_zone).strftime("%Y-%m-%dT%H:%M:%SZ")


def _dttz_to_iso_or_local_marker(dttz: Any) -> str:
    """Read-path wrapper around `_dttz_to_utc_iso`.

    If conversion fails (unknown tz / unparseable datetime), emit an
    explicit non-UTC marker `LOCAL:<datetime> <tz>` so callers can never
    confuse it with a real UTC ISO string.
    """
    if dttz is None:
        return ""
    try:
        return _dttz_to_utc_iso(dttz)
    except ValueError:
        date_time = getattr(dttz, "date_time", None) or ""
        time_zone = getattr(dttz, "time_zone", None) or "UTC"
        if not date_time:
            return ""
        return f"LOCAL:{date_time} {time_zone}"


async def get_auto_reply(graph_client: Any) -> dict:
    """Get the user's auto-reply (out-of-office) configuration.

    Reads /me/mailboxSettings and normalizes the embedded
    automaticRepliesSetting into stable string values. Reply-message bodies
    are passed through `sanitize_output(multiline=True)` to scrub control
    chars/ANSI while preserving newlines.

    Response schema:
        status: "disabled" | "always" | "scheduled"
        internal_message: str
        external_message: str
        external_audience: "none" | "contacts_only" | "all"
        scheduled_start: UTC ISO 8601 string, or "LOCAL:<datetime> <tz>" when
            the timezone can't be translated to UTC.
        scheduled_end: same shape as scheduled_start.
    """
    settings = await graph_client.me.mailbox_settings.get()
    ar = getattr(settings, "automatic_replies_setting", None) if settings else None

    if ar is None:
        return {
            "status": "disabled",
            "internal_message": "",
            "external_message": "",
            "external_audience": "all",
            "scheduled_start": "",
            "scheduled_end": "",
        }

    status_str_by_enum = {enum.name: s for s, enum in _status_pairs()}
    audience_str_by_enum = {enum.name: s for s, enum in _audience_pairs()}

    status_enum = getattr(ar, "status", None)
    status_name = getattr(status_enum, "name", None)
    status = status_str_by_enum.get(status_name, "disabled") if status_name else "disabled"

    audience_enum = getattr(ar, "external_audience", None)
    audience_name = getattr(audience_enum, "name", None)
    audience = audience_str_by_enum.get(audience_name, "all") if audience_name else "all"

    internal = getattr(ar, "internal_reply_message", None) or ""
    external = getattr(ar, "external_reply_message", None) or ""

    return {
        "status": status,
        "internal_message": sanitize_output(internal, multiline=True),
        "external_message": sanitize_output(external, multiline=True),
        "external_audience": audience,
        "scheduled_start": _dttz_to_iso_or_local_marker(
            getattr(ar, "scheduled_start_date_time", None)
        ),
        "scheduled_end": _dttz_to_iso_or_local_marker(getattr(ar, "scheduled_end_date_time", None)),
    }


_VALID_STATUS = ("disabled", "always", "scheduled")
_VALID_AUDIENCE = ("none", "contacts_only", "all")


async def set_auto_reply(
    graph_client: Any,
    status: str,
    internal_message: str = "",
    external_message: str | None = None,
    external_audience: str = "all",
    start: str | None = None,
    end: str | None = None,
    *,
    config: Config,
) -> dict:
    """Set the user's auto-reply (out-of-office) configuration via PATCH.

    Validates inputs strictly: rejects unknown status/audience values, rejects
    control characters in message bodies (other than tab/newline), and
    requires `start`/`end` when `status=="scheduled"`. Datetimes are
    normalized to UTC and sent with `time_zone="UTC"` so Graph doesn't have
    to interpret named zones.
    """
    check_permission(config, CATEGORY_MAILBOX_SETTINGS, "outlook_set_auto_reply")

    if status not in _VALID_STATUS:
        raise ValueError(f"status must be one of {_VALID_STATUS!r}, got {status!r}")
    if external_audience not in _VALID_AUDIENCE:
        raise ValueError(
            f"external_audience must be one of {_VALID_AUDIENCE!r}, got {external_audience!r}"
        )

    if status in ("always", "scheduled"):
        if not internal_message or not internal_message.strip():
            raise ValueError(
                "internal_message must be non-empty when status is 'always' or 'scheduled'"
            )

    if _CONTROL_CHARS_BODY.search(internal_message):
        raise ValueError("internal_message contains control characters")
    if external_message is not None and _CONTROL_CHARS_BODY.search(external_message):
        raise ValueError("external_message contains control characters")

    normalized_start = ""
    normalized_end = ""
    if status == "scheduled":
        if not start:
            raise ValueError("start is required when status is 'scheduled'")
        if not end:
            raise ValueError("end is required when status is 'scheduled'")
        normalized_start = validate_datetime(start)
        normalized_end = validate_datetime(end)

    if external_message is None:
        external_message = internal_message

    from msgraph.generated.models.automatic_replies_setting import AutomaticRepliesSetting
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.mailbox_settings import MailboxSettings

    status_map = dict(_status_pairs())
    audience_map = dict(_audience_pairs())

    ar = AutomaticRepliesSetting()
    ar.status = status_map[status]
    ar.external_audience = audience_map[external_audience]
    ar.internal_reply_message = internal_message
    ar.external_reply_message = external_message
    if status == "scheduled":
        ar.scheduled_start_date_time = DateTimeTimeZone(date_time=normalized_start, time_zone="UTC")
        ar.scheduled_end_date_time = DateTimeTimeZone(date_time=normalized_end, time_zone="UTC")

    body = MailboxSettings()
    body.automatic_replies_setting = ar
    await graph_client.me.mailbox_settings.patch(body)

    return {
        "status": "updated",
        "auto_reply_status": status,
        "external_audience": external_audience,
        "scheduled_start": normalized_start,
        "scheduled_end": normalized_end,
    }
