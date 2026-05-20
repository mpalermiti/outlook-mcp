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


_STATUS_ENUM_TO_STR = {
    "Disabled": "disabled",
    "AlwaysEnabled": "always",
    "Scheduled": "scheduled",
}

_AUDIENCE_ENUM_TO_STR = {
    "None_": "none",
    "ContactsOnly": "contacts_only",
    "All": "all",
}


def _dttz_to_utc_iso(dttz: Any) -> str:
    """Convert a Graph DateTimeTimeZone object to a UTC ISO 8601 string.

    Tries `zoneinfo.ZoneInfo` for the named zone. If the name is unknown
    (e.g. a Windows display name like "Pacific Standard Time" that the
    runtime doesn't map), falls back to emitting the raw timestamp with a
    "+00:00" suffix — callers should treat that as best-effort UTC-ish.
    """
    if dttz is None:
        return ""
    date_time = getattr(dttz, "date_time", None)
    time_zone = getattr(dttz, "time_zone", None) or "UTC"
    if not date_time:
        return ""

    # Graph emits 7-digit fractional seconds (e.g. .0000000). Python's
    # fromisoformat (3.11+) accepts arbitrary fractional precision but
    # strip just in case by truncating to 6 digits.
    s = date_time
    if "." in s:
        head, _, tail = s.partition(".")
        digits = ""
        rest = tail
        for ch in tail:
            if ch.isdigit():
                digits += ch
            else:
                rest = tail[len(digits) :]
                break
        else:
            rest = ""
        digits = digits[:6]
        s = f"{head}.{digits}{rest}" if digits else f"{head}{rest}"

    try:
        naive = datetime.fromisoformat(s)
    except ValueError:
        return f"{date_time}+00:00"

    try:
        tz = ZoneInfo(time_zone)
    except (ZoneInfoNotFoundError, ValueError):
        return f"{date_time}+00:00"

    aware = naive.replace(tzinfo=tz)
    utc_zone = ZoneInfo("UTC")
    return aware.astimezone(utc_zone).strftime("%Y-%m-%dT%H:%M:%SZ")


async def get_auto_reply(graph_client: Any) -> dict:
    """Get the user's auto-reply (out-of-office) configuration.

    Reads /me/mailboxSettings and normalizes the embedded
    automaticRepliesSetting into stable string values. Reply-message bodies
    are passed through `sanitize_output(multiline=True)` to scrub control
    chars/ANSI while preserving newlines.
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

    status_enum = getattr(ar, "status", None)
    status_name = getattr(status_enum, "name", None)
    status = _STATUS_ENUM_TO_STR.get(status_name, "disabled") if status_name else "disabled"

    audience_enum = getattr(ar, "external_audience", None)
    audience_name = getattr(audience_enum, "name", None)
    audience = _AUDIENCE_ENUM_TO_STR.get(audience_name, "all") if audience_name else "all"

    internal = getattr(ar, "internal_reply_message", None) or ""
    external = getattr(ar, "external_reply_message", None) or ""

    return {
        "status": status,
        "internal_message": sanitize_output(internal, multiline=True),
        "external_message": sanitize_output(external, multiline=True),
        "external_audience": audience,
        "scheduled_start": _dttz_to_utc_iso(getattr(ar, "scheduled_start_date_time", None)),
        "scheduled_end": _dttz_to_utc_iso(getattr(ar, "scheduled_end_date_time", None)),
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
    from msgraph.generated.models.automatic_replies_status import AutomaticRepliesStatus
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.external_audience_scope import ExternalAudienceScope
    from msgraph.generated.models.mailbox_settings import MailboxSettings

    status_map = {
        "disabled": AutomaticRepliesStatus.Disabled,
        "always": AutomaticRepliesStatus.AlwaysEnabled,
        "scheduled": AutomaticRepliesStatus.Scheduled,
    }
    audience_map = {
        "none": ExternalAudienceScope.None_,
        "contacts_only": ExternalAudienceScope.ContactsOnly,
        "all": ExternalAudienceScope.All,
    }

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
