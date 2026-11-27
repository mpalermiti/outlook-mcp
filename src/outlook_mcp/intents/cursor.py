"""The ``catch_up`` cursor (contracts §7): where the last call left off, held by the agent.

``cursor = "cu1." + base64url_nopad(canonical JSON of the state)``. The state holds
when it was made, a short hash of the account it belongs to, and two Graph delta
links: the Inbox's (started created-only, §7.3) and the primary calendar's
(§7.4). No subjects, names or addresses: nothing anyone else wrote.

It is not signed. It comes back from the agent, which reads mail, so it is
untrusted input: each link must be Graph's own delta endpoint for its resource
(``require_graph_url``) before any token goes to it, and a link that fails is an
error, never a silent reset. An unknown version or another account's cursor
starts again from the window and says so.

An empty link means "not started yet": that source bootstraps on the next call.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import logging
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any
from urllib.parse import urlsplit

from outlook_mcp.errors import ForeignCursorError
from outlook_mcp.tools._delta import require_graph_url

logger = logging.getLogger(__name__)

PREFIX = "cu1."
VERSION = 1
WARN_BYTES = 8 * 1024
MAX_BYTES = 16 * 1024
# Clock skew a cursor's times may have over now before they are refused.
FUTURE_ALLOWANCE = timedelta(minutes=5)
_UTC = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
INVALID = "Invalid catch_up cursor. Call again without cursor to start from a window."

# `understood.cursor`
NONE = "none"
OK = "ok"
RESET_UNRECOGNISED = "reset_unrecognised"
RESET_OTHER_ACCOUNT = "reset_other_account"
TOO_LARGE = "too_large"


def account_tag(username: str | None) -> str:
    """``sha256(casefold(username))``, first 16 hex digits: which account, not who."""
    return hashlib.sha256((username or "").strip().casefold().encode()).hexdigest()[:16]


@dataclass
class MailState:
    link: str  # the Inbox delta's deltaLink, or its nextLink mid-round; "" to bootstrap
    created_only: bool  # started with changeType=created (§7.3); false on P-DL1's fallback
    # The arrival threshold to use instead of `made` (§7.4): set while a round is
    # unfinished, or when a call did not read mail, so nothing between is lost.
    bootstrap_from: str | None


@dataclass
class CalState:
    link: str  # the primary calendar's delta link; "" to bootstrap
    start: str
    end: str
    # As built: changes are kept only when modified after this time, while a round
    # that began as a fresh snapshot (a first read, a roll or a resync) is unfinished.
    since: str | None = None


@dataclass
class CursorState:
    acct: str
    made: str
    mail: MailState
    cal: CalState
    v: int = VERSION

    def as_dict(self) -> dict:
        return {
            "v": self.v,
            "acct": self.acct,
            "made": self.made,
            "mail": {
                "link": self.mail.link,
                "created_only": self.mail.created_only,
                "bootstrap_from": self.mail.bootstrap_from,
            },
            "cal": {
                "link": self.cal.link,
                "start": self.cal.start,
                "end": self.cal.end,
                "since": self.cal.since,
            },
        }


def encode(state: CursorState) -> str | None:
    """The cursor string, or ``None`` when it would pass ``MAX_BYTES`` (§7.1)."""
    canonical = json.dumps(state.as_dict(), sort_keys=True, separators=(",", ":"))
    text = PREFIX + base64.urlsafe_b64encode(canonical.encode()).decode().rstrip("=")
    if len(text) > MAX_BYTES:
        return None
    if len(text) > WARN_BYTES:
        logger.warning("catch_up cursor is %d bytes, over %d", len(text), WARN_BYTES)
    return text


def _decode(cursor: object) -> dict:
    # Longer than any cursor this server issues: refused unread (review nit).
    if not isinstance(cursor, str) or not cursor.startswith(PREFIX) or len(cursor) > MAX_BYTES:
        raise ValueError(INVALID)
    body = cursor[len(PREFIX) :]
    if not body or not re.fullmatch(r"[A-Za-z0-9_-]+", body):
        raise ValueError(INVALID)
    try:
        raw = base64.urlsafe_b64decode(body + "=" * (-len(body) % 4))
        data = json.loads(raw.decode())
    except (binascii.Error, ValueError, UnicodeDecodeError, RecursionError):
        raise ValueError(INVALID) from None
    if not isinstance(data, dict):
        raise ValueError(INVALID)
    return data


def _text(node: Any, key: str, *, optional: bool = False, utc: bool = False) -> Any:
    value = node.get(key) if isinstance(node, dict) else None
    if value is None and optional:
        return None
    if not isinstance(value, str) or (utc and not _UTC.fullmatch(value)):
        raise ValueError(INVALID)
    return value


_INBOX = re.compile(r"/v1\.0/me/mailFolders(?:/inbox|\('inbox'\))/messages/delta$", re.I | re.A)


def _link(value: str, resource: str) -> str:
    """A stored link, checked to be Graph's own delta endpoint before anything uses it.

    A mail link must also be the Inbox's: the cursor is unsigned, so a link to another
    folder's delta (Junk, say) would have that folder read as Inbox arrivals (review
    D3). Graph answers in the spelling it was asked in, ``mailFolders('inbox')``."""
    if value == "":
        return value
    checked = require_graph_url(value, source="cursor", resource=resource)
    if resource == "mail" and not _INBOX.search(urlsplit(checked).path):
        raise ForeignCursorError("cursor", checked, resource)
    return checked


def _not_after(value: str | None, now: datetime | None) -> str | None:
    """A time from the cursor, refused when it lies past now: a cursor from the future
    would hide every arrival (review D3)."""
    if value is not None and now is not None:
        moment = datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        if moment > now + FUTURE_ALLOWANCE:
            raise ValueError(INVALID)
    return value


def read(
    cursor: str | None, *, username: str | None, now: datetime | None = None
) -> tuple[CursorState | None, str]:
    """The state in ``cursor`` and ``understood.cursor``: ``none`` without one,
    ``reset_*`` when it can't be used (start from the window), ``ok`` otherwise.

    Malformed text raises ``ValueError``; a link that isn't Graph's delta endpoint
    raises ``UntrustedURLError`` or ``ForeignCursorError``.
    """
    if cursor is None:
        return None, NONE
    data = _decode(cursor)
    version = data.get("v")
    if type(version) is not int or version != VERSION:  # `True == 1`, but it is no version
        return None, RESET_UNRECOGNISED
    if data.get("acct") != account_tag(username):
        return None, RESET_OTHER_ACCOUNT
    mail, cal = data.get("mail"), data.get("cal")
    if not isinstance(mail, dict) or not isinstance(cal, dict):
        raise ValueError(INVALID)
    created_only = mail.get("created_only")
    if not isinstance(created_only, bool):
        raise ValueError(INVALID)
    state = CursorState(
        acct=_text(data, "acct"),
        made=_not_after(_text(data, "made", utc=True), now),
        mail=MailState(
            link=_link(_text(mail, "link"), "mail"),
            created_only=created_only,
            bootstrap_from=_not_after(_text(mail, "bootstrap_from", optional=True, utc=True), now),
        ),
        cal=CalState(
            link=_link(_text(cal, "link"), "calendar"),
            start=_text(cal, "start", utc=True),
            end=_text(cal, "end", utc=True),
            since=_not_after(_text(cal, "since", optional=True, utc=True), now),
        ),
    )
    return state, OK
