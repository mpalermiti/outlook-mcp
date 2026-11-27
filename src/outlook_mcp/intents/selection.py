"""Structured mail selection (contracts §4): search mode, the part ``find`` uses.

A ``MailSelection`` names mail by fields, never by a hand-written query. With
words it is a search: everything goes out as one KQL ``$search`` value, and
nothing else, because whether Graph accepts ``$search`` next to ``$filter`` or
``$orderby`` on messages is unsettled (probes P-F2a/b). What KQL says only
roughly is checked again on what comes back (``matches_locally``):

- dates: KQL ``received`` takes whole days in a zone Graph doesn't state, so the
  query asks for one more day on each side and the exact instants are applied
  locally;
- the sender and attachments: applied locally too, so a search Graph evaluates
  loosely cannot return mail the selection excludes.

The filter and scan modes, for selections without words, arrive with the first
tool that needs them (``tidy``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta

from outlook_mcp.intents.people import parse_time, sender_of
from outlook_mcp.validation import sanitize_kql, validate_email

# KQL's operators are uppercase words; lowercase "and" and "or" are ordinary words.
_KQL_OPERATORS = frozenset({"AND", "OR", "NOT", "NEAR"})
_GROUPING = re.compile(r"[()]")


def kql_words(words: str) -> str:
    """The agent's words as KQL that can only narrow a search.

    ``sanitize_kql`` (reuse) takes out quotes, backslashes and ``& | ! *``. Here
    the parentheses and the operator words go too, so the words can't close a
    group or OR their way past the fields joined after them. A property
    restriction such as ``subject:budget`` is kept: ANDed with the rest, it only
    narrows.
    """
    try:
        inner = sanitize_kql(words)[1:-1]
    except ValueError:
        return ""
    tokens = _GROUPING.sub(" ", inner).split()
    return " ".join(t for t in tokens if t not in _KQL_OPERATORS)


def _day(moment: datetime) -> str:
    return moment.date().isoformat()


@dataclass
class MailSelection:
    """Mail named by fields (§4.1), the fields ``find`` passes."""

    words: str | None = None
    folder: str | None = None  # a folder reference; None means all mail
    from_: list[str] = field(default_factory=list)  # addresses, OR'd
    newer_than: datetime | None = None  # receivedDateTime ≥
    older_than: datetime | None = None  # receivedDateTime <
    has_attachments: bool | None = None

    def echo(self) -> dict:
        """``understood.selection``: the normalized selection, its mode and the query sent."""
        from outlook_mcp.intents.envelope import utc

        return {
            "mode": "search",
            "words": self.words,
            "folder": self.folder,
            "from": list(self.from_),
            "newer_than": utc(self.newer_than) if self.newer_than else None,
            "older_than": utc(self.older_than) if self.older_than else None,
            "has_attachments": self.has_attachments,
            "query": search_kql(self),
        }


def search_kql(selection: MailSelection) -> str:
    """The ``$search`` value for a selection with words, quoted as Graph needs it.

    Words first, then the other fields (``_field_parts``).
    """
    if not selection.words:
        raise ValueError("A search needs words.")
    words = kql_words(selection.words)
    if not words:
        raise ValueError(
            f"The words {selection.words[:50]!r} leave nothing to search once quotes, "
            "brackets and the KQL operators AND, OR, NOT and NEAR are taken out."
        )
    return '"' + " ".join([words, *_field_parts(selection)]) + '"'


def _field_parts(selection: MailSelection) -> list[str]:
    """The KQL for every field but the words: the sender (several OR'd in one group),
    the widened dates and ``hasAttachments:true``. Addresses are checked whole, so none
    can carry KQL of its own."""
    parts = []
    senders = [validate_email(a) for a in selection.from_]
    if len(senders) == 1:
        parts.append(f"from:{senders[0]}")
    elif senders:
        parts.append("(" + " OR ".join(f"from:{a}" for a in senders) + ")")
    # Two bounds are joined with AND: live, KQL ORs two restrictions on one property
    # that only a space separates, so `received>=X received<=Y` bounded nothing
    # (live session 2). One bound needs no join.
    dates = []
    if selection.newer_than is not None:
        dates.append(f"received>={_day(selection.newer_than - timedelta(days=1))}")
    if selection.older_than is not None:
        dates.append(f"received<={_day(selection.older_than + timedelta(days=1))}")
    if dates:
        parts.append(" AND ".join(dates))
    if selection.has_attachments:
        parts.append("hasAttachments:true")
    return parts


def matches_locally(row: dict, selection: MailSelection) -> bool:
    """Whether a Graph message row meets the selection's exact fields."""
    if selection.from_:
        address, _ = sender_of(row)
        if address is None or address not in {a.lower() for a in selection.from_}:
            return False
    if selection.newer_than is not None or selection.older_than is not None:
        received = parse_time(row.get("receivedDateTime"))
        if received is None:
            return False
        if selection.newer_than is not None and received < selection.newer_than:
            return False
        if selection.older_than is not None and received >= selection.older_than:
            return False
    if selection.has_attachments is not None:
        if bool(row.get("hasAttachments")) is not selection.has_attachments:
            return False
    return True
