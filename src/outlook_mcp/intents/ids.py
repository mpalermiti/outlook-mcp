"""Immutable ids: one id family per process (contracts §6.3, decision D2).

Outlook ids change when an item moves between folders, unless the request asks
for immutable ids with ``Prefer: IdType="ImmutableId"``. Messages, attachments,
events, meeting messages and contacts support them; folders and calendars keep
constant ids anyway, and To Do tasks and inbox rules are not covered.

A process that loads intent tools sends the preference on every Outlook request,
so one process hands out one id family and an agent never compares two ids for
the same message. Whether Graph honours it on a personal account is not
documented, so the family is learned at run time from ``Preference-Applied``:

- the first id-bearing reply sets it, ``"immutable"`` or ``"rest"``;
- a later reply that disagrees makes it ``"mixed"`` for the rest of the process,
  with a warning logged.

The family flag is short-lived action state (the "no local caching" carve-out in
contracts §0): one word, never mailbox data, gone at exit.
"""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)

IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"'


def _preferences(header: str | None) -> list[str]:
    return [part.strip() for part in (header or "").split(",") if part.strip()]


def _is_immutable_preference(part: str) -> bool:
    name, _, value = part.partition("=")
    return name.strip().lower() == "idtype" and value.strip().strip('"').lower() == "immutableid"


def combine_prefer(existing: str | None) -> str:
    """``existing``'s preferences with ``IdType="ImmutableId"`` added once, comma-separated.

    Never assigned over another preference: ``odata.maxpagesize`` sizes delta
    pages and ``outlook.timezone`` sets the zone of event times, and replacing
    either changes the answer without an error.
    """
    parts = _preferences(existing)
    if not any(_is_immutable_preference(p) for p in parts):
        parts.append(IMMUTABLE_ID_PREFERENCE)
    return ", ".join(parts)


def immutable_applied(preference_applied: str | None) -> bool:
    """Whether a ``Preference-Applied`` value says Graph returned immutable ids."""
    return any(_is_immutable_preference(p) for p in _preferences(preference_applied))


class IdFamily:
    """The id family this process has been handed: ``None`` until the first id-bearing reply."""

    def __init__(self) -> None:
        self.family: str | None = None

    def observe(self, preference_applied: str | None) -> str:
        """Record one successful id-bearing reply's ``Preference-Applied`` header."""
        seen = "immutable" if immutable_applied(preference_applied) else "rest"
        if self.family is None:
            self.family = seen
        elif self.family not in (seen, "mixed"):
            self.family = "mixed"
            logger.warning(
                "Graph returned both immutable and regular ids in this process; ids from "
                "earlier and later answers may not compare equal (id family: mixed)."
            )
        return self.family


# Process-wide: every intent request reports to it.
ID_FAMILY = IdFamily()
