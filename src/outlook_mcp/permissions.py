"""Granular permission control for outlook-mcp write operations.

Each write tool is tagged with one of the categories defined below. At call
time, `check_permission` consults the active Config:

- If `read_only=True`, every write is blocked with `ReadOnlyError`.
- If `allow_categories` is an empty list, the server is "fully open" and all
  categories are permitted (legacy behavior).
- If `allow_categories` is non-empty, only tools whose category appears in the
  list are permitted; everything else raises `PermissionDeniedError`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError

if TYPE_CHECKING:
    from outlook_mcp.config import Config

# ---------------------------------------------------------------------------
# Category constants
# ---------------------------------------------------------------------------

CATEGORY_MAIL_SEND = "mail_send"
CATEGORY_MAIL_DRAFTS = "mail_drafts"
CATEGORY_MAIL_TRIAGE = "mail_triage"
CATEGORY_MAIL_FOLDERS = "mail_folders"
CATEGORY_CALENDAR_WRITE = "calendar_write"
CATEGORY_CONTACTS_WRITE = "contacts_write"
CATEGORY_TODO_WRITE = "todo_write"

VALID_CATEGORIES: set[str] = {
    CATEGORY_MAIL_SEND,
    CATEGORY_MAIL_DRAFTS,
    CATEGORY_MAIL_TRIAGE,
    CATEGORY_MAIL_FOLDERS,
    CATEGORY_CALENDAR_WRITE,
    CATEGORY_CONTACTS_WRITE,
    CATEGORY_TODO_WRITE,
}


def check_permission(config: Config, category: str, tool_name: str) -> None:
    """Enforce permission policy for a write tool.

    Raises:
        ReadOnlyError: if the server is in read-only mode.
        PermissionDeniedError: if `allow_categories` is a non-empty whitelist
            that does not include `category`.

    Returns:
        None if the call is permitted.
    """
    if config.read_only:
        raise ReadOnlyError(tool_name)

    # Empty list = fully open (legacy behavior, no whitelist).
    if not config.allow_categories:
        return None

    if category not in config.allow_categories:
        raise PermissionDeniedError(tool_name, category)

    return None


def mail_send_withheld(config: Config) -> bool:
    """True when a whitelist is in force and ``mail_send`` is not on it."""
    return bool(config.allow_categories) and CATEGORY_MAIL_SEND not in config.allow_categories


def check_sends_mail(config: Config, tool_name: str, doing: str) -> None:
    """Enforce ``mail_send`` on a write that emails other people as a side effect.

    A tool's own category covers what it is named for. A meeting invitation or
    an RSVP comment is also an email — caller-written text, delivered by
    Exchange to an address the call names — and that is the thing ``mail_send``
    exists to gate. Called *after* the tool's own ``check_permission``, for the
    part of the call that sends; ``doing`` names that part for the refusal.

    Raises:
        ReadOnlyError: if the server is in read-only mode.
        PermissionDeniedError: if ``allow_categories`` is a non-empty whitelist
            without ``mail_send``.
    """
    if config.read_only:
        raise ReadOnlyError(tool_name)
    if mail_send_withheld(config):
        raise PermissionDeniedError(tool_name, CATEGORY_MAIL_SEND, doing=doing)
    return None
