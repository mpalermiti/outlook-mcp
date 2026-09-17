"""MCP server for Microsoft Outlook."""

from __future__ import annotations

import functools
import logging
import os
from collections.abc import Callable
from contextlib import asynccontextmanager
from typing import Any

from mcp.server.caching import CacheHint
from mcp.server.mcpserver import Context, MCPServer

from outlook_mcp import __version__, toolsets
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config
from outlook_mcp.errors import (
    OutlookMCPError,
    ToolInputError,
    UnencryptedTokenCacheError,
    wrap_graph_error,
)
from outlook_mcp.graph import GraphClient
from outlook_mcp.tools import (
    admin,
    batch,
    calendar_delta,
    calendar_read,
    calendar_write,
    contacts,
    contacts_delta,
    digest,
    inference_overrides,
    mail_attachments,
    mail_delta,
    mail_drafts,
    mail_folders,
    mail_read,
    mail_thread,
    mail_triage,
    mail_write,
    todo,
    todo_attachments,
    user,
)

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(server):
    """Initialize server state: config, auth, and cached token."""
    config = load_config()
    auth = AuthManager(config)
    # Try to load cached token silently — if this fails, tools will
    # return an error telling the user to run `outlook-mcp auth`.
    try:
        auth.try_cached_token()
    except UnencryptedTokenCacheError as exc:
        # This host cannot store a token safely. That is worth refusing, but
        # not worth killing the server over: dying here shows the client a
        # dead process and leaves the one-line fix on stderr, where no agent
        # reads it. Boot unauthenticated and let every tool call carry the
        # remedy instead.
        logger.warning("%s %s", exc.message, exc.action)
        auth.startup_error = exc
    yield {"config": config, "auth": auth}


# Sent once at connect, to every client, in every session — so it is the right
# home for the conventions that span tools. Each rule below answers a mistake
# visible in two months of real trajectories: 115 calls spent re-checking an
# identity that never changes, and 196 folder listings before folder-scoped
# scans that resolve display names on their own. That is roughly a quarter of
# all traffic, bought back for the cost of sending this string.
INSTRUCTIONS = """\
Microsoft Outlook (personal accounts: outlook.com, hotmail.com, live.com) via Microsoft Graph.

Working rules, each of which saves a round trip:

- You are already signed in. Do not call outlook_whoami, outlook_list_accounts or
  outlook_auth_status to check before doing something — just call the tool you need. If a call
  does fail on authentication, its error says exactly what to run.
- Folder and calendar parameters take display names directly ("Junk Email", "Purchases",
  "Work"), as well as well-known folder names ("inbox", "drafts") and Graph IDs. Do not list
  folders or calendars first to find an ID. Call outlook_list_folders / outlook_list_calendars
  only when you genuinely need to discover what exists.
- Dates accept ISO 8601 (2026-10-22, or 2026-10-22T14:30:00Z) or a relative offset: `7d` is
  seven days ago, `+7d` is seven days from now, `now` is this moment. Units: m, h, d, w.
- Scanning mail or events? Pass concise=True — roughly ten times fewer tokens. To read several
  messages, call outlook_read_messages once with the IDs, never outlook_read_message in a loop.
- Polling on a schedule? The delta tools (outlook_list_inbox_delta and friends) return only what
  changed since the token they handed you last time, and outlook_changes_since composes all
  three into one digest.
- Contact categories come back from outlook_list_contacts and outlook_get_contact, and are
  absent — not empty — from outlook_search_contacts, because Graph's contact $search does not
  return them. Never report a searched contact as uncategorised; read it back if you need to
  know. Addresses are on outlook_get_contact only, and writing one replaces it wholesale, so
  giving a new contact an address is create, then read, then update with every part.
"""

# SEP-2549: tell the client how long `tools/list` stays fresh, so it can stop
# re-fetching ~8.6k tokens of schemas. The set is fixed at import — toolset
# gating reads its env var once — so it cannot change while the process lives.
# Five minutes rather than an hour because it *can* change across a restart,
# which is exactly what someone editing OUTLOOK_MCP_TOOLSETS just did. Private:
# the advertised set depends on this install's configuration, so it is not a
# shared intermediary's to hand to someone else.
TOOL_LIST_CACHE = CacheHint(ttl_ms=5 * 60 * 1000, scope="private")

mcp = MCPServer(
    "outlook-mcp",
    instructions=INSTRUCTIONS,
    lifespan=lifespan,
    version=__version__,
    cache_hints={"tools/list": TOOL_LIST_CACHE},
)


# ── Helpers ─────────────────────────────────────────────


def _get_auth(ctx: Context) -> AuthManager:
    """Extract AuthManager from lifespan context."""
    return ctx.request_context.lifespan_context["auth"]


def _get_config(ctx: Context):
    """Extract Config from lifespan context."""
    return ctx.request_context.lifespan_context["config"]


def _get_graph_client(ctx: Context) -> GraphClient:
    """Return a Graph client, reused across tool calls for the same credential.

    Building a ``GraphServiceClient`` (auth provider, request adapter, TLS
    connection pool) on every tool call is wasteful on recurring agent loops.
    Cache one in the lifespan context and reuse it while the credential is
    unchanged. A ``switch_account`` / re-auth swaps ``AuthManager.credential``
    for a different object, so an identity check rebuilds the client
    automatically — no explicit invalidation needed.
    """
    auth = _get_auth(ctx)
    credential = auth.get_credential()  # raises AuthRequiredError if unauthenticated
    lifespan_ctx = ctx.request_context.lifespan_context
    cached = lifespan_ctx.get("graph_client")
    if cached is None or cached.credential is not credential:
        cached = GraphClient(credential)
        lifespan_ctx["graph_client"] = cached
    return cached


def _wrap_tool_errors(func: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap an async tool function to translate Graph SDK errors.

    - Pass through ``OutlookMCPError`` subclasses (already structured).
    - Re-raise ``ValueError`` as ``ToolInputError`` — same message, but now an
      anticipated failure, so the SDK forwards the text to the model instead of
      withholding it as a crash. Still a ``ValueError`` for existing callers.
    - Convert Graph SDK errors (``ODataError`` / ``APIError``) into
      ``GraphAPIError`` via :func:`wrap_graph_error` so agents see
      ``{code, message, action}`` instead of raw SDK exception text.
    - Re-raise anything else unchanged: a crash is not something we saw coming,
      and its text stays on the server.
    """

    @functools.wraps(func)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return await func(*args, **kwargs)
        except OutlookMCPError:
            raise
        except ValueError as exc:
            raise ToolInputError(str(exc)) from exc
        except Exception as exc:
            try:
                raise wrap_graph_error(exc) from exc
            except TypeError:
                # Not a Graph SDK error — bubble the original.
                raise exc from None

    return wrapper


# ── Auth Tools ──────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_auth_status(ctx: Context) -> dict:
    """Check authentication status. Run `outlook-mcp auth` on the host if needed."""
    auth = _get_auth(ctx)
    result = {
        "authenticated": auth.is_authenticated(),
        "read_only": auth.config.read_only,
    }
    if not auth.is_authenticated():
        if auth.startup_error is not None:
            # Re-running auth would fail identically; say what actually needs
            # changing.
            result["action_required"] = str(auth.startup_error)
        else:
            result["action_required"] = "Run `outlook-mcp auth` on the host to authenticate."
    return result


# ── Mail Read Tools ─────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_inbox(
    ctx: Context,
    folder: str = "inbox",
    count: int = 25,
    unread_only: bool = False,
    from_address: str | None = None,
    after: str | None = None,
    before: str | None = None,
    skip: int = 0,
    cursor: str | None = None,
    classification: str | None = None,
    concise: bool = False,
    uncategorized_only: bool = False,
) -> dict:
    """List messages in one folder with structured filters (read, sender, date, category, Focused).

    Use this for folder-scoped browsing; use outlook_search_mail for KQL full-text search across
    all folders. For polling/recurring agents use outlook_list_inbox_delta (typically 10x cheaper
    after the first call).

    Example: outlook_list_inbox(folder="Junk Email", unread_only=True, count=5)
    `folder` accepts display names, well-known names ("inbox", "junkemail"), or Graph IDs — prefer
    names. Pass concise=True to drop large fields (preview, categories) — ~10x fewer tokens.
    Pass uncategorized_only=True to return only messages with no categories assigned.
    `after`/`before` take ISO 8601 or a relative offset — `7d` is seven days ago, `+7d` is seven
    days from now.
    """
    client = _get_graph_client(ctx)
    return await mail_read.list_inbox(
        client.sdk_client,
        folder,
        count,
        unread_only,
        from_address,
        after,
        before,
        skip,
        cursor=cursor,
        classification=classification,
        concise=concise,
        uncategorized_only=uncategorized_only,
        timezone=_get_config(ctx).timezone,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_read_message(
    ctx: Context,
    message_id: str,
    format: str = "text",
    include_deferred_send: bool = False,
    concise: bool = False,
) -> dict:
    """Get one full message by ID. `format` is "text", "html", or "full" (both).

    Pass include_deferred_send=True to also return the scheduled-send time (PR_DEFERRED_SEND_TIME)
    as deferred_send_datetime — useful when recreating a delayed draft.
    Pass concise=True to drop large fields (body, body_html) and return a 200-char body_preview —
    ~10x fewer tokens for triage scans.
    """
    client = _get_graph_client(ctx)
    return await mail_read.read_message(
        client.sdk_client,
        message_id,
        format,
        include_deferred_send,
        concise=concise,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_read_messages(
    ctx: Context,
    message_ids: list[str],
    format: str = "text",
    concise: bool = False,
    include_deferred_send: bool = False,
) -> dict:
    """Bulk read up to 20 messages by ID via $batch — use NOT N outlook_read_message calls.

    Per-message shape in `messages` matches outlook_read_message byte-for-byte for the same
    (format, concise, include_deferred_send). Ordering follows input `message_ids`. Returns
    `{messages, failures, requested, succeeded, failed}` — 404s on some IDs are surfaced in
    `failures` without failing the whole call (partial-failure tolerant).

    Example: outlook_read_messages(message_ids=[id1, id2, id3], concise=True)
    Hard cap of 20 (Graph $batch limit).
    """
    client = _get_graph_client(ctx)
    return await mail_read.read_messages(
        client,
        message_ids,
        format=format,
        concise=concise,
        include_deferred_send=include_deferred_send,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_search_mail(
    ctx: Context,
    query: str,
    count: int = 25,
    folder: str | None = None,
    cursor: str | None = None,
    concise: bool = False,
) -> dict:
    """Full-text search mail with KQL across all folders (or one, if `folder` is set).

    Use this for "find emails about X"; use outlook_list_inbox for structured filters scoped to a
    single folder.

    Example: outlook_search_mail(query="from:sarah@acme.com received>=2026-01-01", count=10)
    `query` is Microsoft KQL (from:, subject:, received>=, hasattachment:true, AND/OR/NOT).
    Operators must be UPPERCASE — lowercase `and` is matched as a literal term. Two terms
    with no operator between them broaden the search; use AND explicitly to narrow.
    Pass concise=True to drop large fields (preview, categories) — ~10x fewer tokens.
    """
    client = _get_graph_client(ctx)
    return await mail_read.search_mail(
        client.sdk_client, query, count, folder, cursor=cursor, concise=concise
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_folders(
    ctx: Context,
    cursor: str | None = None,
    recursive: bool = False,
) -> dict:
    """List mail folders with message counts, parent_id, and child count.

    Default is top-level only; pass recursive=True to walk the full tree and resolve subfolder
    names.
    """
    client = _get_graph_client(ctx)
    return await mail_read.list_folders(client.sdk_client, cursor=cursor, recursive=recursive)


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_inbox_delta(
    ctx: Context,
    folder: str = "inbox",
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List only inbox changes since the last call.

    Use this for polling/recurring agents — typically 10x cheaper than outlook_list_inbox after
    the first call. Use outlook_list_inbox for one-shot snapshots.

    Example: first call: outlook_list_inbox_delta(); next:
    outlook_list_inbox_delta(delta_token=<token from prior response>).
    is_deleted=True items are tombstones (drop cached payload). has_more=True means drain
    immediately by passing the returned delta_token back.
    """
    client = _get_graph_client(ctx)
    return await mail_delta.list_inbox_delta(client, folder, page_size, delta_token)


# ── Mail Write Tools ────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_send_message(
    ctx: Context,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    request_read_receipt: bool = False,
    reply_to: list[str] | None = None,
) -> dict:
    """Send an email immediately, no human review.

    For human-review workflows use outlook_create_draft + outlook_send_draft instead.
    For replying to an existing message use outlook_reply; for calendar invites use outlook_rsvp.
    Pass reply_to to route recipient replies to a different address (e.g. a shared team alias).
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.send_message(
        client.sdk_client,
        to,
        subject,
        body,
        cc,
        bcc,
        is_html,
        importance,
        request_read_receipt=request_read_receipt,
        reply_to=reply_to,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_reply(
    ctx: Context,
    message_id: str,
    body: str,
    reply_all: bool = False,
    is_html: bool = False,
) -> dict:
    """Reply (or reply-all) to an email message.

    Use this for email; use outlook_rsvp for calendar meeting invites.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.reply(
        client.sdk_client, message_id, body, reply_all, is_html, config=config
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_forward(
    ctx: Context,
    message_id: str,
    to: list[str],
    comment: str | None = None,
) -> dict:
    """Forward an existing message to new recipients, with optional comment."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.forward(client.sdk_client, message_id, to, comment, config=config)


# ── Mail Triage Tools ───────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_move_message(
    ctx: Context,
    message_id: str,
    folder: str,
) -> dict:
    """Move a message to another folder (removes from source).

    Use outlook_copy_message to duplicate without removing the source. For deletion use
    outlook_delete_message (not move to "deleteditems"). `folder` accepts display names,
    well-known names ("inbox", "archive", "deleteditems"), or Graph IDs — prefer names.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.move_message(client.sdk_client, message_id, folder, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_message(
    ctx: Context,
    message_id: str,
    permanent: bool = False,
) -> dict:
    """Delete a message — soft delete (to Deleted Items) by default; permanent=True to hard-delete.

    This is the canonical way to delete a message. Do NOT use
    outlook_move_message(folder="deleteditems") for deletion.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.delete_message(client.sdk_client, message_id, permanent, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_flag_message(
    ctx: Context,
    message_id: str,
    status: str,
) -> dict:
    """Set the follow-up flag on a message. `status` is "flagged", "complete", or "notFlagged"."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.flag_message(client.sdk_client, message_id, status, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_categorize_message(
    ctx: Context,
    message_id: str,
    categories: list[str],
) -> dict:
    """Set categories on a message (replaces the full list).

    Example: outlook_categorize_message(message_id=..., categories=["Follow-up", "Pricing"])
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.categorize_message(
        client.sdk_client, message_id, categories, config=config
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_mark_read(
    ctx: Context,
    message_id: str,
    is_read: bool,
) -> dict:
    """Mark a single message as read or unread (set is_read=True or False)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.mark_read(client.sdk_client, message_id, is_read, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_reclassify_message(
    ctx: Context,
    message_id: str,
    classification: str,
) -> dict:
    """Reclassify ONE message's Focused/Other placement. `classification` is "focused" or "other".

    Use this to fix a single message; use outlook_set_inbox_override for a sticky rule that affects
    future messages from the same sender.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.reclassify_message(
        client.sdk_client, message_id, classification, config=config
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_inbox_overrides(ctx: Context) -> dict:
    """List the user's Focused Inbox per-sender override rules.

    Each override forces mail from a given sender into Focused or Other regardless of Graph's
    inference.
    """
    client = _get_graph_client(ctx)
    return await inference_overrides.list_inbox_overrides(client.sdk_client)


@mcp.tool()
@_wrap_tool_errors
async def outlook_set_inbox_override(
    ctx: Context,
    sender_email: str,
    classify_as: str,
) -> dict:
    """Create or update a sticky Focused/Other rule for a sender (upsert, case-insensitive).

    Use this to permanently change classification for FUTURE messages from a sender; use
    outlook_reclassify_message to fix ONE existing message.

    Example: outlook_set_inbox_override(sender_email="marketing@acme.com", classify_as="other")
    Returns status: "created" or "updated".
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await inference_overrides.set_inbox_override(
        client.sdk_client, sender_email, classify_as, config=config
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_inbox_override(
    ctx: Context,
    override_id: str,
) -> dict:
    """Delete a Focused Inbox per-sender override rule by its ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await inference_overrides.delete_inbox_override(
        client.sdk_client, override_id, config=config
    )


# ── Calendar Read Tools ─────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_events(
    ctx: Context,
    days: int = 7,
    after: str | None = None,
    before: str | None = None,
    count: int = 50,
    cursor: str | None = None,
    concise: bool = False,
    calendar: str | None = None,
) -> dict:
    """List calendar events in a date range (expands recurring instances).

    Use for one-shot queries; use outlook_list_events_delta for polling/recurring agents.

    Pass concise=True to drop large fields (body, attendees, organizer, categories) — ~10x fewer
    tokens for day-at-a-glance scans.

    `calendar`: a display name or an ID from outlook_list_calendars; omit for the default calendar.
    A cursor continues the listing it came from.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_read.list_events(
        client.sdk_client,
        days,
        after,
        before,
        count,
        config.timezone,
        cursor=cursor,
        concise=concise,
        calendar=calendar,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_get_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Get one event by ID: body, attendees, organizer, recurrence, type, show_as, anchor zone.

    `recurrence` comes back in the same shape outlook_create_event accepts; `type` is
    "singleInstance", "seriesMaster", "occurrence" or "exception". `start`/`end` are UTC;
    `original_start_time_zone` is the zone the event is anchored in.
    """
    client = _get_graph_client(ctx)
    return await calendar_read.get_event(client.sdk_client, event_id)


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_events_delta(
    ctx: Context,
    start: str | None = None,
    end: str | None = None,
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List only calendar event changes within a window since the last call.

    Use this for polling/recurring agents — typically 10x cheaper than outlook_list_events after
    the first call. Use outlook_list_events for one-shot queries.

    Example: first call: outlook_list_events_delta(start="2026-05-22T00:00:00Z",
    end="2026-05-29T00:00:00Z"); next: outlook_list_events_delta(delta_token=<token>).
    start/end (ISO 8601) required on first call only; the cursor encodes the window thereafter.
    is_deleted=True items are tombstones. has_more=True means drain immediately.
    """
    client = _get_graph_client(ctx)
    return await calendar_delta.list_events_delta(
        client, start, end, page_size, delta_token, timezone=_get_config(ctx).timezone
    )


# ── Calendar Write Tools ────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_create_event(
    ctx: Context,
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
) -> dict:
    """Create a calendar event with optional attendees, recurrence, busy status, online meeting.

    Example: outlook_create_event(subject="Q3 review", start="2026-08-15T14:00:00Z",
    end="2026-08-15T15:00:00Z", attendees=["alice@acme.com"])
    `is_online` is accepted but has no effect on personal accounts — Graph silently ignores
    isOnlineMeeting for consumer mailboxes (it returns isOnlineMeeting: False,
    onlineMeetingProvider: "unknown"). Teams meetings require a work/school account.
    `start`/`end` are ISO 8601. Passing `recurrence` creates a series, not a single event.
    It takes either a shorthand — "daily", "weekdays", "weekly", "monthly", "yearly", all
    anchored on `start` and open-ended — or a full Microsoft Graph recurrence object for
    anything else, e.g. every other Mon+Fri for 10 occurrences:
    {"pattern": {"type": "weekly", "interval": 2, "daysOfWeek": ["monday", "friday"]},
     "range": {"type": "numbered", "numberOfOccurrences": 10}}
    `range.startDate` defaults to the event's start date. Prefer a bounded range
    ("endDate"/"numbered") when the event has attendees — a "noEnd" series invites them
    to every future occurrence.
    `timezone` is the IANA zone the event is anchored in, which is what a recurring
    series is expanded against (default: the configured zone). A zone name like
    America/Los_Angeles, never an abbreviation like PDT.
    `show_as` is Outlook's "Show as": "free", "tentative", "busy", "oof" (out of office),
    "workingElsewhere", or "unknown". Omitted, Graph defaults the event to busy.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.create_event(
        client.sdk_client,
        subject,
        start,
        end,
        location,
        body,
        attendees,
        is_all_day,
        is_online,
        recurrence,
        timezone,
        show_as,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_update_event(
    ctx: Context,
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
) -> dict:
    """Update fields on an existing event (partial patch — only provided fields change).

    `recurrence` takes the same shapes as outlook_create_event and converts a single
    event into a series, or replaces an existing series' pattern. Omit it to leave any
    recurrence alone; pass remove_recurrence=True to turn a series back into a single
    event, keeping the first occurrence's time (the two are mutually exclusive).
    `attendees` REPLACES the whole guest list (Graph has no add-one operation) and sends
    invitations to everyone on it plus cancellations to anyone dropped — pass the full
    intended list; [] removes everyone. `is_all_day` REQUIRES start and end in the same
    call, both on midnight boundaries. Patching a time keeps the zone the event is
    anchored in. Omitting an argument leaves it unchanged, so False and [] are
    instructions, not absences.
    `show_as` is Outlook's "Show as" — same values as outlook_create_event — and patches
    on its own; unlike is_all_day it needs nothing resent alongside it.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.update_event(
        client.sdk_client,
        event_id,
        subject,
        start,
        end,
        location,
        body,
        recurrence,
        remove_recurrence,
        attendees,
        is_all_day,
        show_as,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Delete a calendar event by ID (cancels and notifies attendees if you're the organizer)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.delete_event(client.sdk_client, event_id, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_rsvp(
    ctx: Context,
    event_id: str,
    response: str,
    message: str | None = None,
) -> dict:
    """RSVP to a calendar meeting invite. `response` is "accept", "decline", or "tentative".

    Use this for meeting invites; use outlook_reply to reply to a regular email message.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.rsvp(client.sdk_client, event_id, response, message, config=config)


# ── Contact Tools ──────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_contacts(
    ctx: Context,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List contacts with cursor pagination.

    Use for one-shot queries; use outlook_list_contacts_delta for polling/recurring agents.
    """
    client = _get_graph_client(ctx)
    return await contacts.list_contacts(client.sdk_client, count, cursor=cursor)


@mcp.tool()
@_wrap_tool_errors
async def outlook_search_contacts(
    ctx: Context,
    query: str,
    count: int = 25,
) -> dict:
    """Search contacts by name or email using KQL query syntax."""
    client = _get_graph_client(ctx)
    return await contacts.search_contacts(client.sdk_client, query, count)


@mcp.tool()
@_wrap_tool_errors
async def outlook_get_contact(ctx: Context, contact_id: str) -> dict:
    """Get one full contact by ID."""
    client = _get_graph_client(ctx)
    return await contacts.get_contact(client.sdk_client, contact_id)


@mcp.tool()
@_wrap_tool_errors
async def outlook_create_contact(
    ctx: Context,
    first_name: str,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    company: str | None = None,
    title: str | None = None,
) -> dict:
    """Create a new contact with name and optional email, phone, company, title."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.create_contact(
        client.sdk_client,
        first_name,
        last_name,
        email,
        phone,
        company,
        title,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_update_contact(
    ctx: Context,
    contact_id: str,
    first_name: str | None = None,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    home_address: dict | None = None,
    business_address: dict | None = None,
    other_address: dict | None = None,
) -> dict:
    """Update an existing contact (partial patch — only provided fields change).

    An address takes the shape outlook_get_contact returns — any subset of
    {"street", "city", "state", "postal_code", "country_or_region"} — and REPLACES that
    whole address, so pass back every part you want to keep. Omit it to leave it untouched.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.update_contact(
        client.sdk_client,
        contact_id,
        first_name,
        last_name,
        email,
        phone,
        home_address,
        business_address,
        other_address,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_contact(ctx: Context, contact_id: str) -> dict:
    """Delete a contact by ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.delete_contact(client.sdk_client, contact_id, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_contacts_delta(
    ctx: Context,
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List only contact changes since the last call.

    Use this for polling/recurring agents — typically 10x cheaper than outlook_list_contacts after
    the first call. Use outlook_list_contacts for one-shot queries.

    Example: first call: outlook_list_contacts_delta(); next:
    outlook_list_contacts_delta(delta_token=<token from prior response>).
    is_deleted=True items are tombstones (drop cached payload). has_more=True means drain
    immediately by passing the returned delta_token back.
    """
    client = _get_graph_client(ctx)
    return await contacts_delta.list_contacts_delta(client, page_size, delta_token)


# ── Digest Tool ─────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_changes_since(
    ctx: Context,
    delta_tokens: dict | None = None,
    fallback_window_hours: int = 24,
) -> dict:
    """One structured "since last call" digest across mail, events, and contacts.

    Use this for recurring agent loops (morning brief, hourly inbox sweep) — one call
    returns counts, urgent_flagged mail, by-sender rollup, plus new/cancelled events and
    contacts counts. Use the three individual delta tools (outlook_list_inbox_delta,
    outlook_list_events_delta, outlook_list_contacts_delta) when you need raw item lists
    or per-resource control.

    Example: first call: outlook_changes_since(); next:
    outlook_changes_since(delta_tokens=<delta_tokens from prior response>).
    First call returns a snapshot filtered to the last `fallback_window_hours` (default 24)
    so the digest doesn't surface thousands of historical items; subsequent calls (tokens
    passed back) return only what changed. Each resource's token is independent — drop
    one stale token without re-syncing the others. If Graph 410s on a token
    (`syncStateNotFound`), that resource auto-resyncs and `_meta.resync` lists which one.
    `urgent_flagged` = high-importance OR flagged mail. `by_sender` = top 5 senders.
    Calendar `modified[]` is reserved for future use — modified events surface in `new[]`
    today (Graph delta doesn't distinguish them). Calendar `organizer_email` is also
    currently empty (the v1.9.0 delta formatter surfaces the organizer name only).
    """
    client = _get_graph_client(ctx)
    return await digest.changes_since(client, delta_tokens, fallback_window_hours)


# ── To Do Tools ────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_task_lists(ctx: Context) -> dict:
    """List all Microsoft To Do task lists for the current user."""
    client = _get_graph_client(ctx)
    return await todo.list_task_lists(client.sdk_client)


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_tasks(
    ctx: Context,
    list_id: str | None = None,
    status: str | None = None,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List tasks in a To Do list with optional `status` filter.

    `status`: "notStarted", "inProgress", or "completed".
    """
    client = _get_graph_client(ctx)
    return await todo.list_tasks(client.sdk_client, list_id, status, count, cursor=cursor)


@mcp.tool()
@_wrap_tool_errors
async def outlook_get_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Get full To Do task details: notes (`body`), checklist items, due, recurrence flag.

    Use this for one task's sub-steps and notes; use outlook_list_tasks for
    overviews. `checklist_items` are ordered unchecked-first, matching the To Do
    client. `list_id` is only needed when the task lives in a non-default list.
    """
    client = _get_graph_client(ctx)
    return await todo.get_task(client.sdk_client, task_id, list_id)


@mcp.tool()
@_wrap_tool_errors
async def outlook_create_task(
    ctx: Context,
    title: str,
    list_id: str | None = None,
    due: str | None = None,
    importance: str | None = None,
    body: str | None = None,
    reminder: bool | None = None,
    recurrence: dict | None = None,
) -> dict:
    """Create a Microsoft To Do task with optional due date, importance, body, and recurrence.

    Example: outlook_create_task(title="Send invoice", due="2026-09-01", importance="high")
    `reminder=True` requires `due` and sets the reminder to the due time — Graph silently
    drops a reminder that has no time.
    `due` takes ISO 8601 or a relative offset — note `+7d` is seven days from now, while a bare
    `7d` means seven days *ago*. `importance` is "low", "normal", or "high". Defaults to the
    user's default list when `list_id` is omitted.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.create_task(
        client.sdk_client,
        title,
        list_id,
        due,
        importance,
        body,
        reminder,
        recurrence,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_update_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
    title: str | None = None,
    due: str | None = None,
    body: str | None = None,
    importance: str | None = None,
) -> dict:
    """Update fields on a To Do task (partial patch — only provided fields change)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.update_task(
        client.sdk_client,
        task_id,
        list_id,
        title,
        due,
        body,
        importance,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_complete_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Mark a To Do task as completed."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.complete_task(
        client.sdk_client,
        task_id,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Delete a task from a Microsoft To Do list."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.delete_task(
        client.sdk_client,
        task_id,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_add_checklist_item(
    ctx: Context,
    task_id: str,
    display_name: str,
    list_id: str | None = None,
) -> dict:
    """Add a checklist item (sub-step) to a To Do task."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.add_checklist_item(
        client.sdk_client,
        task_id,
        display_name,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_update_checklist_item(
    ctx: Context,
    task_id: str,
    checklist_item_id: str,
    display_name: str | None = None,
    is_checked: bool | None = None,
    list_id: str | None = None,
) -> dict:
    """Update a checklist item (partial patch — only provided fields change).

    `is_checked=True` marks a sub-step done; Graph maintains the checked
    timestamp from it. Renaming passes `display_name`.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.update_checklist_item(
        client.sdk_client,
        task_id,
        checklist_item_id,
        display_name,
        is_checked,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_checklist_item(
    ctx: Context,
    task_id: str,
    checklist_item_id: str,
    list_id: str | None = None,
) -> dict:
    """Delete a checklist item from a To Do task."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.delete_checklist_item(
        client.sdk_client,
        task_id,
        checklist_item_id,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_task_attachments(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List attachments on a To Do task (id, name, size, content_type) with pagination.

    Use `outlook_download_task_attachment` with an id from here to save the content.
    has_more=True means pass next_cursor back for the next page.
    """
    client = _get_graph_client(ctx)
    return await todo_attachments.list_task_attachments(
        client.sdk_client, task_id, list_id, count, cursor=cursor
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_download_task_attachment(
    ctx: Context,
    task_id: str,
    attachment_id: str,
    save_path: str,
    list_id: str | None = None,
) -> dict:
    """Download a To Do task attachment's content to a local file.

    `save_path` resolves inside the configured attachments directory; the
    write is atomic (temp file + replace), so a failed download never
    truncates a file already staged there.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo_attachments.download_task_attachment(
        client.sdk_client,
        task_id,
        attachment_id,
        save_path,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_upload_task_attachment(
    ctx: Context,
    task_id: str,
    file_path: str,
    list_id: str | None = None,
) -> dict:
    """Attach a local file to a To Do task via inline base64 POST (1 byte – 20 MiB).

    `file_path` resolves inside the configured attachments directory. Larger
    files are refused up front: Graph caps the request body at 30 MB and base64
    inflates the file 4/3, so 20 MiB is the honest ceiling.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo_attachments.upload_task_attachment(
        client.sdk_client,
        task_id,
        file_path,
        list_id,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_task_attachment(
    ctx: Context,
    task_id: str,
    attachment_id: str,
    list_id: str | None = None,
) -> dict:
    """Remove an attachment from a To Do task."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo_attachments.delete_task_attachment(
        client.sdk_client,
        task_id,
        attachment_id,
        list_id,
        config=config,
    )


# ── Mail Draft Tools ──────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_drafts(
    ctx: Context,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List messages in the Drafts folder with cursor pagination."""
    client = _get_graph_client(ctx)
    return await mail_drafts.list_drafts(client.sdk_client, count, cursor=cursor)


@mcp.tool()
@_wrap_tool_errors
async def outlook_create_draft(
    ctx: Context,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    reply_to: list[str] | None = None,
    deferred_send_datetime: str | None = None,
) -> dict:
    """Create a draft email for later review/send (pair with outlook_send_draft).

    Use this when a human should review before sending; use outlook_send_message to send
    immediately without review.
    Pass deferred_send_datetime (ISO 8601, e.g. "2026-05-06T08:00:00Z") to schedule delayed
    delivery — Exchange holds the message server-side after outlook_send_draft.
    Pass reply_to to pre-populate the Reply-To header.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.create_draft(
        client.sdk_client,
        to,
        subject,
        body,
        cc,
        bcc,
        is_html,
        importance,
        reply_to=reply_to,
        deferred_send_datetime=deferred_send_datetime,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_update_draft(
    ctx: Context,
    draft_id: str,
    subject: str | None = None,
    body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    reply_to: list[str] | None = None,
    is_html: bool = False,
    deferred_send_datetime: str | None = None,
) -> dict:
    """Update an existing draft (partial patch).

    Pass is_html=True when body is HTML — required when overwriting a draft originally composed
    as HTML (consumer Outlook rejects Text-over-HTML PATCH).
    Pass reply_to=[...] to overwrite Reply-To; reply_to=[] to clear it.
    Pass deferred_send_datetime (ISO 8601) to set the scheduled-send time; empty string clears it.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.update_draft(
        client.sdk_client,
        draft_id,
        subject,
        body,
        to,
        cc,
        reply_to=reply_to,
        is_html=is_html,
        deferred_send_datetime=deferred_send_datetime,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_send_draft(ctx: Context, draft_id: str) -> dict:
    """Send an existing draft (pair with outlook_create_draft for human-review send flow)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.send_draft(client.sdk_client, draft_id, config=config)


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_draft(ctx: Context, draft_id: str) -> dict:
    """Delete a draft message by ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.delete_draft(client.sdk_client, draft_id, config=config)


# ── Mail Attachment Tools ─────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_attachments(ctx: Context, message_id: str) -> dict:
    """List attachments on a message — returns IDs, names, sizes, and content types."""
    client = _get_graph_client(ctx)
    return await mail_attachments.list_attachments(client.sdk_client, message_id)


@mcp.tool()
@_wrap_tool_errors
async def outlook_download_attachment(
    ctx: Context,
    message_id: str,
    attachment_id: str,
    save_path: str,
) -> dict:
    """Download an attachment and write the decoded bytes to `save_path` on the host.

    `save_path` is resolved inside the configured attachments directory
    (`attachments_dir`, default ~/.outlook-mcp/attachments) — a bare filename lands
    there; a path outside it is refused. Same directory for reads and writes.
    """
    client = _get_graph_client(ctx)
    return await mail_attachments.download_attachment(
        client.sdk_client,
        message_id,
        attachment_id,
        save_path,
        config=_get_config(ctx),
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_send_with_attachments(
    ctx: Context,
    to: list[str],
    subject: str,
    body: str,
    attachment_paths: list[str],
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
    reply_to: list[str] | None = None,
) -> dict:
    """Send an email with file attachments; auto-switches to upload-session for files >3MB.

    `attachment_paths` resolve inside the configured attachments directory
    (`attachments_dir`, default ~/.outlook-mcp/attachments) — a bare filename is looked up
    there, and a path outside it is refused. Pass reply_to to
    route replies to a different address.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_attachments.send_with_attachments(
        client.sdk_client,
        to,
        subject,
        body,
        attachment_paths,
        cc,
        bcc,
        is_html,
        importance,
        reply_to=reply_to,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_attach_to_draft(
    ctx: Context,
    draft_id: str,
    attachment_paths: list[str],
) -> dict:
    """Add attachments to an existing draft; auto-switches to upload-session for files >3MB.

    `attachment_paths` resolve inside the configured attachments directory
    (`attachments_dir`, default ~/.outlook-mcp/attachments) — a bare filename is looked up
    there, and a path outside it is refused. Returns new
    attachment IDs for later removal via outlook_remove_draft_attachment.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_attachments.attach_to_draft(
        client.sdk_client,
        draft_id,
        attachment_paths,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_remove_draft_attachment(
    ctx: Context,
    draft_id: str,
    attachment_id: str,
) -> dict:
    """Remove a single attachment from a draft message by attachment ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_attachments.remove_draft_attachment(
        client.sdk_client,
        draft_id,
        attachment_id,
        config=config,
    )


# ── Mail Folder Management Tools ─────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_create_folder(
    ctx: Context,
    name: str,
    parent_folder: str | None = None,
) -> dict:
    """Create a mail folder; pass `parent_folder` (name or ID) to nest under an existing folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.create_folder(
        client.sdk_client,
        name,
        parent_folder,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_rename_folder(
    ctx: Context,
    folder_id: str,
    name: str,
) -> dict:
    """Rename a user-created mail folder by ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.rename_folder(
        client.sdk_client,
        folder_id,
        name,
        config=config,
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_delete_folder(ctx: Context, folder_id: str) -> dict:
    """Delete a user-created mail folder by ID; refuses well-known folders (inbox, sentitems)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.delete_folder(
        client.sdk_client,
        folder_id,
        config=config,
    )


# ── Mail Thread Tools ─────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_thread(
    ctx: Context,
    conversation_id: str,
    count: int = 50,
    concise: bool = False,
) -> dict:
    """List all messages in a conversation thread, chronological order.

    Needs `conversation_id` from a message's metadata. Pass concise=True to drop large fields
    (quoted prior-message text in each preview) — ~10x fewer tokens on long reply chains.
    """
    client = _get_graph_client(ctx)
    return await mail_thread.list_thread(client.sdk_client, conversation_id, count, concise=concise)


@mcp.tool()
@_wrap_tool_errors
async def outlook_copy_message(
    ctx: Context,
    message_id: str,
    folder: str,
) -> dict:
    """Copy a message to another folder (duplicates; source is unchanged).

    Use outlook_move_message to remove from source. `folder` accepts display names, well-known
    names, or Graph IDs — prefer names.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_thread.copy_message(
        client.sdk_client,
        message_id,
        folder,
        config=config,
    )


# ── Batch Tools ────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_batch_triage(
    ctx: Context,
    message_ids: list[str],
    action: str,
    value: str,
) -> dict:
    """Triage up to 20 messages in one $batch call.

    `action` is "move", "flag", "categorize", or "mark_read".

    Example: outlook_batch_triage(message_ids=[id1, id2], action="move", value="Archive")
    `value` is the action target (folder name for move, status for flag/mark_read, category name
    for categorize). Hard cap of 20.
    """
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await batch.batch_triage(
        client.sdk_client,
        message_ids,
        action,
        value,
        config=config,
    )


# ── User Tools ─────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_whoami(ctx: Context) -> dict:
    """Get the authenticated user's profile (display name, email, ID)."""
    client = _get_graph_client(ctx)
    return await user.whoami(client.sdk_client)


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_calendars(ctx: Context) -> dict:
    """List all calendars available to the authenticated user (primary + secondary)."""
    client = _get_graph_client(ctx)
    return await user.list_calendars(client.sdk_client)


# ── Admin Tools ────────────────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_categories(ctx: Context) -> dict:
    """List the user's master category definitions (names + colors).

    Provides the valid values for outlook_categorize_message.
    """
    client = _get_graph_client(ctx)
    return await admin.list_categories(client.sdk_client)


@mcp.tool()
@_wrap_tool_errors
async def outlook_get_mail_tips(ctx: Context, emails: list[str]) -> dict:
    """Pre-send check for recipients: out-of-office, delivery limits, mailbox-full warnings."""
    client = _get_graph_client(ctx)
    return await admin.get_mail_tips(client.sdk_client, emails)


# ── Multi-Account Tools ───────────────────────────────


@mcp.tool()
@_wrap_tool_errors
async def outlook_list_accounts(ctx: Context) -> dict:
    """List all configured Outlook accounts and their authentication status."""
    auth = _get_auth(ctx)
    return {"accounts": auth.list_accounts()}


@mcp.tool()
@_wrap_tool_errors
async def outlook_switch_account(ctx: Context, name: str) -> dict:
    """Switch the active Outlook account by configured `name` (from outlook_list_accounts)."""
    auth = _get_auth(ctx)
    return auth.switch_account(name)


# ── Annotations + config-gated toolsets ───────────────────────────────
# Applied once, after every @mcp.tool above has registered. Sets read-only /
# destructive annotations on all tools, and — when OUTLOOK_MCP_TOOLSETS is set
# (e.g. "mail,calendar,digest,delta") — loads only those groups so clients that
# don't need all 70 tools don't pay the per-turn context cost. Unset = all.
toolsets.configure(mcp, toolsets.parse_toolsets(os.environ.get("OUTLOOK_MCP_TOOLSETS")))


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()


# ── Workflow prompts ────────────────────────────────────
#
# Sequencing guidance is not per-tool knowledge, so it does not belong in 70
# docstrings that every client pays for on every turn. A prompt costs one line
# in `prompts/list` until it is invoked — and unlike a SKILL.md, which is not
# even shipped in the wheel, it reaches every MCP client. These three are the
# workflows the trajectory archive shows being assembled by hand, call by call.


@mcp.prompt(
    title="Morning brief",
    description="Today's calendar, unread mail and tasks due, in one pass.",
)
def morning_brief(folder: str = "inbox") -> str:
    """Compose the day-ahead summary from the cheapest calls that answer it."""
    return (
        f"Give me a brief for today, in this order:\n"
        f"1. outlook_list_events(days=1, concise=True) — what is scheduled.\n"
        f"2. outlook_list_inbox(folder={folder!r}, unread_only=True, concise=True, count=25) — "
        f"what arrived and is unread.\n"
        f"3. outlook_list_tasks() — what is due.\n"
        f"Then summarise in plain language: what is on, what needs a reply, what is overdue. "
        f"Lead with anything time-critical. Do not open individual messages unless the preview "
        f"is genuinely ambiguous — and if you need several, use outlook_read_messages once "
        f"rather than outlook_read_message in a loop."
    )


@mcp.prompt(
    title="Triage a folder",
    description="Scan one folder cheaply and act on it in a single batch.",
)
def triage_folder(folder: str = "inbox", count: int = 50) -> str:
    """Scan-then-batch: one read, one write, instead of a call per message."""
    return (
        f"Triage {folder!r}:\n"
        f"1. outlook_list_inbox(folder={folder!r}, count={count}, concise=True) — one scan. "
        f"The folder name goes straight in; there is no need to list folders first.\n"
        f"2. Sort what you find into: needs a reply, read and archive, junk, ignore.\n"
        f"3. Apply the result with outlook_batch_triage in ONE call rather than a "
        f"mark-read/move/delete per message.\n"
        f"Tell me what you did and what you left for me. Ask before deleting anything you are "
        f"not confident about — deletes are recoverable from Deleted Items, but replies are not."
    )


@mcp.prompt(
    title="Catch up since",
    description="What changed in mail, calendar and contacts since a point in time.",
)
def catch_up(since: str = "24h") -> str:
    """Steer to the delta path, which is an order of magnitude cheaper on a poll."""
    return (
        f"Tell me what changed since {since}.\n"
        f"Use outlook_changes_since(fallback_window_hours=...) — it composes the mail, calendar "
        f"and contacts delta queries into one digest and hands back per-resource delta tokens. "
        f"Keep those tokens and pass them back next time; that call then returns only what "
        f"changed, which is roughly ten times cheaper than re-scanning.\n"
        f"Summarise: new mail worth my attention (it flags high-importance and flagged items "
        f"for you), calendar changes, and anything cancelled."
    )
