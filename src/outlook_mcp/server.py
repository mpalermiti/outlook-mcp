"""FastMCP server for Microsoft Outlook."""

from __future__ import annotations

import functools
from collections.abc import Callable
from contextlib import asynccontextmanager
from typing import Any

from mcp.server.fastmcp import Context, FastMCP

from outlook_mcp import __version__
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config
from outlook_mcp.errors import OutlookMCPError, wrap_graph_error
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
    user,
)


@asynccontextmanager
async def lifespan(server):
    """Initialize server state: config, auth, and cached token."""
    config = load_config()
    auth = AuthManager(config)
    # Try to load cached token silently — if this fails, tools will
    # return an error telling the user to run `outlook-mcp auth`.
    auth.try_cached_token(auth.get_token_scopes())
    yield {"config": config, "auth": auth}


mcp = FastMCP(
    "outlook-mcp",
    instructions="MCP server for Microsoft Outlook via Microsoft Graph API",
    lifespan=lifespan,
)
# FastMCP doesn't expose a `version` kwarg, so set it on the underlying server
# directly. Otherwise serverInfo.version reports the MCP SDK's version.
mcp._mcp_server.version = __version__


# ── Helpers ─────────────────────────────────────────────


def _get_auth(ctx: Context) -> AuthManager:
    """Extract AuthManager from lifespan context."""
    return ctx.request_context.lifespan_context["auth"]


def _get_config(ctx: Context):
    """Extract Config from lifespan context."""
    return ctx.request_context.lifespan_context["config"]


def _get_graph_client(ctx: Context) -> GraphClient:
    """Create Graph client from auth context."""
    auth = _get_auth(ctx)
    return GraphClient(auth.get_credential())


def _wrap_tool_errors(func: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap an async tool function to translate Graph SDK errors.

    - Pass through ``OutlookMCPError`` subclasses (already structured).
    - Pass through ``ValueError`` (validation errors carry useful messages).
    - Convert Graph SDK errors (``ODataError`` / ``APIError``) into
      ``GraphAPIError`` via :func:`wrap_graph_error` so agents see
      ``{code, message, action}`` instead of raw SDK exception text.
    - Re-raise anything else unchanged.
    """

    @functools.wraps(func)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return await func(*args, **kwargs)
        except OutlookMCPError:
            raise
        except ValueError:
            raise
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
) -> dict:
    """List messages in one folder with structured filters (read, sender, date, Focused class).

    Use this for folder-scoped browsing; use outlook_search_mail for KQL full-text search across
    all folders. For polling/recurring agents use outlook_list_inbox_delta (typically 10x cheaper
    after the first call).

    Example: outlook_list_inbox(folder="Junk Email", unread_only=True, count=5)
    `folder` accepts display names, well-known names ("inbox", "junkemail"), or Graph IDs — prefer
    names. Pass concise=True to drop large fields (preview, categories) — ~10x fewer tokens.
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
    sensitivity: str = "normal",
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
        sensitivity=sensitivity,
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
) -> dict:
    """List calendar events in a date range (expands recurring instances).

    Use for one-shot queries; use outlook_list_events_delta for polling/recurring agents.

    Pass concise=True to drop large fields (body, attendees, organizer, categories) — ~10x fewer
    tokens for day-at-a-glance scans.
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
    )


@mcp.tool()
@_wrap_tool_errors
async def outlook_get_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Get one full calendar event by ID, including body, attendees, organizer, and recurrence."""
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
    return await calendar_delta.list_events_delta(client, start, end, page_size, delta_token)


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
    recurrence: str | None = None,
) -> dict:
    """Create a calendar event with optional attendees, recurrence, and Teams online meeting.

    Example: outlook_create_event(subject="Q3 review", start="2026-08-15T14:00:00Z",
    end="2026-08-15T15:00:00Z", attendees=["alice@acme.com"], is_online=True)
    `start`/`end` are ISO 8601. `recurrence` accepts a simple string ("daily", "weekly", "monthly").
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
) -> dict:
    """Update fields on an existing event (partial patch — only provided fields change)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.update_event(
        client.sdk_client, event_id, subject, start, end, location, body, config=config
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
) -> dict:
    """Update an existing contact (partial patch — only provided fields change)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.update_contact(
        client.sdk_client,
        contact_id,
        first_name,
        last_name,
        email,
        phone,
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
    `due` is ISO 8601. `importance` is "low", "normal", or "high". Defaults to the user's default
    list when `list_id` is omitted.
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
    """Download an attachment from a message and write decoded bytes to `save_path` on the host."""
    client = _get_graph_client(ctx)
    return await mail_attachments.download_attachment(
        client.sdk_client,
        message_id,
        attachment_id,
        save_path,
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

    `attachment_paths` must be absolute paths to files that exist on the host. Pass reply_to to
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

    `attachment_paths` must be absolute paths to files that exist on the host. Returns new
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


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
