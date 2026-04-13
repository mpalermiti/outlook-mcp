"""FastMCP server for Microsoft Outlook."""

from __future__ import annotations

from contextlib import asynccontextmanager

from mcp.server.fastmcp import Context, FastMCP

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config
from outlook_mcp.graph import GraphClient
from outlook_mcp.tools import calendar_read, calendar_write, mail_read, mail_triage, mail_write


@asynccontextmanager
async def lifespan(server):
    """Initialize server state: config and auth manager."""
    config = load_config()
    auth = AuthManager(config)
    yield {"config": config, "auth": auth}


mcp = FastMCP(
    "outlook-mcp",
    instructions="MCP server for Microsoft Outlook via Microsoft Graph API",
    lifespan=lifespan,
)


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


# ── Auth Tools ──────────────────────────────────────────


@mcp.tool()
async def outlook_login(ctx: Context, read_only: bool = False) -> dict:
    """Start device-code OAuth2 flow. Opens browser for Microsoft sign-in."""
    auth = _get_auth(ctx)
    if read_only:
        auth.config.read_only = read_only
    return auth.login()


@mcp.tool()
async def outlook_logout(ctx: Context) -> dict:
    """Remove stored credentials."""
    auth = _get_auth(ctx)
    return auth.logout()


@mcp.tool()
async def outlook_auth_status(ctx: Context) -> dict:
    """Check authentication status."""
    auth = _get_auth(ctx)
    return {
        "authenticated": auth.is_authenticated(),
        "read_only": auth.config.read_only,
    }


# ── Mail Read Tools ─────────────────────────────────────


@mcp.tool()
async def outlook_list_inbox(
    ctx: Context,
    folder: str = "inbox",
    count: int = 25,
    unread_only: bool = False,
    from_address: str | None = None,
    after: str | None = None,
    before: str | None = None,
    skip: int = 0,
) -> dict:
    """List messages in a folder with filtering by read status, sender, and date range."""
    client = _get_graph_client(ctx)
    return await mail_read.list_inbox(
        client.sdk_client, folder, count, unread_only, from_address, after, before, skip
    )


@mcp.tool()
async def outlook_read_message(
    ctx: Context,
    message_id: str,
    format: str = "text",
) -> dict:
    """Get full message by ID. Format: text, html, or full (both)."""
    client = _get_graph_client(ctx)
    return await mail_read.read_message(client.sdk_client, message_id, format)


@mcp.tool()
async def outlook_search_mail(
    ctx: Context,
    query: str,
    count: int = 25,
    folder: str | None = None,
) -> dict:
    """Search mail using KQL query. Query is automatically sanitized."""
    client = _get_graph_client(ctx)
    return await mail_read.search_mail(client.sdk_client, query, count, folder)


@mcp.tool()
async def outlook_list_folders(ctx: Context) -> dict:
    """List all mail folders with message counts."""
    client = _get_graph_client(ctx)
    return await mail_read.list_folders(client.sdk_client)


# ── Mail Write Tools ────────────────────────────────────


@mcp.tool()
async def outlook_send_message(
    ctx: Context,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
) -> dict:
    """Send email with recipients, CC, BCC, HTML support, and importance level."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.send_message(
        client.sdk_client, to, subject, body, cc, bcc, is_html, importance, config.read_only
    )


@mcp.tool()
async def outlook_reply(
    ctx: Context,
    message_id: str,
    body: str,
    reply_all: bool = False,
    is_html: bool = False,
) -> dict:
    """Reply or reply-all to a message."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.reply(
        client.sdk_client, message_id, body, reply_all, is_html, config.read_only
    )


@mcp.tool()
async def outlook_forward(
    ctx: Context,
    message_id: str,
    to: list[str],
    comment: str | None = None,
) -> dict:
    """Forward a message to recipients."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.forward(
        client.sdk_client, message_id, to, comment, config.read_only
    )


# ── Mail Triage Tools ───────────────────────────────────


@mcp.tool()
async def outlook_move_message(
    ctx: Context,
    message_id: str,
    folder: str,
) -> dict:
    """Move a message to a folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.move_message(
        client.sdk_client, message_id, folder, config.read_only
    )


@mcp.tool()
async def outlook_delete_message(
    ctx: Context,
    message_id: str,
    permanent: bool = False,
) -> dict:
    """Delete a message. Moves to Deleted Items by default. Set permanent=true to hard delete."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.delete_message(
        client.sdk_client, message_id, permanent, config.read_only
    )


@mcp.tool()
async def outlook_flag_message(
    ctx: Context,
    message_id: str,
    status: str,
) -> dict:
    """Set follow-up flag. Status: flagged, complete, or notFlagged."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.flag_message(
        client.sdk_client, message_id, status, config.read_only
    )


@mcp.tool()
async def outlook_categorize_message(
    ctx: Context,
    message_id: str,
    categories: list[str],
) -> dict:
    """Set categories on a message."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.categorize_message(
        client.sdk_client, message_id, categories, config.read_only
    )


@mcp.tool()
async def outlook_mark_read(
    ctx: Context,
    message_id: str,
    is_read: bool,
) -> dict:
    """Mark a message as read or unread."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.mark_read(
        client.sdk_client, message_id, is_read, config.read_only
    )


# ── Calendar Read Tools ─────────────────────────────────


@mcp.tool()
async def outlook_list_events(
    ctx: Context,
    days: int = 7,
    after: str | None = None,
    before: str | None = None,
    count: int = 50,
) -> dict:
    """List calendar events in a date range. Expands recurring events."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_read.list_events(
        client.sdk_client, days, after, before, count, config.timezone
    )


@mcp.tool()
async def outlook_get_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Get full event details by ID."""
    client = _get_graph_client(ctx)
    return await calendar_read.get_event(client.sdk_client, event_id)


# ── Calendar Write Tools ────────────────────────────────


@mcp.tool()
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
    """Create a calendar event with attendees, recurrence, and online meeting support."""
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
        config.read_only,
    )


@mcp.tool()
async def outlook_update_event(
    ctx: Context,
    event_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
) -> dict:
    """Update event fields."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.update_event(
        client.sdk_client, event_id, subject, start, end, location, body, config.read_only
    )


@mcp.tool()
async def outlook_delete_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Delete a calendar event."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.delete_event(client.sdk_client, event_id, config.read_only)


@mcp.tool()
async def outlook_rsvp(
    ctx: Context,
    event_id: str,
    response: str,
    message: str | None = None,
) -> dict:
    """RSVP to an event. Response: accept, decline, or tentative."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.rsvp(
        client.sdk_client, event_id, response, message, config.read_only
    )


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
