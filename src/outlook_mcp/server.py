"""FastMCP server for Microsoft Outlook."""

from __future__ import annotations

from contextlib import asynccontextmanager

from mcp.server.fastmcp import Context, FastMCP

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config
from outlook_mcp.graph import GraphClient
from outlook_mcp.tools import (
    admin,
    batch,
    calendar_read,
    calendar_write,
    contacts,
    mail_attachments,
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
async def outlook_auth_status(ctx: Context) -> dict:
    """Check authentication status. Run `outlook-mcp auth` on the host if needed."""
    auth = _get_auth(ctx)
    result = {
        "authenticated": auth.is_authenticated(),
        "read_only": auth.config.read_only,
    }
    if not auth.is_authenticated():
        result["action_required"] = (
            "Run `outlook-mcp auth` on the host to authenticate."
        )
    return result


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
    cursor: str | None = None,
    classification: str | None = None,
) -> dict:
    """List messages. Filters: read status, sender, date, Focused Inbox classification."""
    client = _get_graph_client(ctx)
    return await mail_read.list_inbox(
        client.sdk_client, folder, count, unread_only, from_address, after, before, skip,
        cursor=cursor, classification=classification,
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
    cursor: str | None = None,
) -> dict:
    """Search mail using KQL query. Query is automatically sanitized."""
    client = _get_graph_client(ctx)
    return await mail_read.search_mail(client.sdk_client, query, count, folder, cursor=cursor)


@mcp.tool()
async def outlook_list_folders(ctx: Context, cursor: str | None = None) -> dict:
    """List all mail folders with message counts."""
    client = _get_graph_client(ctx)
    return await mail_read.list_folders(client.sdk_client, cursor=cursor)


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
    sensitivity: str = "normal",
    request_read_receipt: bool = False,
) -> dict:
    """Send email with recipients, CC, BCC, HTML support, importance, and sensitivity."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_write.send_message(
        client.sdk_client, to, subject, body, cc, bcc, is_html, importance,
        sensitivity=sensitivity, request_read_receipt=request_read_receipt,
        config=config,
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
        client.sdk_client, message_id, body, reply_all, is_html, config=config
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
        client.sdk_client, message_id, to, comment, config=config
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
        client.sdk_client, message_id, folder, config=config
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
        client.sdk_client, message_id, permanent, config=config
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
        client.sdk_client, message_id, status, config=config
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
        client.sdk_client, message_id, categories, config=config
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
        client.sdk_client, message_id, is_read, config=config
    )


@mcp.tool()
async def outlook_reclassify_message(
    ctx: Context,
    message_id: str,
    classification: str,
) -> dict:
    """Reclassify a message's Focused Inbox placement. classification: "focused" or "other"."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_triage.reclassify_message(
        client.sdk_client, message_id, classification, config=config
    )


# ── Calendar Read Tools ─────────────────────────────────


@mcp.tool()
async def outlook_list_events(
    ctx: Context,
    days: int = 7,
    after: str | None = None,
    before: str | None = None,
    count: int = 50,
    cursor: str | None = None,
) -> dict:
    """List calendar events in a date range. Expands recurring events."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_read.list_events(
        client.sdk_client, days, after, before, count, config.timezone, cursor=cursor,
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
        config=config,
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
        client.sdk_client, event_id, subject, start, end, location, body, config=config
    )


@mcp.tool()
async def outlook_delete_event(
    ctx: Context,
    event_id: str,
) -> dict:
    """Delete a calendar event."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await calendar_write.delete_event(client.sdk_client, event_id, config=config)


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
        client.sdk_client, event_id, response, message, config=config
    )


# ── Contact Tools ──────────────────────────────────────


@mcp.tool()
async def outlook_list_contacts(
    ctx: Context,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List contacts with pagination."""
    client = _get_graph_client(ctx)
    return await contacts.list_contacts(client.sdk_client, count, cursor=cursor)


@mcp.tool()
async def outlook_search_contacts(
    ctx: Context,
    query: str,
    count: int = 25,
) -> dict:
    """Search contacts using KQL query."""
    client = _get_graph_client(ctx)
    return await contacts.search_contacts(client.sdk_client, query, count)


@mcp.tool()
async def outlook_get_contact(ctx: Context, contact_id: str) -> dict:
    """Get full contact details by ID."""
    client = _get_graph_client(ctx)
    return await contacts.get_contact(client.sdk_client, contact_id)


@mcp.tool()
async def outlook_create_contact(
    ctx: Context,
    first_name: str,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    company: str | None = None,
    title: str | None = None,
) -> dict:
    """Create a new contact."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.create_contact(
        client.sdk_client, first_name, last_name, email, phone, company, title,
        config=config,
    )


@mcp.tool()
async def outlook_update_contact(
    ctx: Context,
    contact_id: str,
    first_name: str | None = None,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
) -> dict:
    """Update an existing contact (partial patch)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.update_contact(
        client.sdk_client, contact_id, first_name, last_name, email, phone,
        config=config,
    )


@mcp.tool()
async def outlook_delete_contact(ctx: Context, contact_id: str) -> dict:
    """Delete a contact by ID."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await contacts.delete_contact(client.sdk_client, contact_id, config=config)


# ── To Do Tools ────────────────────────────────────────


@mcp.tool()
async def outlook_list_task_lists(ctx: Context) -> dict:
    """List all To Do task lists."""
    client = _get_graph_client(ctx)
    return await todo.list_task_lists(client.sdk_client)


@mcp.tool()
async def outlook_list_tasks(
    ctx: Context,
    list_id: str | None = None,
    status: str | None = None,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List tasks in a To Do list with optional status filter."""
    client = _get_graph_client(ctx)
    return await todo.list_tasks(client.sdk_client, list_id, status, count, cursor=cursor)


@mcp.tool()
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
    """Create a task in a To Do list."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.create_task(
        client.sdk_client, title, list_id, due, importance, body, reminder, recurrence,
        config=config,
    )


@mcp.tool()
async def outlook_update_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
    title: str | None = None,
    due: str | None = None,
    body: str | None = None,
    importance: str | None = None,
) -> dict:
    """Update a task in a To Do list (partial patch)."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.update_task(
        client.sdk_client, task_id, list_id, title, due, body, importance,
        config=config,
    )


@mcp.tool()
async def outlook_complete_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Mark a task as completed."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.complete_task(
        client.sdk_client, task_id, list_id, config=config,
    )


@mcp.tool()
async def outlook_delete_task(
    ctx: Context,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Delete a task from a To Do list."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await todo.delete_task(
        client.sdk_client, task_id, list_id, config=config,
    )


# ── Mail Draft Tools ──────────────────────────────────


@mcp.tool()
async def outlook_list_drafts(
    ctx: Context,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List messages in the Drafts folder."""
    client = _get_graph_client(ctx)
    return await mail_drafts.list_drafts(client.sdk_client, count, cursor=cursor)


@mcp.tool()
async def outlook_create_draft(
    ctx: Context,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    is_html: bool = False,
    importance: str = "normal",
) -> dict:
    """Create a draft message in the Drafts folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.create_draft(
        client.sdk_client, to, subject, body, cc, bcc, is_html, importance,
        config=config,
    )


@mcp.tool()
async def outlook_update_draft(
    ctx: Context,
    draft_id: str,
    subject: str | None = None,
    body: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
) -> dict:
    """Update an existing draft message."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.update_draft(
        client.sdk_client, draft_id, subject, body, to, cc, config=config,
    )


@mcp.tool()
async def outlook_send_draft(ctx: Context, draft_id: str) -> dict:
    """Send an existing draft message."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.send_draft(client.sdk_client, draft_id, config=config)


@mcp.tool()
async def outlook_delete_draft(ctx: Context, draft_id: str) -> dict:
    """Delete a draft message."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_drafts.delete_draft(client.sdk_client, draft_id, config=config)


# ── Mail Attachment Tools ─────────────────────────────


@mcp.tool()
async def outlook_list_attachments(ctx: Context, message_id: str) -> dict:
    """List attachments on a message."""
    client = _get_graph_client(ctx)
    return await mail_attachments.list_attachments(client.sdk_client, message_id)


@mcp.tool()
async def outlook_download_attachment(
    ctx: Context,
    message_id: str,
    attachment_id: str,
    save_path: str | None = None,
) -> dict:
    """Download an attachment. Returns base64 content or saves to file."""
    client = _get_graph_client(ctx)
    return await mail_attachments.download_attachment(
        client.sdk_client, message_id, attachment_id, save_path,
    )


@mcp.tool()
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
) -> dict:
    """Send a message with file attachments. Handles large files automatically."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_attachments.send_with_attachments(
        client.sdk_client, to, subject, body, attachment_paths, cc, bcc, is_html, importance,
        config=config,
    )


# ── Mail Folder Management Tools ─────────────────────


@mcp.tool()
async def outlook_create_folder(
    ctx: Context,
    name: str,
    parent_folder: str | None = None,
) -> dict:
    """Create a mail folder, optionally under a parent folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.create_folder(
        client.sdk_client, name, parent_folder, config=config,
    )


@mcp.tool()
async def outlook_rename_folder(
    ctx: Context,
    folder_id: str,
    name: str,
) -> dict:
    """Rename a mail folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.rename_folder(
        client.sdk_client, folder_id, name, config=config,
    )


@mcp.tool()
async def outlook_delete_folder(ctx: Context, folder_id: str) -> dict:
    """Delete a user-created mail folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_folders.delete_folder(
        client.sdk_client, folder_id, config=config,
    )


# ── Mail Thread Tools ─────────────────────────────────


@mcp.tool()
async def outlook_list_thread(
    ctx: Context,
    conversation_id: str,
    count: int = 50,
) -> dict:
    """List messages in a conversation thread, chronological order."""
    client = _get_graph_client(ctx)
    return await mail_thread.list_thread(client.sdk_client, conversation_id, count)


@mcp.tool()
async def outlook_copy_message(
    ctx: Context,
    message_id: str,
    folder: str,
) -> dict:
    """Copy a message to a folder."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await mail_thread.copy_message(
        client.sdk_client, message_id, folder, config=config,
    )


# ── Batch Tools ────────────────────────────────────────


@mcp.tool()
async def outlook_batch_triage(
    ctx: Context,
    message_ids: list[str],
    action: str,
    value: str,
) -> dict:
    """Triage up to 20 messages in one call. Actions: move, flag, categorize, mark_read."""
    client = _get_graph_client(ctx)
    config = _get_config(ctx)
    return await batch.batch_triage(
        client.sdk_client, message_ids, action, value, config=config,
    )


# ── User Tools ─────────────────────────────────────────


@mcp.tool()
async def outlook_whoami(ctx: Context) -> dict:
    """Get current user profile: display name, email, and ID."""
    client = _get_graph_client(ctx)
    return await user.whoami(client.sdk_client)


@mcp.tool()
async def outlook_list_calendars(ctx: Context) -> dict:
    """List all calendars for the current user."""
    client = _get_graph_client(ctx)
    return await user.list_calendars(client.sdk_client)


# ── Admin Tools ────────────────────────────────────────


@mcp.tool()
async def outlook_list_categories(ctx: Context) -> dict:
    """List master categories with names and colors."""
    client = _get_graph_client(ctx)
    return await admin.list_categories(client.sdk_client)


@mcp.tool()
async def outlook_get_mail_tips(ctx: Context, emails: list[str]) -> dict:
    """Get mail tips for recipients: delivery restrictions, OOF messages, etc."""
    client = _get_graph_client(ctx)
    return await admin.get_mail_tips(client.sdk_client, emails)


# ── Multi-Account Tools ───────────────────────────────


@mcp.tool()
async def outlook_list_accounts(ctx: Context) -> dict:
    """List configured accounts with authentication status."""
    auth = _get_auth(ctx)
    return {"accounts": auth.list_accounts()}


@mcp.tool()
async def outlook_switch_account(ctx: Context, name: str) -> dict:
    """Switch to a different configured account."""
    auth = _get_auth(ctx)
    return auth.switch_account(name)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
