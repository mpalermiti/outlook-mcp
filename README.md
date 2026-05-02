<!-- mcp-name: io.github.mpalermiti/outlook-mcp -->

# outlook-mcp

MCP server for Microsoft Outlook personal accounts via Microsoft Graph API.

[![PyPI](https://img.shields.io/pypi/v/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![MCP Registry](https://img.shields.io/badge/MCP_Registry-listed-green)](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti)

> **Personal Microsoft accounts only** — `@outlook.com`, `@hotmail.com`, `@live.com`. Work/school accounts (Entra ID) are not supported in v1.

> **Disclaimer:** Independent open-source project. Not affiliated with, endorsed by, or supported by Microsoft Corporation. "Outlook" and "Microsoft Graph" are trademarks of Microsoft.

---

## Who this is for

You'll like this if you're:

- An **agent builder** wiring Outlook into your own infra (OpenClaw, Claude Code, Cursor, custom MCP host) and want a typed tool surface — not stdout you have to parse
- Building on **personal Microsoft accounts** (Outlook.com / Hotmail / Live) and want full control: BYO Azure app, no enterprise consent flow, no shared client ID
- Looking for **real coverage** — mail, calendar, contacts, to-do, drafts, folders, batch ops, threading — instead of a mail-only or calendar-only wrapper
- Security-conscious: tokens in the OS keyring (Keychain on macOS), granular `allow_categories`, optional `read_only` mode, zero telemetry

This **isn't for you** if you need work/school M365 accounts (use Microsoft's official tooling — Entra ID auth and admin-consent flows are out of scope here), or if a basic mail-only client would suffice (this has 54 tools — way more than you need for "read my inbox").

### How it differs from other Outlook tools you'll find

This is the only **first-class MCP server** in the personal-Outlook space — most alternatives are bash scripts or skill-shaped CLI wrappers the agent shells out to. That distinction matters: the agent gets typed tool schemas with structured args/returns, not stdout it has to parse. Other things you won't find elsewhere: `/$batch`-optimized triage (10-20× faster on bulk ops), recursive folder ops with name resolution, granular per-category permissions, multi-account support, and full attachment write paths including >3MB upload sessions for drafts.

---

## What This Enables

Give your AI agent full Outlook access. Example prompts that just work:

- *"Summarize my unread email from the past 24 hours and flag anything time-sensitive."*
- *"What's in my Focused Inbox right now? Anything in Other that looks like it belongs up top?"*
- *"Any shipping updates in my inbox? Track what I'm waiting on and when it's supposed to arrive."*
- *"Scan my email for upcoming subscription renewals — what's about to auto-charge in the next two weeks?"*
- *"I've got a trip to Seattle next week — check my calendar for the itinerary and create a To Do task with a packing checklist."*
- *"Draft a reply to the last message from my sister saying I'll call her this weekend."*
- *"Move all newsletter and promotional email from this week to a 'Read Later' folder — batch 20 at a time."*

The server exposes 54 discrete tools so the agent can compose its own workflow — read, triage, write, schedule, track tasks — without hardcoded macros.

## Works With

- **[OpenClaw](https://openclaw.ai)** — native MCP support, available via [ClawHub](https://clawhub.ai/skills?q=outlook-mcp)
- **[Claude Code](https://claude.com/claude-code)** — add to `~/.claude/settings.json` under `mcpServers`
- **[Cursor](https://cursor.com)** — MCP-compatible
- **Any MCP client** — it's a standard stdio MCP server

Listed on the [official MCP Registry](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti) as `io.github.mpalermiti/outlook-mcp`.

---

## Features

**54 tools** across 13 categories:

- **Auth (1)** -- auth status check (login is via CLI)
- **Mail Read (4)** -- list inbox (with Focused Inbox filter), read message, search (KQL), list folders
- **Mail Write (3)** -- send, reply/reply-all, forward
- **Mail Triage (6)** -- move, delete (soft by default), flag, categorize, mark read/unread, reclassify (Focused Inbox)
- **Calendar Read (2)** -- list events (with recurring expansion), get event details
- **Calendar Write (4)** -- create, update, delete, RSVP (accept/decline/tentative)
- **Contacts (6)** -- list, search, get, create, update, delete
- **To Do (6)** -- list task lists, list/create/update/complete/delete tasks
- **Drafts (5)** -- list, create, update, send, delete
- **Attachments (5)** -- list, download, send-with-attachments, attach-to-draft, remove-draft-attachment
- **Folder Management (3)** -- create, rename, delete mail folders
- **Threading and Batch (3)** -- list thread, copy message, batch triage
- **User and Admin (6)** -- whoami, list calendars, list categories, mail tips, accounts

**Design principles:**

- **BYOID** -- Bring Your Own ID. You register your own Azure AD app. No shared client ID.
- **Zero telemetry** -- no analytics, no local caching, no third-party calls.
- **Token storage** -- OS keyring via `azure-identity` (macOS Keychain, Windows Credential Store, Linux Secret Service).
- **Input validation** -- all inputs validated (email, Graph IDs, OData, KQL, datetimes) before any API call.
- **Read-only mode** -- set `read_only: true` in config to block all write operations.
- **Soft delete** -- delete moves to Deleted Items by default. Hard delete requires explicit `permanent: true`.
- **Timezone-aware** -- calendar operations respect your configured IANA timezone.

---

## Azure AD App Registration

You need to register a free Azure AD app to get a client ID.

### Prerequisites (Personal Microsoft Accounts)

Microsoft has deprecated app registration for personal accounts without an Azure AD tenant. You need to create a free Azure account first:

1. Go to [azure.microsoft.com/free](https://azure.microsoft.com/free) and sign up with your personal `@outlook.com` account. Requires a credit card for identity verification but **won't charge you**. This creates a proper Azure AD tenant.

### Register the App

1. Go to [App Registrations](https://go.microsoft.com/fwlink/?linkid=2083908) and sign in with your `@outlook.com` account.

2. Click **"+ New registration"** and fill in:
   - **Name:** anything except Microsoft-branded terms (e.g. `mp-outlook-mcp` — names like "Outlook MCP" will be rejected)
   - **Supported account types:** select **"Personal Microsoft accounts only"**
   - **Redirect URI:** leave blank

3. Click **Register**. Copy the **Application (client) ID** from the overview page.

4. Go to **Authentication (Preview)** → **Settings** tab → toggle **"Allow public client flows"** to **Yes** → **Save**.

5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → add:
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`, `Tasks.ReadWrite`
   - `User.Read`, `offline_access`

No client secret is needed. The device code flow uses public client auth.

---

## Quick Start

### Install

**Option A — from PyPI (recommended):**

```bash
uv tool install outlook-graph-mcp
# or: pipx install outlook-graph-mcp
# or: pip install outlook-graph-mcp
```

**Option B — from source:**

```bash
git clone https://github.com/mpalermiti/outlook-mcp.git
cd outlook-mcp
uv sync
```

### Configure

Create `~/.outlook-mcp/config.json`:

```json
{
  "client_id": "YOUR_APPLICATION_CLIENT_ID",
  "tenant_id": "consumers",
  "timezone": "America/Los_Angeles",
  "read_only": true
}
```

The only required field is `client_id`. Everything else has sensible defaults. Start with `read_only: true` — flip to `false` when you're comfortable.

### Register with your MCP client

**If installed from PyPI:**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

**If installed from source:**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["--directory", "/path/to/outlook-mcp", "run", "outlook-mcp"]
    }
  }
}
```

**For OpenClaw**, use the `openclaw mcp` CLI — it writes to `mcp.servers` in `~/.openclaw/openclaw.json` for you:

```bash
# If installed from PyPI:
openclaw mcp set outlook '{"command":"outlook-mcp"}'

# If installed from source:
openclaw mcp set outlook '{"command":"uv","args":["--directory","/path/to/outlook-mcp","run","outlook-mcp"]}'

# Verify:
openclaw mcp list
openclaw mcp show outlook --json
```

Restart the OpenClaw gateway after registering. See the [OpenClaw MCP docs](https://docs.openclaw.ai/cli/mcp) for SSE/HTTP transport variants.

### Authenticate

Run this once on the machine where the MCP server will run:

```bash
uv run outlook-mcp auth
```

You'll get a URL and a code. Open the URL in any browser, enter the code, and sign in with your Microsoft account. Tokens are cached in the OS keyring — the MCP server picks them up automatically.

Other CLI commands:

```bash
uv run outlook-mcp status   # Check auth status
uv run outlook-mcp logout   # Clear credentials
uv run outlook-mcp serve    # Start MCP server (default, used by OpenClaw/Claude)
```

---

## Tool Reference

### Auth

| Tool | Description |
|------|-------------|
| `outlook_auth_status` | Check if authenticated and whether read-only mode is active. |

> **Note:** Authentication is handled via the CLI (`outlook-mcp auth`), not through MCP tools. See [Authenticate](#authenticate) above.

### Mail Read

| Tool | Description |
|------|-------------|
| `outlook_list_inbox` | List messages in a folder. `folder` accepts display names, well-known names, or Graph IDs. Filter by read status, sender, date range, Focused Inbox classification. Pagination via `skip`. |
| `outlook_read_message` | Get full message by ID. Format: `text`, `html`, or `full` (both). |
| `outlook_search_mail` | Search mail using KQL query. Optionally scope to a folder by name or ID. |
| `outlook_list_folders` | List mail folders with counts, `parent_id`, and `child_count`. Pass `recursive=true` to walk the full folder tree (subfolders included). |

### Mail Write

| Tool | Description |
|------|-------------|
| `outlook_send_message` | Send email. Supports TO/CC/BCC, HTML body, importance level. |
| `outlook_reply` | Reply or reply-all to a message. |
| `outlook_forward` | Forward a message to one or more recipients with optional comment. |

### Mail Triage

| Tool | Description |
|------|-------------|
| `outlook_move_message` | Move a message to a folder by name or ID. |
| `outlook_delete_message` | Delete a message. Soft delete (Deleted Items) by default. `permanent: true` for hard delete. |
| `outlook_flag_message` | Set follow-up flag: `flagged`, `complete`, or `notFlagged`. |
| `outlook_categorize_message` | Set categories on a message. |
| `outlook_mark_read` | Mark a message as read or unread. |
| `outlook_reclassify_message` | Move a message between Focused Inbox and Other (`focused` / `other`). |

### Calendar Read

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List events in a date range. Expands recurring events. Configurable via `days`, `after`, `before`. |
| `outlook_get_event` | Get full event details: attendees, body, online meeting URL, recurrence. |

### Calendar Write

| Tool | Description |
|------|-------------|
| `outlook_create_event` | Create event with location, attendees, recurrence, online meeting support. |
| `outlook_update_event` | Update event fields (subject, time, location, body). Only patches changed fields. |
| `outlook_delete_event` | Delete a calendar event. |
| `outlook_rsvp` | RSVP to an event: `accept`, `decline`, or `tentative`. Optionally include a message. |

### Contacts

| Tool | Description |
|------|-------------|
| `outlook_list_contacts` | List contacts with cursor pagination. |
| `outlook_search_contacts` | Search contacts by name or email. |
| `outlook_get_contact` | Get full contact details by ID. |
| `outlook_create_contact` | Create a new contact. |
| `outlook_update_contact` | Update contact fields. |
| `outlook_delete_contact` | Delete a contact. |

### To Do

| Tool | Description |
|------|-------------|
| `outlook_list_task_lists` | List To Do lists. |
| `outlook_list_tasks` | List tasks with status filter and pagination. |
| `outlook_create_task` | Create task with due date, importance, recurrence. |
| `outlook_update_task` | Update task fields. |
| `outlook_complete_task` | Mark task as completed. |
| `outlook_delete_task` | Delete a task. |

### Drafts

| Tool | Description |
|------|-------------|
| `outlook_list_drafts` | List draft messages with pagination. |
| `outlook_create_draft` | Create a draft for later review and sending. |
| `outlook_update_draft` | Update draft fields. |
| `outlook_send_draft` | Send an existing draft. |
| `outlook_delete_draft` | Delete a draft. |

### Attachments

| Tool | Description |
|------|-------------|
| `outlook_list_attachments` | List attachments on a message. |
| `outlook_download_attachment` | Download attachment and save decoded bytes to a file. |
| `outlook_send_with_attachments` | Send message with file attachments (auto upload session for >3MB). |
| `outlook_attach_to_draft` | Add attachments to an existing draft (auto upload session for >3MB). |
| `outlook_remove_draft_attachment` | Remove a single attachment from a draft. |

### Folder Management

| Tool | Description |
|------|-------------|
| `outlook_create_folder` | Create mail folder (top-level or nested). |
| `outlook_rename_folder` | Rename a mail folder. |
| `outlook_delete_folder` | Delete a mail folder (refuses well-known folders). |

### Threading and Batch

| Tool | Description |
|------|-------------|
| `outlook_list_thread` | Get all messages in a conversation thread. |
| `outlook_copy_message` | Copy a message to another folder. |
| `outlook_batch_triage` | Batch move/flag/categorize/mark_read (max 20 per call). Single Graph `/$batch` round-trip — 10-20× faster than per-message calls for large triage. |

### User and Admin

| Tool | Description |
|------|-------------|
| `outlook_whoami` | Get current user profile. |
| `outlook_list_calendars` | List available calendars. |
| `outlook_list_categories` | List category definitions with colors. |
| `outlook_get_mail_tips` | Pre-send check (OOF, delivery restrictions). |
| `outlook_list_accounts` | List configured accounts. |
| `outlook_switch_account` | Switch active account. |

---

## Configuration

Config lives at `~/.outlook-mcp/config.json` (created with `0600` permissions).

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `client_id` | `string` | `null` | Azure AD application (client) ID. Required for auth. |
| `tenant_id` | `string` | `"consumers"` | Azure AD tenant. Use `"consumers"` for personal Microsoft accounts. |
| `timezone` | `string` | `"UTC"` | IANA timezone (e.g. `"America/New_York"`). Used for relative date computations in calendar tools. |
| `read_only` | `bool` | `false` | When `true`, all write tools (send, reply, move, delete, create, update, RSVP) return an error. |
| `allow_categories` | `list[string]` | `[]` | Optional. Restrict write tools to specific categories (see below). Empty list = all writes allowed when `read_only: false`. |

### Granular Write Permissions (optional)

By default, `read_only: false` unlocks **all** write tools. For finer control, set `allow_categories` to restrict write access to specific categories. Read tools (list, search, get) are always allowed — `allow_categories` only narrows the write surface.

**Available categories:**

| Category | Tools | Risk |
|---|---|---|
| `mail_drafts` | create/update/delete draft | Safe — drafts only, no send |
| `mail_triage` | move, delete (soft), flag, categorize, mark read, copy, batch | Moderate — reversible except hard delete |
| `mail_folders` | create/rename/delete folder | Moderate |
| `mail_send` | send, reply, forward, send_draft, send_with_attachments | **Dangerous** — sends email on your behalf |
| `calendar_write` | create/update/delete event, RSVP | Moderate — creates calendar entries |
| `contacts_write` | create/update/delete contact | Moderate |
| `todo_write` | create/update/complete/delete task | Safe — your own task list |

**Example policies:**

**Draft-only assistant** (agent can compose drafts, you review and send):

```json
{ "read_only": false, "allow_categories": ["mail_drafts", "mail_triage", "todo_write"] }
```

**Calendar-only** (agent can manage your schedule, nothing else):

```json
{ "read_only": false, "allow_categories": ["calendar_write"] }
```

**Full write access** (agent can do everything):

```json
{ "read_only": false }
```

**Read-only** (safest default, no writes):

```json
{ "read_only": true }
```

When `allow_categories` is set, any tool in a non-allowed category returns a permission-denied error (`PermissionDeniedError`) naming the blocked category. When `allow_categories` is empty (or unset) and `read_only` is false, all write tools are permitted. `read_only: true` always takes precedence — if set, all writes are blocked regardless of `allow_categories`. Unknown category names are rejected at config load time with a validation error; only the seven names above are accepted.

---

## Privacy and Security

- **Zero telemetry.** No analytics, no tracking, no usage data collected.
- **Zero local caching.** Every call goes directly to Microsoft Graph. No local email/calendar storage.
- **Zero third-party calls.** The server only talks to `graph.microsoft.com` and `login.microsoftonline.com`.
- **Token storage.** OAuth tokens are stored in your OS keyring via `azure-identity` (`TokenCachePersistenceOptions`). On macOS this is Keychain. No tokens are written to disk in plaintext.
- **No logging of sensitive data.** Message bodies, recipient addresses, and tokens are never logged.
- **Config permissions.** Config directory is `0700`, config file is `0600`. Symlinked configs are rejected.
- **Input validation.** All user inputs (email addresses, Graph IDs, OData filters, KQL queries, datetimes) are validated and sanitized before reaching the Graph API.

---

## Development

```bash
# Install dev dependencies
uv sync --extra dev

# Run tests
uv run pytest

# Lint
uv run ruff check src/ tests/

# Format
uv run ruff format src/ tests/

# Run server locally (stdio)
uv run outlook-mcp
```

**Requirements:** Python 3.10+

---

## Roadmap

- **Out-of-Office** -- get/set automatic replies
- **Inbox Rules** -- list, create, delete rules
- **Focused Inbox overrides** -- per-sender rules (always Focused/Other for X)
- **Advanced mail** -- raw MIME export, internet message headers
- **Calendar** -- cancel event (with attendee notification)
- **Checklists** -- checklist items on To Do tasks
- **Enterprise (Entra ID)** -- work/school account support

---

## License

MIT. See [LICENSE](LICENSE).
