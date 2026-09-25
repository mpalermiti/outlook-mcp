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
- Security-conscious: tokens in the OS keyring (Keychain on macOS, libsecret on Linux -- never cleartext unless you opt in), granular `allow_categories`, optional `read_only` mode, zero telemetry

This **isn't for you** if you need work/school M365 accounts (use Microsoft's official tooling — Entra ID auth and admin-consent flows are out of scope here), or if a basic mail-only client would suffice (this has 70 tools — way more than you need for "read my inbox").

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

The server exposes 70 discrete tools so the agent can compose its own workflow — read, triage, write, schedule, track tasks — without hardcoded macros.

## Works With

- **[OpenClaw](https://openclaw.ai)** — native MCP support, available via [ClawHub](https://clawhub.ai/skills?q=outlook-mcp)
- **[Claude Code](https://claude.com/claude-code)** — add to `~/.claude/settings.json` under `mcpServers`
- **[Cursor](https://cursor.com)** — MCP-compatible
- **Any MCP client** — it's a standard stdio MCP server

Listed on the [official MCP Registry](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti) as `io.github.mpalermiti/outlook-mcp`.

---

## Features

**70 tools** across 13 categories:

- **Auth (1)** -- auth status check (login is via CLI)
- **Mail Read (7)** -- list inbox (with Focused Inbox and uncategorized filters), read message, bulk read by ID via `$batch`, search (KQL), list folders, delta-sync inbox changes, composed "since last call" digest across mail/events/contacts
- **Mail Write (3)** -- send, reply/reply-all, forward
- **Mail Triage (9)** -- move, delete (soft by default), flag, categorize, mark read/unread, reclassify (Focused Inbox), list/set/delete per-sender Focused Inbox overrides
- **Calendar Read (3)** -- list events (with recurring expansion), get event details, delta-sync event changes
- **Calendar Write (4)** -- create, update, delete, RSVP (accept/decline/tentative)
- **Contacts (7)** -- list, search, get, create, update, delete, delta-sync changes
- **To Do (14)** -- task lists, tasks (list/get/create/update/complete/delete), checklist items (add/update/delete), task attachments (list/download/upload/delete)
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
- **Read-only mode** -- set `read_only: true` in config to block all write operations. Note this limits the *tools*, not the *token* -- see [What `read_only` does and does not do](#what-read_only-does-and-does-not-do).
- **Soft delete** -- delete moves to Deleted Items by default. Hard delete requires explicit `permanent: true`.
- **Timezone-aware** -- calendar operations respect your configured IANA timezone.
- **Relative dates** -- every datetime parameter takes ISO 8601 or an offset: `7d` is seven days ago, `+7d` is seven days from now, `now` is this moment. Units: `m`, `h`, `d`, `w`.
- **Bounded attachments** -- attachment reads and writes are confined to `attachments_dir`, so a message that asks an agent to mail a file elsewhere on disk cannot be obeyed.
- **Bounded delta cursors** -- a `delta_token` is caller-held state, so it is untrusted input. Every URL that would carry a Graph bearer token is parsed and required to be https on `graph.microsoft.com`, which is what stops a poisoned cursor from redirecting your mailbox token to someone else.
- **Workflow prompts** -- `morning_brief`, `triage_folder` and `catch_up` ship as MCP prompts, so the common sequences do not have to be reconstructed call by call.

### Agent-friendly shape (1.8.0)

Two pure-code upgrades that make the same 57 tools cheaper and more recoverable for AI agents:

- **Concise mode** — pass `concise=True` to the five high-volume read tools (`outlook_list_inbox`, `outlook_read_message`, `outlook_search_mail`, `outlook_list_events`, `outlook_list_thread`) to drop bulky fields: full message bodies, per-event attendee lists, quoted prior-message text in threads, body previews/categories on inbox listings. Typical payload reduction ~10×. Default `concise=False` preserves the existing response shape — strict backward compat.

- **Structured Graph errors** — every tool wraps msgraph SDK exceptions into `{code, message, action}` responses with operator-friendly recovery hints: re-auth on 401, a link to the repo's [ROADMAP dead-ends list](https://github.com/mpalermiti/outlook-mcp/blob/main/ROADMAP.md#investigated-and-not-viable) on 403/`ErrorAccessDenied`, re-list on 404/`ErrorItemNotFound`, back-off on 429, retry on 503. `OutlookMCPError` subclasses and validation errors pass through unchanged.

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
  "read_only": true,
  "attachments_dir": "~/.outlook-mcp/attachments"
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

## Troubleshooting

### `SSL: CERTIFICATE_VERIFY_FAILED` on Linux

If auth fails with `[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate`, your Python environment can't find the system CA bundle. This is common on minimal/container Linux images and with the isolated venv from `uv tool install`.

Point Python at your system CA bundle. Set **both** variables — auth (via `azure-identity` → `requests`) reads `REQUESTS_CA_BUNDLE`, while the delta/`$batch` paths (via `httpx`) read `SSL_CERT_FILE`:

```bash
export SSL_CERT_FILE=/etc/ssl/certs/ca-certificates.crt      # httpx + Python ssl
export REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt  # azure-identity auth
```

The path varies by distro: Debian/Ubuntu use `/etc/ssl/certs/ca-certificates.crt`; RHEL/Fedora use `/etc/pki/tls/certs/ca-bundle.crt`. If the file is missing, install your distro's CA package (`ca-certificates`). Set these in the same environment your MCP client launches the server from so they apply at runtime, not just to the one-time `auth` command.

### Token cache stored unencrypted (Linux)

A one-time startup warning about the token cache falling back to plaintext means `libsecret`/PyGObject isn't importable — see [Privacy and Security](#privacy-and-security) for the fix.

---

## Tool Reference

**Dates.** Every datetime parameter (`after`, `before`, `start`, `end`, `due`,
`deferred_send_datetime`) accepts ISO 8601 — `2026-10-22` or `2026-10-22T14:30:00Z` — or a
relative offset: `7d` is seven days **ago**, `+7d` is seven days **from now**, and `now` is
this moment. Units are `m`, `h`, `d`, `w`. Bare means *ago*, matching the usual CLI
convention, so a due date in the future needs the `+`. Zone-less input is interpreted in your
configured `timezone`; responses are always UTC.

### Auth

| Tool | Description |
|------|-------------|
| `outlook_auth_status` | Check if authenticated and whether read-only mode is active. |

> **Note:** Authentication is handled via the CLI (`outlook-mcp auth`), not through MCP tools. See [Authenticate](#authenticate) above.

### Mail Read

| Tool | Description |
|------|-------------|
| `outlook_list_inbox` | List messages in a folder. `folder` accepts display names, well-known names, or Graph IDs. Filter by read status, sender, date range, Focused Inbox classification. Pagination via `skip`. |
| `outlook_read_message` | Get full message by ID. Format: `text`, `html`, or `full` (both). Pass `include_deferred_send=True` to also surface the draft's scheduled delivery time. |
| `outlook_read_messages` | Bulk read up to 20 messages by ID via Graph `$batch` in one round-trip. Per-message shape matches `outlook_read_message` byte-for-byte for the same `(format, concise, include_deferred_send)`. Partial-failure tolerant: 404s on some IDs surface in `failures[]` without failing the whole call. Use NOT N `outlook_read_message` calls. |
| `outlook_search_mail` | Search mail using KQL query. Optionally scope to a folder by name or ID. |
| `outlook_list_folders` | List mail folders with counts, `parent_id`, and `child_count`. Pass `recursive=true` to walk the full folder tree (subfolders included). |
| `outlook_list_inbox_delta` | List only inbox changes since the last call. First call returns a full snapshot plus a `delta_token`; subsequent calls (token passed back) return only added/updated/deleted items. Deletes come back as `{id, is_deleted: True}`. Cursor is stateless — agent persists and replays. |
| `outlook_changes_since` | One structured "since last call" digest composing mail/events/contacts deltas. Returns counts + `urgent_flagged` mail + top-5 `by_sender` + new/cancelled events. Each resource has an independent `delta_token`; stale-token recovery (HTTP 410) auto-resyncs that resource and surfaces `_meta.resync`. First-call snapshot is filtered to `fallback_window_hours` (default 24). Designed for recurring agent loops. |

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
| `outlook_list_inbox_overrides` | List Focused Inbox per-sender override rules. |
| `outlook_set_inbox_override` | Upsert a per-sender Focused Inbox override (`focused` / `other`). Case-insensitive sender matching; PATCH-if-exists, else POST. |
| `outlook_delete_inbox_override` | Delete a Focused Inbox override by ID. |

### Calendar Read

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List events in a date range. Expands recurring events. Each event carries `type`, so a series master is distinguishable from a one-off, and `show_as`, the free/busy status Outlook labels "Show as". Configurable via `days`, `after`, `before`. `calendar` selects which calendar to read: omit (or `"primary"`) for the default, otherwise a display name (case-insensitive; not-found and ambiguous errors name what exists) or an ID from `outlook_list_calendars`. A `cursor` continues the listing it came from, so later pages need neither `calendar` nor a second lookup. `concise=True` omits `show_as`. |
| `outlook_get_event` | Get full event details: attendees, body, online meeting URL, recurrence, `type` (`singleInstance` / `seriesMaster` / `occurrence` / `exception`), `show_as`. |
| `outlook_list_events_delta` | List only event changes inside a window since the last call. `start` and `end` (ISO 8601) required on the first call (Graph constraint — no whole-calendar sync). Each changed event carries `show_as`, matching `outlook_list_events`. Deletes come back as `{id, is_deleted: True}`. Cursor is stateless. |

### Calendar Write

| Tool | Description |
|------|-------------|
| `outlook_create_event` | Create event with location and attendees. (`is_online` has no effect on personal accounts — Graph ignores `isOnlineMeeting` for consumer mailboxes.) Pass `recurrence` to create a **series**: a shorthand (`daily`, `weekdays`, `weekly`, `monthly`, `yearly`, anchored on `start`) or a full [Graph recurrence object](https://learn.microsoft.com/graph/api/resources/patternedrecurrence) for anything else. `range.startDate` defaults to the event's start date. `show_as` sets the free/busy status Outlook labels "Show as" — `free`, `tentative`, `busy`, `oof` (out of office), `workingElsewhere`, or `unknown`; omit it and Graph applies its own default of `busy`. |
| `outlook_update_event` | Update event fields (subject, time, location, body, attendees, all-day, `show_as`). Only patches changed fields. Pass `recurrence` to turn a single event into a series, or `remove_recurrence=True` to turn a series back into a single event. `attendees` **replaces** the whole guest list and emails invitations/cancellations; `is_all_day` needs `start`+`end` in the same call, while `show_as` patches on its own. Patching a time keeps the zone the event is anchored in; `timezone` (with `start`+`end`) re-anchors it elsewhere, which is how a series created before zones existed gets repaired. |
| `outlook_delete_event` | Delete a calendar event. |
| `outlook_rsvp` | RSVP to an event: `accept`, `decline`, or `tentative`. Optionally include a message. |

### Contacts

| Tool | Description |
|------|-------------|
| `outlook_list_contacts` | List contacts with cursor pagination. Summaries carry `categories`; `outlook_search_contacts` omits the key because Graph's `$search` does not return it. |
| `outlook_search_contacts` | Search contacts by name or email. |
| `outlook_get_contact` | Get full contact details by ID, including home/business/other addresses, categories and personal notes. |
| `outlook_create_contact` | Create a new contact. |
| `outlook_update_contact` | Update contact fields. `home_address`, `business_address` and `other_address` take the shape `outlook_get_contact` returns — any subset of `street`, `city`, `state`, `postal_code`, `country_or_region` — and **replace** that whole address, so pass back every part you want to keep. Omit one to leave it untouched. |
| `outlook_delete_contact` | Delete a contact. |
| `outlook_list_contacts_delta` | List only contact changes since the last call. Deletes come back as `{id, is_deleted: True}`. Cursor is stateless. |

### To Do

| Tool | Description |
|------|-------------|
| `outlook_list_task_lists` | List To Do lists. |
| `outlook_list_tasks` | List tasks with status filter and pagination. |
| `outlook_get_task` | Get full task details: notes (`body`), checklist items (ordered unchecked-first), due, recurrence flag. |
| `outlook_create_task` | Create task with due date, importance, recurrence. |
| `outlook_update_task` | Update task fields. |
| `outlook_complete_task` | Mark task as completed. |
| `outlook_delete_task` | Delete a task. |
| `outlook_add_checklist_item` | Add a checklist item (sub-step) to a task. |
| `outlook_update_checklist_item` | Update a checklist item — mark done (`is_checked`) or rename (partial patch). |
| `outlook_delete_checklist_item` | Delete a checklist item from a task. |
| `outlook_list_task_attachments` | List attachments on a To Do task (id, name, size, content_type) with pagination. |
| `outlook_download_task_attachment` | Download a task attachment's content to a local file (confined to `attachments_dir`). |
| `outlook_upload_task_attachment` | Attach a local file (from `attachments_dir`) to a task via inline base64 POST (1 byte – 20 MiB). |
| `outlook_delete_task_attachment` | Remove an attachment from a To Do task. |

### Drafts

| Tool | Description |
|------|-------------|
| `outlook_list_drafts` | List draft messages with pagination. |
| `outlook_create_draft` | Create a draft. Supports scheduled delivery via `deferred_send_datetime` (server-side, Outlook-desktop-compatible "Delay Delivery"). |
| `outlook_update_draft` | Update draft fields. Accepts `is_html=True` for HTML bodies and `deferred_send_datetime` to set or clear the scheduled delivery time. |
| `outlook_send_draft` | Send an existing draft. |
| `outlook_delete_draft` | Delete a draft. |

### Attachments

> **Since 1.20.0, these tools only reach `attachments_dir`** (default `~/.outlook-mcp/attachments`).
> A bare filename resolves inside it; a path outside it is refused, including via a symlink.
> To email a file, move it there first — or widen `attachments_dir`, understanding that
> anything reachable from it can be sent. Before 1.20.0 these tools could read any file the
> server process could read, which meant an email asking an agent to attach one could be obeyed.
>
> The same fence covers the To Do attachment tools (`outlook_upload_task_attachment`,
> `outlook_download_task_attachment`): uploads read from `attachments_dir` and downloads
> write into it, so neither can sweep arbitrary files off disk. (Downloads are not
> `todo_write`-gated — they are reads, like `outlook_download_attachment`; the fence, not
> the category, is what confines them.)

| Tool | Description |
|------|-------------|
| `outlook_list_attachments` | List attachments on a message. |
| `outlook_download_attachment` | Download an attachment and save decoded bytes into `attachments_dir`. |
| `outlook_send_with_attachments` | Send a message with attachments read from `attachments_dir` (auto upload session for >3MB). |
| `outlook_attach_to_draft` | Add attachments from `attachments_dir` to an existing draft (auto upload session for >3MB). |
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

## Prompts

Three workflows ship as MCP prompts, so the common sequences do not have to be
reconstructed call by call. Any MCP client that supports prompts will list them; in most
clients they appear as slash commands or a prompt picker.

| Prompt | Arguments | What it does |
|--------|-----------|--------------|
| `morning_brief` | `folder` (default `inbox`) | Today's events, unread mail and tasks due, in the cheapest order — one scan each, `concise=True`, batched reads. |
| `triage_folder` | `folder` (default `inbox`), `count` (default 50) | One cheap scan of a folder, sorted into reply / archive / junk, applied with a single `outlook_batch_triage` call rather than one call per message. |
| `catch_up` | `since` (default `24h`) | What changed in mail, calendar and contacts, via the delta path — roughly ten times cheaper than re-scanning on a schedule. |

They cost nothing until invoked: `prompts/list` carries only a name and one line each, and
the body is fetched on use.

---

## Configuration

Config lives at `~/.outlook-mcp/config.json` (created with `0600` permissions).

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `client_id` | `string` | `null` | Azure AD application (client) ID. Required for auth. |
| `tenant_id` | `string` | `"consumers"` | Azure AD tenant. Use `"consumers"` for personal Microsoft accounts. |
| `timezone` | `string` | `"UTC"` | IANA timezone (e.g. `"America/New_York"`). Interprets zone-less dates, **and anchors every event you create** — a recurring event is expanded in this zone, so on the default `"UTC"` a 09:00 weekly meeting shifts an hour when the clocks change. Set it to where you are. |
| `read_only` | `bool` | `false` | When `true`, all write tools (send, reply, move, delete, create, update, RSVP) return an error. Gates the tools, not the Microsoft token -- see below. |
| `attachments_dir` | `string` | `"~/.outlook-mcp/attachments"` | The only directory the attachment tools may read from or write to. Every path an agent supplies is resolved and must land inside it — a symlink out or a `..` is refused. Widen it only if you understand that anything reachable can be emailed. |
| `allow_categories` | `list[string]` | `[]` | Optional. Restrict write tools to specific categories (see below). Empty list = all writes allowed when `read_only: false`. |
| `allow_unencrypted_token_cache` | `bool` | `false` | Permit the OAuth token cache to be written in cleartext when the platform has no encrypted store (Linux without libsecret). Off by default: authentication stops with an explanation rather than silently persisting a reusable Graph token in plaintext. macOS and Windows always encrypt and are unaffected. |

### Toolset selection (optional) — `OUTLOOK_MCP_TOOLSETS`

All 70 tool schemas load into the client's context every turn (~13.2k tokens by the chars/4 proxy `test_tool_surface_budget.py` measures with — a different yardstick than the ~8.6k o200k figure in ROADMAP; under this one, the 62-tool surface is ~11.7k, so the To Do detail tools add ~12%). A client that only needs part of the surface can set the `OUTLOOK_MCP_TOOLSETS` environment variable to a comma-separated list of tool groups, and only those load. The `account` group (auth / identity) is always available.

```bash
# e.g. a recurring mail + calendar agent: ~30 tools instead of 70 (~57% fewer tool tokens/turn)
OUTLOOK_MCP_TOOLSETS="mail,calendar,digest,delta"
```

Groups: `mail`, `drafts`, `attachments`, `calendar`, `contacts`, `todo`, `folders`, `digest`, `delta`, `admin`. Unset (the default) loads everything — fully backward compatible. This only affects which tools are advertised; enabled tools behave identically.

### What `read_only` does and does not do

`read_only: true` stops outlook-mcp's write tools from running. Ask it to send mail and it
refuses.

**It does not make your Microsoft credential read-only.** When outlook-mcp signs in it
requests the `.default` scope -- "everything this Azure app has been approved for." If you
consented the app to `Mail.ReadWrite` and `Mail.Send` (which the setup steps above tell you
to), the stored token can send mail whether `read_only` is on or off.

Two consequences worth understanding:

- `read_only` is a line in a text file. Anything able to edit `~/.outlook-mcp/config.json`
  turns it off and has write access immediately -- no re-authentication, no new consent
  prompt.
- The enforcement lives in this server's Python code. Any other process holding the cached
  token is unaffected by it.

So treat `read_only` as a guardrail against an agent doing something rash, **not as a
security boundary**. If you want a credential that genuinely cannot write, register a
second Azure app consented only to the read scopes (`Mail.Read`, `Calendars.Read`,
`Contacts.Read`, `Tasks.Read`, `User.Read`) and point `client_id` at that one. Then
Microsoft enforces it rather than us.

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
| `todo_write` | create/update/complete/delete task, checklist items; upload/delete task attachments | Moderate — your own task list, but `outlook_upload_task_attachment` reads local files from `attachments_dir` and pushes their bytes to Graph, and task/checklist/attachment deletes are irreversible. Listing and downloading attachments are plain reads, gated like every other read (not at all) and fenced to `attachments_dir` |

**Example policies:**

**Draft-only assistant** (agent can compose drafts, you review and send):

```json
{ "read_only": false, "allow_categories": ["mail_drafts", "mail_triage", "todo_write"] }
```

(Note that `todo_write` includes the *write-side* task-attachment tools — file reads
from `attachments_dir`, uploads to Graph, and irreversible deletes. Listing and
downloading task attachments are reads and are not write-gated, like the mail
attachment reads — see the table above.)

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
- **Zero local caching.** Every call goes directly to Microsoft Graph. No local email/calendar storage. (One carve-out: the To Do default-list id is resolved once per process and kept for the session — an id, not content; see `outlook_list_tasks`.)
- **Zero third-party calls.** The server only talks to `graph.microsoft.com` and `login.microsoftonline.com`.
- **Token storage.** OAuth tokens are persisted via `azure-identity`'s `TokenCachePersistenceOptions`. On macOS the OS Keychain is used; on Windows, DPAPI; on Linux with PyGObject/libsecret available, gnome-keyring. On Linux *without* libsecret (e.g. the isolated venv created by `uv tool install`), tokens fall back to a `0600` plaintext file at `~/.IdentityService/` and the MCP logs a one-time warning at startup. For encrypted storage on Linux, install `python3-gi gnome-keyring libsecret-1-0` and re-create the venv with `--system-site-packages`.
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

- **Inbox Rules** -- list, create, delete rules
- **Advanced mail** -- raw MIME export, internet message headers
- **Calendar** -- cancel event (with attendee notification)
- **Enterprise (Entra ID)** -- work/school account support

---

## License

MIT. See [LICENSE](LICENSE).
