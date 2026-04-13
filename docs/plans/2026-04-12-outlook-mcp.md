# Outlook MCP Server — Implementation Plan & PRD

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build the definitive open-source MCP server for Microsoft Outlook personal accounts (Outlook.com/Hotmail), published to ClawHub.ai — giving any AI agent framework full programmatic access to mail, calendar, contacts, and tasks via Microsoft Graph API.

**Architecture:** Python FastMCP server using stdio transport. Microsoft Graph SDK (`msgraph-sdk`) for API calls, `azure-identity` for OAuth2 device code auth with persistent token cache. Pydantic models for input validation. Tool responses structured for LLM consumption. Designed for OpenClaw but compatible with any MCP client (Claude Code, Cursor, etc.).

**Tech Stack:** Python 3.10+, FastMCP (mcp v1.27+), msgraph-sdk v1.55+, azure-identity v1.25+, Pydantic v2, pytest, uv (package manager)

---

## Table of Contents

1. [Product Requirements](#product-requirements)
2. [Architecture](#architecture)
3. [Security Model](#security-model)
4. [Tier 1: Core (MVP)](#tier-1-core-mvp)
5. [Tier 2: Power Features](#tier-2-power-features)
6. [Tier 3: Differentiators](#tier-3-differentiators)
7. [Future: Enterprise / Entra ID](#future-enterprise--entra-id)
8. [Implementation Tasks](#implementation-tasks)

---

## Product Requirements

### Problem Statement

There is no Outlook MCP server in the OpenClaw ecosystem or broader MCP community. Gmail has multiple options. Outlook has zero. AI agents that need to read, triage, send, or schedule via Outlook have no standardized integration path.

### Target Users

- **Primary:** OpenClaw users with personal Microsoft accounts (Outlook.com, Hotmail, Live)
- **Secondary:** Claude Code / Cursor / any MCP-compatible agent framework users
- **Future:** Enterprise users with Azure AD / Entra ID work accounts

### Success Criteria

- Ship Tier 1 to ClawHub with working auth, mail, and calendar tools
- Zero telemetry, zero data caching, zero third-party calls beyond Microsoft Graph
- BYOID (Bring Your Own ID) — user registers their own Azure AD app. README walks through app registration step-by-step.
- Tool responses optimized for LLM consumption (structured, concise, relevant fields only)

### Non-Goals (V1)

- Desktop/GUI email client features
- Local mail caching or offline mode
- Enterprise admin consent flows (captured for future)
- Webhook/push notification subscriptions (Tier 3)
- Email rendering (HTML→text conversion for display)
- Shared public client ID (evaluate post-launch if community demand warrants it)
- Multi-account support (Tier 2)
- Checking other people's availability (enterprise — `getSchedule` only returns own schedule on personal accounts)

### Inspirations & Stolen Ideas

From **olkcli** (MIT, GitHub: rlrghb/olkcli):
- Input validation patterns: OData filter injection prevention, KQL sanitization, Graph ID validation, terminal escape stripping
- PKCE on device code flow (defense-in-depth)
- `--read-only` scope concept → `read_only` config option
- Focused Inbox awareness (`inferenceClassification`)
- Well-known folder name mapping
- Timezone handling (IANA timezone in config, UTC in API, local in display)
- Response size limiting on HTTP reads
- Atomic file writes for config
- Datetime parse/re-serialize for OData filter safety

From **Gmail MCP** (@gongrzhe/server-gmail-autoauth-mcp):
- Tool granularity: one tool = one action (not CRUD-grouped)
- Zod/Pydantic schema-first tool definitions → auto-generated JSON schema
- Batch processing with graceful fallback (batch → per-item on failure)
- Thread/conversation support via `threadId` / `conversationId`
- Label/category management as first-class tools
- `get_or_create` idempotent patterns

---

## Architecture

### Project Structure

```
outlook-mcp/
├── src/
│   └── outlook_mcp/
│       ├── __init__.py              # Version, package metadata
│       ├── server.py                # FastMCP server entry point, lifespan, tool wiring
│       ├── auth.py                  # Device code OAuth2, token persistence
│       ├── graph.py                 # Graph client factory, request helpers
│       ├── config.py                # Config file management (~/.outlook-mcp/)
│       ├── validation.py            # Input validation (ported from olkcli patterns)
│       ├── errors.py                # Exception hierarchy for consistent error handling
│       ├── pagination.py            # Cursor-based pagination helpers (Tier 2)
│       ├── models/
│       │   ├── __init__.py
│       │   ├── mail.py              # Mail Pydantic models
│       │   ├── calendar.py          # Calendar Pydantic models
│       │   ├── contacts.py          # Contact Pydantic models
│       │   ├── todo.py              # To Do Pydantic models
│       │   └── common.py            # Shared models (pagination, errors)
│       └── tools/
│           ├── __init__.py          # Tool registration
│           ├── auth_tools.py        # login, logout, status
│           ├── mail_read.py         # list_inbox, read_message, search_mail, list_folders
│           ├── mail_write.py        # send_message, reply, forward
│           ├── mail_triage.py       # move, delete (soft), flag, categorize, mark_read
│           ├── mail_thread.py       # list_thread, copy_message (Tier 2)
│           ├── mail_folders.py      # create_folder, rename, delete (Tier 2)
│           ├── mail_drafts.py       # list, create, update, send, delete (Tier 2)
│           ├── mail_attachments.py  # list, download, send_with (Tier 2)
│           ├── calendar_read.py     # list_events, get_event
│           ├── calendar_write.py    # create, update, delete, rsvp
│           ├── contacts.py          # list, search, get, create, update, delete (Tier 2)
│           ├── todo.py              # lists, tasks, checklists (Tier 2)
│           ├── focused.py           # list_focused, move_to_focused (Tier 3)
│           ├── admin.py             # OOF, inbox rules, categories (Tier 2-3)
│           ├── user.py              # whoami (Tier 2)
│           └── notifications.py     # subscribe, unsubscribe (Tier 3)
├── tests/
│   ├── conftest.py                  # Shared fixtures, mock Graph client
│   ├── test_validation.py
│   ├── test_auth.py
│   ├── test_config.py
│   ├── test_errors.py
│   ├── test_mail_read.py
│   ├── test_mail_write.py
│   ├── test_mail_triage.py
│   ├── test_calendar_read.py
│   ├── test_calendar_write.py
│   └── ...
├── pyproject.toml                   # Package config, dependencies, scripts
├── README.md
├── LICENSE                          # MIT
├── CLAUDE.md                        # Claude Code instructions
├── SKILL.md                         # ClawHub manifest
├── SECURITY.md                      # Vulnerability disclosure policy
└── .github/
    └── workflows/
        ├── ci.yml                   # Test + lint on PR
        └── release.yml              # PyPI publish on tag
```

### Data Flow

```
Agent (OpenClaw/Claude/etc.)
  ↓ MCP tool call (JSON over stdio)
FastMCP Server (long-running process)
  ↓ Validates input (Pydantic)
  ↓ Checks auth (azure-identity credential)
  ↓ Silent token refresh if needed
Graph API Client (msgraph-sdk)
  ↓ HTTPS to graph.microsoft.com
  ↓ Response → structured dict (sanitized)
  ↑ Returns to agent as MCP tool result
```

### Server State Management (FastMCP Lifespan)

Server state is managed via FastMCP's `lifespan` context pattern. Config and AuthManager are created at startup, yielded as a context dict, and accessed by tools via the `Context` parameter. The Graph client is created per-request (lightweight — the credential handles token refresh internally).

```python
from contextlib import asynccontextmanager
from mcp.server.fastmcp import FastMCP, Context

@asynccontextmanager
async def lifespan(server):
    config = load_config()
    auth = AuthManager(config)
    yield {"config": config, "auth": auth}

mcp = FastMCP("outlook-mcp", version="0.1.0", lifespan=lifespan)

@mcp.tool()
async def outlook_list_inbox(ctx: Context, folder: str = "inbox", ...):
    auth: AuthManager = ctx.request_context.lifespan_context["auth"]
    client = GraphClient(auth.get_credential())
    return await list_inbox(client.sdk_client, folder, ...)
```

### Error Handling

Exception-based, not return-value dicts. Custom hierarchy with structured fields for LLM consumption. FastMCP converts raised exceptions into MCP error responses automatically.

```python
# src/outlook_mcp/errors.py

class OutlookMCPError(Exception):
    """Base exception for all outlook-mcp errors."""
    def __init__(self, code: str, message: str, action: str | None = None):
        self.code = code
        self.message = message
        self.action = action
        super().__init__(message)

class AuthRequiredError(OutlookMCPError):
    """Raised when a tool is called without authentication."""
    def __init__(self):
        super().__init__(
            "auth_required",
            "Not authenticated. No valid credential found.",
            "Call outlook_login to authenticate with your Microsoft account.",
        )

class ReadOnlyError(OutlookMCPError):
    """Raised when a write tool is called in read-only mode."""
    def __init__(self, tool_name: str):
        super().__init__(
            "read_only",
            f"Cannot use {tool_name} — server is in read-only mode.",
            "Set read_only to false in ~/.outlook-mcp/config.json to enable write operations.",
        )

class NotFoundError(OutlookMCPError):
    """Raised when a requested resource doesn't exist."""
    def __init__(self, resource: str, resource_id: str):
        super().__init__(
            "not_found",
            f"{resource} '{resource_id}' not found.",
            None,
        )

class GraphAPIError(OutlookMCPError):
    """Raised when the Graph API returns an error."""
    def __init__(self, status_code: int, error_code: str, message: str):
        action = None
        if status_code == 401:
            action = "Token may have expired. Try outlook_login to re-authenticate."
        elif status_code == 429:
            action = "Rate limited by Microsoft Graph. Wait a moment and retry."
        super().__init__(
            f"graph_api_{error_code}",
            message,
            action,
        )
        self.status_code = status_code
```

Tools raise these exceptions; they never return error dicts. This gives agents a clear error/success signal and produces consistent error structure across all 21+ tools.

### Timezone Contract

- **Config:** `timezone` field stores IANA timezone (default: `"UTC"`). Set by user during setup.
- **Graph API calls:** Always UTC (mandatory per Microsoft Graph).
- **Tool response datetimes:** Always ISO 8601 with UTC suffix (`Z`). No ambiguity.
- **Tool input datetimes:** If a datetime string lacks timezone info, it is interpreted in the config timezone and converted to UTC before sending to Graph.
- **`days` parameter** (on `list_events`): "now" is computed in the config timezone. `list_events(days=7)` means "from now in user's timezone through 7 days."
- **Rationale:** Agents always receive UTC (predictable), but user-facing inputs respect their local timezone (ergonomic).

### Key Design Decisions

| Decision | Choice | Why |
|----------|--------|-----|
| Language | Python | MCP ecosystem is Python-first, FastMCP is Python, ClawHub expects Python |
| Auth library | azure-identity | Microsoft's own library, handles token refresh/cache/PKCE automatically, battle-tested |
| Graph library | msgraph-sdk | Official SDK, typed, maintained by Microsoft, matches azure-identity |
| Token storage | azure-identity `TokenCachePersistenceOptions` | Uses OS keyring (macOS Keychain, etc.) automatically, no hand-rolled crypto |
| Transport | stdio | OpenClaw standard, works with all MCP clients |
| Input validation | Pydantic + custom validators | Schema-first, auto-generates JSON schema for MCP, catches bad input before API calls |
| No MSAL | Correct — azure-identity wraps MSAL internally | We get MSAL's token management without the complexity of direct MSAL usage |
| No local cache | V1 correct | Privacy-first, every call hits Graph fresh. Evaluate caching in Tier 3 |
| Tool granularity | One tool = one operation | Matches Gmail MCP pattern, easier for agents to reason about |
| Auth model | BYOID (Bring Your Own ID) | No shared client ID dependency. Users register their own Azure AD app. |
| Error handling | Exception-based (`OutlookMCPError` hierarchy) | FastMCP converts to MCP errors automatically. Clean error/success signal for agents. |
| Server state | FastMCP lifespan context | Config + AuthManager at startup, Graph client per-request. Standard pattern. |
| Timezone | UTC in/out, config timezone for relative computations | Predictable for agents, ergonomic for users |
| Delete semantics | Soft delete (move to Deleted Items) | Matches user expectation. Hard delete available via `permanent` flag. |

---

## Security Model

### Authentication

```
┌─────────────────────────────────────────────────────┐
│ First Run: Device Code Flow                         │
│                                                     │
│ 1. Server requests device code from Azure AD        │
│ 2. Returns URL + code to agent via MCP result       │
│ 3. User opens browser, enters code, consents        │
│ 4. Server polls for token completion                │
│ 5. Tokens cached to OS keyring via azure-identity   │
│                                                     │
│ Subsequent Runs: Silent Token Refresh               │
│                                                     │
│ 1. azure-identity loads cached refresh token        │
│ 2. Silently exchanges for new access token          │
│ 3. No user interaction required                     │
│                                                     │
│ Expired Refresh Token (>90 days inactive):          │
│                                                     │
│ 1. Silent refresh fails                             │
│ 2. AuthRequiredError raised with re-login action    │
│ 3. User calls outlook_login to start new flow       │
└─────────────────────────────────────────────────────┘
```

### Scopes

**Personal account (default):**
```
Mail.ReadWrite
Mail.Send
Calendars.ReadWrite
Contacts.ReadWrite
Tasks.ReadWrite
User.Read
offline_access
```

**Read-only mode** (config: `read_only: true`):
```
Mail.Read
Calendars.Read
Contacts.Read
Tasks.Read
User.Read
offline_access
```

**Enterprise scopes** (future, additive):
```
MailboxSettings.ReadWrite    # OOF, inbox rules
People.Read                  # Directory search
Calendars.Read.Shared        # Shared calendars
Mail.Read.Shared             # Shared mailboxes
```

### App Registration (BYOID)

- **User registers their own Azure AD app** — README provides step-by-step walkthrough
- **Config stores:** `client_id` (required) and `tenant_id` (default: `consumers`)
- **No client secret** — public client, device code flow only
- **Tenant:** `consumers` for personal accounts (not `common` — we explicitly restrict to personal in V1)
- **Future consideration:** If community demand warrants, evaluate a shared public client ID for zero-config setup

### Input Validation (Ported from olkcli)

| Vector | Validation | Source |
|--------|-----------|--------|
| Graph IDs (message, event, folder) | Regex: `^[a-zA-Z0-9_=+/-]{1,1024}$` | olkcli `validateID()` |
| Email addresses | Regex + Pydantic `EmailStr` | olkcli `ValidateEmail()` |
| **Datetimes (OData filters)** | **Parse as ISO 8601, re-serialize to UTC. Reject unparseable input.** | olkcli `buildMailFilter()` |
| KQL search queries | Strip `":()&\|!*\` then wrap in quotes | olkcli `sanitizeKQL()` |
| Folder names | Well-known whitelist + ID validation for custom | olkcli folder mapping |
| Phone numbers | Regex: `^[0-9 ()+.\-]{1,30}$` | olkcli `ValidatePhone()` |
| Response sizes | Limit API response body reads | olkcli `io.LimitReader` pattern |
| Tool output | Strip control characters from all string fields | olkcli `sanitizeStr()` |

### File Permissions

| Path | Permissions | Contents |
|------|------------|----------|
| `~/.outlook-mcp/` | 0700 | Config directory |
| `~/.outlook-mcp/config.json` | 0600 | Client ID, tenant, timezone, default account |
| Token cache | OS keyring | Managed by azure-identity, not our files |

### Privacy Guarantees

- **Zero telemetry.** No analytics, crash reporting, or phone-home.
- **Zero local data.** No email bodies, calendar events, or contacts cached to disk.
- **Zero third-party calls.** Only `login.microsoftonline.com` and `graph.microsoft.com`.
- **Token in OS keyring only.** Never written to a plain file (unless no keyring, then azure-identity's encrypted fallback).
- **No logging of sensitive data.** Verbose/debug mode logs request URLs and status codes, never tokens or message content.

---

## Tier 1: Core MVP (21 tools)

**Goal:** Minimum viable Outlook MCP. Auth + mail read/write/triage + calendar read/write. Ship to ClawHub.

| # | Tool | Description | Graph Endpoint | Parameters |
|---|------|-------------|----------------|------------|
| | **Auth** | | | |
| 1 | `outlook_login` | Device code OAuth2 flow | `POST /oauth2/v2.0/devicecode` | `read_only?: bool` |
| 2 | `outlook_logout` | Remove stored credentials | Local only | — |
| 3 | `outlook_auth_status` | Check token validity | `GET /me` | — |
| | **Mail — Read** | | | |
| 4 | `outlook_list_inbox` | List messages with filters (folder, unread, sender, date) | `GET /me/mailFolders/{id}/messages` | `folder?: str`, `count?: int (25, max 100)`, `unread_only?: bool`, `from_address?: str`, `after?: str`, `before?: str`, `skip?: int` |
| 5 | `outlook_read_message` | Get full message by ID (text/html/full) | `GET /me/messages/{id}` | `message_id: str`, `format?: "text"\|"html"\|"full"` |
| 6 | `outlook_search_mail` | Search mail using KQL | `GET /me/messages?$search=` | `query: str`, `count?: int (25, max 100)`, `folder?: str` |
| 7 | `outlook_list_folders` | List all mail folders | `GET /me/mailFolders` | — |
| | **Mail — Write** | | | |
| 8 | `outlook_send_message` | Send with to/cc/bcc, HTML, importance | `POST /me/sendMail` | `to: list[str]`, `subject: str`, `body: str`, `cc?: list[str]`, `bcc?: list[str]`, `is_html?: bool`, `importance?: "low"\|"normal"\|"high"` |
| 9 | `outlook_reply` | Reply or reply-all | `POST /me/messages/{id}/reply` | `message_id: str`, `body: str`, `reply_all?: bool`, `is_html?: bool` |
| 10 | `outlook_forward` | Forward to recipients | `POST /me/messages/{id}/forward` | `message_id: str`, `to: list[str]`, `comment?: str` |
| | **Mail — Triage** | | | |
| 11 | `outlook_move_message` | Move to a folder | `POST /me/messages/{id}/move` | `message_id: str`, `folder: str` |
| 12 | `outlook_delete_message` | Soft delete (move to Deleted Items). Optional hard delete. | `POST /me/messages/{id}/move` (soft) or `DELETE /me/messages/{id}` (hard) | `message_id: str`, `permanent?: bool (false)` |
| 13 | `outlook_flag_message` | Set follow-up flag | `PATCH /me/messages/{id}` | `message_id: str`, `status: "flagged"\|"complete"\|"notFlagged"` |
| 14 | `outlook_categorize_message` | Set categories | `PATCH /me/messages/{id}` | `message_id: str`, `categories: list[str]` |
| 15 | `outlook_mark_read` | Mark read or unread | `PATCH /me/messages/{id}` | `message_id: str`, `is_read: bool` |
| | **Calendar — Read** | | | |
| 16 | `outlook_list_events` | List events in date range (expands recurring). `days` computed relative to "now" in config timezone. | `GET /me/calendarView` | `days?: int (7)`, `after?: str`, `before?: str`, `count?: int (25, max 100)` |
| 17 | `outlook_get_event` | Get full event details | `GET /me/events/{id}` | `event_id: str` |
| | **Calendar — Write** | | | |
| 18 | `outlook_create_event` | Create with attendees, recurrence, online meeting, location | `POST /me/events` | `subject: str`, `start: str`, `end: str`, `location?: str`, `body?: str`, `attendees?: list[str]`, `is_all_day?: bool`, `is_online?: bool`, `recurrence?: "daily"\|"weekdays"\|"weekly"\|"monthly"\|"yearly"` |
| 19 | `outlook_update_event` | Update event fields | `PATCH /me/events/{id}` | `event_id: str`, `subject?: str`, `start?: str`, `end?: str`, `location?: str`, `body?: str` |
| 20 | `outlook_delete_event` | Delete event | `DELETE /me/events/{id}` | `event_id: str` |
| 21 | `outlook_rsvp` | Accept / decline / tentative | `POST /me/events/{id}/accept` etc. | `event_id: str`, `response: "accept"\|"decline"\|"tentative"`, `message?: str` |

### Tool Response Format

All tools return structured dicts optimized for LLM consumption:

```python
# Success — list
{
    "messages": [
        {
            "id": "AAMkAG...",
            "subject": "Q2 Planning",
            "from": "boss@company.com",
            "received": "2026-04-12T10:30:00Z",
            "is_read": False,
            "importance": "high",
            "preview": "Let's sync on the Q2 roadmap..."
        }
    ],
    "count": 25,
    "has_more": True
}

# Success — single item
{
    "id": "AAMkAG...",
    "subject": "Q2 Planning",
    "from": {"name": "Boss Name", "email": "boss@company.com"},
    "to": [{"name": "You", "email": "you@outlook.com"}],
    "received": "2026-04-12T10:30:00Z",
    "body": "Plain text body content...",
    "body_html": "<p>HTML body if requested...</p>",
    "attachments": [{"id": "...", "name": "file.pdf", "size": 1024}],
    "conversation_id": "AAQkAG...",
    "is_read": True,
    "importance": "normal",
    "categories": ["Blue Category"],
    "flag": "notFlagged"
}

# Success — action
{
    "status": "sent",
    "message_id": "AAMkAG..."
}

# Errors are raised as exceptions (OutlookMCPError), not returned as dicts.
# FastMCP converts them to MCP error responses automatically.
```

---

## Tier 2: Power Features (31 tools)

**Goal:** Complete the tool surface for serious agent workflows. Attachments, To Do, contacts, drafts, threading, pagination, batch operations, plus promoted Tier 3 tools.

| # | Tool | Description | Graph Endpoint | Parameters |
|---|------|-------------|----------------|------------|
| | **Auth** | | | |
| 22 | `outlook_list_accounts` | List authenticated accounts (multi-account support) | Local only | — |
| | **Mail — Attachments** | | | |
| 23 | `outlook_list_attachments` | List attachments on a message | `GET /me/messages/{id}/attachments` | `message_id: str` |
| 24 | `outlook_download_attachment` | Download attachment to disk | `GET /me/messages/{id}/attachments/{att_id}` | `message_id: str`, `attachment_id: str`, `save_path?: str` |
| 25 | `outlook_send_with_attachments` | Send with files (auto upload session >3MB) | `POST /me/sendMail` + `createUploadSession` | Same as `send_message` + `attachment_paths: list[str]` |
| | **Mail — Drafts** | | | |
| 26 | `outlook_list_drafts` | List draft messages | `GET /me/mailFolders/drafts/messages` | `count?: int (25)` |
| 27 | `outlook_create_draft` | Create draft (supports reply/forward-with-attach flow) | `POST /me/messages` | Same as `send_message` params |
| 28 | `outlook_update_draft` | Update draft fields | `PATCH /me/messages/{id}` | `draft_id: str`, `subject?: str`, `body?: str`, `to?: list[str]`, `cc?: list[str]` |
| 29 | `outlook_send_draft` | Send existing draft | `POST /me/messages/{id}/send` | `draft_id: str` |
| 30 | `outlook_delete_draft` | Delete draft | `DELETE /me/messages/{id}` | `draft_id: str` |
| | **Mail — Folders** | | | |
| 31 | `outlook_create_folder` | Create mail folder | `POST /me/mailFolders` | `name: str`, `parent_folder?: str` |
| 32 | `outlook_rename_folder` | Rename mail folder | `PATCH /me/mailFolders/{id}` | `folder_id: str`, `name: str` |
| 33 | `outlook_delete_folder` | Delete mail folder | `DELETE /me/mailFolders/{id}` | `folder_id: str` |
| | **Mail — Threading** | | | |
| 34 | `outlook_list_thread` | Get all messages in a conversation | `GET /me/messages?$filter=conversationId eq '...'` | `conversation_id: str`, `count?: int (50)` |
| | **Mail — Pre-Send** | | | |
| 35 | `outlook_get_mail_tips` | Pre-send check (OOF, mailbox full, invalid recipient) | `POST /me/getMailTips` | `emails: list[str]` |
| 36 | `outlook_copy_message` | Copy message to a folder | `POST /me/messages/{id}/copy` | `message_id: str`, `folder: str` |
| | **Mail — Send Options (additive params)** | | | |
| | `outlook_send_message` gains | `sensitivity?: "normal"\|"personal"\|"private"\|"confidential"`, `request_read_receipt?: bool` | | Added to existing tool |
| | **Contacts** | | | |
| 37 | `outlook_list_contacts` | List contacts | `GET /me/contacts` | `count?: int (25)`, `skip?: int` |
| 38 | `outlook_search_contacts` | Search by name or email | `GET /me/contacts?$search=` | `query: str`, `count?: int` |
| 39 | `outlook_get_contact` | Get full contact details | `GET /me/contacts/{id}` | `contact_id: str` |
| 40 | `outlook_create_contact` | Create contact | `POST /me/contacts` | `first_name: str`, `last_name?: str`, `email?: str`, `phone?: str`, `company?: str`, `title?: str` |
| 41 | `outlook_update_contact` | Update contact fields | `PATCH /me/contacts/{id}` | `contact_id: str`, `first_name?: str`, `last_name?: str`, `email?: str`, `phone?: str` |
| 42 | `outlook_delete_contact` | Delete contact | `DELETE /me/contacts/{id}` | `contact_id: str` |
| | **To Do** | | | |
| 43 | `outlook_list_task_lists` | List To Do lists | `GET /me/todo/lists` | — |
| 44 | `outlook_list_tasks` | List tasks with status filter | `GET /me/todo/lists/{id}/tasks` | `list_id?: str`, `status?: "notStarted"\|"inProgress"\|"completed"`, `count?: int` |
| 45 | `outlook_create_task` | Create task (due, reminder, importance, recurrence) | `POST /me/todo/lists/{id}/tasks` | `title: str`, `list_id?: str`, `due?: str`, `importance?: str`, `body?: str`, `reminder?: str`, `recurrence?: str` |
| 46 | `outlook_update_task` | Update task fields | `PATCH /me/todo/lists/{id}/tasks/{id}` | `task_id: str`, `list_id?: str`, `title?: str`, `due?: str`, `body?: str`, `importance?: str` |
| 47 | `outlook_complete_task` | Mark completed | `PATCH /me/todo/lists/{id}/tasks/{id}` | `task_id: str`, `list_id?: str` |
| 48 | `outlook_delete_task` | Delete task | `DELETE /me/todo/lists/{id}/tasks/{id}` | `task_id: str`, `list_id?: str` |
| | **Calendar** | | | |
| 49 | `outlook_list_calendars` | List available calendars | `GET /me/calendars` | — |
| | **Batch** | | | |
| 50 | `outlook_batch_triage` | Batch move/flag/categorize/mark read (max 20 per call) | `POST /$batch` | `message_ids: list[str]`, `action: str`, `value: str` |
| | **Promoted from Tier 3** | | | |
| 51 | `outlook_list_categories` | List category definitions with colors | `GET /me/outlook/masterCategories` | — |
| 52 | `outlook_whoami` | Current user profile (name, email) | `GET /me` | — |

### Pagination Design

Tier 2 adds cursor-based pagination to all list tools:

```python
# Request
outlook_list_inbox(count=25, cursor="eyJza2lwIjoyNX0=")

# Response
{
    "messages": [...],
    "count": 25,
    "cursor": "eyJza2lwIjo1MH0=",  # base64-encoded skip token
    "has_more": True
}
```

Implementation: Wrap Graph API's `@odata.nextLink` into opaque base64 cursors. The agent passes cursor back to get the next page. No raw URLs exposed.

### Batch Operations Design

Use Graph API `$batch` endpoint for `outlook_batch_triage`:

```
POST https://graph.microsoft.com/v1.0/$batch
{
    "requests": [
        {"id": "1", "method": "PATCH", "url": "/me/messages/{id1}", "body": {"isRead": true}},
        {"id": "2", "method": "PATCH", "url": "/me/messages/{id2}", "body": {"isRead": true}},
        ...
    ]
}
```

- Max 20 requests per batch (Graph API limit)
- Graceful fallback: if batch fails, retry per-item (Gmail MCP pattern)
- Return per-item success/failure status

---

## Tier 3: Differentiators (17 tools)

**Goal:** Features that no other Outlook integration has. Real-time notifications, intelligent inbox awareness, advanced management, checklists.

| # | Tool | Description | Graph Endpoint | Parameters |
|---|------|-------------|----------------|------------|
| | **Focused Inbox** | | | |
| 53 | `outlook_list_focused` | List Focused or Other tab | `GET /me/mailFolders/inbox/messages?$filter=inferenceClassification eq '...'` | `tab: "focused"\|"other"`, `count?: int`, `unread_only?: bool` |
| 54 | `outlook_move_to_focused` | Override classification (move to Focused/Other) | `PATCH /me/messages/{id}` | `message_id: str`, `classification: "focused"\|"other"` |
| | **Categories** | | | |
| 55 | `outlook_create_category` | Create custom category | `POST /me/outlook/masterCategories` | `name: str`, `color?: str (preset0-preset24)` |
| 56 | `outlook_delete_category` | Delete category | `DELETE /me/outlook/masterCategories/{id}` | `category_id: str` |
| | **Out-of-Office** (personal accounts: basic support) | | | |
| 57 | `outlook_get_ooo` | Get auto-reply settings | `GET /me/mailboxSettings/automaticRepliesSetting` | — |
| 58 | `outlook_set_ooo` | Set or disable auto-reply | `PATCH /me/mailboxSettings` | `enabled: bool`, `message?: str`, `start?: str`, `end?: str`, `external_message?: str` |
| | **Inbox Rules** (personal accounts: limited support) | | | |
| 59 | `outlook_list_rules` | List inbox rules | `GET /me/mailFolders/inbox/messageRules` | — |
| 60 | `outlook_create_rule` | Create inbox rule | `POST /me/mailFolders/inbox/messageRules` | `name: str`, `from_address?: str`, `subject_contains?: str`, `move_to?: str`, `mark_read?: bool`, `forward_to?: str` |
| 61 | `outlook_delete_rule` | Delete inbox rule | `DELETE /me/mailFolders/inbox/messageRules/{id}` | `rule_id: str` |
| | **Advanced Mail** | | | |
| 62 | `outlook_get_raw_message` | Get raw MIME content (for archiving/forwarding-as-attachment) | `GET /me/messages/{id}/$value` | `message_id: str` |
| 63 | `outlook_get_message_headers` | Get internet message headers (spam/phishing analysis) | `GET /me/messages/{id}?$select=internetMessageHeaders` | `message_id: str` |
| | **Calendar — Advanced** | | | |
| 64 | `outlook_cancel_event` | Cancel event (sends cancellation to attendees) | `POST /me/events/{id}/cancel` | `event_id: str`, `message?: str` |
| | **To Do — Checklists** | | | |
| 65 | `outlook_list_checklist_items` | List checklist items on a task | `GET /me/todo/lists/{id}/tasks/{id}/checklistItems` | `task_id: str`, `list_id?: str` |
| 66 | `outlook_create_checklist_item` | Create checklist item | `POST /me/todo/lists/{id}/tasks/{id}/checklistItems` | `task_id: str`, `list_id?: str`, `display_name: str` |
| 67 | `outlook_toggle_checklist_item` | Toggle checked/unchecked | `PATCH /me/todo/lists/{id}/tasks/{id}/checklistItems/{id}` | `task_id: str`, `item_id: str`, `list_id?: str` |
| | **Notifications** | | | |
| 68 | `outlook_subscribe` | Change notifications (poll-based V1, webhook V2) | `POST /subscriptions` or poll loop | `resource: "mail"\|"calendar"`, `change_type: "created"\|"updated"\|"deleted"` |
| 69 | `outlook_unsubscribe` | Remove notification subscription | `DELETE /subscriptions/{id}` | `subscription_id: str` |

**Notification implementation note:** Graph API subscriptions require a publicly reachable webhook URL. For MCP, recommended approach is poll-based notifications — a background task polls Graph every N seconds and surfaces changes via MCP notification mechanism. Less real-time, but zero infrastructure requirements.

---

## Future: Enterprise / Entra ID (10 tools)

> This section captures enterprise requirements for future implementation. NOT in scope for V1-V3.

### Authentication Changes

```python
# Personal (V1 - current)
credential = DeviceCodeCredential(
    client_id=CLIENT_ID,
    tenant_id="consumers"        # Personal accounts only
)

# Enterprise (future)
credential = DeviceCodeCredential(
    client_id=CLIENT_ID,          # BYOID
    tenant_id="organizations"     # Or specific tenant UUID
)
```

### Additional Scopes

| Scope | Feature |
|-------|---------|
| `MailboxSettings.ReadWrite` | Full OOF, inbox rules |
| `People.Read` | Organization directory search |
| `People.Read.All` | Full directory access |
| `Calendars.Read.Shared` | View shared/delegate calendars |
| `Mail.Read.Shared` | Access shared mailboxes |
| `User.ReadBasic.All` | People picker / attendee lookup |

### Enterprise-Only Tools

| # | Tool | Description | Graph Endpoint |
|---|------|-------------|----------------|
| 70 | `outlook_check_availability` | Free/busy for email addresses (requires org permissions) | `POST /me/calendar/getSchedule` |
| 71 | `outlook_search_directory` | Search org directory for people | `GET /me/people` |
| 72 | `outlook_find_meeting_times` | Find available times across attendees | `POST /me/findMeetingTimes` |
| 73 | `outlook_list_shared_mailboxes` | List shared mailboxes user has access to | `GET /me/mailFolders` (shared) |
| 74 | `outlook_read_shared_inbox` | Read from shared mailbox | `GET /users/{id}/messages` |
| 75 | `outlook_list_shared_calendars` | List calendars shared with user | `GET /me/calendars?$filter=canShare` |
| 76 | `outlook_send_as` | Send on behalf of shared mailbox | `POST /users/{id}/sendMail` |
| 77 | `outlook_manage_delegates` | Manage calendar delegate permissions | `PATCH /me/calendar/calendarPermissions` |
| 78 | `outlook_create_task_list` | Create To Do list | `POST /me/todo/lists` |
| 79 | `outlook_delete_task_list` | Delete To Do list | `DELETE /me/todo/lists/{id}` |

### Admin Consent Requirements

Enterprise deployments will require:
1. Azure AD admin consent for org-wide permissions
2. Documentation for IT admins on app registration
3. Conditional Access Policy compatibility
4. Token binding / CAE (Continuous Access Evaluation) support via azure-identity

### Multi-Tenant App Registration

For ClawHub distribution to enterprise users:
- Register app as multi-tenant in Azure AD
- Provide admin consent URL in docs
- Support per-tenant restrictions via config
- Document required Graph API permissions per tool

---

## Not Tiered — Evaluate Later

These exist in the Graph API but may not justify dedicated tools. Revisit based on community feedback.

| Capability | Graph Endpoint | Notes |
|---|---|---|
| Task attachments (upload/download/list/delete) | `/me/todo/lists/{id}/tasks/{id}/attachments` | olkcli has full CRUD. Low agent demand? |
| Task linked resources | `/me/todo/lists/{id}/tasks/{id}/linkedResources` | olkcli has it. Niche. |
| Extended properties on messages | `singleValueExtendedProperties` | Agent metadata storage on messages. Power use. |
| Open extensions | `/me/messages/{id}/extensions` | Custom data on messages. |
| Update inbox rule | `PATCH /me/mailFolders/inbox/messageRules/{id}` | Create + delete may be enough. |
| Contact folders | `/me/contactFolders` | Most personal accounts don't use these. |
| Calendar groups | `/me/calendarGroups` | Rare on personal accounts. |
| Schedule send / delayed delivery | `singleValueExtendedProperties` (deferred send) | Hacky via extended properties. Not clean API support. |
| Focused Inbox sender overrides (permanent) | `POST /me/inferenceClassification/overrides` | Train Focused per-sender. Different from one-off move. |
| Message importance update (on existing) | `PATCH /me/messages/{id}` | Already possible via generic PATCH. Add as param to triage? |

---

## Implementation Tasks

### Task 1: Project Scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `src/outlook_mcp/__init__.py`
- Create: `src/outlook_mcp/server.py`
- Create: `src/outlook_mcp/errors.py`
- Create: `tests/conftest.py`
- Create: `CLAUDE.md`
- Create: `LICENSE`
- Create: `.github/workflows/ci.yml`

**Step 1: Initialize project with uv**

```bash
cd ~/ClaudeCode/outlook-mcp
git init
uv init --lib --name outlook-mcp
```

**Step 2: Configure pyproject.toml**

```toml
[project]
name = "outlook-mcp"
version = "0.1.0"
description = "MCP server for Microsoft Outlook via Microsoft Graph API"
readme = "README.md"
license = "MIT"
requires-python = ">=3.10"
authors = [
    { name = "Michael Palermiti" }
]
keywords = ["mcp", "outlook", "microsoft-graph", "email", "calendar", "openclaw"]
classifiers = [
    "Development Status :: 3 - Alpha",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
]
dependencies = [
    "mcp[cli]>=1.27",
    "msgraph-sdk>=1.55",
    "azure-identity>=1.25",
    "pydantic>=2.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-asyncio>=0.24",
    "pytest-cov>=5.0",
    "ruff>=0.4",
]

[project.scripts]
outlook-mcp = "outlook_mcp.server:main"

[tool.ruff]
target-version = "py310"
line-length = 100

[tool.ruff.lint]
select = ["E", "F", "I", "N", "W", "UP"]

[tool.pytest.ini_options]
testpaths = ["tests"]
asyncio_mode = "auto"
```

**Step 3: Create minimal server entry point with lifespan**

```python
# src/outlook_mcp/__init__.py
"""Outlook MCP Server — Microsoft Outlook integration via Graph API."""

__version__ = "0.1.0"
```

```python
# src/outlook_mcp/server.py
"""FastMCP server for Microsoft Outlook."""

from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP


@asynccontextmanager
async def lifespan(server):
    """Initialize server state: config and auth manager."""
    # Will be populated when config.py and auth.py are built
    yield {}


mcp = FastMCP(
    "outlook-mcp",
    version="0.1.0",
    description="MCP server for Microsoft Outlook via Microsoft Graph API",
    lifespan=lifespan,
)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
```

**Step 4: Create error hierarchy**

```python
# src/outlook_mcp/errors.py
"""Exception hierarchy for outlook-mcp."""


class OutlookMCPError(Exception):
    """Base exception for all outlook-mcp errors."""

    def __init__(self, code: str, message: str, action: str | None = None):
        self.code = code
        self.message = message
        self.action = action
        super().__init__(message)


class AuthRequiredError(OutlookMCPError):
    """Raised when a tool is called without authentication."""

    def __init__(self):
        super().__init__(
            "auth_required",
            "Not authenticated. No valid credential found.",
            "Call outlook_login to authenticate with your Microsoft account.",
        )


class ReadOnlyError(OutlookMCPError):
    """Raised when a write tool is called in read-only mode."""

    def __init__(self, tool_name: str):
        super().__init__(
            "read_only",
            f"Cannot use {tool_name} — server is in read-only mode.",
            "Set read_only to false in ~/.outlook-mcp/config.json to enable write operations.",
        )


class NotFoundError(OutlookMCPError):
    """Raised when a requested resource doesn't exist."""

    def __init__(self, resource: str, resource_id: str):
        super().__init__(
            "not_found",
            f"{resource} '{resource_id}' not found.",
            None,
        )


class GraphAPIError(OutlookMCPError):
    """Raised when the Graph API returns an error."""

    def __init__(self, status_code: int, error_code: str, message: str):
        action = None
        if status_code == 401:
            action = "Token may have expired. Try outlook_login to re-authenticate."
        elif status_code == 429:
            action = "Rate limited by Microsoft Graph. Wait a moment and retry."
        super().__init__(
            f"graph_api_{error_code}",
            message,
            action,
        )
        self.status_code = status_code
```

**Step 5: Create test skeleton**

```python
# tests/conftest.py
"""Shared test fixtures for outlook-mcp."""

import pytest


@pytest.fixture
def mock_graph_client():
    """Mock Microsoft Graph client for unit tests."""
    # Will be implemented when graph.py is built
    pass
```

**Step 6: Create CLAUDE.md**

```markdown
# Outlook MCP Server

## What This Is
MCP server for Microsoft Outlook personal accounts (Outlook.com/Hotmail) via Microsoft Graph API.
Published to ClawHub.ai for the OpenClaw community.

## Tech Stack
- Python 3.10+, FastMCP, msgraph-sdk, azure-identity, Pydantic v2
- Package manager: uv
- Testing: pytest + pytest-asyncio

## Commands
- `uv run pytest` — run tests
- `uv run ruff check src/ tests/` — lint
- `uv run ruff format src/ tests/` — format
- `uv run outlook-mcp` — start server (stdio)

## Architecture
- `src/outlook_mcp/server.py` — FastMCP entry point, lifespan context
- `src/outlook_mcp/auth.py` — Device code OAuth2 via azure-identity
- `src/outlook_mcp/graph.py` — Graph client factory
- `src/outlook_mcp/config.py` — Config file management (~/.outlook-mcp/)
- `src/outlook_mcp/validation.py` — Input validation (OData, KQL, IDs, datetimes)
- `src/outlook_mcp/errors.py` — Exception hierarchy
- `src/outlook_mcp/tools/` — One file per tool group
- `src/outlook_mcp/models/` — Pydantic models for I/O

## Conventions
- One tool = one operation (not grouped CRUD)
- Tool names prefixed with `outlook_`
- All input validated via Pydantic + validation.py before Graph API calls
- No telemetry, no local caching, no third-party calls
- Tests: TDD, pytest, mock Graph client for unit tests
- Errors: raise OutlookMCPError subclasses, never return error dicts
- Datetimes: UTC in responses, config timezone for input interpretation
- Delete: soft delete (move to Deleted Items) by default
```

**Step 7: Create LICENSE (MIT)**

**Step 8: Install dependencies**

```bash
uv sync --all-extras
```

**Step 9: Verify test runner works**

```bash
uv run pytest -v
```
Expected: 0 tests collected, no errors.

**Step 10: Commit**

```bash
git add -A
git commit -m "feat: project scaffolding with FastMCP entry point, error hierarchy"
```

---

### Task 2: Config Management

**Files:**
- Create: `src/outlook_mcp/config.py`
- Create: `tests/test_config.py`

**Step 1: Write failing tests**

```python
# tests/test_config.py
"""Tests for config management."""

import json
import os

import pytest

from outlook_mcp.config import Config, load_config, save_config


def test_default_config():
    """Default config has sensible values."""
    config = Config()
    assert config.client_id is None  # BYOID — user must provide
    assert config.tenant_id == "consumers"
    assert config.read_only is False
    assert config.timezone == "UTC"


def test_config_requires_client_id_for_auth():
    """Config without client_id should be loadable but flagged."""
    config = Config()
    assert config.client_id is None


def test_config_dir_created(tmp_path, monkeypatch):
    """Config directory is created with 0700 permissions."""
    config_dir = tmp_path / ".outlook-mcp"
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(config_dir))
    save_config(Config(), config_dir=str(config_dir))
    assert config_dir.exists()
    assert oct(config_dir.stat().st_mode & 0o777) == "0o700"


def test_config_file_permissions(tmp_path, monkeypatch):
    """Config file is written with 0600 permissions."""
    config_dir = tmp_path / ".outlook-mcp"
    config_dir.mkdir(mode=0o700)
    save_config(Config(), config_dir=str(config_dir))
    config_file = config_dir / "config.json"
    assert config_file.exists()
    assert oct(config_file.stat().st_mode & 0o777) == "0o600"


def test_config_roundtrip(tmp_path):
    """Config saves and loads correctly."""
    config_dir = str(tmp_path / ".outlook-mcp")
    original = Config(
        client_id="my-app-uuid",
        timezone="America/Los_Angeles",
        read_only=True,
    )
    save_config(original, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert loaded.client_id == "my-app-uuid"
    assert loaded.timezone == "America/Los_Angeles"
    assert loaded.read_only is True


def test_config_rejects_symlink(tmp_path):
    """Config refuses to load from a symlinked file."""
    config_dir = tmp_path / ".outlook-mcp"
    config_dir.mkdir(mode=0o700)
    real_file = tmp_path / "evil_config.json"
    real_file.write_text(json.dumps({"timezone": "Evil/Zone"}))
    symlink = config_dir / "config.json"
    symlink.symlink_to(real_file)
    with pytest.raises(PermissionError, match="symlink"):
        load_config(config_dir=str(config_dir))


def test_config_override_client_id(tmp_path):
    """Client ID set via config."""
    config_dir = str(tmp_path / ".outlook-mcp")
    config = Config(client_id="custom-client-id-uuid")
    save_config(config, config_dir=config_dir)
    loaded = load_config(config_dir=config_dir)
    assert loaded.client_id == "custom-client-id-uuid"
```

**Step 2: Run tests to verify they fail**

```bash
uv run pytest tests/test_config.py -v
```
Expected: ImportError — `outlook_mcp.config` doesn't exist.

**Step 3: Implement config.py**

```python
# src/outlook_mcp/config.py
"""Config file management for outlook-mcp."""

import json
import os
import stat
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field

DEFAULT_TENANT_ID = "consumers"
DEFAULT_CONFIG_DIR = os.path.expanduser("~/.outlook-mcp")


class Config(BaseModel):
    """Outlook MCP server configuration."""

    client_id: str | None = Field(default=None, description="Azure AD app client ID (BYOID)")
    tenant_id: str = Field(default=DEFAULT_TENANT_ID)
    read_only: bool = Field(default=False)
    timezone: str = Field(default="UTC", description="IANA timezone for relative date computations")


def _ensure_dir(dir_path: str) -> Path:
    """Create config directory with 0700 permissions."""
    path = Path(dir_path)
    path.mkdir(parents=True, exist_ok=True)
    path.chmod(0o700)
    return path


def _atomic_write(file_path: Path, data: str) -> None:
    """Write file atomically with fsync, set 0600 permissions."""
    dir_path = file_path.parent
    fd, tmp_path = tempfile.mkstemp(dir=str(dir_path), suffix=".tmp")
    try:
        with os.fdopen(fd, "w") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp_path, stat.S_IRUSR | stat.S_IWUSR)  # 0600
        os.replace(tmp_path, str(file_path))
    except Exception:
        os.unlink(tmp_path)
        raise


def save_config(config: Config, config_dir: str = DEFAULT_CONFIG_DIR) -> None:
    """Save config to disk."""
    dir_path = _ensure_dir(config_dir)
    file_path = dir_path / "config.json"
    _atomic_write(file_path, config.model_dump_json(indent=2))


def load_config(config_dir: str = DEFAULT_CONFIG_DIR) -> Config:
    """Load config from disk. Returns defaults if no config file exists."""
    file_path = Path(config_dir) / "config.json"

    if not file_path.exists():
        return Config()

    # Reject symlinks (olkcli pattern)
    if file_path.is_symlink():
        raise PermissionError(f"Refusing to load symlinked config: {file_path}")

    # Warn and fix if permissions are too wide
    mode = file_path.stat().st_mode & 0o777
    if mode != 0o600:
        file_path.chmod(0o600)

    data = file_path.read_text()
    return Config.model_validate_json(data)
```

**Step 4: Run tests to verify they pass**

```bash
uv run pytest tests/test_config.py -v
```
Expected: All 7 tests pass.

**Step 5: Commit**

```bash
git add src/outlook_mcp/config.py tests/test_config.py
git commit -m "feat: config management with BYOID, file permissions, symlink rejection"
```

---

### Task 3: Input Validation

**Files:**
- Create: `src/outlook_mcp/validation.py`
- Create: `tests/test_validation.py`

**Step 1: Write failing tests**

```python
# tests/test_validation.py
"""Tests for input validation — ported from olkcli patterns."""

import pytest

from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_datetime,
    validate_email,
    validate_graph_id,
    validate_folder_name,
    validate_phone,
)


class TestGraphIdValidation:
    def test_valid_id(self):
        assert validate_graph_id("AAMkAGI2TG93AAA=") == "AAMkAGI2TG93AAA="

    def test_valid_id_with_slashes(self):
        assert validate_graph_id("AAMkAG/test+id=") == "AAMkAG/test+id="

    def test_rejects_empty(self):
        with pytest.raises(ValueError, match="empty"):
            validate_graph_id("")

    def test_rejects_too_long(self):
        with pytest.raises(ValueError, match="too long"):
            validate_graph_id("A" * 1025)

    def test_rejects_special_chars(self):
        with pytest.raises(ValueError, match="invalid"):
            validate_graph_id("id with spaces")

    def test_rejects_injection(self):
        with pytest.raises(ValueError, match="invalid"):
            validate_graph_id("../../etc/passwd")


class TestEmailValidation:
    def test_valid_email(self):
        assert validate_email("user@outlook.com") == "user@outlook.com"

    def test_rejects_no_at(self):
        with pytest.raises(ValueError):
            validate_email("notanemail")

    def test_rejects_injection(self):
        with pytest.raises(ValueError):
            validate_email("user@evil.com' OR 1=1--")


class TestDatetimeValidation:
    def test_valid_iso_utc(self):
        result = validate_datetime("2026-04-12T10:30:00Z")
        assert result == "2026-04-12T10:30:00Z"

    def test_valid_iso_with_offset(self):
        result = validate_datetime("2026-04-12T10:30:00+05:00")
        # Should parse and re-serialize to UTC
        assert "Z" in result or "+" in result  # Valid ISO output

    def test_valid_date_only(self):
        """Date-only input gets interpreted as midnight UTC."""
        result = validate_datetime("2026-04-12")
        assert "2026-04-12" in result

    def test_rejects_garbage(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("not-a-date")

    def test_rejects_injection(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("2026-04-12' OR 1=1--")

    def test_rejects_odata_injection(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("2026-04-12T00:00:00Z eq true")


class TestKqlSanitization:
    def test_simple_query(self):
        assert sanitize_kql("budget report") == '"budget report"'

    def test_strips_dangerous_chars(self):
        result = sanitize_kql('from:boss" OR (hack)')
        assert '"' not in result.strip('"')
        assert "(" not in result.strip('"')
        assert ")" not in result.strip('"')

    def test_preserves_alphanumeric(self):
        result = sanitize_kql("meeting notes 2026")
        assert "meeting" in result
        assert "notes" in result
        assert "2026" in result

    def test_strips_kql_operators(self):
        result = sanitize_kql("test & hack | evil")
        assert "&" not in result
        assert "|" not in result


class TestFolderNameValidation:
    def test_wellknown_folders(self):
        assert validate_folder_name("inbox") == "inbox"
        assert validate_folder_name("drafts") == "drafts"
        assert validate_folder_name("sentitems") == "sentitems"
        assert validate_folder_name("deleteditems") == "deleteditems"
        assert validate_folder_name("junkemail") == "junkemail"
        assert validate_folder_name("archive") == "archive"

    def test_case_insensitive_wellknown(self):
        assert validate_folder_name("Inbox") == "inbox"
        assert validate_folder_name("DRAFTS") == "drafts"

    def test_custom_folder_id(self):
        """Custom folder IDs pass through graph ID validation."""
        assert validate_folder_name("AAMkAGFolderId=") == "AAMkAGFolderId="

    def test_rejects_invalid(self):
        with pytest.raises(ValueError):
            validate_folder_name("../../evil")


class TestPhoneValidation:
    def test_valid_phone(self):
        assert validate_phone("+1 (555) 123-4567") == "+1 (555) 123-4567"

    def test_rejects_letters(self):
        with pytest.raises(ValueError):
            validate_phone("call me maybe")

    def test_rejects_too_long(self):
        with pytest.raises(ValueError):
            validate_phone("1" * 31)


class TestOutputSanitization:
    def test_strips_control_chars(self):
        assert sanitize_output("normal text") == "normal text"
        assert sanitize_output("evil\x1b[31mred\x1b[0m") == "evilred"
        assert sanitize_output("tab\there") == "tab here"

    def test_preserves_newlines_in_multiline(self):
        result = sanitize_output("line1\nline2", multiline=True)
        assert "\n" in result

    def test_strips_null_bytes(self):
        assert sanitize_output("null\x00byte") == "nullbyte"
```

**Step 2: Run tests to verify they fail**

```bash
uv run pytest tests/test_validation.py -v
```

**Step 3: Implement validation.py**

```python
# src/outlook_mcp/validation.py
"""Input validation — patterns ported from olkcli (MIT)."""

import re
from datetime import datetime, timezone

# Graph API entity ID pattern: alphanumeric, =, +, /, -
_GRAPH_ID_RE = re.compile(r"^[a-zA-Z0-9_=+/\-]{1,1024}$")

# Email pattern (simplified but sufficient for validation)
_EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

# Phone pattern
_PHONE_RE = re.compile(r"^[0-9 ()+.\-]{1,30}$")

# KQL dangerous characters to strip
_KQL_DANGEROUS = re.compile(r'[":()&|!*\\]')

# Control characters (C0 + C1 + DEL), excluding \n and \t
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
# ANSI escape sequences
_ANSI_ESCAPE = re.compile(r"\x1b\[[0-9;]*[a-zA-Z]")

# ISO 8601 datetime: strict pattern to reject injection attempts
_ISO_DATETIME_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}"           # date
    r"(?:T\d{2}:\d{2}:\d{2}"        # optional time
    r"(?:\.\d+)?"                    # optional fractional seconds
    r"(?:Z|[+-]\d{2}:\d{2})?"       # optional timezone
    r")?$"
)

WELL_KNOWN_FOLDERS = {
    "inbox", "drafts", "sentitems", "deleteditems",
    "junkemail", "archive", "outbox",
}


def validate_graph_id(value: str) -> str:
    """Validate a Microsoft Graph entity ID."""
    if not value:
        raise ValueError("Graph ID must not be empty")
    if len(value) > 1024:
        raise ValueError("Graph ID too long (max 1024 chars)")
    if not _GRAPH_ID_RE.match(value):
        raise ValueError(f"Graph ID contains invalid characters: {value[:50]}")
    return value


def validate_email(value: str) -> str:
    """Validate an email address."""
    if not _EMAIL_RE.match(value):
        raise ValueError(f"Invalid email address: {value[:50]}")
    return value


def validate_datetime(value: str) -> str:
    """Validate and re-serialize a datetime string.

    Accepts ISO 8601 formats. Rejects injection attempts.
    Returns a safe ISO 8601 string suitable for OData filters.
    """
    # First pass: reject anything that doesn't look like a date
    if not _ISO_DATETIME_RE.match(value):
        raise ValueError(f"Invalid datetime format: {value[:50]}")

    # Second pass: actually parse it to ensure validity
    try:
        if "T" in value:
            # Full datetime
            if value.endswith("Z"):
                dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
            else:
                dt = datetime.fromisoformat(value)
            # Re-serialize to UTC ISO 8601
            utc_dt = dt.astimezone(timezone.utc)
            return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            # Date only — validate by parsing
            dt = datetime.strptime(value, "%Y-%m-%d")
            return f"{value}T00:00:00Z"
    except (ValueError, OverflowError) as e:
        raise ValueError(f"Invalid datetime: {value[:50]}") from e


def sanitize_kql(query: str) -> str:
    """Sanitize a KQL search query to prevent injection."""
    sanitized = _KQL_DANGEROUS.sub("", query)
    return f'"{sanitized}"'


def validate_folder_name(name: str) -> str:
    """Validate a folder name — well-known names or Graph IDs."""
    lower = name.lower()
    if lower in WELL_KNOWN_FOLDERS:
        return lower
    return validate_graph_id(name)


def validate_phone(value: str) -> str:
    """Validate a phone number."""
    if not _PHONE_RE.match(value):
        raise ValueError(f"Invalid phone number: {value[:30]}")
    return value


def sanitize_output(text: str, multiline: bool = False) -> str:
    """Strip control characters and ANSI escapes from output text."""
    text = _ANSI_ESCAPE.sub("", text)
    text = _CONTROL_CHARS.sub("", text)
    if not multiline:
        text = text.replace("\n", " ").replace("\t", " ")
    else:
        text = text.replace("\t", " ")
    return text
```

**Step 4: Run tests**

```bash
uv run pytest tests/test_validation.py -v
```
Expected: All tests pass.

**Step 5: Commit**

```bash
git add src/outlook_mcp/validation.py tests/test_validation.py
git commit -m "feat: input validation with datetime parse/re-serialize, ported from olkcli"
```

---

### Task 4: Auth Module

**Files:**
- Create: `src/outlook_mcp/auth.py`
- Create: `tests/test_auth.py`

**Step 1: Write failing tests**

```python
# tests/test_auth.py
"""Tests for auth module."""

import pytest

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError


def test_auth_manager_init():
    """AuthManager initializes with config."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.config is config
    assert auth.credential is None


def test_auth_scopes_default():
    """Default scopes include read-write."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.ReadWrite" in scopes
    assert "Mail.Send" in scopes
    assert "Calendars.ReadWrite" in scopes
    assert "offline_access" in scopes


def test_auth_scopes_read_only():
    """Read-only mode uses read scopes."""
    config = Config(client_id="test-id", read_only=True)
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.Read" in scopes
    assert "Mail.ReadWrite" not in scopes
    assert "Mail.Send" not in scopes
    assert "Calendars.Read" in scopes
    assert "Calendars.ReadWrite" not in scopes


def test_auth_not_authenticated():
    """is_authenticated returns False before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False


def test_auth_get_credential_raises_when_not_authenticated():
    """get_credential raises AuthRequiredError before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    with pytest.raises(AuthRequiredError):
        auth.get_credential()


def test_auth_requires_client_id():
    """Login raises if client_id is not configured."""
    config = Config()  # No client_id
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="client_id"):
        auth.login()
```

**Step 2: Run tests to verify they fail**

```bash
uv run pytest tests/test_auth.py -v
```

**Step 3: Implement auth.py**

```python
# src/outlook_mcp/auth.py
"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import logging
from typing import Any

from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions

from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError

logger = logging.getLogger(__name__)

SCOPES_READWRITE = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
    "offline_access",
]

SCOPES_READONLY = [
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "Tasks.Read",
    "User.Read",
    "offline_access",
]


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._account_email: str | None = None

    def get_scopes(self) -> list[str]:
        """Return the appropriate scopes based on config."""
        return SCOPES_READONLY if self.config.read_only else SCOPES_READWRITE

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def login(self) -> dict[str, str]:
        """Start device code auth flow.

        Returns:
            Dict with 'status' and 'message' for the agent to display.

        Raises:
            ValueError: If client_id is not configured.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json. See README for setup instructions."
            )

        cache_options = TokenCachePersistenceOptions(name="outlook-mcp")

        # azure-identity prompt_callback receives a single dict argument
        def _on_device_code(device_code_info: dict) -> None:
            self._device_code_message = device_code_info.get("message", "")

        self.credential = DeviceCodeCredential(
            client_id=self.config.client_id,
            tenant_id=self.config.tenant_id,
            cache_persistence_options=cache_options,
            prompt_callback=_on_device_code,
        )

        return {
            "status": "login_started",
            "message": "Device code authentication initiated. "
            "Complete the sign-in when prompted.",
        }

    def get_credential(self) -> DeviceCodeCredential:
        """Get the current credential, raising if not authenticated."""
        if self.credential is None:
            raise AuthRequiredError()
        return self.credential

    def logout(self) -> dict[str, str]:
        """Clear stored credentials."""
        self.credential = None
        self._account_email = None
        return {"status": "logged_out", "message": "Credentials cleared."}
```

**Step 4: Run tests**

```bash
uv run pytest tests/test_auth.py -v
```

**Step 5: Commit**

```bash
git add src/outlook_mcp/auth.py tests/test_auth.py
git commit -m "feat: auth module with BYOID, device code flow, read-only scopes"
```

---

### Task 5: Graph Client Factory

**Files:**
- Create: `src/outlook_mcp/graph.py`
- Create: `tests/test_graph.py`

**Step 1: Write failing tests**

```python
# tests/test_graph.py
"""Tests for Graph client factory."""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.graph import GraphClient
from outlook_mcp.errors import AuthRequiredError


def test_graph_client_requires_credential():
    """GraphClient raises without credential."""
    with pytest.raises(AuthRequiredError):
        GraphClient(credential=None)


def test_graph_client_init():
    """GraphClient initializes with a credential."""
    mock_credential = MagicMock()
    client = GraphClient(credential=mock_credential)
    assert client.sdk_client is not None
```

**Step 2: Implement graph.py**

```python
# src/outlook_mcp/graph.py
"""Microsoft Graph client factory."""

from __future__ import annotations

from typing import Any

from msgraph import GraphServiceClient

from outlook_mcp.errors import AuthRequiredError


class GraphClient:
    """Wrapper around the Microsoft Graph SDK client."""

    def __init__(self, credential: Any) -> None:
        if credential is None:
            raise AuthRequiredError()
        self.sdk_client = GraphServiceClient(credentials=credential)
```

**Step 3: Run tests, commit**

```bash
uv run pytest tests/test_graph.py -v
git add src/outlook_mcp/graph.py tests/test_graph.py
git commit -m "feat: Graph client factory"
```

---

### Task 6: Pydantic Models

**Files:**
- Create: `src/outlook_mcp/models/__init__.py`
- Create: `src/outlook_mcp/models/common.py`
- Create: `src/outlook_mcp/models/mail.py`
- Create: `src/outlook_mcp/models/calendar.py`
- Create: `tests/test_models.py`

**Step 1: Write failing tests**

```python
# tests/test_models.py
"""Tests for Pydantic models."""

import pytest

from outlook_mcp.models.mail import (
    MessageSummary,
    MessageDetail,
    SendMessageInput,
    ReplyInput,
    ForwardInput,
    TriageInput,
)
from outlook_mcp.models.calendar import (
    EventSummary,
    EventDetail,
    CreateEventInput,
    RsvpInput,
)
from outlook_mcp.models.common import ListResponse


class TestMailModels:
    def test_send_message_validates_emails(self):
        msg = SendMessageInput(
            to=["valid@outlook.com"],
            subject="Test",
            body="Hello",
        )
        assert msg.to == ["valid@outlook.com"]

    def test_send_message_rejects_empty_to(self):
        with pytest.raises(ValueError):
            SendMessageInput(to=[], subject="Test", body="Hello")

    def test_send_message_defaults(self):
        msg = SendMessageInput(
            to=["a@b.com"], subject="Test", body="Hello"
        )
        assert msg.is_html is False
        assert msg.importance == "normal"
        assert msg.cc is None
        assert msg.bcc is None

    def test_triage_validates_flag_status(self):
        t = TriageInput(message_id="AAMkAG123=", action="flag", value="flagged")
        assert t.value == "flagged"

    def test_reply_input(self):
        r = ReplyInput(message_id="AAMkAG123=", body="Thanks!", reply_all=False)
        assert r.reply_all is False


class TestCalendarModels:
    def test_create_event_required_fields(self):
        e = CreateEventInput(
            subject="Meeting",
            start="2026-04-15T10:00:00",
            end="2026-04-15T11:00:00",
        )
        assert e.subject == "Meeting"
        assert e.is_all_day is False
        assert e.is_online is False

    def test_rsvp_validates_response(self):
        r = RsvpInput(event_id="AAMkAG123=", response="accept")
        assert r.response == "accept"

    def test_rsvp_rejects_invalid_response(self):
        with pytest.raises(ValueError):
            RsvpInput(event_id="AAMkAG123=", response="maybe")


class TestCommonModels:
    def test_list_response(self):
        resp = ListResponse(items=[{"id": "1"}], count=1, has_more=False)
        assert resp.count == 1
```

**Step 2: Implement models**

```python
# src/outlook_mcp/models/__init__.py
"""Pydantic models for Outlook MCP tool I/O."""
```

```python
# src/outlook_mcp/models/common.py
"""Shared models."""

from __future__ import annotations
from typing import Any

from pydantic import BaseModel, Field


class ListResponse(BaseModel):
    """Standard list response wrapper."""
    items: list[Any]
    count: int
    has_more: bool = False
    cursor: str | None = None
```

```python
# src/outlook_mcp/models/mail.py
"""Mail-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator


class MessageSummary(BaseModel):
    """Compact message representation for list results."""
    id: str
    subject: str
    from_email: str
    from_name: str = ""
    received: str
    is_read: bool
    importance: str = "normal"
    preview: str = ""
    has_attachments: bool = False
    categories: list[str] = Field(default_factory=list)
    flag: str = "notFlagged"
    conversation_id: str = ""


class MessageDetail(BaseModel):
    """Full message representation."""
    id: str
    subject: str
    from_email: str
    from_name: str = ""
    to: list[dict[str, str]] = Field(default_factory=list)
    cc: list[dict[str, str]] = Field(default_factory=list)
    received: str
    body: str = ""
    body_html: str | None = None
    is_read: bool
    importance: str = "normal"
    has_attachments: bool = False
    attachments: list[dict[str, str]] = Field(default_factory=list)
    categories: list[str] = Field(default_factory=list)
    flag: str = "notFlagged"
    conversation_id: str = ""


class SendMessageInput(BaseModel):
    """Input for sending a message."""
    to: list[str] = Field(min_length=1)
    subject: str
    body: str
    cc: list[str] | None = None
    bcc: list[str] | None = None
    is_html: bool = False
    importance: str = "normal"

    @field_validator("importance")
    @classmethod
    def validate_importance(cls, v: str) -> str:
        if v not in ("low", "normal", "high"):
            raise ValueError(f"importance must be low, normal, or high; got {v}")
        return v


class ReplyInput(BaseModel):
    """Input for replying to a message."""
    message_id: str
    body: str
    reply_all: bool = False
    is_html: bool = False


class ForwardInput(BaseModel):
    """Input for forwarding a message."""
    message_id: str
    to: list[str] = Field(min_length=1)
    comment: str | None = None


class TriageInput(BaseModel):
    """Input for triage actions (move, flag, categorize, mark read)."""
    message_id: str
    action: str
    value: str


class DeleteInput(BaseModel):
    """Input for deleting a message (soft or hard)."""
    message_id: str
    permanent: bool = False
```

```python
# src/outlook_mcp/models/calendar.py
"""Calendar-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator


class EventSummary(BaseModel):
    """Compact event representation for list results."""
    id: str
    subject: str
    start: str
    end: str
    location: str = ""
    is_all_day: bool = False
    organizer: str = ""
    response_status: str = ""
    is_online: bool = False


class EventDetail(BaseModel):
    """Full event representation."""
    id: str
    subject: str
    start: str
    end: str
    location: str = ""
    body: str = ""
    is_all_day: bool = False
    organizer: dict[str, str] = Field(default_factory=dict)
    attendees: list[dict[str, str]] = Field(default_factory=list)
    response_status: str = ""
    is_online: bool = False
    online_meeting_url: str | None = None
    recurrence: dict | None = None
    categories: list[str] = Field(default_factory=list)


class CreateEventInput(BaseModel):
    """Input for creating a calendar event."""
    subject: str
    start: str
    end: str
    location: str | None = None
    body: str | None = None
    attendees: list[str] | None = None
    is_all_day: bool = False
    is_online: bool = False
    recurrence: str | None = None

    @field_validator("recurrence")
    @classmethod
    def validate_recurrence(cls, v: str | None) -> str | None:
        if v is not None and v not in ("daily", "weekdays", "weekly", "monthly", "yearly"):
            raise ValueError(f"recurrence must be daily/weekdays/weekly/monthly/yearly; got {v}")
        return v


class UpdateEventInput(BaseModel):
    """Input for updating an event."""
    event_id: str
    subject: str | None = None
    start: str | None = None
    end: str | None = None
    location: str | None = None
    body: str | None = None


class RsvpInput(BaseModel):
    """Input for RSVPing to an event."""
    event_id: str
    response: str
    message: str | None = None

    @field_validator("response")
    @classmethod
    def validate_response(cls, v: str) -> str:
        if v not in ("accept", "decline", "tentative"):
            raise ValueError(f"response must be accept/decline/tentative; got {v}")
        return v
```

**Step 3: Run tests, commit**

```bash
uv run pytest tests/test_models.py -v
git add src/outlook_mcp/models/ tests/test_models.py
git commit -m "feat: Pydantic models for mail and calendar I/O"
```

---

### Task 7: Mail Read Tools

**Files:**
- Create: `src/outlook_mcp/tools/__init__.py`
- Create: `src/outlook_mcp/tools/mail_read.py`
- Create: `tests/test_mail_read.py`

**Step 1: Write failing tests**

```python
# tests/test_mail_read.py
"""Tests for mail read tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.tools.mail_read import list_inbox, read_message, search_mail


class TestListInbox:
    @pytest.mark.asyncio
    async def test_list_inbox_returns_messages(self):
        """list_inbox returns structured message list."""
        mock_message = MagicMock()
        mock_message.id = "AAMkAG123="
        mock_message.subject = "Test Subject"
        mock_message.from_ = MagicMock()
        mock_message.from_.email_address = MagicMock()
        mock_message.from_.email_address.address = "sender@test.com"
        mock_message.from_.email_address.name = "Sender"
        mock_message.received_date_time = "2026-04-12T10:00:00Z"
        mock_message.is_read = False
        mock_message.importance = MagicMock(value="normal")
        mock_message.body_preview = "Preview text..."
        mock_message.has_attachments = False
        mock_message.categories = []
        mock_message.flag = MagicMock()
        mock_message.flag.flag_status = MagicMock(value="notFlagged")
        mock_message.conversation_id = "conv123"

        mock_client = AsyncMock()
        mock_client.me.mail_folders.by_mail_folder_id.return_value.messages.get = AsyncMock(
            return_value=MagicMock(value=[mock_message], odata_next_link=None)
        )

        result = await list_inbox(mock_client, folder="inbox", count=25)
        assert result["count"] == 1
        assert result["messages"][0]["subject"] == "Test Subject"
        assert result["messages"][0]["from_email"] == "sender@test.com"
        assert result["messages"][0]["is_read"] is False

    @pytest.mark.asyncio
    async def test_list_inbox_validates_count(self):
        """Count is clamped to 1-100."""
        mock_client = AsyncMock()
        mock_client.me.mail_folders.by_mail_folder_id.return_value.messages.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )
        result = await list_inbox(mock_client, count=200)
        assert result["count"] == 0

    @pytest.mark.asyncio
    async def test_list_inbox_validates_dates(self):
        """Date params are validated via validate_datetime."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError, match="Invalid datetime"):
            await list_inbox(mock_client, after="not-a-date")


class TestSearchMail:
    @pytest.mark.asyncio
    async def test_search_sanitizes_query(self):
        """Search query is sanitized before sending to Graph."""
        mock_client = AsyncMock()
        mock_client.me.messages.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )
        result = await search_mail(mock_client, query='test" OR (hack)')
        assert result["count"] == 0
```

**Step 2: Implement mail_read.py**

```python
# src/outlook_mcp/tools/__init__.py
"""Tool registration for Outlook MCP."""

# src/outlook_mcp/tools/mail_read.py
"""Mail read tools: list_inbox, read_message, search_mail, list_folders."""

from __future__ import annotations

from typing import Any

from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_datetime,
    validate_email,
    validate_folder_name,
    validate_graph_id,
)


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _format_message_summary(msg: Any) -> dict:
    """Convert Graph SDK message to summary dict."""
    from_addr = ""
    from_name = ""
    if msg.from_ and msg.from_.email_address:
        from_addr = msg.from_.email_address.address or ""
        from_name = msg.from_.email_address.name or ""

    flag_status = "notFlagged"
    if msg.flag and msg.flag.flag_status:
        flag_status = msg.flag.flag_status.value if hasattr(msg.flag.flag_status, "value") else str(msg.flag.flag_status)

    importance = "normal"
    if msg.importance:
        importance = msg.importance.value if hasattr(msg.importance, "value") else str(msg.importance)

    return {
        "id": msg.id,
        "subject": sanitize_output(msg.subject or "(no subject)"),
        "from_email": from_addr,
        "from_name": sanitize_output(from_name),
        "received": str(msg.received_date_time or ""),
        "is_read": bool(msg.is_read),
        "importance": importance,
        "preview": sanitize_output(msg.body_preview or ""),
        "has_attachments": bool(msg.has_attachments),
        "categories": list(msg.categories or []),
        "flag": flag_status,
        "conversation_id": msg.conversation_id or "",
    }


async def list_inbox(
    graph_client: Any,
    folder: str = "inbox",
    count: int = 25,
    unread_only: bool = False,
    from_address: str | None = None,
    after: str | None = None,
    before: str | None = None,
    skip: int = 0,
) -> dict:
    """List messages in a folder."""
    count = _clamp(count, 1, 100)
    folder = validate_folder_name(folder)

    query_params = {
        "$top": count,
        "$skip": skip,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,isRead,importance,bodyPreview,hasAttachments,categories,flag,conversationId",
    }

    # Build filter with validated inputs
    filters = []
    if unread_only:
        filters.append("isRead eq false")
    if from_address:
        validate_email(from_address)
        safe_from = from_address.replace("'", "''")
        filters.append(f"from/emailAddress/address eq '{safe_from}'")
    if after:
        safe_after = validate_datetime(after)
        filters.append(f"receivedDateTime ge {safe_after}")
    if before:
        safe_before = validate_datetime(before)
        filters.append(f"receivedDateTime le {safe_before}")

    if filters:
        query_params["$filter"] = " and ".join(filters)

    response = await graph_client.me.mail_folders.by_mail_folder_id(
        folder
    ).messages.get(query_params=query_params)

    messages = [_format_message_summary(m) for m in (response.value or [])]

    return {
        "messages": messages,
        "count": len(messages),
        "has_more": response.odata_next_link is not None,
    }


async def read_message(
    graph_client: Any,
    message_id: str,
    format: str = "text",
) -> dict:
    """Read a single message by ID."""
    message_id = validate_graph_id(message_id)

    msg = await graph_client.me.messages.by_message_id(message_id).get()

    from_addr = ""
    from_name = ""
    if msg.from_ and msg.from_.email_address:
        from_addr = msg.from_.email_address.address or ""
        from_name = msg.from_.email_address.name or ""

    to_list = []
    for r in (msg.to_recipients or []):
        if r.email_address:
            to_list.append({
                "name": sanitize_output(r.email_address.name or ""),
                "email": r.email_address.address or "",
            })

    cc_list = []
    for r in (msg.cc_recipients or []):
        if r.email_address:
            cc_list.append({
                "name": sanitize_output(r.email_address.name or ""),
                "email": r.email_address.address or "",
            })

    body_text = ""
    body_html = None
    if msg.body:
        content = msg.body.content or ""
        if format == "html" or format == "full":
            body_html = content
        if format == "text" or format == "full":
            body_text = sanitize_output(content, multiline=True)

    attachments = []
    for att in (msg.attachments or []):
        attachments.append({
            "id": att.id,
            "name": sanitize_output(att.name or ""),
            "size": att.size or 0,
        })

    return {
        "id": msg.id,
        "subject": sanitize_output(msg.subject or "(no subject)"),
        "from_email": from_addr,
        "from_name": sanitize_output(from_name),
        "to": to_list,
        "cc": cc_list,
        "received": str(msg.received_date_time or ""),
        "body": body_text,
        "body_html": body_html,
        "is_read": bool(msg.is_read),
        "importance": msg.importance.value if msg.importance and hasattr(msg.importance, "value") else "normal",
        "has_attachments": bool(msg.has_attachments),
        "attachments": attachments,
        "categories": list(msg.categories or []),
        "flag": msg.flag.flag_status.value if msg.flag and msg.flag.flag_status and hasattr(msg.flag.flag_status, "value") else "notFlagged",
        "conversation_id": msg.conversation_id or "",
    }


async def search_mail(
    graph_client: Any,
    query: str,
    count: int = 25,
    folder: str | None = None,
) -> dict:
    """Search mail using KQL."""
    count = _clamp(count, 1, 100)
    safe_query = sanitize_kql(query)

    query_params = {
        "$top": count,
        "$search": safe_query,
        "$select": "id,subject,from,receivedDateTime,isRead,importance,bodyPreview,hasAttachments,categories,flag,conversationId",
    }

    if folder:
        folder = validate_folder_name(folder)
        response = await graph_client.me.mail_folders.by_mail_folder_id(
            folder
        ).messages.get(query_params=query_params)
    else:
        response = await graph_client.me.messages.get(query_params=query_params)

    messages = [_format_message_summary(m) for m in (response.value or [])]

    return {
        "messages": messages,
        "count": len(messages),
        "has_more": response.odata_next_link is not None,
    }


async def list_folders(graph_client: Any) -> dict:
    """List all mail folders."""
    response = await graph_client.me.mail_folders.get()

    folders = []
    for f in (response.value or []):
        folders.append({
            "id": f.id,
            "name": sanitize_output(f.display_name or ""),
            "total": f.total_item_count or 0,
            "unread": f.unread_item_count or 0,
        })

    return {
        "folders": folders,
        "count": len(folders),
    }
```

**Step 3: Run tests, commit**

```bash
uv run pytest tests/test_mail_read.py -v
git add src/outlook_mcp/tools/ tests/test_mail_read.py
git commit -m "feat: mail read tools — list_inbox, read_message, search_mail, list_folders"
```

---

### Task 8: Mail Write Tools

**Files:**
- Create: `src/outlook_mcp/tools/mail_write.py`
- Create: `tests/test_mail_write.py`

**Step 1: Write failing tests for send, reply, forward**

Tests follow same pattern as Task 7 — mock Graph SDK, verify correct API calls and response format. Key tests:
- `send_message`: validates emails, builds Message with recipients/body/importance, calls `graph_client.me.send_mail.post()`
- `reply`: validates message_id, calls `.reply.post()` or `.reply_all.post()` based on `reply_all` flag
- `forward`: validates message_id + to addresses, calls `.forward.post()`
- All write tools raise `ReadOnlyError` when config `read_only=True`

**Step 2: Implement mail_write.py**

Core operations:
- `send_message`: Build `Message` object with recipients, body, importance → `graph_client.me.send_mail.post()`
- `reply`: Validate message_id → `graph_client.me.messages.by_message_id(id).reply.post()` or `.reply_all.post()`
- `forward`: Validate message_id + to addresses → `graph_client.me.messages.by_message_id(id).forward.post()`

**Step 3: Run tests, commit**

---

### Task 9: Mail Triage Tools

**Files:**
- Create: `src/outlook_mcp/tools/mail_triage.py`
- Create: `tests/test_mail_triage.py`

**Step 1: Write failing tests for move, delete (soft/hard), flag, categorize, mark_read**

Key test cases:
- `delete_message(permanent=False)`: calls `move` to `deleteditems` folder
- `delete_message(permanent=True)`: calls `DELETE` endpoint
- All triage tools validate message_id via `validate_graph_id`
- All write tools raise `ReadOnlyError` when `read_only=True`

**Step 2: Implement mail_triage.py**

Core operations:
- `move_message`: Validate folder → `graph_client.me.messages.by_message_id(id).move.post()`
- `delete_message`: If `permanent=False`, move to `deleteditems`. If `permanent=True`, `graph_client.me.messages.by_message_id(id).delete()`
- `flag_message`: PATCH message with `flag.flagStatus`
- `categorize_message`: PATCH message with `categories`
- `mark_read`: PATCH message with `isRead`

**Step 3: Run tests, commit**

---

### Task 10: Calendar Read Tools

**Files:**
- Create: `src/outlook_mcp/tools/calendar_read.py`
- Create: `tests/test_calendar_read.py`

**Step 1: Write failing tests for list_events, get_event**

Key test cases:
- `list_events(days=7)`: computes `startDateTime` = now (in config timezone), `endDateTime` = now + 7 days, both as UTC
- `list_events(after="2026-04-12", before="2026-04-19")`: validates and uses explicit dates
- `get_event`: validates event_id

**Step 2: Implement calendar_read.py**

Core operations:
- `list_events`: Use `calendarView` endpoint with computed `startDateTime`/`endDateTime` → returns expanded recurring events
- `get_event`: `graph_client.me.events.by_event_id(id).get()`

The `days` → datetime computation:
```python
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

def _compute_calendar_range(days: int, after: str | None, before: str | None, timezone: str) -> tuple[str, str]:
    """Compute start/end datetimes for calendarView."""
    tz = ZoneInfo(timezone)
    if after:
        start_utc = validate_datetime(after)
    else:
        now_local = datetime.now(tz)
        start_utc = now_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

    if before:
        end_utc = validate_datetime(before)
    else:
        now_local = datetime.now(tz)
        end_local = now_local + timedelta(days=days)
        end_utc = end_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

    return start_utc, end_utc
```

**Step 3: Run tests, commit**

---

### Task 11: Calendar Write Tools

**Files:**
- Create: `src/outlook_mcp/tools/calendar_write.py`
- Create: `tests/test_calendar_write.py`

**Step 1: Write failing tests for create, update, delete, rsvp**

**Step 2: Implement calendar_write.py**

Core operations:
- `create_event`: Build `Event` object → `graph_client.me.events.post()`
- `update_event`: PATCH event fields
- `delete_event`: `graph_client.me.events.by_event_id(id).delete()`
- `rsvp`: `.accept.post()` / `.decline.post()` / `.tentatively_accept.post()`

All write tools check `read_only` and raise `ReadOnlyError` if set.

**Step 3: Run tests, commit**

---

### Task 12: Wire Tools to FastMCP Server

**Files:**
- Modify: `src/outlook_mcp/server.py`
- Create: `src/outlook_mcp/tools/auth_tools.py`
- Create: `tests/test_server.py`

**Step 1: Write failing test that server exposes tools**

```python
# tests/test_server.py
"""Tests for MCP server tool registration."""

from outlook_mcp.server import mcp


def test_server_has_tools():
    """Server registers all Tier 1 tools."""
    tool_names = [t.name for t in mcp.list_tools()]
    expected = [
        "outlook_login", "outlook_logout", "outlook_auth_status",
        "outlook_list_inbox", "outlook_read_message", "outlook_search_mail", "outlook_list_folders",
        "outlook_send_message", "outlook_reply", "outlook_forward",
        "outlook_move_message", "outlook_delete_message", "outlook_flag_message",
        "outlook_categorize_message", "outlook_mark_read",
        "outlook_list_events", "outlook_get_event",
        "outlook_create_event", "outlook_update_event", "outlook_delete_event",
        "outlook_rsvp",
    ]
    for name in expected:
        assert name in tool_names, f"Missing tool: {name}"
```

**Step 2: Wire lifespan with real config and auth**

```python
@asynccontextmanager
async def lifespan(server):
    config = load_config()
    auth = AuthManager(config)
    yield {"config": config, "auth": auth}
```

**Step 3: Wire all tools using `@mcp.tool()` decorators**

Each tool function:
1. Accepts Pydantic-validated input params
2. Gets auth manager from lifespan context via `Context`
3. Creates Graph client per-request
4. Calls the appropriate tool function
5. Returns structured dict

```python
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
    """List messages in inbox or a specific folder. Supports filtering by read status, sender, and date range."""
    auth = ctx.request_context.lifespan_context["auth"]
    client = GraphClient(auth.get_credential())
    return await list_inbox(client.sdk_client, folder, count, unread_only, from_address, after, before, skip)
```

**Step 4: Implement auth_tools.py with login/logout/status**

**Step 5: Run full test suite**

```bash
uv run pytest -v
```

**Step 6: Commit**

```bash
git add -A
git commit -m "feat: wire all Tier 1 tools to FastMCP server with lifespan context"
```

---

### Task 13: Error Handling Tests

**Files:**
- Create: `tests/test_errors.py`

Test the exception hierarchy:
- `AuthRequiredError` has correct code, message, action
- `ReadOnlyError` includes the tool name
- `NotFoundError` includes resource type and ID
- `GraphAPIError` maps status codes to actions (401 → re-auth, 429 → rate limit)
- All inherit from `OutlookMCPError`

```bash
git add tests/test_errors.py
git commit -m "test: error hierarchy coverage"
```

---

### Task 14: SKILL.md (ClawHub Manifest)

**Files:**
- Create: `SKILL.md`

```markdown
---
name: outlook-mcp
description: MCP server for Microsoft Outlook via Microsoft Graph API. Mail, calendar, contacts, and tasks.
homepage: https://github.com/mpalermiti/outlook-mcp
metadata:
  openclaw:
    emoji: "\U0001F4EC"
    requires:
      python: ">=3.10"
    install:
      - id: pip
        kind: pip
        package: outlook-mcp
        bins: ["outlook-mcp"]
        label: "Install outlook-mcp (pip)"
      - id: uv
        kind: shell
        command: "uv tool install outlook-mcp"
        bins: ["outlook-mcp"]
        label: "Install outlook-mcp (uv)"
---

# outlook-mcp

MCP server for Microsoft Outlook personal accounts (Outlook.com, Hotmail, Live).
Provides AI agents with full access to mail, calendar, contacts, and tasks via Microsoft Graph API.

## Setup

1. **Register an Azure AD app** (one-time, see README for step-by-step)
2. **Configure:** Create `~/.outlook-mcp/config.json`:
   ```json
   {
     "client_id": "YOUR-APP-CLIENT-ID",
     "tenant_id": "consumers",
     "timezone": "America/Los_Angeles"
   }
   ```
3. **Install:** `uv tool install outlook-mcp` or `pip install outlook-mcp`
4. **Register in MCP config:**
   ```json
   {
     "mcp": {
       "servers": {
         "outlook": {
           "command": "outlook-mcp",
           "args": []
         }
       }
     }
   }
   ```
5. **Authenticate:** use the `outlook_login` tool — it returns a URL and code for device-code sign-in.

## Tools

### Auth
- `outlook_login` — Start device-code OAuth2 flow
- `outlook_logout` — Remove stored credentials
- `outlook_auth_status` — Check authentication status

### Mail — Read
- `outlook_list_inbox` — List messages with filters (folder, unread, sender, date)
- `outlook_read_message` — Get full message by ID
- `outlook_search_mail` — Search mail using KQL query
- `outlook_list_folders` — List all mail folders

### Mail — Write
- `outlook_send_message` — Send email with recipients, CC, BCC, HTML, importance
- `outlook_reply` — Reply or reply-all to a message
- `outlook_forward` — Forward a message

### Mail — Triage
- `outlook_move_message` — Move to a folder
- `outlook_delete_message` — Delete (moves to Deleted Items; use permanent=true for hard delete)
- `outlook_flag_message` — Set follow-up flag
- `outlook_categorize_message` — Set categories
- `outlook_mark_read` — Mark read or unread

### Calendar
- `outlook_list_events` — List events in date range
- `outlook_get_event` — Get event details
- `outlook_create_event` — Create event with attendees, recurrence, online meeting
- `outlook_update_event` — Update event fields
- `outlook_delete_event` — Delete event
- `outlook_rsvp` — Accept, decline, or tentatively accept

## Privacy
- Zero telemetry
- Zero local caching of email/calendar data
- Only connects to login.microsoftonline.com and graph.microsoft.com
- Tokens stored in OS keyring (macOS Keychain, etc.)

## Notes
- BYOID: you register your own Azure AD app (see README)
- IDs are opaque Graph strings — get them from list/search tools, never guess
- Dates are ISO 8601, always UTC in responses
- Mail search uses KQL syntax
- Personal accounts only in V1. Enterprise (Entra ID) planned for future.
```

**Commit:**

```bash
git add SKILL.md
git commit -m "feat: ClawHub SKILL.md manifest"
```

---

### Task 15: SECURITY.md

**Files:**
- Create: `SECURITY.md`

Model after olkcli's professional security policy:
- Private vulnerability reporting via GitHub Security Advisories
- 48-hour acknowledgment SLA
- Scope: token leakage, auth bypass, injection, path traversal, supply chain
- Out of scope: social engineering, DoS on Microsoft endpoints

**Commit:**

```bash
git add SECURITY.md
git commit -m "docs: security policy"
```

---

### Task 16: README.md

**Files:**
- Create: `README.md`

Sections:
1. One-line description + badges
2. Features (Tier 1 tool list)
3. **Azure AD App Registration** — step-by-step walkthrough with screenshots description
4. Quick Start (install, configure `config.json`, register MCP, auth)
5. Tool Reference (table of all tools with params)
6. Privacy & Security (zero telemetry promise, token storage)
7. Configuration (config file, timezone, read-only mode)
8. Development (clone, install, test)
9. Roadmap (Tier 2, Tier 3 bullet points)
10. License (MIT)

**Commit:**

```bash
git add README.md
git commit -m "docs: README with BYOID setup guide and full tool reference"
```

---

### Task 17: CI Pipeline

**Files:**
- Create: `.github/workflows/ci.yml`

```yaml
name: CI
on: [push, pull_request]
jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        python-version: ["3.10", "3.11", "3.12", "3.13"]
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v5
      - run: uv sync --all-extras
      - run: uv run ruff check src/ tests/
      - run: uv run pytest --cov=outlook_mcp -v
```

**Commit:**

```bash
git add .github/
git commit -m "ci: test + lint pipeline"
```

---

### Task 18: Integration Smoke Test

**Files:**
- Create: `tests/test_integration.py`

A manual/integration test (marked `@pytest.mark.integration`, skipped by default) that:
1. Loads real config (requires `client_id` in `~/.outlook-mcp/config.json`)
2. Attempts silent token refresh
3. Lists inbox (1 message)
4. Lists today's events
5. Verifies response shapes match expected structure

Run with: `uv run pytest -m integration -v` (requires prior `outlook_login`)

**Commit:**

```bash
git add tests/test_integration.py
git commit -m "test: integration smoke test for manual verification"
```

---

### Task 19: Azure AD App Registration Docs

**Not code — documented in README (Task 16).**

Step-by-step for users:
1. Go to https://entra.microsoft.com → App registrations → New registration
2. Name: anything (e.g. `my-outlook-mcp`)
3. Supported account types: **Personal Microsoft accounts only** (tenant: `consumers`)
4. Redirect URI: leave blank (device code flow doesn't need one)
5. Under **Authentication** → Allow public client flows: **Yes**
6. Under **API Permissions**, add Microsoft Graph delegated permissions:
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `Tasks.ReadWrite`
   - `User.Read`
   - `offline_access`
7. Copy the **Application (client) ID** → put in `~/.outlook-mcp/config.json`
8. No client secret needed (public client)

---

## Summary

| Tier | Tools | Cumulative | Status |
|------|-------|------------|--------|
| **Tier 1: Core MVP** | 21 | 21 | Tasks 1-18 above |
| **Tier 2: Power** | 31 | 52 | Extend after Tier 1 ships |
| **Tier 3: Differentiators** | 17 | 69 | After Tier 2 |
| **Enterprise (future)** | 10 | 79 | After Tier 3 |
| **Untiered** | ~10 | ~89 | Evaluate based on community feedback |

**Total tool surface at full build: 79 tools (+ ~10 untiered)**

The implementation follows TDD throughout. Each task is a commit boundary. The Graph SDK mocking pattern established in Task 7 carries through all tool tasks.

Ship Tier 1 to ClawHub. Gather community feedback. Iterate on Tiers 2-3 based on actual agent usage patterns.
