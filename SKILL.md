---
name: outlook-mcp
description: MCP server for Microsoft Outlook personal accounts. 51 tools — mail, calendar, contacts, to-do, drafts, attachments, folders, threading, batch ops.
homepage: https://github.com/mpalermiti/outlook-mcp
metadata:
  openclaw:
    emoji: "\U0001F4EC"
    requires:
      python: ">=3.10"
    install:
      - id: uv
        kind: shell
        command: "git clone https://github.com/mpalermiti/outlook-mcp.git && cd outlook-mcp && uv sync"
        bins: ["outlook-mcp"]
        label: "Clone and install (uv)"
---

# outlook-mcp

MCP server for Microsoft Outlook personal accounts (Outlook.com, Hotmail, Live).
Provides AI agents with full access to mail, calendar, contacts, and tasks via Microsoft Graph API.

> Independent open-source project. Not affiliated with Microsoft.

## Important

- **Personal Microsoft accounts only** (`@outlook.com`, `@hotmail.com`, `@live.com`). Work/school accounts (Entra ID) are not supported in v1.
- **Requires Azure AD app registration** — free, takes ~5 minutes, but you need a free Azure account first. See README.
- **Auth is CLI-based** — run `outlook-mcp auth` on the host before the agent can use it. No interactive auth through MCP tools.

## Setup

1. **Create a free Azure account** at [azure.microsoft.com/free](https://azure.microsoft.com/free) (sign up with your `@outlook.com` address)
2. **Register an Azure AD app** (see README for step-by-step)
3. **Configure:** Create `~/.outlook-mcp/config.json`:
   ```json
   {
     "client_id": "YOUR-APP-CLIENT-ID",
     "tenant_id": "consumers",
     "timezone": "America/Los_Angeles",
     "read_only": true
   }
   ```
4. **Register in OpenClaw config:**
   ```json
   {
     "mcp": {
       "servers": {
         "outlook": {
           "command": "uv",
           "args": ["--directory", "/path/to/outlook-mcp", "run", "outlook-mcp"]
         }
       }
     }
   }
   ```
5. **Authenticate on the host:**
   ```bash
   cd /path/to/outlook-mcp && uv run outlook-mcp auth
   ```
6. **Restart the gateway:** `openclaw gateway restart`

## Tools (51)

### Auth
- `outlook_auth_status` — Check authentication status and read-only mode

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
- `outlook_delete_message` — Delete (soft by default, permanent optional)
- `outlook_flag_message` — Set follow-up flag
- `outlook_categorize_message` — Set categories
- `outlook_mark_read` — Mark read or unread

### Calendar
- `outlook_list_events` — List events in date range (expands recurring)
- `outlook_get_event` — Get event details
- `outlook_create_event` — Create event with attendees, recurrence, online meeting
- `outlook_update_event` — Update event fields
- `outlook_delete_event` — Delete event
- `outlook_rsvp` — Accept, decline, or tentatively accept

### Contacts
- `outlook_list_contacts` — List with cursor pagination
- `outlook_search_contacts` — Search by name or email
- `outlook_get_contact` — Get full details
- `outlook_create_contact` — Create
- `outlook_update_contact` — Update
- `outlook_delete_contact` — Delete

### To Do
- `outlook_list_task_lists` — List To Do lists
- `outlook_list_tasks` — List tasks with status filter and pagination
- `outlook_create_task` — Create with due date, importance, recurrence
- `outlook_update_task` — Update
- `outlook_complete_task` — Mark completed
- `outlook_delete_task` — Delete

### Drafts
- `outlook_list_drafts` — List with pagination
- `outlook_create_draft` — Create for later review
- `outlook_update_draft` — Update
- `outlook_send_draft` — Send
- `outlook_delete_draft` — Delete

### Attachments
- `outlook_list_attachments` — List on a message
- `outlook_download_attachment` — Download (base64 or save to file)
- `outlook_send_with_attachments` — Send with files (auto upload session for >3MB)

### Folder Management
- `outlook_create_folder` — Create (top-level or nested)
- `outlook_rename_folder` — Rename
- `outlook_delete_folder` — Delete (refuses well-known folders)

### Threading and Batch
- `outlook_list_thread` — Get all messages in a conversation
- `outlook_copy_message` — Copy to another folder
- `outlook_batch_triage` — Batch move/flag/categorize/mark_read (max 20)

### User and Admin
- `outlook_whoami` — Current user profile
- `outlook_list_calendars` — Available calendars
- `outlook_list_categories` — Category definitions with colors
- `outlook_get_mail_tips` — Pre-send check (OOF, delivery restrictions)
- `outlook_list_accounts` — Configured accounts
- `outlook_switch_account` — Switch active account

## Privacy
- Zero telemetry, zero local caching
- Only connects to `login.microsoftonline.com` and `graph.microsoft.com`
- Tokens stored in OS keyring (macOS Keychain, Windows Credential Store)
- BYOID: you register your own Azure AD app — no shared client ID

## Notes
- IDs are opaque Graph strings — get them from list/search tools, never guess
- Dates are ISO 8601, UTC in responses, config timezone for input interpretation
- Mail search uses KQL syntax
- Start with `read_only: true`, flip when comfortable
