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

> This is an independent open-source project. Not affiliated with, endorsed by, or supported by Microsoft Corporation.

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
5. **Authenticate:** use the `outlook_login` tool.

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

### Contacts
- `outlook_list_contacts` — List contacts with cursor pagination
- `outlook_search_contacts` — Search contacts by name or email
- `outlook_get_contact` — Get full contact details by ID
- `outlook_create_contact` — Create a new contact
- `outlook_update_contact` — Update contact fields
- `outlook_delete_contact` — Delete a contact

### To Do
- `outlook_list_task_lists` — List To Do lists
- `outlook_list_tasks` — List tasks with status filter and pagination
- `outlook_create_task` — Create task with due date, importance, recurrence
- `outlook_update_task` — Update task fields
- `outlook_complete_task` — Mark task as completed
- `outlook_delete_task` — Delete a task

### Drafts
- `outlook_list_drafts` — List draft messages with pagination
- `outlook_create_draft` — Create a draft for later review and sending
- `outlook_update_draft` — Update draft fields
- `outlook_send_draft` — Send an existing draft
- `outlook_delete_draft` — Delete a draft

### Attachments
- `outlook_list_attachments` — List attachments on a message
- `outlook_download_attachment` — Download attachment (base64 or save to file)
- `outlook_send_with_attachments` — Send message with file attachments

### Folder Management
- `outlook_create_folder` — Create mail folder (top-level or nested)
- `outlook_rename_folder` — Rename a mail folder
- `outlook_delete_folder` — Delete a mail folder

### Threading and Batch
- `outlook_list_thread` — Get all messages in a conversation thread
- `outlook_copy_message` — Copy a message to another folder
- `outlook_batch_triage` — Batch move/flag/categorize/mark_read (max 20 per call)

### User and Admin
- `outlook_whoami` — Get current user profile
- `outlook_list_calendars` — List available calendars
- `outlook_list_categories` — List category definitions with colors
- `outlook_get_mail_tips` — Pre-send check (OOF, delivery restrictions)
- `outlook_list_accounts` — List configured accounts
- `outlook_switch_account` — Switch active account

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
