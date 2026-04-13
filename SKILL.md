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
