# outlook-mcp

MCP server for Microsoft Outlook personal accounts via Microsoft Graph API.

> **Disclaimer:** This is an independent open-source project. Not affiliated with, endorsed by, or supported by Microsoft Corporation. "Outlook" and "Microsoft Graph" are trademarks of Microsoft.

---

## Features

**21 tools** across 6 categories:

- **Auth (3)** -- device-code OAuth2 login, logout, status check
- **Mail Read (4)** -- list inbox, read message, search (KQL), list folders
- **Mail Write (3)** -- send, reply/reply-all, forward
- **Mail Triage (5)** -- move, delete (soft by default), flag, categorize, mark read/unread
- **Calendar Read (2)** -- list events (with recurring expansion), get event details
- **Calendar Write (4)** -- create, update, delete, RSVP (accept/decline/tentative)

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

You need to register a free Azure AD app to get a client ID. This takes about 2 minutes.

### Steps

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your personal Microsoft account (Outlook.com, Hotmail, Live).

2. Search for **"App registrations"** in the top search bar. Click it.

3. Click **"New registration"**.

4. Fill in:
   - **Name:** anything you want (e.g. `my-outlook-mcp`)
   - **Supported account types:** select **"Personal Microsoft accounts only"**
   - **Redirect URI:** leave blank

5. Click **Register**.

6. Under **Authentication** (left sidebar):
   - Scroll to **"Allow public client flows"**
   - Set to **Yes**
   - Click **Save**

7. Under **API permissions** (left sidebar):
   - Click **"Add a permission"**
   - Select **"Microsoft Graph"**
   - Select **"Delegated permissions"**
   - Add these permissions:
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.ReadWrite`
     - `Contacts.ReadWrite`
     - `Tasks.ReadWrite`
     - `User.Read`
     - `offline_access`

8. Go back to **Overview**. Copy the **Application (client) ID**. You will need this for the config file.

No client secret is needed. The device code flow uses public client auth.

---

## Quick Start

### Install

```bash
# Clone
git clone https://github.com/mpalermiti/outlook-mcp.git
cd outlook-mcp

# Install with uv
uv sync
```

### Configure

Create `~/.outlook-mcp/config.json`:

```json
{
  "client_id": "YOUR_APPLICATION_CLIENT_ID",
  "tenant_id": "consumers",
  "timezone": "America/Los_Angeles",
  "read_only": false
}
```

The only required field is `client_id`. Everything else has sensible defaults.

### Register with your MCP client

Add to your MCP client config (e.g. Claude Code `settings.json`, OpenClaw, Cursor):

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

### Authenticate

Once the server is running, call the `outlook_login` tool. It will start a device-code flow -- you will be given a URL and a code to enter in your browser. After sign-in, tokens are cached in your OS keyring.

---

## Tool Reference

### Auth

| Tool | Description |
|------|-------------|
| `outlook_login` | Start device-code OAuth2 flow. Optionally set `read_only: true`. |
| `outlook_logout` | Remove stored credentials. |
| `outlook_auth_status` | Check if authenticated and whether read-only mode is active. |

### Mail Read

| Tool | Description |
|------|-------------|
| `outlook_list_inbox` | List messages in a folder. Filter by read status, sender, date range. Pagination via `skip`. |
| `outlook_read_message` | Get full message by ID. Format: `text`, `html`, or `full` (both). |
| `outlook_search_mail` | Search mail using KQL query. Optionally scope to a folder. |
| `outlook_list_folders` | List all mail folders with total and unread counts. |

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

---

## Configuration

Config lives at `~/.outlook-mcp/config.json` (created with `0600` permissions).

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `client_id` | `string` | `null` | Azure AD application (client) ID. Required for auth. |
| `tenant_id` | `string` | `"consumers"` | Azure AD tenant. Use `"consumers"` for personal Microsoft accounts. |
| `timezone` | `string` | `"UTC"` | IANA timezone (e.g. `"America/New_York"`). Used for relative date computations in calendar tools. |
| `read_only` | `bool` | `false` | When `true`, all write tools (send, reply, move, delete, create, update, RSVP) return an error. |

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

### Tier 2: Power Features

- **Contacts** -- list, search, get, create, update, delete
- **To Do** -- task lists, tasks (create, update, complete, delete)
- **Drafts** -- list, create, update, send, delete
- **Attachments** -- list, download, send-with-attachments (upload session for large files)
- **Threading** -- list all messages in a conversation
- **Folders** -- create, rename, delete mail folders
- **Batch** -- batch triage operations (move/flag/categorize up to 20 per call)
- **Pagination** -- cursor-based pagination on all list tools
- **Categories** -- list category definitions with colors
- **User profile** -- `whoami` tool

### Tier 3: Differentiators

- **Focused Inbox** -- list Focused/Other tab, override classification
- **Out-of-Office** -- get/set automatic replies
- **Inbox Rules** -- list, create, delete rules
- **Advanced mail** -- raw MIME export, internet message headers
- **Calendar** -- cancel event (with attendee notification), list calendars
- **Checklists** -- checklist items on To Do tasks
- **Notifications** -- poll-based change notifications for mail and calendar

---

## License

MIT. See [LICENSE](LICENSE).
