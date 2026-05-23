# Outlook MCP Server

## What This Is
MCP server for Microsoft Outlook personal accounts (Outlook.com/Hotmail) via Microsoft Graph API.
Works with any MCP client (OpenClaw, Claude Code, Cursor).

## Tech Stack
- Python 3.10+, FastMCP, msgraph-sdk, azure-identity, Pydantic v2
- Package manager: uv
- Testing: pytest + pytest-asyncio

## Commands
- `uv run pytest` — run tests
- `uv run ruff check src/ tests/` — lint
- `uv run ruff format src/ tests/` — format
- `uv run outlook-mcp` — start server (stdio)
- `uv run python scripts/preflight.py` — pre-release Graph smoke test (must pass before tagging; see `RELEASING.md`)

## Architecture
- `src/outlook_mcp/server.py` — FastMCP entry point, lifespan context
- `src/outlook_mcp/auth.py` — Device code OAuth2 via azure-identity
- `src/outlook_mcp/graph.py` — Graph client factory
- `src/outlook_mcp/config.py` — Config file management (~/.outlook-mcp/)
- `src/outlook_mcp/validation.py` — Input validation (OData, KQL, IDs, datetimes)
- `src/outlook_mcp/errors.py` — Exception hierarchy
- `src/outlook_mcp/pagination.py` — Cursor-based pagination
- `src/outlook_mcp/tools/` — One file per tool group:
  - `auth_tools.py`, `mail_read.py`, `mail_write.py`, `mail_triage.py` — Tier 1
  - `calendar_read.py`, `calendar_write.py` — Tier 1
  - `contacts.py` — Contact CRUD
  - `todo.py` — To Do task management
  - `mail_drafts.py` — Draft management
  - `mail_attachments.py` — Attachment handling
  - `mail_folders.py` — Folder management
  - `mail_thread.py` — Threading and copy
  - `batch.py` — Batch operations
  - `user.py` — User profile, calendars
  - `admin.py` — Categories, mail tips
  - `inference_overrides.py` — Focused Inbox per-sender override CRUD
  - `mail_delta.py` — Mail delta-sync queries (`outlook_list_inbox_delta`)
  - `calendar_delta.py` — Calendar delta-sync queries (`outlook_list_events_delta`)
  - `contacts_delta.py` — Contacts delta-sync queries (`outlook_list_contacts_delta`)
  - `_delta.py` — Shared httpx-backed delta helper (raw HTTP bypasses the SDK)
  - `digest.py` — Composed "since last call" digest (`outlook_changes_since`) wrapping the three delta tools
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
