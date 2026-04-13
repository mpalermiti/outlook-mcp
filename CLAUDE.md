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
