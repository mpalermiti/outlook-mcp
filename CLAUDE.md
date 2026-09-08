# Outlook MCP Server

## What This Is
MCP server for Microsoft Outlook personal accounts (Outlook.com/Hotmail) via Microsoft Graph API.
Works with any MCP client (OpenClaw, Claude Code, Cursor).

## Tech Stack
- Python 3.10+, MCP Python SDK 2.x (`MCPServer`), msgraph-sdk, azure-identity, Pydantic v2
- Package manager: uv
- Testing: pytest + pytest-asyncio

## Commands
- `uv run pytest` — run tests (offline unit suite; `integration`/`live` markers are deselected by default)
- `uv run pytest -m live -v` — live query-shape guards; run before tagging if you changed any `$filter`/`$orderby`/`$search` construction (see `RELEASING.md` 1b)
- `uv run pytest -m integration -v` — live response-shape smoke tests
- `uv run ruff check src/ tests/` — lint
- `uv run ruff format src/ tests/` — format
- `uv run outlook-mcp` — start server (stdio)
- `uv run python scripts/preflight.py` — pre-release Graph smoke test (must pass before tagging; see `RELEASING.md`)

## Releasing
Publishing is automated — do **not** run `uv publish` or `mcp-publisher` by hand.
Publishing a GitHub release triggers `.github/workflows/publish.yml`, which re-checks
the version lockstep, runs tests and lint, builds, and publishes to PyPI and the MCP
registry via GitHub OIDC (no stored credentials). Full process in `RELEASING.md`.
Still manual by design: the live tier (run it *before* tagging) and ClawHub.

## Architecture
- `src/outlook_mcp/server.py` — `MCPServer` entry point, lifespan context
- `src/outlook_mcp/auth.py` — Device code OAuth2 via azure-identity
- `src/outlook_mcp/graph.py` — Graph client factory
- `src/outlook_mcp/config.py` — Config file management (~/.outlook-mcp/)
- `src/outlook_mcp/validation.py` — Input validation (OData, KQL, IDs, datetimes)
- `src/outlook_mcp/errors.py` — Exception hierarchy
- `src/outlook_mcp/pagination.py` — Cursor-based pagination
- `src/outlook_mcp/throttle.py` — Retry-After honoring for the raw-httpx delta/`$batch` paths (SDK path already retries via kiota)
- `src/outlook_mcp/toolsets.py` — Tool annotations + config-gated toolset selection (`OUTLOOK_MCP_TOOLSETS`); `configure()` runs once after registration
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
  - `_recurrence.py` — Shared recurrence conversion for calendar events and To Do tasks (Graph models both identically)
  - `digest.py` — Composed "since last call" digest (`outlook_changes_since`) wrapping the three delta tools

## Conventions
- One tool = one operation (not grouped CRUD)
- Tool names prefixed with `outlook_`
- All input validated in `validation.py` before Graph API calls; tool-argument types are
  enforced by the schemas `MCPServer` generates from the annotations. There is deliberately
  no hand-written Pydantic I/O layer — one existed until 1.16.0, was wired to nothing, and
  is why #41 went unnoticed for fourteen releases: a validator that looked authoritative and
  never ran. If you add one, wire it to the tool path in the same commit.
- No telemetry, no local caching, no third-party calls
- Tests: TDD, pytest, mock Graph client for unit tests. `tests/test_no_dead_parameters.py` fails the build on any parameter declared and never read — the #41 shape; use it, remove it, or justify it in that file's ALLOWED set. Mocks assert what we *send* — they cannot see a query Graph rejects or silently mis-evaluates, so anything that builds a `$filter`/`$orderby`/`$search` string also needs a `@pytest.mark.live` guard
- Errors: raise OutlookMCPError subclasses, never return error dicts
- Datetimes: UTC in responses, config timezone for input interpretation
- Delete: soft delete (move to Deleted Items) by default
- Dependency bounds: an unbounded requirement can break every fresh install without a single commit. `mcp[cli]` with no upper bound shipped a package that could not be installed for five weeks (2026-07-28 → 09-03) while CI stayed green — `uv sync` resolves through `uv.lock`, so the `test` job never sees what a new user actually gets. The `fresh-install` (per push) and `published-install` (weekly cron) jobs in `ci.yml` are the guard against this class; keep them working
