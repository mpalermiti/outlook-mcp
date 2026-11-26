# Outlook MCP Server

## What This Is
MCP server for Microsoft Outlook personal accounts (Outlook.com/Hotmail) via Microsoft Graph API.
Works with any MCP client (OpenClaw, Claude Code, Cursor).

## Tech Stack
- Python 3.10+, MCP Python SDK 2.x (`MCPServer`), msgraph-sdk, azure-identity, Pydantic v2
- uvicorn for the opt-in HTTP transport; the optional `documents` extra adds pypdf, openpyxl, psutil and defusedxml for `outlook_read_attachment`
- Package manager: uv
- Testing: pytest + pytest-asyncio

## Commands
- `uv run pytest` — run tests (offline unit suite; the `integration`/`live`/`live_write`/`live_send` markers are deselected by default). The PDF/XLSX parser tests skip unless `uv sync --all-extras` installed the `documents` extra
- `OUTLOOK_MCP_TOOLSETS="$(uv run python -c 'from outlook_mcp import toolsets; print(",".join(sorted(set(toolsets.TOOL_GROUPS.values()))))')" uv run pytest -q` — the same suite with every toolset loaded, the opt-in groups included, as CI also runs it
- `scripts/live-tests.sh <tier>` — the live tiers, in a container, against the project's two test accounts (`RELEASING.md` 1–1e; needs Docker, so the sandbox off, and an env file outside the repo, `OUTLOOK_MCP_LIVE_ENV_FILE`). Tiers: `preflight`; `live`, the query-shape guards (run before tagging if you changed any `$filter`/`$orderby`/`$search` construction); `integration`, the response-shape smoke tests; `live_write`; `live_send --sends N`; plus `status`, `leftovers`, `clean`, and `offline` (`--python X.Y` for another Python)
- `OUTLOOK_MCP_LIVE_OWN_MAILBOX=1 uv run pytest -m live -v` — optional: the read-only `live` tier against your own signed-in mailbox, on the host (no account check; anything but a read is refused before it leaves). The only live test run that may happen outside the container. `uv run python scripts/preflight.py` on the host also still works, GET-only, against whatever account is signed in there, but the release step is the runner's `preflight`
- Without the test accounts: run the offline suite (`uv run pytest`, or `scripts/live-tests.sh --python X.Y offline` with Docker) and, optionally, the own-mailbox `live` check. `preflight`, `integration`, `live_write` and `live_send` need the accounts, and refuse on the host; say in the PR they were not run. Setting up a pair: `RELEASING.md` §1
- `uv run ruff check src/ tests/ scripts/` — lint
- `uv run ruff format src/ tests/ scripts/` — format (CI runs `ruff format --check` on the same paths and fails on any file it would change)
- `uv run outlook-mcp` — start server (stdio; `OUTLOOK_MCP_TRANSPORT=streamable-http` serves HTTP instead, see the README)
- `scripts/live-tests.sh preflight` — pre-release Graph smoke test (must pass before tagging; see `RELEASING.md`)

## Releasing
Publishing is automated — do **not** run `uv publish` or `mcp-publisher` by hand.
Publishing a GitHub release triggers `.github/workflows/publish.yml`, which re-checks
the version lockstep, runs tests and lint, builds, and publishes to PyPI and the MCP
registry via GitHub OIDC (no stored credentials). Full process in `RELEASING.md`.
Still manual by design: the live tiers (run them *before* tagging) and ClawHub.

## Live tests and the test accounts
The live tiers run against two throwaway Outlook.com accounts the project owns: a main account and its pair, the only address mail may be sent to. They are signed in only inside Docker volumes, and `scripts/live-tests.sh` is the way in. The image copies an allow-list of paths from the checkout, never the whole of it (`docker/`, held by `tests/test_live_image.py`). Rules for any session working here:
- **No test-account address in the repository**: not in code, tests, docs, commit messages or PRs. The addresses reach a run only from the env file outside the repo (`tests/live_guard.py` reads `OUTLOOK_MCP_TEST_ACCOUNT`/`OUTLOOK_MCP_TEST_PAIR`); output from the guards and helpers names roles (`account`, `pair`), never addresses. Scratch files that hold one stay outside the repo too.
- **Never sign a test account in on the host**, and never run a live tier there (the one exception is the own-mailbox `live` opt-in above). On macOS, with azure-identity below 1.27, every outlook-mcp token cache on a machine shares one Keychain item, so a test-account sign-in on the host would replace the maintainer's own.
- Each live test refuses to start unless `/me` names one of the two accounts and the saved sign-in record agrees. A failed `/me` refuses, and the record alone never stands in for it. The exception is the own-mailbox `live` opt-in. While a live test runs, every HTTP request the process makes, through the SDK, a tool's own client, or raw `httpx` or `requests`, is held to its tier's guard before it leaves (`live_guard.guarding`):
  - `live` and `integration` may only read (`ReadOnly`).
  - `live_write` may write but never send (`NeverSends`). That rules out sends, replies, forwards, RSVPs, invitations, changes to meetings with attendees, forwarding rules, automatic replies and calendar sharing.
  - `live_send` sends only to the pair, within the run's `--sends` budget of at most 10 (`SendGuard`).
  - `tests/test_live_tiers_never_send.py` also checks the files.
  - A child process is not covered: `tests/live_mailbox.py` keeps its own rules.
  - New accounts have tiny sending limits, and Outlook.com locks one that sends too fast. Keep budgets small, and never send to the fictional people in seeded mail.
  - If the env file names a stop file (`OUTLOOK_MCP_LIVE_STOP_FILE`) and it exists, a test account is locked or blocked: every tier but `offline` refuses (`tests/test_live_tests_stop_file.py`). Never delete it to get a run through; a person deletes it after unlocking the account.
- Everything a writing tier creates carries a marker (`LIVE_WRITE_SUBJECT`, `LIVE_WRITE_NAME`, `LIVE_SEND_SUBJECT`); the runner lists leftovers after the run and `scripts/live-tests.sh clean` removes them. Live tests must leave the accounts as they found them: the mailboxes also hold a seeded scenario that other evaluation work relies on.

## Architecture
- `src/outlook_mcp/server.py` — `MCPServer` entry point, lifespan context
- `src/outlook_mcp/http_transport.py` — Opt-in streamable-HTTP transport (`OUTLOOK_MCP_TRANSPORT`): settings checked before anything binds, bearer-token gate, Host/Origin allow-list, uvicorn config. stdio stays the default
- `src/outlook_mcp/auth.py` — Device code OAuth2 via azure-identity
- `src/outlook_mcp/graph.py` — Graph client factory
- `src/outlook_mcp/config.py` — Config file management (`~/.outlook-mcp/`, or `OUTLOOK_MCP_CONFIG_DIR` — one directory per server instance, one instance per account)
- `src/outlook_mcp/_config_env.py` — Optional `OUTLOOK_MCP_*` environment input over the config file (#180): `client_id` only when the file has none, `tenant_id`, `timezone` and `show_tips` as overrides; the security settings only tighten the file's policy, and a value that would loosen it is refused at startup (README "Environment configuration")
- `src/outlook_mcp/validation.py` — Input validation (OData, KQL, IDs, datetimes, time zones)
- `src/outlook_mcp/_windows_zones.py` — Generated Windows zone name → IANA key table (CLDR `windowsZones.xml`, territory `001`), used only to convert a `Z`/offset event time into a zone Graph names in Windows terms (#236); the name sent back stays Graph's. Regenerate with `scripts/generate_windows_zones.py` (its docstring has the steps); never edit by hand
- `src/outlook_mcp/errors.py` — Exception hierarchy. `OutlookMCPError` inherits the SDK's `ToolError`; this is load-bearing, not cosmetic (see Conventions)
- `src/outlook_mcp/pagination.py` — Cursor-based pagination
- `src/outlook_mcp/cli.py` — `outlook-mcp` entry point: `serve` (default), `auth`, `status`, `logout`, `toolsets` (show the selection in force and its source; with names, write config.json's `toolsets` key alone, keeping the rest of the file as written)
- `src/outlook_mcp/permissions.py` — `read_only` and `allow_categories` write gating (the eight write categories); `check_sends_mail` also gates a write whose content reaches someone else — an invitation, or content written into a calendar or To Do list another person owns
- `src/outlook_mcp/folder_resolver.py`, `calendar_resolver.py` — Display name / well-known name → Graph id resolution for mail folders and calendars; a mail folder can also be a path, `Inbox/Travel/Flights`, with `//` for a `/` in a name (#144). `calendar_resolver.py` also decides whether a calendar is the user's own (owner address vs the default calendar's), which `calendar_write`/`calendar_attachments` use to require `mail_send` for someone else's
- `src/outlook_mcp/throttle.py` — Retry-After honoring for the raw-httpx delta/`$batch` paths (the MIME export reports a 429 instead; SDK path already retries via kiota), and `call_without_replay` for SDK writes that must not be repeated (sends, event attachments, createReply/createForward, rule creation; To Do attachments use their own retry-off option): kiota's retry off, 429 retried here, never a 5xx (#131)
- `src/outlook_mcp/toolsets.py` — Tool annotations, lean advertised input schemas (`lean_input_schema`), and config-gated toolset selection (`OUTLOOK_MCP_TOOLSETS` when set and not blank, else config.json's `toolsets`, else the default: `resolve_selection`; `classic` is an alias for the default; a name that is no group loads nothing and is warned about at import, a selection from config.json is named on stderr at every start, and a `toolsets` value that is not a string is ignored with a warning); the off-by-default groups (`OFF_BY_DEFAULT`: `documents`, `rules`, `mime`, `contact_photos`, `intents`) load only when named, and `mailbox_settings_write: true` adds `rules` when nothing is named; `configure()` runs once after registration. Also `TIP_MAP`/`tip_for`, the one-time `server_tip` that `server._wrap_tool_errors` may add to a classic result; off (`TIPS_LIVE`) until the owner's go decision on the intent tools, held by `test_intent_tip.py::test_tips_stay_off_until_the_go_decision`
- `src/outlook_mcp/intents/` — The intent tools (opt-in `intents` group; `docs/design/intent-contracts.md`). `find.py` is `outlook_find`'s implementation; `resolver.py` and `people.py` resolve names to things (ranked only from Sent Items and contacts; `single` only for someone written to); `selection.py` builds the KQL; `envelope.py` the answer shape (other people's text only under `others_wrote`); `budget.py` the per-call request budget, deadline and four-unit limiter; `transport.py` the counted raw Graph path; `ids.py` the immutable-id switch, on in a process that loads an intent tool (SDK hooks in `graph.py`, `with_preference` on the classic raw paths), and the id family read from `Preference-Applied`. Probe ids (`P-xx`, contracts §14) mark where the code relies on unproven Graph behaviour; the live guards are `tests/test_live_intent_queries.py`
- `src/outlook_mcp/sender_auth.py` — Microsoft's `Authentication-Results` verdict on a message's From domain, for `include_sender_auth` (#130). Pure: no Graph, no I/O
- `src/outlook_mcp/tools/` — One file per tool group:
  - `mail_read.py`, `mail_write.py`, `mail_triage.py` — Tier 1 (auth tools live directly in `server.py`)
  - `calendar_read.py`, `calendar_write.py` — Tier 1
  - `calendar_attachments.py` — Event attachments (attach from `attachments_dir`, with the mail tools' upload session for >3MB; remove one)
  - `contacts.py` — Contact CRUD
  - `contact_photos.py` — `outlook_set_contact_photo` (opt-in `contact_photos` group): a JPEG, PNG or GIF from `attachments_dir`, type read from the bytes, capped at 20 MiB by a bounded read, PUT through the SDK's `photo.content`; set only
  - `todo.py` — To Do task management
  - `todo_attachments.py` — To Do task attachments (inline base64 uploads ≤20 MiB, contentBytes downloads)
  - `mail_drafts.py` — Draft management
  - `mail_attachments.py` — Mail attachments: list, download (to `attachments_dir`, or `as_content`), send and attach to drafts; home of `resolve_attachment_path`, the download helpers the To Do tools share, and the >3MB upload session the event tools reuse
  - `mail_attachment_read.py`, `_document_worker.py` — `outlook_read_attachment` (opt-in `documents` group): PDF, XLSX and text attachments as paged text. PDF/XLSX parse in `_document_worker.py`, a short-lived clean-environment subprocess the parent bounds by memory, wall time and output — containment, not a sandbox
  - `mail_mime.py` — `outlook_export_message` (opt-in `mime` group): a message's original bytes from a raw-httpx GET of `$value`, uncompressed, no redirects, streamed to a 10 MiB cap; returned as a JSON block (`size_bytes`, `sha256`) and a `message/rfc822` resource blob
  - `mail_folders.py` — Folder management
  - `mail_rules.py` — Inbox rule CRUD; opt-in toolset, and a rule that forwards needs `mail_send`
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
  - `_mail_body.py` — Shared mail-body conversion: plain text and Markdown (`is_markdown`, a safe subset) as HTML, and a reply's or forward's text above the quote Graph wrote
  - `digest.py` — Composed "since last call" digest (`outlook_changes_since`) wrapping the three delta tools

## Conventions
- One tool = one operation (not grouped CRUD)
- Tool names prefixed with `outlook_`
- All input validated in `validation.py` before Graph API calls; tool-argument types are
  enforced by the schemas `MCPServer` generates from the annotations. There is deliberately
  no hand-written Pydantic I/O layer — one existed until 1.16.0, was wired to nothing, and
  is why #41 went unnoticed for fourteen releases: a validator that looked authoritative and
  never ran. If you add one, wire it to the tool path in the same commit.
- No telemetry, no local caching, no third-party calls (carve-outs: the To Do default-list
  id is resolved once per Graph client and kept — an id, not content; and the intent tools'
  process-wide action state, never mailbox data: the id-family word and the request limiter in
  `intents/`, and the shown-once tip flag)
- Tests: TDD, pytest, mock Graph client for unit tests. Four offline guards against the silent-no-op class that produced #41 — a call that succeeds and does nothing: `test_no_dead_parameters.py` (parameter declared, never read), `test_no_dead_modules.py` (module nothing imports), `test_sdk_fields_exist.py` (attribute assigned on an SDK model that has no such field — the SDK drops it silently), `test_write_payloads_reach_the_wire.py` (each write argument must appear in the *serialized* payload, not just on the model). Fix the finding or justify an allowlist entry in the file; never weaken the guard. The live tiers' guards (whose mailbox, who may get mail) are tested offline in `test_live_guards.py`. Separately, `tests/conftest.py` keeps the offline suite off the real settings and sign-in (#267): a run whose `-m` selects no live tier gets an empty `OUTLOOK_MCP_CONFIG_DIR` before `outlook_mcp` is imported, and each test without a live marker gets a fresh empty directory of its own (every `DEFAULT_CONFIG_DIR` in the package, the `load_config`/`save_config` defaults and child processes point there); building a real `DeviceCodeCredential` or MSAL client, opening the token cache, a DNS lookup, or connecting or sending off loopback fails the test (`real_credential` marks the few that build a credential on purpose, its `_get_app` replaced). Mocks assert what we *send* — they cannot see a query Graph rejects or silently mis-evaluates, so anything that builds a `$filter`/`$orderby`/`$search` string also needs a `@pytest.mark.live` guard
- Errors: raise OutlookMCPError subclasses, never return error dicts. They inherit the SDK's
  `ToolError` — an *anticipated* failure, whose text the SDK forwards to the model. Anything
  inheriting plain `Exception` is treated as a crash and reaches the model as
  `Error executing tool <name>` with the message withheld. That is not a detail: it silently
  suppressed the entire hierarchy from 1.14.0 to 1.19.0 while every type assertion stayed green.
  A new error type inherits from `OutlookMCPError`, and `__str__` carries the `action` hint
  because that string is what the agent reads. Guarded by `test_error_text_reaches_client.py`.
  `_wrap_tool_errors` also converts the two non-SDK failures every raw-httpx path can raise:
  `httpx.HTTPStatusError` (from `raise_for_status()`, rebuilt without the URL) and
  azure-identity's `ClientAuthenticationError`. A raw path reads its reply with
  `errors.graph_json(response)`, never a bare `response.json()`: otherwise its
  `JSONDecodeError`, a `ValueError`, is reported as invalid input. The wrapper also replaces
  any lone surrogate left in a tool's result, and `OutlookMCPError.__str__` does the same for
  error text: one in a content block or in error text ends the stdio server (#186)
- Cross-tool guidance goes in `INSTRUCTIONS` (sent once per session) or a prompt, never into 71
  docstrings — a docstring is paid for on every turn by every client. A docstring stays
  self-sufficient for using *that* tool; sequencing across tools does not belong there
- Anything taking a host filesystem path routes through `resolve_attachment_path`. Paths come
  from the model, and the model reads email — treat them as untrusted input, and confine by
  resolving, never by string comparison
- Tool schemas are a per-turn cost with a measured baseline — two yardsticks, never compared
  with each other: ~8,644 **o200k** tokens for the 62-tool surface (real tokenizer, ROADMAP
  2026-07) and the budget test's own **chars/4 proxy** measure for the current surface (the
  To Do detail tools added ~+12% per turn; the budget test header records the measured value).
  Metadata that is correct but inert — `openWorldHint`, which is `true` by default anyway, or
  titles that restate the tool name — is not free. Pydantic's share of that is removed for
  every tool at registration: `toolsets.lean_input_schema` drops the `title` it writes on each
  input and tool, and an optional input's `null` branch and `"default": null` (a quarter of the
  surface). A new tool gets it with nothing to do; advertising only, since the SDK validates
  arguments against the function's own model (`test_lean_tool_schemas.py`).
  `test_tool_surface_budget.py` holds the line:
  one ceiling for the default surface, and separate ones with the opt-in `documents`,
  `rules`, `mime`, `contact_photos` or `intents` group loaded, so an opt-in tool cannot raise every user's per-turn cost
- Datetimes: UTC in responses, config timezone for input interpretation
- Delete: soft delete (move to Deleted Items) by default
- Dependency bounds: an unbounded requirement, direct or transitive, can break every fresh install without a single commit, and the `test` job cannot see it because `uv sync` resolves through `uv.lock`. The `fresh-install` (per push) and `published-install` (weekly cron) jobs in `ci.yml` cover what a new user gets; keep them working. They check that the package imports and registers its tools, not what reaches Graph, so behaviour a dependency controls on the wire needs its own test against whatever versions resolve (`tests/test_me_rewrite_reaches_the_wire.py` is the model). When a dependency's new release breaks us, cap it with a comment naming the upstream fix that lifts the cap. Dependabot moves `uv.lock` weekly inside these bounds (`versioning-strategy: lockfile-only`) and never edits them; a floor or cap moves by hand. Current cap: `azure-identity<1.27` (#244) — on macOS, 1.27 keeps a named cache under a new Keychain entry and loses the existing sign-in; it lifts when `auth.py` moves that sign-in across, and `tests/test_azure_identity_cap.py` holds it until then
