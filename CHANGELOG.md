# Changelog

All notable changes to outlook-graph-mcp are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Documentation

- **`read_only` is documented as what it is: a tool gate, not a token scope.** It blocks
  this server's write tools; it does not narrow the OAuth token, which is acquired with
  `.default` and carries whatever the Azure app was consented for. A `read_only` server
  still holds a write-capable Graph token, and the setting is a line in `config.json` rather
  than something Microsoft enforces. README gains a section spelling this out, SECURITY.md
  names it in the design list, and ROADMAP carries the real fix (a separately consented
  read-only app). Raised by the ClawHub scanner as `[T05]`; predates v1 and is not a
  regression.

## [1.21.0] — 2026-09-11

### Security

- **A delta cursor can no longer redirect your mailbox token.** `fetch_delta_pages`
  used its `delta_token` argument as the request URL verbatim and attached
  `Authorization: Bearer <Graph token>` to it, with no check on scheme or host;
  `@odata.nextLink` values read back out of a response body were followed the same
  way. `_delta.py` deliberately does not persist cursors — "that's the caller's
  job" — so the cursor is agent-held state, and an agent takes instructions from
  the mail it reads. A string like `https://evil.example/collect` arriving in a
  message body was therefore enough to send a live, full-mailbox access token to
  a stranger. Every URL that would carry the token is now parsed and required to
  be https on `graph.microsoft.com` (`require_graph_url`), including the
  `deltaLink` handed back as the next cursor, so a poisoned link is never stored
  and replayed either.

  Parsed, not string-matched, for the same reason `resolve_attachment_path`
  resolves instead of comparing prefixes: `startswith("https://graph.microsoft.com")`
  accepts `https://graph.microsoft.com@evil.example/` — whose real host is
  `evil.example` — and `https://graph.microsoft.com.evil.example/`.

  Guarded live: `tests/test_live_delta_cursors.py` drives a real delta round
  against Graph and replays the returned cursor through the guard, plus pins
  that a real `deltaLink` carries no port and no userinfo — the two assumptions
  the `netloc` equality rests on. The mocked suite cannot check this, because
  its cursors are ones we wrote; preflight only checks that the delta endpoints
  answer, and never routes a cursor through `fetch_delta_pages`.

  Hardened further after review: the validator refuses control and space
  characters outright (`urlsplit` deletes tab/CR/LF before parsing while an
  HTTP client does not, so one string could read as two different hosts),
  compares the whole `netloc` rather than `hostname` so userinfo and an
  explicit port are refused with it, and returns the string it actually
  checked rather than the one passed in.

  Reported by the ClawHub security scanner (`[T09]`) against 1.19.0.

- **The token cache is no longer written in cleartext without being asked.**
  `allow_unencrypted_storage=True` was unconditional, so on Linux without
  libsecret a reusable Graph refresh token was persisted to disk in plaintext
  with only a log line to mark it. It is now opt-in via
  `allow_unencrypted_token_cache` (default `false`); without it, authentication
  stops and explains the two ways forward instead. macOS and Windows always had
  an encrypted store and are unaffected.

  **Breaking** on Linux hosts without libsecret that relied on the silent
  fallback: set `allow_unencrypted_token_cache: true`, or install
  `gnome-keyring libsecret-1-0 python3-gi` and re-create the venv with
  `--system-site-packages`.

### Fixed

- **Startup no longer blocks for fifteen minutes on an expired token.**
  `try_cached_token` promises a token "without user interaction", but built its
  credential with azure-identity's default
  `disable_automatic_authentication=False`. A cache miss therefore did not
  return False — it opened a device-code flow and polled for a human until
  `timeout` (900s), inside `lifespan`, on server startup. The silent path now
  forbids the interactive flow and reports failure, which is the answer the
  caller was always documented to get. `outlook-mcp auth` is unaffected; the
  prompt belongs to it.

  This also explains why `uv run pytest` — documented as the offline unit
  suite — hung on developer machines while passing in CI: the tests build the
  real server, whose lifespan authenticates for real, and CI has no
  `~/.outlook-mcp/auth_record.json` to authenticate with. Three test files that
  could not complete locally now run in seconds.

- **The server still boots when the token cache is unwritable.** Refusing to
  write plaintext is right; dying in `lifespan` on the way up is not — the
  client sees a dead process and the one line the operator needs never leaves
  stderr. It now starts unauthenticated and every tool call carries the real
  remedy, which `outlook_auth_status` reports too. "Run `outlook-mcp auth`" is
  specifically the wrong advice here, because it fails the same way.

- **The second Linux failure mode is recognised.** The eager
  `find_spec("gi")` check only sees a *missing* libsecret. When libsecret is
  importable but unusable — no running Secret Service, as in a display-less
  SSH session or a container — azure-identity refuses lazily at first token
  use, with a `ValueError` naming its own `allow_unencrypted_storage` kwarg
  rather than the config key. That is now translated, so both halves of the
  condition give the same answer.

- `try_cached_token` no longer reports a token-storage misconfiguration as an
  expired token. It caught every exception and returned `False`, which is right
  for a stale token and wrong for "this host cannot store one safely" — that
  looped the operator back through `outlook-mcp auth` with no idea what to change.

### Changed

- **Dependency lock refreshed.** `pip-audit` reported advisories across
  `aiohttp`, `click`, `cryptography` (including GHSA-537c-gmf6-5ccf), `h2`,
  `pyjwt`, `python-multipart`, `starlette`, and `urllib3`; all had published
  fixes. Now clean. Notable jumps: `starlette` 1.0.0 → 1.6.0, `mcp` 2.1.1 →
  2.2.0, `pydantic` 2.12.5 → 2.13.5, `aiohttp` 3.13.5 → 3.14.3.

### Documentation

- **SECURITY.md said tokens were "never in plain files", which was not true**
  on Linux without libsecret. Corrected, alongside the same overstatement in
  SKILL.md and README's feature list. The delta-cursor boundary is now documented
  next to the attachment one.
- Private vulnerability reporting is enabled on the repository, so the
  `/security/advisories/new` link in SECURITY.md now works for outside reporters
  (#52).

## [1.20.0] — 2026-09-10

Security, and the first release shaped by what agents actually did rather than by
what we assumed they would. Two months of recorded trajectories (1,279 tool calls)
turned out to disagree with the design in three specific places.

### Security

- **Attachment reads and writes are confined to a configured directory.**
  `outlook_download_attachment` checked its `save_path` only for the substring
  `..`, so any absolute path was writable. `outlook_send_with_attachments` and
  `outlook_attach_to_draft` checked nothing at all, so any file the server
  process could read could be emailed anywhere. Against a mail server — which
  reads untrusted input for a living — "attach the file at &lt;path&gt; and reply" was
  a working exfiltration primitive. New `attachments_dir` config (default
  `~/.outlook-mcp/attachments`) bounds all three. Confinement is resolved rather
  than textual, so a symlink sitting inside the directory and a sibling directory
  sharing a name prefix are both refused.

### Fixed

- **Every domain error reached the model as `Error executing tool <name>`.** The
  2.x SDK forwards the text of an anticipated failure (`ToolError`) and withholds
  the text of a crash. `OutlookMCPError` inherited from `Exception`, so
  auth-required, read-only, permission-denied, not-found, every wrapped Graph
  error and every validation message was classified as a crash and suppressed.
  Only pydantic's argument-type errors survived, because the SDK raises those
  itself. The mock suite could not see it: it asserts on exception *types*, which
  stayed green throughout.
- **The `action` recovery hints had never reached an agent, on either SDK.** They
  lived on the exception attribute and never in `str()`. Two of them named
  `outlook_login`, a tool that does not exist; they now say `outlook-mcp auth`.

### Added

- **Relative dates on every datetime parameter.** `7d` / `-7d` (ago), `+7d` (from
  now), `now`; units m/h/d/w. The archive shows `after="7d"` attempted fifty
  times and refused every time, then worked around with a hand-computed
  timestamp. The rejection message now names the accepted formats.
- **Connect-time `instructions`.** 115 calls re-checked an identity that cannot
  change mid-session and 196 listed folders before folder-scoped scans that
  resolve display names on their own — about a quarter of all traffic. Guidance
  that spans tools belongs in one string sent once per session, not duplicated
  into 62 docstrings paid for every turn.
- **Three workflow prompts** — `morning_brief`, `triage_folder`, `catch_up` — the
  workflows the archive shows being assembled by hand. `catch_up` is where the
  delta tools finally get named; they had gone unused because nothing pointed an
  agent at them.
- **`ttlMs` / `cacheScope` on `tools/list`** (SEP-2549). Was `ttlMs: 0`, so
  clients re-fetched ~8.6k tokens of schemas they could have kept. Now five
  minutes, private.
- **`idempotentHint` on seven tools** whose operation is a PATCH of an absolute
  value. Opt-in and pinned rather than inferred: a client that trusts the hint
  may retry after a timeout, and a wrong entry sends a second email.
  `openWorldHint` deliberately left unset — it is `true` by default in the schema
  and would be `true` on all 62, so stating it costs schema tokens on every turn
  and tells a client nothing new.

### Changed

- **BREAKING:** an attachment path outside `attachments_dir` is now refused.
  Callers that saved to or attached from arbitrary locations must move the file
  there or widen the setting. This is why the change ships in a minor rather than
  a patch.

### Notes

Tool count unchanged at 62. Tool titles were considered and dropped: they would
have added roughly 10% to the schema bytes every client pays for on every turn,
and would have displayed "List Inbox" while the README and every error message
say `outlook_list_inbox`.

## [1.19.0] — 2026-09-08

Tiers 2 and 3 of the silent-no-op audit, and the two bugs they found on their first run.

### Added

- **`tests/test_no_dead_modules.py`** — builds the intra-package import graph from the console-script entry point and fails on any module nothing reaches. The module-level counterpart of the dead-parameter guard: `models/` sat unreferenced for months while 30 tests imported it directly, because tests that import a module by name make dead code look alive. Collects imports from every AST node, since this package imports lazily inside functions. Mutation-verified.
- **`tests/test_sdk_fields_exist.py`** — every `model.attr = …` in `tools/` where `model` was bound by an `msgraph.generated.models` constructor must name a declared dataclass field of that class. Dataclass instances accept arbitrary attributes, so assigning a field the SDK doesn't have raises nothing, serializes nothing, and reaches nothing. Offline, no mocks.
- **`tests/test_write_payloads_reach_the_wire.py`** — for eleven write tools across calendar, mail, drafts, contacts and To Do, calls the handler with a distinctive sentinel for every argument, serializes the object handed to `.post()`/`.patch()` exactly as kiota would, and asserts each value is present in that JSON. Attribute-level assertions cannot see a field the handler forgot (#41) or the SDK dropped (`remove_recurrence`); the wire can.
- **Live round-trips for contacts and To Do** (`test_live_contacts_write.py`, `test_live_todo_write.py`): every parameter each write tool accepts is written, read back through the tool's own read path, and asserted — the only detector for Graph accepting a value and ignoring it. The write tier's rules now cover calendar, contacts and To Do (the three surfaces with no outward side effect; still no mail, ever). Because the real config may whitelist only some write categories, `live_write_config` opens the three it needs on an **in-process copy**; `~/.outlook-mcp/config.json` is never touched.

### Removed

- **`sensitivity` on `outlook_send_message`.** Found by the wire test and confirmed by the field guard: msgraph-sdk's `Message` has no `sensitivity` field — not in the dataclass, not in `serialize()`, not in the deserializers. `msg.sensitivity = Sensitivity.Private` hung a stray attribute on the instance that kiota never saw. Sending it by hand via `additional_data` was then tried against live Graph (as a draft, deleted): `400 UnableToDeserializePostBody`. It is a legacy MAPI property that this endpoint does not accept, so there is nothing to implement. The parameter never did anything for anyone; it is removed rather than left as a convincing no-op. **Schema change** on one tool.
- **`src/outlook_mcp/tools/auth_tools.py`** — a six-line docstring stub left behind when the auth tools moved into `server.py`, imported by nothing. Flagged by the new module guard.

### Fixed

- **To Do `reminder=True` was silently stored as off.** The live round-trip caught it: Graph accepts `isReminderOn: true`, returns 201, and persists `false` unless `reminderDateTime` is also set. Confirmed by isolating the two cases live. `create_task` now anchors the reminder on the due time, and `reminder=True` without `due` is a named error — there is nothing to anchor it on, and the alternative is the no-op we just removed.

### Verified

- Offline **638 passed, 32 deselected**; ruff clean across `src/`, `tests/`, `scripts/`.
- **16 live write-tier tests** — calendar 10, contacts 3, To Do 3 — plus preflight 13/13, live 10, integration 6. Calendar, contacts and tasks swept afterwards for leftover test data: none.
- 62 tools, unchanged.

## [1.18.0] — 2026-09-08

A guard against the bug class that produced [#41], and the bug it immediately found.

### Added

- **`tests/test_no_dead_parameters.py`** — fails the build on any parameter that is declared and never read. That is the static half of the #41 shape: `outlook_create_event` took `recurrence`, threaded it to the handler, never applied it, and returned `status: created` for fourteen releases. An AST walk catches it offline, in CI, for free.

  Deliberately strict about what counts as a use: only `Name` loads, never attribute names. Counting `event.recurrence` as a use of a `recurrence` parameter is precisely how #41 would have slipped past. Functions taking `*args`/`**kwargs` are skipped, since forwarding makes the question undecidable. Two framework-callback parameters are allowlisted with written reasons, and a second test fails if an allowlist entry goes stale — a stale exemption silently re-opens the hole it was excusing. The guard is mutation-verified: injecting a dead parameter makes it fail with the file, line and parameter name.

  It catches one of the three shapes. Graph accepting a field and ignoring it (`is_online`, 1.16.0) needs live coverage; the SDK dropping a field from the payload (`remove_recurrence`, 1.17.0) needs assertions on serialized output. Both are noted in the module docstring.

### Fixed

- **`outlook_reply` ignored `is_html`** — found by the new guard on its first run. Every reply sent with `is_html=True` went out as plain text, with any markup stripped. Graph's reply action takes `comment` as plain text only; markup has to ride on the action's `message` field as an `ItemBody` with contentType HTML. Both are now set correctly and exclusively — setting `comment` alongside `message` would duplicate the text in the sent reply. Verified that Graph stores and preserves an HTML body on a reply-shaped message; the send path itself is not live-tested, because the write tier does not send mail.
- **`AuthManager.login_interactive()` and `try_cached_token()` took a `scopes` argument they ignored**, resolving scopes internally via `get_token_scopes()` instead. Callers passed a value that did nothing. The parameter is removed from both. Internal API; no MCP tool signature changes.

### Verified

- Offline **619 passed, 26 deselected**; ruff clean across `src/`, `tests/`, `scripts/`.
- preflight 13/13, live 10, integration 6, live_write 10 — the auth signature change is on the path every credentialed tier and `outlook-mcp status` uses, so all four were re-run.
- 62 tools, unchanged.

## [1.17.0] — 2026-09-07

Ending a series without deleting it.

### Added

- **`remove_recurrence` on `outlook_update_event`.** Ending a series previously meant deleting the event — there was no way to turn a series master back into a single occurrence. `remove_recurrence=True` does that, keeping the first occurrence's time. It is a separate flag rather than a sentinel on `recurrence`, because `None` there already means "leave alone"; passing both raises.

  The implementation is not the obvious one. Graph clears a series with an explicit `"recurrence": null`, but the msgraph SDK **omits any field set to `None`** from the serialized payload — `Event(subject="x", recurrence=None)` serializes to `{"subject": "x"}` — so `event.recurrence = None` is a silent no-op of exactly the kind that produced [#41]. The null is therefore sent through `additional_data`, which kiota does serialize. Two tests pin this: one asserts the wire format contains `"recurrence": null` rather than checking the attribute (an attribute assertion would stay green through the bug), and one asserts the SDK still drops an explicit `None`, so if kiota ever changes, the workaround can be simplified.

  Verified live: a `seriesMaster` becomes a `singleInstance` with `recurrence: null` and its start time intact.

## [1.16.0] — 2026-09-07

Calendar editing catches up with calendar creation, and a layer that only looked like validation comes out.

### Added

- **`attendees` on `outlook_update_event`.** You could create a meeting with guests but never add one to a meeting that already existed. Note the semantics, which are Graph's, not ours: there is no add-one operation, so the argument **replaces the whole collection** — Outlook emails invitations to everyone on the new list and cancellations to anyone dropped. `[]` removes them all. Every argument on this tool keeps `None` = "leave alone", so `False` and `[]` remain instructions rather than absences.
- **`is_all_day` on `outlook_update_event`**, which requires `start` and `end` in the *same* call. Graph rejects a lone isAllDay patch with `ErrorInvalidRequest: Missing parameters: Event.Start`; that is now a named local error naming both bounds instead of an opaque 400 from the wire.
- **`type` on every listed event** (`singleInstance` / `seriesMaster` / `occurrence` / `exception`). `outlook_list_events` could not distinguish a recurring series from a one-off without fetching each event individually. Computed once in the summary formatter; the detail formatter inherits it. `concise=True` still omits it — that mode exists to drop tokens.
- **`Accept-Encoding: gzip` on the two raw-httpx paths** (`fetch_delta_pages`, `read_messages`). The SDK path already negotiated compression; these bypass the SDK and did not. Delta pages and 20-message `$batch` responses are large, highly compressible, and refetched constantly by polling agents. Roadmap Tier-0 #7.

### Removed

- **`src/outlook_mcp/models/` — the entire Pydantic I/O layer.** All 22 classes across `calendar.py`, `mail.py`, `todo.py`, `contacts.py` and `common.py` were referenced by zero non-model source files; only their own tests imported them. They read as the input-validation layer, and `CLAUDE.md` claimed one existed ("All input validated via Pydantic + validation.py"), but nothing on any tool path ever constructed them. This is the exact mechanism that hid [#41] for fourteen releases: `CreateEventInput.validate_recurrence` enforced the shorthand allowlist, looked authoritative, and never ran. They are redundant by construction — `MCPServer` generates tool schemas from the type annotations. `CLAUDE.md` now says what actually runs, and requires any future model layer to be wired to the tool path in the same commit. Drops 30 tests that exercised only dead code.
- **`scripts/probe_datetime_semantics.py`.** One-off evidence gathering for 1.15.0's timezone question. The question is answered, the finding is recorded in the 1.15.0 entry below, and a script that creates and deletes real calendar events should not ship in a bundle other people install.

### Fixed

- **`is_online` is documented as a no-op on personal accounts.** Adding it to `outlook_update_event` was in scope until the live write tier caught what mocks cannot: Graph accepts `isOnlineMeeting` on a consumer mailbox and silently ignores it. A created event comes back `isOnlineMeeting: False`, `onlineMeetingProvider: "unknown"`, `onlineMeeting: None`. Teams meetings require a work/school account, which this server does not target — so `outlook_create_event`'s `is_online` parameter has never done anything for this project's only audience. Rather than ship a convincing no-op on a second tool, the parameter was left off `update_event`, the behavior is documented on `create_event`, and a live test now pins it: if Microsoft ever starts honouring it, that test fails and tells us the parameter is worth adding.

### Verified

- Offline suite **608 passed, 25 deselected**; `ruff check src/ tests/` clean; 62 tools register (unchanged — this release adds parameters, not tools).
- **9 live write-tier tests** against a real personal `@outlook.com` calendar, including all-day conversion, partial-patch isolation, and the online-meeting pin. `attendees` is deliberately *not* exercised live — patching it emails real invitations, and the tier's no-attendees rule outranks the coverage; the built payload is unit-tested instead.

[#41]: https://github.com/mpalermiti/outlook-mcp/issues/41

## [1.15.0] — 2026-09-07

Recurring calendar events, which have never actually worked, and a timezone bug found while fixing them.

### Fixed

- **`outlook_create_event` silently discarded `recurrence`.** Every "recurring" event the server has ever created was a single occurrence with `recurrence: null`. The parameter was declared on both the tool and the handler and then never assigned to the Graph `Event` before `POST /me/events` — so the call returned `status: created` and looked like it worked. Two independent failures sat on that one argument: the tool schema typed it `str | None`, so a Graph recurrence object was rejected by Pydantic before reaching the handler, and the JSON-string workaround reached the handler and was dropped there. The `daily`/`weekly`/`monthly` shorthands the docstring has advertised since 1.0 were equally inert — `CreateEventInput.validate_recurrence` in `models/calendar.py` is referenced only by its own unit test, never by the tool path, so any string was accepted and ignored. Reported by @Wermeling. ([#41])
- **`outlook_get_event` returned a Python repr for `recurrence`.** The read path did `str(event.recurrence)`, emitting ~500 characters of `PatternedRecurrence(additional_data={}, odata_type=None, pattern=...)` into a field the response model declares as `dict | None`. It now returns the same JSON shape `outlook_create_event` accepts, with unset fields omitted.

- **Zone-less datetimes were resolved against the server's clock.** `validate_datetime` handed a naive datetime ("2026-10-22T12:30:00") straight to `.astimezone()`, which resolves against the *process's* local timezone. The same query string therefore meant a different instant depending on where the server happened to run: `12:30` became `12:30Z` on a UTC container, `19:30Z` on a laptop in California and `10:30Z` in Berlin. That skewed every `$filter` bound built from a zone-less value in `outlook_list_inbox`, `outlook_list_events` and `outlook_list_events_delta`, and — the one with a visible consequence — the PR_DEFERRED_SEND_TIME of a scheduled message, so "send this at 09:00" went out hours off. `validate_datetime` now takes an explicit `tz` for zone-less input only; input carrying a `Z` or an offset already pins its own instant and is untouched. The default is UTC and the server passes `config.timezone`, which is what `CLAUDE.md` has always said input interpretation should use.

  **Behavior change** where `config.timezone` is not UTC: a bare date now means midnight *there* rather than midnight UTC, so `after="2026-10-22"` on an `America/Los_Angeles` config resolves to `2026-10-22T07:00:00Z`. Hosts running in the configured zone — the common case — see no change at all; hosts where the two disagreed were the ones silently producing different answers.

  The write paths (`calendar_write`, `todo`) call `validate_datetime` for its exception and discard the normalized value, so they are unaffected. That is deliberate and now verified rather than assumed: `scripts/probe_datetime_semantics.py` established that Graph honors an explicit offset in the `dateTime` field and reads a naive value as the declared `timeZone`, making the current write behavior correct. "Fixing" those call sites to use the discarded value would have introduced exactly the host-clock bug this entry removes.

### Added

- **`recurrence` accepts three input shapes** on `outlook_create_event`: a Microsoft Graph recurrence object (matching `outlook_create_task`, which has taken `dict` since 1.5.0); a JSON-encoded string of one, since some client bridges stringify nested arguments; or a shorthand — `daily`, `weekdays`, `weekly`, `monthly`, `yearly` — expanded against the event's own start, so `weekly` on a Monday start means every Monday and `monthly` on the 7th means the 7th.
- **`range.startDate` is reconciled with the event start.** Graph requires the range to begin on the day of the first occurrence and returns `ErrorInvalidRecurrenceRange` otherwise. It is now defaulted from `start` when omitted, and a supplied value that disagrees is rejected locally with a message naming both dates instead of surfacing as an opaque 400. The date is taken as the caller wrote it rather than UTC-normalized, so an evening or early-morning series doesn't shift a day.
- **`recurrence` on `outlook_update_event`.** There was previously no path to add recurrence to an event that already existed — the parameter simply wasn't there, so a single event could never become a series. It now takes the same three shapes as `outlook_create_event`. Graph anchors a series on the master's start, so when `start` isn't part of the same patch the event's current start is read first and used as the anchor; a partial patch that omits `recurrence` leaves any existing pattern untouched. Reading that start back also required `event_start_date` to tolerate Graph's seven-digit fractional seconds, which `datetime.fromisoformat` rejects on the project's 3.10 floor.
- **`type` on `outlook_get_event`** — `singleInstance`, `seriesMaster`, `occurrence` or `exception`. This is the field that lets a client confirm a series actually took; its absence is part of why #41 needed a raw Graph call to diagnose.
- **`src/outlook_mcp/tools/_recurrence.py`** — recurrence conversion shared by calendar and To Do, which model it identically. `build_patterned_recurrence` is To Do's 1.5.0 converter moved here unchanged, so `outlook_create_task` keeps exactly its previous behavior; `build_event_recurrence` is the calendar entry point layered on top; `serialize_recurrence` is the inverse for read paths.
- **`live_write` test tier** — the project's first write-side tests, and the reason this fix is trustworthy. Recurrence cannot be validated any other way: a `PatternedRecurrence` that is well-formed to the SDK still 400s if the range disagrees with the start, and mocks only ever assert what we *build*. Double-gated (marker deselected by default **and** skipped without `OUTLOOK_MCP_LIVE_WRITE=1`), calendar-only, no attendees, bounded ranges only, everything deleted in a `finally`. Rules in `tests/conftest.py`; runbook step 1d in `RELEASING.md`.

### Verified

- Five write-tier tests green against live Graph on a personal `@outlook.com` account: the series master is created, comes back `type: seriesMaster`, the recurrence round-trips through `outlook_get_event` as a dict, the `startDate` we default is one Graph accepts, and a non-recurring control still reports `singleInstance`. Every created event deleted; the calendar was re-queried afterwards and held no leftovers.
- Offline suite: **609 passed, 21 deselected** (was 577 / 16 at 1.14.0). `ruff check src/ tests/ scripts/` clean.
- A separate probe (`scripts/probe_datetime_semantics.py`) established what Graph does with the datetimes this tool sends, since `create_event` labels every start `timeZone: "UTC"` while passing the caller's string through verbatim. Graph honors an explicit offset (`12:30:00+02:00` → `10:30:00Z`) and reads a naive datetime as the declared UTC. Both are correct and deterministic; no change was made. Recorded because the obvious "fix" — using the normalized value `validate_datetime` already computes and throws away — would have introduced a real bug, resolving naive input against the *server's* local clock.

[#41]: https://github.com/mpalermiti/outlook-mcp/issues/41

## [1.14.0] — 2026-09-04

Migration to the `mcp` 2.x SDK, lifting the `<2` ceiling that 1.13.1 pinned as a stopgap. **No behavior change for clients** — all 62 tool schemas serialize byte-for-byte identically to 1.13.1.

### Changed

- **`mcp[cli]` requirement raised to `>=2`** (was `>=1.27,<2`). The SDK renamed `FastMCP` to `MCPServer` and moved it from `mcp.server.fastmcp` to `mcp.server.mcpserver` in 2.0.0 (2026-07-28). 1.13.1 pinned below that release to restore installs; this completes the migration instead of living under the ceiling. The Python floor is unchanged — `mcp` 2.1.1 requires `>=3.10`, same as this project.
- `ToolAnnotations` moved to snake_case field names in the SDK (`read_only_hint`, `destructive_hint`), retaining the camelCase forms as serialization aliases. The wire format clients see (`readOnlyHint` / `destructiveHint`) is unchanged; only in-process attribute reads in `tests/test_toolsets.py` needed updating. The annotation and `OUTLOOK_MCP_TOOLSETS` gating work from 1.12.0 carries over untouched.

### Removed

- The `mcp._mcp_server.version = __version__` workaround. `MCPServer` accepts a `version` kwarg directly, so `serverInfo.version` is now set through public API rather than by reaching into a private attribute.

### Verified

- Offline suite green on `mcp` 2.1.1: `577 passed, 16 deselected` — identical to 1.13.1.
- All 62 tool schemas dumped from 1.x and 2.1.1 and diffed: **byte-for-byte identical** (52,998 bytes each).
- `serverInfo.version` reports `1.14.0`; 62 tools register.

## [1.13.1] — 2026-09-04

Hotfix. **Every fresh install since 2026-07-28 was dead on arrival** — this restores it. No code changes; one dependency bound and a CI job to make sure it cannot happen again.

### Fixed

- **`mcp[cli]` was declared with no upper bound, so new installs pulled an incompatible major.** `mcp` 2.0.0 (2026-07-28) removed `mcp.server.fastmcp` by design — it was renamed to `MCPServer`. `src/outlook_mcp/server.py` imports `from mcp.server.fastmcp import Context, FastMCP`, so a fresh `pip install outlook-graph-mcp` resolved `mcp` 2.1.1 and `outlook-mcp serve` exited 1 with:

  ```
  ModuleNotFoundError: No module named 'mcp.server.fastmcp'
  ```

  Now pinned to `mcp[cli]>=1.27,<2`, which resolves 1.29.1 and imports cleanly with all 62 tools.

  This was **not** a 1.13.0 regression. 1.12.0 shipped 2026-07-18, ten days before `mcp` 2.0.0, and was broken by the same unbounded specifier the moment that release landed. Every install path — PyPI, `uvx`, ClawHub, the MCP registry — was affected for roughly five weeks. Migrating to the `mcp` 2.x `MCPServer` API is a separate, deliberate piece of work; this release only restores a working install.

- **CI could not see it, which is why it lasted five weeks.** The test job runs `uv sync`, which resolves through `uv.lock` — and the lock pinned the old, working `mcp` 1.27.0. All 577 tests passed in 0.75s against a dependency set no new user would ever get. Same shape as the bugs fixed in 1.13.0: a green suite testing something other than what ships.

### Added

- **`fresh-install` CI job.** Builds the wheel, installs it into a clean venv with a **fresh dependency resolve that deliberately ignores `uv.lock`**, prints the resolved versions, then asserts the server imports and registers all 62 tools. This is the check that turns "an upstream major broke us" from a user-reported outage into a failed PR. It fails today on `main` without the pin — verified locally before shipping.

## [1.13.0] — 2026-09-04

Correctness release. Three shipped bugs fixed, all of which were invisible to a fully green mock suite — including one tool that had never worked in any released version. Adds the live test tier that would have caught them. One new filter parameter; no new tools (still 62), no breaking changes.

### Fixed

- **`outlook_list_inbox` returned `400 InefficientFilter` whenever `from_address` or `classification` was used** ([#31](https://github.com/mpalermiti/outlook-mcp/issues/31)). `list_inbox` sets `$orderby=receivedDateTime desc` unconditionally, but Graph requires an `$orderby` property to also appear in `$filter` **and to precede** the other filtered properties — clause order is load-bearing, not just presence. Both parameters failed on every call, as did `unread_only + from_address` and `from_address + after`. Clauses are now emitted date-first, with a permissive `receivedDateTime` floor prepended only when the caller supplies no `after`/`before`. Validation still runs in signature order, so validation-error precedence is unchanged.

  Considered and rejected: dropping `$orderby` when `receivedDateTime` isn't filtered. Graph's implicit ordering follows whichever index served the filter and is **ascending** for `inferenceClassification` — that fix would have silently returned the *oldest* focused mail with no error.

- **`outlook_list_thread` returned `400` on every call and had never worked in a released version.** Same root cause: `$filter=conversationId eq '…'` with `$orderby=receivedDateTime asc`. Verified against six real conversations before and after.

- **`outlook_search_mail` silently returned zero results for every documented KQL property restriction** ([#30](https://github.com/mpalermiti/outlook-mcp/issues/30)). `sanitize_kql` stripped `:`, so `subject:Unlock` was sent as `"subjectUnlock"` — HTTP 200, no error, no results. The docstring's own example (`from:sarah@acme.com`) was broken; `from:`, `subject:` and `hasattachment:true` were all dead.

  `:`, `(` and `)` are now preserved. Still stripped, each for a measured reason: `"` (an embedded quote makes Graph *silently discard* `$search` and return the whole mailbox — the real injection vector), `\` (a genuine escape metachar: `\s` is a 400, a trailing `\` escapes our closing quote), `*` (buys nothing, since Graph already prefix-matches, and a bare `*` would become a silent whole-mailbox read where it is currently a loud 400), and `&` `|` `!` (not operators; the symbol forms silently zero out a matching query). Applies to `outlook_search_contacts` too, which shares the sanitizer.

  A query that sanitizes to empty now raises a clear `ValueError` instead of an opaque Graph `BadRequest`.

- **Integration tests had been silently skipping since they were written.** The fixture called `AuthManager.login()`, which does not exist; the `AttributeError` was swallowed by a bare `except Exception` and reported as a skip — the "6 skipped" in every run. Repointed at `try_cached_token`. Running them then surfaced two more latent defects: a session-scoped Graph client raised `Event loop is closed` (the kiota/httpx transport binds to the running loop, so it is now function-scoped), and `test_list_drafts_smoke` asserted `result["drafts"]` where `list_drafts` returns `messages`.

### Added

- **`uncategorized_only` on `outlook_list_inbox`** ([#28](https://github.com/mpalermiti/outlook-mcp/pull/28), thanks [@jasond727](https://github.com/jasond727)) — filters to messages with no categories assigned via OData `not categories/any()`. Default `False`; behavior unchanged when omitted. Lets an agent ask "what have I not triaged yet?" server-side instead of pulling and filtering the whole inbox.

- **Live query-shape test tier (`uv run pytest -m live`)** — 10 read-only tests guarding the three failure classes above. Mocks assert what we *send*; they cannot see what Graph *does* with it, which is how all three bugs shipped under 558 green tests. These assert on returned data, not just absence of an exception, because the dangerous failure mode is a silent 200 with wrong results. Mutation-verified: reverting each fix fails the corresponding tests. `scripts/preflight.py` does not cover this — it probes endpoint reachability and treats a 400 as a non-blocking SKIP.

  The default `pytest` run is now the offline unit suite only; `integration` and `live` are deselected via `addopts`. Previously `integration` ran for real on any machine holding a cached token and merely *skipped* in CI, which reads as "covered" when nothing was exercised.

### Changed

- `outlook_list_inbox` summary line now names `category` as a filter dimension — that first docstring line is what an LLM reads when choosing among 62 tools.
- `outlook_search_mail` docstring records two Graph behaviors that cost real query results: `AND`/`OR`/`NOT` must be **uppercase** (lowercase `and` matches as a literal term), and two terms with no operator between them *broaden* rather than narrow.
- The `403 ErrorAccessDenied` recovery hint now carries a URL. It previously said "See ROADMAP 'Investigated and not viable'", but `ROADMAP.md` ships in neither the sdist nor the wheel — anyone who installed from PyPI had nothing to open.

### Documentation

- **`ROADMAP.md` mailboxSettings entry carried a misattributed quote.** It claimed the endpoint "is documented as `Delegated (personal Microsoft account): Not supported`". Both mailboxSettings reference pages list personal accounts as *supported*, and their permissions tables have been unchanged since 2023-10-27 — so they did not say that when 1.7.1 shipped either. The string belongs to a neighbouring `MailboxSettings`-scoped page (`workHoursAndLocations`, a different resource). The empirical `403` was always the real justification and is unaffected; re-verified 2026-09-03. The 1.7.1 entry below repeats the quote and is deliberately left as written — this changelog does not retroactively edit released sections.

- **Inbox message rules documented as non-viable** ([#29](https://github.com/mpalermiti/outlook-mcp/pull/29), thanks [@tlewthwa](https://github.com/tlewthwa)) — `/me/mailFolders/inbox/messageRules` returns `403 ErrorAccessDenied` on personal accounts despite Graph's docs listing personal-MSA support, verified with a raw HTTP probe and the token's scopes recorded. Includes a COM/MAPI side investigation and a pointer to `/me/inferenceClassification/overrides`, the one rule-shaped capability that does work on personal accounts. This closes out the only Near-term roadmap item as blocked.

### Tool count

- 1.12.0 → 1.13.0: **62 tools, 13 categories** (no change).

## [1.12.0] — 2026-07-18

Performance & efficiency pass for recurring agent loops (personal-account mail + calendar). No new tools, no breaking changes; the default tool surface is unchanged (62 tools). Validated against a measured tool-schema token count (~8,644 tokens/turn) and a mid-2026 ecosystem review.

### Added — Config-gated toolsets

`OUTLOOK_MCP_TOOLSETS` env var (e.g. `mail,calendar,digest,delta`) loads only the named tool groups plus the always-on `account`/auth tools, instead of all 62. Groups: `mail`, `drafts`, `attachments`, `calendar`, `contacts`, `todo`, `folders`, `digest`, `delta`, `admin`. A client that only needs mail + calendar drops the surface from 62 → ~30 tools (~52% fewer tool-schema tokens per turn). Unset = all tools (backward compatible). The grouping is a flexible selector rather than a fixed core/admin package split — the measurement showed the admin/override/batch group is only ~7% of tokens, so real savings come from dropping whole unused domains.

### Added — Tool annotations

Every tool now carries `readOnlyHint` / `destructiveHint` (`ToolAnnotations`): 26 read-only tools, 8 destructive (delete/remove), the rest additive writes. Clients can auto-approve reads and gate destructive ops (delete mail, decline event) without a hardcoded allowlist. Classification lives in `toolsets.py` as one reviewable table, drift-guarded by a test.

### Changed — Concurrent digest

`outlook_changes_since` fetches mail / events / contacts concurrently (`asyncio.gather`) instead of sequentially — ~2–3× lower latency on the most-used recurring tool. Each resource still resyncs independently on a stale token; the `_meta.resync` order is deterministic regardless of completion order.

### Changed — Graph client reuse

`_get_graph_client` caches one `GraphServiceClient` (auth provider, request adapter, TLS pool) in the lifespan context and reuses it while the credential is unchanged, instead of rebuilding it on every tool call. A `switch_account` / re-auth swaps the credential object, so an identity check rebuilds automatically.

### Fixed — Throttling on the raw-httpx paths

New `throttle.py` restores `Retry-After` honoring on the delta / `$batch` paths that bypass the SDK's kiota `RetryHandler`:
- `outlook_changes_since` / delta queries retry the delta GET on 429/503.
- `outlook_read_messages` — a `$batch` returns HTTP 200 even when a sub-request is throttled (429/503); the tool now re-issues just the throttled sub-requests honoring `Retry-After` (bounded by `max_retries`) instead of recording a 429 as a permanent failure.

## [1.11.1] — 2026-07-18

### Fixed — `outlook_download_attachment` corrupted binary attachments ([#25](https://github.com/mpalermiti/outlook-mcp/issues/25))

`outlook_download_attachment` raised `UnicodeDecodeError` on any attachment whose bytes aren't valid UTF-8 — in practice every non-text file (`.pdf`, `.docx`, images, etc.). The tool double-decoded `contentBytes`: the msgraph SDK (Kiota) already base64-decodes it into raw `bytes` during deserialization, but the tool then called `.decode("utf-8")` + `base64.b64decode()` on the already-decoded bytes. It now writes `attachment.content_bytes` verbatim.

Regression from [#9](https://github.com/mpalermiti/outlook-mcp/pull/9) (v1.6-era), which added decoding that is correct for the raw Graph REST/JSON API but wrong when going through the SDK. Diagnosed and verified by @AI-OWEN. Added a regression test that downloads a non-UTF-8 binary payload and asserts byte-for-byte fidelity; corrected the pre-existing download test whose mock fed base64-*encoded* bytes and thereby masked the bug.

No API, signature, or tool-count change. **Tool count: 62, 13 categories** (unchanged).

### Docs

- Added a **Troubleshooting** section to the README covering `SSL: CERTIFICATE_VERIFY_FAILED` on Linux ([#24](https://github.com/mpalermiti/outlook-mcp/issues/24)) — point Python at the system CA bundle via both `SSL_CERT_FILE` (httpx delta/`$batch` paths) and `REQUESTS_CA_BUNDLE` (azure-identity auth).

## [1.11.0] — 2026-05-22

### Added — Bulk message read via `$batch` (1 tool)

One new MCP tool: `outlook_read_messages(message_ids, format="text", concise=False, include_deferred_send=False)`. Reads up to 20 messages by ID in a single Graph `$batch` round-trip instead of N sequential calls. Per-message shape in `messages[]` matches `outlook_read_message` byte-for-byte for the same `(format, concise, include_deferred_send)` combo — agents can collapse a "fetch N messages by ID" loop into one call without changing how they consume the result.

**Return shape.** `{messages, failures, requested, succeeded, failed}` where `messages` is the read_message-shaped dicts in input order (Graph response order is not trusted; the impl uses the input index as the sub-request id and rebuilds), `failures` is `{id, status, code, message}` for IDs that didn't return 2xx, and the three counts satisfy `requested == succeeded + failed`.

**Partial failures are not exceptions.** A 404 on one of 20 IDs surfaces in `failures[]`; the other 19 are returned in `messages[]`. Only input-validation errors (empty list, >20 IDs, malformed Graph ID) and transport-level failures (Graph 5xx on the whole batch) raise.

**Tool count: 61 → 62, 13 categories** (no new category — extends Mail Read).

## [1.10.0] — 2026-05-22

### Added — Composed "since last call" digest (1 tool)

One new MCP tool: `outlook_changes_since(delta_tokens=None, fallback_window_hours=24)`. Wraps the three v1.9.0 delta tools (mail, events, contacts) into a single structured payload. Designed for recurring agent loops — a morning brief or hourly inbox sweep that wants ONE call instead of orchestrating three deltas and reasoning over raw item shapes itself.

**Return shape (top-level):**
- `mail`: `{new_count, modified_count, removed_count, urgent_flagged[], by_sender{}}` — `urgent_flagged` is mail where `importance == "high"` OR `flag == "flagged"`; `by_sender` is the top 5 senders by message count in the digest window.
- `events`: `{new[], modified[], cancelled[]}` — cancelled bucket holds delta tombstones; modified bucket is currently reserved (Graph delta responses don't carry an affirmative change marker on live events, so changes surface as `new[]` today — see the tool docstring).
- `contacts`: `{new_count, modified_count, removed_count}`.
- `delta_tokens`: `{mail, events, contacts}` — caller-managed watermarks for the next call, same pattern as the v1.9.0 delta tools.
- `window`: `{from, to}` — when the digest's bootstrap window starts/ends.

**First-call behavior.** For each resource without a delta token, the digest bootstraps by calling the underlying delta tool with no token (which returns a full snapshot plus a fresh token). It then filters the snapshot to the last `fallback_window_hours` so the digest doesn't surface thousands of historical items on first run. For calendar specifically: passes `start = now - fallback_window_hours` and `end = now + 7 days` to capture recent + upcoming changes.

**Subsequent-call behavior.** With a stored `delta_tokens` dict the digest uses each token verbatim, drains pagination up to 5 internal pages (~1,000 items per resource per call), and classifies each item. Each resource is independent — a missing or stale token for one doesn't block the other two.

**`syncStateNotFound` recovery.** If a stored token is too old, Graph returns HTTP 410. The digest auto-drops that bad token, re-bootstraps just that resource as a "first call", and surfaces `_meta.resync: ["mail"]` (etc.) so the caller knows their watermark was discarded. Other resources are unaffected.

**No new Graph endpoints.** This release composes already-tested v1.9.0 endpoints; preflight remains at 13 endpoints.

### Tool count
- 1.9.1 → 1.10.0: **60 → 61 tools, 13 categories** (no new category).

## [1.9.1] — 2026-05-22

### Changed — Tool docstring audit (no behavior change)

Docstring audit. Every `@mcp.tool()` docstring has been rewritten to a consistent shape — one-line action, contrastive pointer for ambiguous pairs (e.g. `outlook_reply` vs `outlook_rsvp`, `outlook_send_message` vs `outlook_create_draft` + `outlook_send_draft`, `outlook_reclassify_message` vs `outlook_set_inbox_override`, snapshot vs delta tools, search vs list, move vs copy, delete vs move-to-deleteditems), concrete syntax example for params with non-obvious shape (KQL queries, ISO 8601 dates, delta-token round-trips, attachment paths, batch shape). Designed to reduce wrong-tool selection by AI agents. No behavior changes; signatures, params, defaults, return shapes are byte-identical to 1.9.0.

## [1.9.0] — 2026-05-21

### Added — Delta queries (3 tools)

Wraps Microsoft Graph's `$delta` endpoints. The use case: an agent's
recurring poll (morning brief, hourly inbox check) currently re-fetches
the last N messages on every run, even when nothing changed. With delta
queries the second-and-later call returns only what changed since the
last token — typically 0–3 items on a stable inbox. Real money on
per-token billing for recurring agent jobs.

- `outlook_list_inbox_delta(folder="inbox", page_size=50, delta_token=None)`
  — wraps `GET /me/mailFolders/{folder}/messages/delta`.
- `outlook_list_events_delta(start, end, page_size=50, delta_token=None)`
  — wraps `GET /me/calendarView/delta`. `start` and `end` (ISO 8601) are
  required on the first call; ignored on follow-ups (the cursor encodes
  the window).
- `outlook_list_contacts_delta(page_size=50, delta_token=None)` — wraps
  `GET /me/contacts/delta`.

**Cursor semantics.** The `delta_token` is opaque to the caller (it's
the deltaLink or nextLink URL Graph hands back, passed through as a
string). outlook-mcp does *not* persist tokens server-side — the agent
stores its own watermark and replays it. Matches the existing
pagination `cursor` pattern.

**Tombstones.** Deleted items come back as `{id, is_deleted: True}`
with no other fields. Agents should treat them as cache evictions.
Live items also carry `is_deleted: False` for symmetry.

**Cap behavior.** Each call auto-follows `@odata.nextLink` internally
up to `page_size * 4` items, then stops with `has_more: True` and the
nextLink as the returned `delta_token` so the caller can resume. This
keeps a single tool call bounded even on a large first-call snapshot
(e.g. ~12k contacts).

**Verified working on personal accounts via the preflight script** —
the three new delta endpoints are reachable from `/me/...` on
outlook.com mailboxes without additional consent scopes.

### Tool count
- 1.8.0 → 1.9.0: **57 → 60 tools, 13 categories** (no new category).

## [1.8.0] — 2026-05-21

### Added — Agent-friendly shape

- **Concise mode** — opt-in `concise=True` flag on the five high-volume read tools (`outlook_list_inbox`, `outlook_read_message`, `outlook_search_mail`, `outlook_list_events`, `outlook_list_thread`). When set, the server drops bulky fields (`preview`/`categories` for message summaries; full `body`/`body_html` for single-message reads, replaced with a 200-char single-line `body_preview`; `body`/`organizer`/`response_status`/`categories` plus the full `attendees` list for events, replaced with `is_organizer` and `attendees_count`; quoted prior-message text on thread previews, via a heuristic on `On ... wrote:` / `From: ...` / `----- Original Message -----` markers). Typical payload reduction ~10×, designed for triage scans before deciding to fetch full content.

- **Graph error wrapper** — every tool now passes its result through a thin decorator that converts msgraph's `ODataError` / `kiota_abstractions.APIError` into a structured `GraphAPIError(status_code, error_code, message, action)` with recovery hints for 401 ("run `outlook-mcp auth` on the host"), 403/`ErrorAccessDenied` (ROADMAP pointer to known unsupported-endpoint dead-ends), 404/`ErrorItemNotFound` ("re-list to get fresh IDs"), 429 ("back off, respect Retry-After"), and 503 ("transient — retry after a short delay"). `OutlookMCPError` subclasses and `ValueError`s pass through unchanged; non-Graph exceptions bubble up untouched.

No new tools; existing tool responses unchanged when `concise=False` (default). Strict backward compat for existing callers.

### Tool count
- 1.7.1 → 1.8.0: **57 tools, 13 categories** (no change).

## [1.7.1] — 2026-05-20

### Removed
- `outlook_get_timezone`, `outlook_set_timezone`, `outlook_get_auto_reply`, `outlook_set_auto_reply` and the `mailbox_settings` permission category. Microsoft Graph's `/me/mailboxSettings` resource is documented as `Delegated (personal Microsoft account): Not supported`; every sub-path returns `ErrorAccessDenied` on outlook.com / hotmail.com / live.com mailboxes regardless of granted scopes. The project supports personal accounts only, so these four tools shipped in 1.7.0 are non-viable for every user. Verified directly against the live Graph API. The original brainstorm claim that `mailboxSettings` "works fully on outlook.com" was incorrect, and the v1.7.0 PR shipped without an integration smoke test that would have caught it.
- `MailboxSettings.Read` / `MailboxSettings.ReadWrite` Graph consent scopes (no longer requested).

### Fixed
- `outlook-mcp auth` device-code polling timeout raised from 300s to 900s to match Microsoft's 15-minute device-code TTL. The shorter default was failing real users whose sign-in flow takes more than 5 minutes (2FA, passkey prompts, browser delays).

### Unchanged from 1.7.0
- `outlook_list_inbox_overrides`, `outlook_set_inbox_override`, `outlook_delete_inbox_override` — Focused Inbox per-sender override CRUD via `/me/inferenceClassification/overrides`. Verified working on personal accounts.

### Migration from 1.7.0
- If you upgraded to 1.7.0 and re-authed: no action needed. The removed tools simply disappear; nothing else is affected.
- If you added `mailbox_settings` to `allow_categories` in `~/.outlook-mcp/config.json`: the server will refuse to load until you remove that string. The category no longer exists.

### Tool count
- 1.7.0: 61 tools, 14 categories.
- 1.7.1: **57 tools, 13 categories.**

## [1.7.0] — 2026-05-19

### Added — Mail Triage / Inference Overrides (3 tools)
- `outlook_list_inbox_overrides` — List Focused Inbox per-sender override rules.
- `outlook_set_inbox_override` — Upsert a per-sender override (`focused`/`other`). Case-insensitive sender matching; PATCH-if-exists, else POST. Gated by the existing `mail_triage` permission category.
- `outlook_delete_inbox_override` — Delete an override by ID.

These are the rule-level parallel of `outlook_reclassify_message`, which only fixes a single message in place.

### Added — Mailbox Settings (4 tools, new category)
- `outlook_get_timezone` / `outlook_set_timezone` — read/write `/me/mailboxSettings/timeZone`. Accepts IANA (`America/Los_Angeles`) or Windows (`Pacific Standard Time`) zone names; Graph validates.
- `outlook_get_auto_reply` / `outlook_set_auto_reply` — read/write the out-of-office / auto-reply configuration via `/me/mailboxSettings/automaticRepliesSetting`. Supports `disabled` / `always` / `scheduled` status, internal + external messages (with mirror-from-internal default), `none` / `contacts_only` / `all` external-audience scopes, and scheduled start/end datetimes.

All four are gated by a new `mailbox_settings` permission category.

### Auth scopes
- `MailboxSettings.Read` (read-only mode) and `MailboxSettings.ReadWrite` (full mode) added to the consent scope lists. **Existing users must re-run `outlook-mcp auth`** so the cached token picks up the new scopes — otherwise the four mailbox-settings tools will fail with an auth error.

### Notes
- `outlook_get_auto_reply` returns scheduled datetimes as UTC ISO 8601 (`YYYY-MM-DDTHH:MM:SSZ`) after translating common Windows zone names via a built-in CLDR mapping. When a zone can't be translated (rare, e.g. an obscure Windows display name not in the mapping), the value is emitted as `LOCAL:<datetime> <tz>` — explicitly non-UTC so callers don't silently treat local time as UTC.
- Tool count: **54 → 61**.

## [1.6.1] — 2026-05-17

Documentation-only release. Functionally identical to 1.6.0; no code changes. Cut to refresh the README that ships on the PyPI project page.

### Documentation
- Privacy & Security section corrected to describe the actual platform-by-platform token storage behavior after 1.6.0's libsecret-fallback fix (#8, #12). The previous unconditional "no tokens are written to disk in plaintext" claim is accurate on macOS Keychain, Windows DPAPI, and Linux with libsecret — but not on Linux without (e.g. `uv tool install` on Ubuntu, the failure mode reported in #7). The new text spells out each path and the one-time warning behavior.
- Tool Reference tables in README updated to document the parameters added in 1.6.0: `deferred_send_datetime` on create/update draft, `is_html` on update draft, and `include_deferred_send` on read message.

## [1.6.0] — 2026-05-17

### Added
- `outlook_create_draft` and `outlook_update_draft` accept a
  `deferred_send_datetime: str` parameter (ISO 8601). The value is set as
  the legacy MAPI `PR_DEFERRED_SEND_TIME` extended property (`SystemTime
  0x3FEF`); once the draft is sent (e.g. via `outlook_send_draft`),
  Exchange holds the message server-side until the given UTC instant.
  This is the same mechanism Outlook desktop uses for "Delay Delivery"
  and runs server-side, so the client doesn't need to be online at the
  scheduled time. On `update_draft`, passing an empty string clears the
  property. Inputs are validated and normalized to UTC ISO 8601 (`Z`
  form). ([#10])
- `outlook_update_draft` accepts `is_html: bool = False`. Required when
  overwriting a draft originally composed as HTML in the Outlook
  web/desktop UI — PATCHing such a draft with a Text body is rejected
  by the consumer-Outlook MAPI store with `ErrorAccessDenied` /
  `MapiSetProperties`. Default stays Text for back-compat. ([#10])
- `outlook_read_message` accepts `include_deferred_send: bool = False`.
  When `True`, surfaces the `PR_DEFERRED_SEND_TIME` extended property as
  `deferred_send_datetime` in the response (`null` if not set). Enables
  read-then-recreate workflows that preserve a draft's scheduled send
  time. ([#10])

### Fixed
- `outlook_download_attachment` was writing Microsoft Graph's
  base64-encoded `contentBytes` straight to disk, producing a base64
  text file instead of the actual binary. The tool now decodes the
  content before writing. Reported and fixed by @andylokandy. ([#9])

### Changed (breaking)
- `outlook_download_attachment` no longer accepts `save_path=None` and no
  longer returns `content_base64` in the response. `save_path` is now
  required and the tool always writes decoded bytes to that path,
  returning `{saved_to, name, size, content_type}`. The in-memory base64
  return path was the wrong paradigm for attachments — large binaries
  through the MCP message channel burn LLM context tokens — and the
  on-disk path was the buggy one (see Fixed). If you were relying on
  `content_base64`, switch to passing `save_path` and reading the file.
  ([#9])

### Changed
- The unencrypted token-cache fallback enabled in #8 now emits a
  one-time `logger.warning(...)` on the first credential build when
  PyGObject/libsecret isn't importable. macOS Keychain and Windows
  DPAPI still cache encrypted as before; only Linux installs without
  PyGObject (e.g. `uv tool install` on Ubuntu — issue #7) take the
  plaintext path, and now surface that fact to operators instead of
  silently writing tokens to disk.

[#9]: https://github.com/mpalermiti/outlook-mcp/pull/9
[#10]: https://github.com/mpalermiti/outlook-mcp/pull/10

## [1.5.2] — 2026-04-29

### Documentation / positioning
- Sharpened SKILL.md `description:` (drives the ClawHub search snippet) to lead with positioning ("MCP server, not a CLI wrapper") instead of a generic feature list. Highlights granular permissions, OS-keyring auth, batch optimization, and BYO Azure app — the differentiators against other Outlook skills in the registry.
- Added a "Who this is for / How it differs from other Outlook tools" section near the top of README.md so users browsing the listing can self-select in 30 seconds.

No code changes from 1.5.1.

## [1.5.1] — 2026-04-29

### Documentation
- Corrected stale `## Tools (51)` heading in `SKILL.md` to `## Tools (54)`. The frontmatter description was already correct, but the body heading rendered on the ClawHub skill page was missed when PRs #3 and #4 added new tools. Functionally identical to 1.5.0; no code changes.

## [1.5.0] — 2026-04-29

### Added
- `outlook_send_message`, `outlook_send_with_attachments`, `outlook_create_draft`, and `outlook_update_draft` accept a `reply_to: list[str]` parameter that maps 1:1 to Microsoft Graph's `message.replyTo`. On `update_draft`, `reply_to=[]` clears the field. ([#3])
- `outlook_attach_to_draft(draft_id, attachment_paths)` adds files to an existing draft, reusing the 3 MB inline / upload-session split from `outlook_send_with_attachments`. Returns the new attachment IDs for inline (small-file) attachments. ([#4])
- `outlook_remove_draft_attachment(draft_id, attachment_id)` deletes a single attachment from a draft. ([#4])
- Tool count: 52 → 54.

### Fixed
- **Tasks (`outlook_create_task` / `outlook_update_task` / `outlook_complete_task`):** request payloads were being built as raw `dict`s, but the Microsoft Graph SDK calls `.serialize()` on the payload — so every call failed with `'dict' object has no attribute 'serialize'`. All three tools now build typed `TodoTask` SDK models with `DateTimeTimeZone`, `ItemBody`, `Importance` enum, and `TaskStatus` enum. The `recurrence` dict input is converted to a typed `PatternedRecurrence` (with `RecurrencePattern` / `RecurrenceRange` and strict enum validation across `RecurrencePatternType`, `DayOfWeek`, `WeekIndex`, `RecurrenceRangeType`). Reported by @waynegault. ([#2], [#5])
- **Contacts (`outlook_list_contacts` / `outlook_search_contacts` / `outlook_get_contact` / `outlook_create_contact` / `outlook_update_contact`):** the consumer Outlook (Outlook.com / Hotmail) Graph endpoint does not expose the unified `phones` aggregate property — only `mobilePhone` (single string), `homePhones` (list), and `businessPhones` (list). Reads requested `phones` via `$select` and got 400; writes set `Contact.phones = [Phone()]` and would have hit the same 400 on consumer accounts. The whole module is migrated to the consumer Graph schema. Reported by @waynegault. ([#1], [#6])

### Changed (potentially breaking response shape)
- `outlook_get_contact` no longer returns `phones: [{number, type}]`. It now returns three separate fields: `mobile_phone: str`, `home_phones: list[str]`, `business_phones: list[str]`. The old field was always empty on consumer accounts, so any consumer parsing it was already getting `[]` — but if you have code reading `phones[0].number`, switch to `mobile_phone`.
- `outlook_list_contacts` and `outlook_search_contacts` summary responses keep their top-level `phone: str` field, but it is now correctly populated via mobile → first home → first business fallback (was previously empty on consumer accounts).
- Tool *inputs* are unchanged: `outlook_create_contact(phone=...)` and `outlook_update_contact(phone=...)` still take a single phone string, now stored as `mobilePhone`.

[#1]: https://github.com/mpalermiti/outlook-mcp/issues/1
[#2]: https://github.com/mpalermiti/outlook-mcp/issues/2
[#3]: https://github.com/mpalermiti/outlook-mcp/pull/3
[#4]: https://github.com/mpalermiti/outlook-mcp/pull/4
[#5]: https://github.com/mpalermiti/outlook-mcp/pull/5
[#6]: https://github.com/mpalermiti/outlook-mcp/pull/6

## [1.4.1] — 2026-04-22

### Fixed
- Both `outlook_list_folders(recursive=True)` and the folder name resolver were stopping at Microsoft Graph's default page size (10) when walking subfolders, silently dropping any user folder sorted after the 10th child. Fix paginates via `@odata.nextLink` and requests `$top=100` up front.

## [1.4.0] — 2026-04-21

### Added
- Recursive folder listing and subfolder name resolution.
