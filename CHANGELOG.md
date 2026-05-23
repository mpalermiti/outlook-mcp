# Changelog

All notable changes to outlook-graph-mcp are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
