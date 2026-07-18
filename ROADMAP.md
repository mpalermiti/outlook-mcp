# Roadmap

Planned work for `outlook-graph-mcp`. Items here are committed-to direction; timing depends on demand. Community PRs welcome.

## Near-term

### Mail rules CRUD
Programmatic management of Outlook inbox rules via `/me/mailFolders/inbox/messageRules`. No other MCP I'm aware of exposes this.

**Shape:** `outlook_list_rules`, `outlook_create_rule`, `outlook_update_rule`, `outlook_delete_rule`. Rule definitions follow Graph's `messageRule` resource (conditions, actions, exceptions, sequence, isEnabled).

**Impact:** unlocks natural-language rule creation ("auto-move all Audi emails to TLDR") and programmatic inbox shaping. Strong demo surface.

---

## Performance & efficiency

Stack-ranked agent-optimization work from a 2026-07 review. The server already implements the standard Graph playbook — `$select` on every query, cursor pagination, `concise=True` payloads, delta queries, `$batch` bulk ops, dual (name + ID) identifiers. These items target the remaining gaps: static tool-context cost, connection reuse, and cross-resource latency for recurring agent loops. Ranked by impact-for-agents; each notes effort so they can be re-sorted.

1. **Config-gated toolsets** — 62 tool schemas load into client context every turn (large fixed token cost; tool-selection accuracy degrades past ~30–50 tools). Gate registration behind config so a client loads only the groups it uses (e.g. `mail,calendar,todo,digest`), or split into `outlook-core` + `outlook-admin` servers. Additive; no behavior change to enabled tools. Highest recurring-cost lever, and the only one fixable purely server-side. **Impact: high · Effort: med.**
2. **Parallelize `outlook_changes_since`** — `digest.py` awaits mail → events → contacts sequentially though they're independent. `asyncio.gather` gives ~2–3× lower latency on the most-used recurring-agent tool. **Impact: high · Effort: low.**
3. **Persistent Graph connection reuse** — `_get_graph_client` builds a new `GraphServiceClient` (and TLS pool) per tool call; the raw-httpx `$batch`/delta paths (`read_messages`, `fetch_delta_pages`) open ephemeral clients. Cache the client in the lifespan context (one per account, invalidate on `switch_account`) and share one long-lived httpx client for the raw paths. Compounds with #2. **Impact: high · Effort: med.**
4. **Tool annotations** — set `readOnlyHint` / `destructiveHint` (`ToolAnnotations`, supported by the SDK) on all tools so clients can auto-approve reads and gate destructive ops (delete mail, decline event). Currently unset on all 62. **Impact: med · Effort: low.**
5. **Folder name→ID memoization** — `resolve_folder_id` re-fetches the full `/me/mailFolders` tree (plus a BFS subfolder walk) on every display-name resolution; well-known names and Graph IDs already short-circuit with no network call. Add an in-memory, session-scoped name→ID cache (bust on lookup failure). Removes a full folder-tree fetch per custom-folder triage op. **Impact: med · Effort: low–med.**
6. **Throttling hardening on raw-httpx paths** — the SDK path already retries 429/503 via kiota's `RetryHandler`; `read_messages` and `fetch_delta_pages` don't retry the batch/delta envelope. Separately, `$batch` returns 200 even when sub-requests are throttled — `batch_triage` and `read_messages` currently record a 429'd sub-request as a permanent failure instead of retrying it with `Retry-After`. **Impact: med · Effort: low–med.**
7. **gzip on raw-httpx paths** — send `Accept-Encoding: gzip` (httpx auto-decompresses). Smaller payloads on large bodies. **Impact: low · Effort: trivial.**
8. **Structured output schemas** — tools return untyped `dict`; typed Pydantic returns would emit `outputSchema` / `structuredContent`. Gate on confirming the target client consumes it (otherwise pure cost). **Impact: low · Effort: med.**

**Suggested sequencing:** #2 + #3 + #4 as one test-first PR (all high-certainty, ~half a day), then decide #1's grouping from a measured tool-schema token count, then #5 / #6.

### Deferred (revisit if the hosting model changes)
- **Change-notification webhooks** as the delta trigger (eliminates fixed-interval polling) — needs a public HTTPS endpoint; N/A for stdio/local.
- **Code-execution-with-MCP** progressive tool disclosure (large token reduction) — needs a client that presents tools as a sandboxed code API.
- **FastMCP 2.x migration** for middleware/tags — its `ResponseCachingMiddleware` conflicts with the no-local-caching principle; config-gating (#1) delivers the tag benefit without the dependency swap.

---

## Ideas (not committed)

- **Shared / delegated mailboxes** — `/users/{id}/messages` path for delegated access
- **Calendar find-meeting-times** — `/me/findMeetingTimes` for availability queries
- **Category CRUD with colors** — first-class category management, not just assignment
- **Multi-account support** — `config.accounts` array already exists but is unused; wire up account-scoped tool calls

---

## Investigated and not viable

- **Mailbox settings (timezone, auto-reply, working hours, etc.)** — `/me/mailboxSettings/*` is documented as "Delegated (personal Microsoft account): Not supported" and verified to return `ErrorAccessDenied` on outlook.com mailboxes regardless of granted scopes. The resource is Exchange Online-only; consumer Outlook.com uses a different backend that Microsoft never bridged to Graph for this endpoint. Re-investigate if Microsoft publishes a consumer-account path for these settings.

---

## Done

- **1.11.1** — Hotfix: `outlook_download_attachment` corrupted binary attachments (#25). The tool double-decoded `contentBytes` — the msgraph SDK already base64-decodes it to raw bytes, so the extra `.decode("utf-8")` + `base64.b64decode()` raised `UnicodeDecodeError` on every non-UTF-8 file (.pdf/.docx/images). Now writes the SDK bytes verbatim. Regression from #9. Added a binary-fidelity regression test; corrected the pre-existing test whose base64-encoded mock masked the bug. No tool-count change.
- **1.11.0** — Bulk message read via `$batch` (1 new tool): `outlook_read_messages(message_ids, format, concise, include_deferred_send)`. Reads up to 20 messages by ID in one Graph `$batch` round-trip instead of N sequential calls. Per-message shape matches `outlook_read_message` byte-for-byte for the same `(format, concise, include_deferred_send)`. Returns `{messages, failures, requested, succeeded, failed}` — partial-failure tolerant (a 404 on one of 20 IDs surfaces in `failures[]` while the rest succeed). Input ordering preserved regardless of Graph's response ordering. Raises `ValueError` on input-validation errors (empty list, >20 IDs, malformed Graph ID) and `httpx.HTTPStatusError` on a transport-level 5xx (not swallowed into `failures[]`). Tool count: 61 → 62.
- **1.10.0** — Composed "since last call" digest: `outlook_changes_since(delta_tokens, fallback_window_hours)`. One MCP call wraps the three v1.9.0 delta tools (mail/events/contacts) and returns a structured payload — mail counts + `urgent_flagged[]` (high-importance OR flagged) + top-5 `by_sender{}`; events `new[] / modified[] / cancelled[]`; contacts counts; per-resource `delta_tokens` for caller-managed watermarks; `window` for the digest range. First call filters the bootstrap snapshot to `fallback_window_hours` (default 24) so the digest doesn't surface thousands of historical items. Each resource is independent — a stale token on one auto-resyncs that resource only, surfaced via `_meta.resync`. Internal pagination drains up to 5 pages (~1,000 items) per resource per call. Designed for recurring agent loops (morning brief, hourly inbox sweep). No new Graph endpoints — composes already-tested v1.9.0 delta endpoints. Tool count: 60 → 61.
- **1.9.1** — Tool docstring audit for AI agent clarity (no behavior change). Rewrote every `@mcp.tool()` docstring to a consistent shape: one-line action, contrastive pointer for ambiguous pairs, concrete syntax example. Designed to reduce wrong-tool selection. Signatures, params, defaults, return shapes byte-identical to 1.9.0.
- **1.9.0** — Delta queries (3 new tools): `outlook_list_inbox_delta`, `outlook_list_events_delta`, `outlook_list_contacts_delta`. Wraps Graph's `$delta` endpoints. First call returns a snapshot plus an opaque `delta_token`; subsequent calls (token passed back) return only added/updated/deleted items. Tombstones are `{id, is_deleted: True}` with no other fields — agents drop cached payloads cleanly. Stateless cursor (server doesn't persist tokens, matching the existing pagination `cursor` pattern). Per-call safety cap auto-follows `@odata.nextLink` up to `page_size * 4` items, then surfaces `has_more: True` plus the nextLink so the caller resumes. Massive token savings for recurring agent jobs polling a stable inbox/calendar/contacts. Tool count: 57 → 60.
- **1.8.0** — Agent-friendly shape, pure code. Two upgrades, no new tools: (a) `concise=True` opt-in on the five high-volume read tools (`outlook_list_inbox`, `outlook_read_message`, `outlook_search_mail`, `outlook_list_events`, `outlook_list_thread`) — drops bulky body / attendee / categories / quoted-text fields for ~10× smaller payloads on triage scans; (b) Graph SDK error wrapper that translates `ODataError`/`APIError` into a structured `GraphAPIError(code, message, action)` with recovery hints (re-auth on 401, ROADMAP pointer on 403/`ErrorAccessDenied`, re-list on 404, back-off on 429, retry on 503). Strict backward compat — defaults preserve the existing response shapes.
- **1.7.1** — Yanked the four mailbox-settings tools added in 1.7.0 (see CHANGELOG). Microsoft Graph's `/me/mailboxSettings` resource isn't supported for personal accounts, which is the project's only target. Auth timeout raised from 5 to 15 minutes.
- **1.7.0** — Focused Inbox per-sender override CRUD (upsert by sender, case-insensitive match): `outlook_list_inbox_overrides`, `outlook_set_inbox_override`, `outlook_delete_inbox_override`. Tool count: 54 → 57.
- **1.6.1** — Documentation-only refresh; corrected the Linux token-storage claim and updated tool-reference tables.
- **1.6.0** — Schedule-send / deferred-delivery via the `PR_DEFERRED_SEND_TIME` extended property: `outlook_create_draft` and `outlook_update_draft` accept `deferred_send_datetime`; `outlook_update_draft` accepts `is_html`; `outlook_read_message` accepts `include_deferred_send`. No new tools.
- **1.5.2** — Docs/positioning-only: sharpened SKILL.md description and added a "Who this is for / How it differs from other Outlook tools" section to README to compete more clearly with the other Outlook skills in the registry
- **1.5.1** — Docs-only: corrected stale `## Tools (51)` → `## Tools (54)` heading in SKILL.md (the frontmatter was already correct; ClawHub renders the body)
- **1.5.0** — `reply_to` parameter on send/draft tools (#3); `outlook_attach_to_draft` + `outlook_remove_draft_attachment` (#4); typed-model fix for `outlook_create_task` / `outlook_update_task` / `outlook_complete_task` plus dict→`PatternedRecurrence` conversion (#2, #5); consumer Graph phone-field migration for all contact tools — `mobilePhone` / `homePhones` / `businessPhones` instead of the unsupported `phones` aggregate (#1, #6). Tool count: 52 → 54.
- **1.4.1** — Paginate `childFolders` calls so parents with >10 subfolders return the full set
- **1.4.0** — Recursive folder tree listing (`recursive=true`) + subfolder name resolution
- **1.3.1** — Graph `/$batch` endpoint for `outlook_batch_triage` (10–20× perf)
- **1.3.0** — Transparent folder name resolution across all folder-accepting tools
- **1.2.0** — Focused Inbox classification filter
- **1.1.0** — Granular write permissions via `allow_categories`
- **1.0.0** — Initial 51-tool release
