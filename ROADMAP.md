# Roadmap

Planned work for `outlook-graph-mcp`. Items here are committed-to direction; timing depends on demand. Community PRs welcome.

## Near-term

### Mail rules CRUD
Programmatic management of Outlook inbox rules via `/me/mailFolders/inbox/messageRules`. No other MCP I'm aware of exposes this.

**Shape:** `outlook_list_rules`, `outlook_create_rule`, `outlook_update_rule`, `outlook_delete_rule`. Rule definitions follow Graph's `messageRule` resource (conditions, actions, exceptions, sequence, isEnabled).

**Impact:** unlocks natural-language rule creation ("auto-move all Audi emails to TLDR") and programmatic inbox shaping. Strong demo surface.

---

## Performance & efficiency

Stack-ranked agent-optimization work from a 2026-07 review, re-validated against a mid-2026 market scan (MCP spec evolution, competing email/calendar MCPs, Microsoft's first-party moves). The scan's verdict: every ecosystem signal points *into* this perf/cost work, not away from it — depth (delta, `$batch`, connection reuse, throttling, gating) is the durable differentiator, since the leanest competitors ship none of it. The server already implements the standard Graph playbook (`$select`, cursor pagination, `concise=True`, delta queries, `$batch`, dual name+ID identifiers); the items below target the remaining gaps.

**Strategic context (de-risks the whole program):** Microsoft's first-party Outlook MCP surfaces (Work IQ, GA 2026-06-16; Agent 365 "Outlook Mail" / "Outlook Calendar" servers) are gated to M365-Copilot-licensed / Frontier *enterprise* tenants over org data — **no consumer Outlook.com (personal MSA) support**. The personal-account niche this server owns is therefore uncontested by Microsoft today. **Watch item / kill-switch:** a future Microsoft first-party MCP for personal accounts is the one event that would materially threaten this project — monitor.

**Measured baseline (2026-07-18):** the 62 tool schemas serialize to **~8,644 tokens/turn** (o200k proxy; Claude ±~10%), avg 139 tok/tool — real but far from the debunked "50–120K context tax." Biggest domains: mail 25%, drafts 12%, calendar 10%. This settles the config-gating design (Tier 0 #4).

### Tier 0 — Neo personal-account perf/cost (fixed priority) — ✅ shipped in v1.12.0

Latency + token/API cost for the persistent single-agent recurring mail+calendar loop. All externally re-validated by the scan; none displaced by it. Items #1–#5 shipped in v1.12.0; #6–#7 remain as cheap follow-ups.

1. **Parallelize `outlook_changes_since`** — `digest.py` awaits mail → events → contacts sequentially though they're independent. `asyncio.gather` → ~2–3× lower latency on the most-used recurring tool. **Impact: high · Effort: low.**
2. **Persistent Graph connection reuse** — `_get_graph_client` builds a new `GraphServiceClient` (+ TLS pool) per call; raw-httpx `$batch`/delta paths (`read_messages`, `fetch_delta_pages`) open ephemeral clients. Cache the client in the lifespan context (one per account, invalidate on `switch_account`) and share one long-lived httpx client for the raw paths. Compounds with #1. **Impact: high · Effort: med.**
3. **Tool annotations** — set `readOnlyHint` / `destructiveHint` (`ToolAnnotations`, SDK-supported) on all 62 so clients auto-approve reads and gate destructive ops. Aligns with the 2025-11-25 spec. **Impact: med · Effort: low.**
4. **Config-gated toolsets** — highest recurring-cost lever and the only one fixable purely server-side; validated by Microsoft's own Work-IQ 10-verb design. **Decision (from the measurement): a flexible toolset selector, NOT a two-package `core`/`admin` split** — the admin/override/batch group is only ~7% of tokens, so a binary split barely helps; real reduction comes from dropping whole *domains* a client doesn't use. Gate registration behind config (e.g. `OUTLOOK_MCP_TOOLSETS=mail,calendar,digest,delta`). For Neo's mail+calendar slice that's ~4,155 tok — **~52% off every turn.** Additive; no behavior change to enabled tools. **Impact: high · Effort: med.**
5. **Throttling hardening on raw-httpx paths** *(promoted from "med" — the scan reframes it as a correctness bug, not just perf)*. The SDK path retries 429/503 via kiota's `RetryHandler`, but `read_messages` / `fetch_delta_pages` don't retry the batch/delta envelope, and `$batch` returns 200 even when sub-requests are throttled — currently recorded as a *permanent failure* instead of retried with `Retry-After`. Graph enforces a global 130,000 req/10s ceiling on top of per-mailbox limits; direct (non-SDK) callers must implement `Retry-After` + backoff themselves. **Impact: med–high · Effort: low–med.**
6. **Folder name→ID memoization** — `resolve_folder_id` re-fetches the full `/me/mailFolders` tree (+ BFS subfolder walk) per display-name resolution; well-known names and Graph IDs already short-circuit. Add a session-scoped name→ID cache (bust on lookup failure). **Impact: med · Effort: low–med.**
7. **gzip on raw-httpx paths** — send `Accept-Encoding: gzip` (httpx auto-decompresses). **Impact: low · Effort: trivial.**

**Sequencing:** #1 + #2 + #3 as one test-first PR (high-certainty, ~half a day) → #5 (correctness) → #4 (selector; grouping now settled by the measurement) → #6 / #7.

### Tier 1 — public multi-agent registry audience (additive; zero impact on Neo's stdio loop; gated on client support)

For the population installing this from the MCP registry, not for Neo. stdio stays the default and unchanged.

- **Stateless Streamable-HTTP deployment** — the 2026-07-28 spec RC removed `Mcp-Session-Id`, so a remote server can scale behind a plain round-robin LB with no session store. Optional remote transport alongside stdio.
- **OAuth discovery hardening** — OIDC Discovery, RFC 9728 Protected-Resource-Metadata, incremental scope consent (SEP-835), Client ID Metadata Documents (SEP-991). Load-bearing only when exposed as a remote OAuth resource.
- **`tools/list` caching** — SEP-2549 `ttlMs` / `cacheScope` once SDK + clients honor it.
- **Cross-provider / multi-account** — a competing server already unifies M365 + Outlook.com + Google in one MCP. The unused `config.accounts` array is the hook. Real but new; secondary to Tier 0.

**Caveat:** several Tier-1 surfaces are release-candidate / draft spec (statelessness RC, SEP-2549) — don't build against them until Claude / Cursor / OpenClaw actually honor them.

### Tier 2 — deferred (hosting- or client-gated; revisit when the precondition lands)
- **Change-notification webhooks** as the delta trigger (near-real-time, avoids polling/throttling) — needs a public HTTPS endpoint; N/A for stdio/local. Rich-mode notifications carry the changed object inline (a token lever) but need an encryption cert. Recommended end-state is delta-tokens **+** webhooks.
- **Code-execution-with-MCP / progressive tool disclosure** — needs a client that presents tools as a sandboxed code API; still theoretical for the single-personal-agent case.
- **Structured output schemas** — typed Pydantic returns → `outputSchema` / `structuredContent`. Gate on confirming the target client consumes it (else pure cost).
- **FastMCP 2.x migration** for middleware/tags — its `ResponseCachingMiddleware` conflicts with the no-local-caching principle; the Tier-0 selector delivers the tag benefit without the dependency swap.

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

- **1.12.0** — Tier-0 perf/cost pass (no new tools, no breaking changes). Concurrent `outlook_changes_since` digest (`asyncio.gather`, ~2–3× lower latency); persistent Graph client reuse (cached in lifespan, rebuilt on credential change); throttling hardening on the raw-httpx delta/`$batch` paths (`throttle.py` — `Retry-After` retry on the delta GET, and per-sub-request retry so a throttled `$batch` sub-request isn't recorded as a permanent failure); tool annotations (`readOnlyHint`/`destructiveHint` on all 62); and config-gated toolsets (`OUTLOOK_MCP_TOOLSETS` — a flexible selector that drops Neo's mail+calendar surface 62 → ~30 tools, ~52%/turn). Grouping settled by a measured ~8,644-tok baseline. Roadmap items Tier-0 #1–#5.
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
