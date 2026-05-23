# Roadmap

Planned work for `outlook-graph-mcp`. Items here are committed-to direction; timing depends on demand. Community PRs welcome.

## Near-term

### Mail rules CRUD
Programmatic management of Outlook inbox rules via `/me/mailFolders/inbox/messageRules`. No other MCP I'm aware of exposes this.

**Shape:** `outlook_list_rules`, `outlook_create_rule`, `outlook_update_rule`, `outlook_delete_rule`. Rule definitions follow Graph's `messageRule` resource (conditions, actions, exceptions, sequence, isEnabled).

**Impact:** unlocks natural-language rule creation ("auto-move all Audi emails to TLDR") and programmatic inbox shaping. Strong demo surface.

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
