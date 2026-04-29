# Roadmap

Planned work for `outlook-graph-mcp`. Items here are committed-to direction; timing depends on demand. Community PRs welcome.

## Near-term

### Delta queries for `list_inbox` / folder scans
Swap repeated `/me/mailFolders/inbox/messages` polls for `/me/mailFolders/inbox/messages/delta`. Returns only what changed since the last delta token. Makes recurring agent workflows (morning briefs, daily junk scans, weekly digests) near-free instead of re-fetching 25 messages on every run.

**Shape:** new `outlook_list_inbox_delta` tool (or `delta=true` flag on `list_inbox`) that persists the delta token per-folder in the config dir and returns only new/changed/removed messages on subsequent calls.

**Impact:** 10–100× fewer tokens transferred per poll for stable inboxes. Meaningful for agent cost + latency.

---

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

## Done

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
