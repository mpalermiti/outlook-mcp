# docs-truth-audit — 2026-10-13

README.md, SKILL.md, ROADMAP.md, CHANGELOG.md (Unreleased) and CLAUDE.md checked against
`src/`. The registered surface was read from the running server (`Client(mcp).list_tools()`,
`list_prompts()`), not from memory: 68 tools, 3 prompts.

## Checks and results

| # | Check | Source of truth | Result |
|---|-------|-----------------|--------|
| 1 | README Tool Reference rows = registered tools, no duplicates | `list_tools()` | OK (68/68). **Pinned.** |
| 2 | SKILL.md tool bullets = registered tools, no duplicates | `list_tools()` | OK (68/68). **Pinned.** |
| 3 | README "Features" per-category counts = rows in the matching table | README tables (checked against #1) | OK, all 13. **Pinned.** |
| 4 | "68 tools … across 13 categories" | `list_tools()`, feature list | OK. **Pinned.** |
| 5 | Every stated total (`N tools`, `Tools (N)`, `all N`, `instead of N`) | `list_tools()` | **Fixed:** README "the same 57 tools" (1.8.0-era count, read as current). **Pinned.** |
| 6 | Toolset example `mail,calendar,digest,delta` tool count | `toolsets.TOOL_GROUPS` | **Fixed:** "~30 tools … ~55% fewer" → 28 tools, "about half" (measured 6,536 of 13,224 chars/4 proxy tokens, 51% fewer). **Pinned** (count; the percentage is prose). |
| 7 | README "Groups:" list = selectable groups | `toolsets.TOOL_GROUPS` minus `ALWAYS_ON_GROUPS` | OK (10). **Pinned.** |
| 8 | Concise-mode tool lists (README, SKILL) = tools with a `concise` parameter | tool schemas | **Fixed:** both named five; six take it (`outlook_read_messages` missing). **Pinned.** |
| 9 | Error shape claim | `errors.OutlookMCPError.__str__`, SDK `ToolError` path | **Fixed:** README and SKILL said `{code, message, action}` *responses*; the client receives a tool error whose text is message + hint (`code` is not in it). Hint list (401/403/404/429/503) matches `_HINT_TABLE`. Not pinned (prose). |
| 10 | `allow_categories` table rows and "only the seven names" = accepted set | `permissions.VALID_CATEGORIES` | OK (7). **Pinned.** |
| 11 | Which tools each category gates (table column) | `check_permission(CATEGORY_*, "outlook_…")` call sites | **Fixed:** `mail_drafts` omitted attach-to/remove-from draft; `mail_triage` omitted reclassify and Focused Inbox override set/delete. Not pinned — the column uses short prose names, not tool names. |
| 12 | Prompts: names and arguments (README table, SKILL list) | `list_prompts()` | OK. **Pinned.** |
| 13 | Every backticked lowercase identifier in a README row / SKILL bullet is a parameter of that tool | tool schemas | OK — every non-parameter hit is an enum value (`busy`, `home`), a return field (`type`, `parent_id`, `original_start_time_zone`, all present in formatters) or a config key (`attachments_dir`). No stale parameter names. Ad hoc, not pinned (too many legitimate non-parameter tokens). |
| 14 | SKILL `outlook_list_inbox` filter list | `list_inbox` params | **Fixed:** "category" → "uncategorized only" (the only category filter is `uncategorized_only`). |
| 15 | CLI commands named in README (`serve`, `auth`, `status`, `logout`) | `cli.py` dispatch | OK. |
| 16 | README roadmap vs ROADMAP.md | ROADMAP "Investigated and not viable" | **Fixed:** removed "Inbox Rules" (403 on personal accounts, recorded in ROADMAP); added a pointer to ROADMAP.md. Remaining items: see open question 1. |
| 17 | CLAUDE.md architecture list = modules in `src/outlook_mcp/` | `ls` | **Fixed:** `cli.py`, `permissions.py`, `folder_resolver.py`, `calendar_resolver.py` were missing. Every listed file exists. |
| 18 | CLAUDE.md "68 docstrings", "62-tool surface ~8,644 o200k" | `list_tools()`, ROADMAP | OK (68 current; 62 is the dated measurement, labelled as such). |
| 19 | CHANGELOG Unreleased entries vs code | tool schemas, `errors._HINT_TABLE` | OK: `job_title`, `phone_type`, `title` on update, the 412 hint all present. |
| 20 | ROADMAP "Done" release lines | — | Historical; tool counts there are per-release and correct as history. Not edited (no release lines touched, per MISSION). |

Guard: `tests/test_docs_match_the_surface.py` (10 tests). Red first on #5, #6, #8; green after
the fixes. Mutation check: deleting one SKILL bullet, changing a feature count, and dropping a
group from the README each reddened the matching test.

## Open questions for Michael (no queued work depends on them — footnotes, not cards)

1. **What the README roadmap promises.** After removing inbox rules it still lists raw MIME
   export / internet message headers, calendar cancel-with-notification, and Enterprise
   (Entra ID). None of the three is in ROADMAP.md. Entra contradicts the README's own
   "Who this is for" line, which calls work/school accounts out of scope. And `outlook_delete_event`
   on a meeting you organise already notifies attendees (the README's `calendar_write` row says
   so), so "cancel" would only add a cancellation message. Options: move what is still meant
   into ROADMAP.md and leave the README with the pointer only, or keep the list. Not decided
   here — it is a promise, not a fact.
2. **ROADMAP "Near-term → Mail rules CRUD"** sits under Near-term while its own status line and
   the "not viable" section say it is blocked on personal accounts. Moving it would be an
   editorial call on committed direction; left as is.
3. **`is_online`** — SKILL's `outlook_create_event` line says "online meeting", the README says
   `is_online` has no effect on personal accounts, and ROADMAP 1.23.0 records it honoured on two
   contributors' mailboxes. That is issue #70 (live-only, deliberately not queued); not touched.
