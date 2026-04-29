# Changelog

All notable changes to outlook-graph-mcp are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
