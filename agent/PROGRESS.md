# Progress — outlook-mcp-dev

Newest first. Every worker session prepends: date, task, key outcomes, files/test evidence, notes for the next worker.

## 2026-10-15 — contact-clear-phone (#120 follow-up)

- No new steering on PR #122 (every comment already addressed; no review yet of 8fdcea7).
- #120 re-checked: open, unassigned, no comments; PR search hit #67 unrelated.
- `outlook_update_contact` takes `clear_phones: list[str] | None` (`mobile`/`business`/`home`). `_clear_phones` sends `"mobilePhone": null` via `additional_data` (not `mobile_phone = None`, which leaks the Python name) and `[]` for the lists. `_check_clear_phones` refuses unknown types and clearing the type `phone` is written to in the same call, before any Graph call. Create unchanged (nothing to clear).
- TDD: `test_write_payloads_reach_the_wire.py::TestContacts` +2 (move mobile→business in one PATCH, `mobilePhone` is the only null; lists empty as `[]`), `test_contacts.py` +2 (refusals): red (TypeError) → green. Mutation: `contact.mobile_phone = None` instead of `additional_data` → the move test fails.
- Live: `tests/test_live_contacts_write.py::TestPhoneTypeRoundTrip` +2 (move, empty a list), NOT run. Only `--collect-only` was used on that file to confirm the two new tests collect (12 collected) — nothing executed. No `[?]`: Michael closed the contact live run with B on 2026-10-13. ASSUMED: Graph clears a list on `[]`.
- Surface 13,203 → 13,261 proxy (Python 3.13), ceiling 13,300, no raise; budget header records it. ~39 tokens left: the next new argument anywhere needs a deliberate raise.
- Docs: README + SKILL update_contact rows, update_contact docstrings, CHANGELOG Unreleased → Fixed (#120 entry).
- Suite: 1377 passed, 89 deselected; ruff check + format clean.

## 2026-10-14 — kiota-cap-watch, plus review fix on live-gap-98

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABb3Ho4A (`[construct review]` on live-gap-98: `_assert_moved` raised `_SeriesDidNotMoveError` for any three Mondays, ignoring time of day and which Mondays, so a different defect would XFAIL; and `-rxX` never prints dates for an XFAIL). Agreed. New offline `tests/test_live_gap_98_classifier.py` (8 tests, imports the live module as a module so its class is not re-collected) red on exactly the reviewer's three cases → green after `_assert_moved` requires the original three Mondays at `time_of_day`. Findings file's "how it stays out of the way" and `-rxX` sentence corrected; live-gap-98-run Notes updated. Small, done in-session.
- kiota-cap-watch: msgraph-sdk-python-core#1129 is a merged PR (2026-10-06); msgraph-core 1.5.2 (same day) contains it and requires kiota >=1.13; msgraph-sdk 1.64.0 still accepts core >=1.5.1. So the cap can lift only together with a direct `msgraph-core>=1.5.2` floor. Cap untouched. `agent/findings/kiota-cap-watch.md` has every source, date and label. Parked `[?]` kiota-cap-lift.
- Evidence: suite 1373 passed, 87 deselected; ruff check + format clean.
- Environment notes for the next worker: `uv` is at `~/.local/bin` (not on PATH) and its default cache is outside the sandbox — `export PATH=$HOME/.local/bin:$PATH UV_CACHE_DIR=$TMPDIR/uvcache` then `uv run --frozen ...` works offline. `gh` fails TLS inside the sandbox (OSStatus -26276); read-only `gh` calls need the sandbox off. A throwaway-venv install of newer kiota to run the /me wire test was declined by the permission gate; the fix mechanism is VERIFIED by reading only.

## 2026-10-14 — live-gap-98 (#98), plus review fix on pypi-environment

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABb2MBvw (`[construct review]` on actions-sha-pin: the pypi-environment `[?]` omitted that option A breaks the documented `gh workflow run publish.yml` retry, which dispatches on the default branch and would be refused by a tag-only environment; and that with no environment on PyPI's entry, `registry` could also mint a PyPI token). Both verified by reading RELEASING.md, publish.yml and the findings file; the `[?]` card's Why/A/Recommend/Notes and the findings file now say so. Small fix, done in-session.
- #98 re-checked: open, unassigned, no comments; PR hits #103/#83/#63 leave it out.
- `tests/test_live_calendar_write.py::TestMovingAMastersDateMovesItsOccurrences`: timed (weekly Mon 09:00 America/New_York ×3, moved to Tue with no `timezone`/`recurrence`) and all-day (×3, numbered only). Each: premise three Mondays via `/instances` → `updated` → three Tuesdays at the same time of day. Strict `xfail(raises=_SeriesDidNotMoveError)`: only #98's exact shape is absorbed; any other failure is red; a fix is XPASS(strict) → red → remove the marker. Behaviour unchanged (card said don't fix). No CHANGELOG entry: test-only.
- Offline verification: default suite 1365 passed, 87 deselected (+2), ruff check + format clean. Scratch (uncommitted) dry run of both bodies against fakes, 7/7: bug→`_SeriesDidNotMoveError`, Tuesdays→pass, Wednesdays→plain AssertionError, delete always runs, update carries no tz/recurrence/attendees. xfail(strict, raises=subclass) semantics checked on a scratch file.
- NOT run live. Parked `[?]` live-gap-98-run. Findings: `agent/findings/live-gap-98.md` (assertions, run command, result table).
- Note for next worker: ruff's N818 wants exception classes to end in `Error` — in tests too.

## 2026-10-14 — actions-sha-pin

- No new steering on PR #122 (no unaddressed comments).
- Pinned all 9 `uses:` in ci.yml/publish.yml to full SHAs, `# vX.Y.Z` comments, each the commit the current major tag resolves to (no bumps): checkout v4.4.0, setup-uv v5.4.2 (annotated `v5` dereferenced); upload-artifact v4.6.2 / download-artifact v4.3.0 added for the job split. Table in `agent/findings/actions-sha-pin.md`.
- `.github/dependabot.yml`: github-actions only, weekly.
- `publish.yml` split: `build` (contents: read, no token; lockstep, tests, lint, build, upload `dist`) → `pypi` (id-token only; download, setup-uv `enable-cache: false`, `uv publish`) → `registry` (id-token + contents; checkout, wait for PyPI, mcp-publisher, summary; version via `needs.build.outputs.version`). No `environment:` — parked as `[?]` pypi-environment.
- Guard `tests/test_workflows_are_pinned.py` (4 tests, indentation-parsed — no PyYAML in the env): red 3/4 → green; mutations (unpin one checkout; add `environment:`) each redden.
- Suite: 1365 passed, 85 deselected; ruff check + format clean. actionlint 1.7.12 clean (ephemeral `uvx --from actionlint-py`, sandbox off for the download). RELEASING.md + CHANGELOG Unreleased → Security.
- Not verifiable offline: a real release run. First release after merge is the test; `gh workflow run publish.yml` re-runs idempotently.

## 2026-10-13 — docs-truth-audit, plus closing contact-job-title-live

- No new steering on PR #122 (the only `[construct review]` comment was already addressed). Michael answered contact-job-title-live with **B** (ship without the live run): card checked off, live test left in place unrun.
- Audit: README, SKILL, ROADMAP, CHANGELOG Unreleased, CLAUDE.md against the running server's `list_tools()` / `list_prompts()`, `toolsets`, `permissions`, `errors`, `cli.py`. 20 checks, each with its result, in `agent/findings/docs-truth-audit.md`.
- Fixed: README "same 57 tools"; toolset example ~30 → 28 tools ("about half"; measured 6,536 of 13,224 proxy tokens); concise lists five → six (`outlook_read_messages`) in README and SKILL; error shape described as tool-error text, not `{code, message, action}` responses (README, SKILL); `allow_categories` table now names draft attachments (mail_drafts) and reclassify + Focused overrides (mail_triage); README roadmap drops inbox rules and points at ROADMAP.md; SKILL `list_inbox` "category" → "uncategorized only"; CLAUDE.md architecture gains cli/permissions/folder_resolver/calendar_resolver. CHANGELOG Unreleased → Documentation.
- Guard: `tests/test_docs_match_the_surface.py` (10 tests: README rows, SKILL bullets, per-section and total counts, toolset example + groups, concise list, category set, prompts). Red on 3 before the fixes, green after; mutation-checked (drop a SKILL bullet / bump a feature count / drop a group → each reddens).
- Suite: 1360 passed, 85 deselected; ruff check + format clean. Tool surface unchanged (no src edits).
- Left as findings footnotes (no queued work depends on them): README roadmap's remaining promises (MIME, cancel, Entra — Entra contradicts the README's own scope line); ROADMAP lists blocked rules CRUD under Near-term; SKILL "online meeting" vs #70.

## 2026-10-13 — hint-412-pinned-write (#96), plus review fix on contact `title`

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABbgaGew (`[construct review]`: `outlook_update_contact`'s client-visible description never said `title` is the name prefix). Fixed: added create's line to the update docstring; new `test_contacts.py::test_title_is_described_as_the_name_prefix_where_the_model_reads_it` (both write tools, via `Client(mcp).list_tools()`) red for update → green. Surface 13,201 → 13,224 proxy (ceiling 13,300, no raise; budget header records it). ~76 tokens of headroom left — contact-clear-phone may need a deliberate raise.
- #96 re-checked: open, unassigned, no comments; PR search hits unrelated.
- `errors._HINT_TABLE` 412 hint reworded — no longer "Nothing was modified"; says the change may already have been applied after a retried request, re-read and check before re-sending. Retry behaviour unchanged. Tests red→green: `test_error_wrapper.py::test_wrap_412_conflict_says_the_change_may_have_landed_and_to_re_read`, `test_error_text_reaches_client.py::test_pinned_write_conflict_hint_reaches_the_client`. Stale docstring in `test_calendar_write.py` updated. CHANGELOG Unreleased → Fixed.
- Findings: `agent/findings/hint-412-pinned-write.md` — kiota-http 1.12.3 retries PATCH on 429/503/504; disabling retry on the pinned PATCH loses every row of the trade-off once the hint is honest. No `[?]`.
- Suite: 1350 passed, 85 deselected; ruff check + format clean.

## 2026-10-13 — contact-phone-type (#120), with contact-title-on-update folded in

- Re-checked #120: open, unassigned, no comments; the PR search hit (#67) is an unrelated To Do PR. No PR comments on #122 to address.
- `outlook_create_contact` / `outlook_update_contact` take `phone_type` (`mobile` default, `business`, `home`); `_assign_phone` files `phone` under `mobilePhone`, `businessPhones=[phone]` or `homePhones=[phone]` and leaves the other two unset (never `None`). `_check_phone_type` refuses an unknown type and a non-mobile type with no phone, before any Graph call. Plain `str` with values in the docstring, matching `importance`/`show_as`.
- `outlook_update_contact` takes `title` (name prefix) — the queued #119 follow-up, same signature, so folded in per its card.
- TDD: 7 wire tests (each type × create/update, lands only in its own field; update `title`) + 3 validation tests failed first (TypeError), then green.
- Live: `TestPhoneTypeRoundTrip` (create business; update home + title, mobile untouched) added to `tests/test_live_contacts_write.py`, NOT run. The existing `[?]` contact-job-title-live now covers it — same command.
- Docs: README + SKILL contact rows, update docstring (replace-list rule), CHANGELOG Unreleased → Fixed.
- Suite: 1347 passed, 85 deselected; ruff check + format clean. Tool surface 13,100 → 13,201 proxy tokens (ceiling 13,300, no raise; budget header now records it). ~99 tokens of headroom left.
- Design note: `phone_type` picks one field and does not clear the others, so "move mobile → business" still leaves the mobile number. Clearing would need an explicit null via `additional_data` (the backing store emits a `None` under its Python name). Not requested; left as is.

## 2026-10-13 — contact-job-title (#119)

- Re-checked #119: open, unassigned, no comments, no linked PR (search hit #63 is an older, unrelated merged PR).
- `outlook_create_contact` and `outlook_update_contact` take `job_title` → Graph `jobTitle`; `outlook_get_contact` returns `job_title`. `title` kept as-is, described as the name prefix in the create docstring, README and SKILL.md. CHANGELOG under Unreleased → Fixed.
- TDD: wire tests (create sentinel lands under `"jobTitle"`, `"title"` absent when only a job is given; update carries `jobTitle`) and the detail-format tests failed first (TypeError / KeyError), then green.
- `tests/test_live_contacts_write.py` edited (prefix in `title`, job in `job_title`, display name not prefixed by the job; update persists `job_title`). NOT run — parked as `[?]` contact-job-title-live.
- Suite: 1336 passed, 83 deselected; ruff check + format clean. Tool surface 13,024 → 13,100 proxy tokens (ceiling 13,300, no raise; note the budget header's last recorded 13,245 is stale-high).
- Queued follow-up contact-title-on-update (update cannot fix a wrong prefix in place).
- Env note for next worker: `uv` is at `~/.local/bin/uv` (not on PATH) and needs `UV_CACHE_DIR=$TMPDIR/uv-cache` under the sandbox; `gh` needs the sandbox off (TLS).

## 2026-10-12 — initializer

- Mission created: close the gap between what outlook-mcp promises and what it does, offline only. Queue seeded with 7 tasks: two contact fixes Neil filed (#119, #120), the 412 hint (#96), a docs truth audit, Actions SHA pinning, the live test for #98, and a kiota cap check.
- Left out on purpose: #74 (Michael asked Neil to take it on 2026-09-22), #121 (a new tool that needs live verification), #97/#70/#72 (live-only; a `live-gap-*` card each once #98's pattern is proven).
- Guard verified on Neo before the first session: with `OUTLOOK_MCP_CONFIG_DIR` pointed at the sealed directory the server's own startup check finds no client id and no token, and even a credential handed the real client id is refused (`AuthenticationRequiredError`) because no account record exists. Offline suite under the guard: 1335 passed, 83 deselected; ruff check and format clean.
- Note for first worker: read MISSION.md's Boundaries in full before anything else. The suite command is the first line of MISSION.md.
