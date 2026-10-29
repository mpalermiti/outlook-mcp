# Progress — outlook-mcp-dev

Newest first. Every worker session prepends: date, task, key outcomes, files/test evidence, notes for the next worker.

## 2026-10-28 — pypi-environment-settings (read-only re-check; Michael chose A, both halves done 2026-10-23)

- No steering on PR #171 (no conversation or inline comments).
- Recorded in `agent/findings/actions-sha-pin.md` § Re-check. VERIFIED via `gh api`: environments = only `pypi`; `custom_branch_policies: true`, one `{name: v*, type: tag}` policy, no required reviewers; rulesets `Protect main`, `Protect release tags` (unchanged: deletion/update/non_fast_forward, no bypass), `Release tag creation` (25761696: creation only, bypass RepositoryRole 5, always), all active. PyPI side recorded as Michael-confirmed. Supporting: deployments to `pypi` for v1.26.0–v1.30.1 all `success`, each version on PyPI seconds later.
- No code change. Suite: 2642 passed, 121 deselected; ruff check + format clean.
- Env notes: `gh` and the suite need the sandbox off here (gh: TLS cert error; suite: local port bind + worker subprocess). A plain `uv run` after the 1.30.x merge re-synced and dropped the `documents` extra (13 failures); `uv sync --all-extras --locked` then `uv run --offline pytest` is green.
- Pushed the local merge of `origin/main` (deadbfd) with this commit; it was committed but not pushed.

## 2026-10-23 — pr122-final-fixes (steering: Michael's maintainer review + checker review of c97ecde)

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABewh78g (`[construct review]`, two CHANGELOG findings), IC_kwDOSA_0Lc8AAAABewxvxQ (Michael, `[maintainer review]`, two README lines + "no further changes on this PR"). No inline comments.
- All four verified before editing: httpx 0.28.1 `_config.py` reads `SSL_CERT_FILE` before `certifi.where()`; mcp `lowlevel/server.py` defaults `middleware = [OpenTelemetryMiddleware()]`, `server/_otel.py` passes `ctx.meta` to `extract_trace_context` → `opentelemetry.propagate.extract`; `src/outlook_mcp/server.py` never touches `middleware` and runs `transport="stdio"`; `_format_contact_detail` returns `job_title`; pyproject has `msgraph-core>=1.5.2` and no kiota cap.
- CHANGELOG Unreleased → Security: certifi line adds the system-CA clause; aiohttp/opentelemetry line now gives stdio-only as the reason, not "never parses". README: `outlook_get_contact` row names `job_title`; #80 troubleshooting says 1.26.0 swaps the cap for the msgraph-core floor (Michael named the version).
- Docs only, no test change (no behaviour moved). Suite: 1382 passed, 89 deselected; ruff check + format clean.
- **Next worker:** see the Hold note at the top of QUEUE.md — do not push to #122; wait for it to merge, then start large-upload-checks-each-chunk on a fresh branch. #125 re-checked today: open, unassigned, no comments, no PR.
- Env note: `uv` needs `PATH=$HOME/.local/bin:$PATH UV_CACHE_DIR=$TMPDIR/uvcache` and `--offline` in the sandbox.

## 2026-10-22 — aiohttp-lock-refresh (Michael answered A, widened to certifi + opentelemetry-api)

- No new steering on PR #122 (newest comment is the worker's reply IC_kwDOSA_0Lc8AAAABeZZcBQ; no inline comments).
- Open Dependabot alerts on `main`: ten (aiohttp #2–4, #6–10; certifi #5 critical; opentelemetry-api #11). VERIFIED via `gh api`; both advisories read via `gh api /advisories/<ghsa>`.
- `uv lock --upgrade-package certifi --upgrade-package aiohttp --upgrade-package opentelemetry-api`: aiohttp 3.14.5, certifi 2026.10.19, opentelemetry-api/sdk 1.45.1, semconv 0.66b1, yarl 1.25.1 — the six Michael's dry run predicted.
- `pyproject.toml`: `certifi>=2026.10.19` with a comment naming GHSA-6q48-6c7w-9gw2 and the test.
- TDD: `tests/test_trust_store_excludes_compromised_roots.py` (httpx default context + requests' bundle, env overrides cleared; floor present in pyproject as text, since 3.10 has no tomllib). Red before the change (3 failed). First draft matched any "emSign" and stayed red on 2026.10.19: the advisory names only `emSign ECC Root CA - G3` and 2026.10.19 restores G1/C1/ECC C3, so the test now names G3. Mutation: `uv run --with certifi==2026.7.22` → 2 failed; lock → green.
- CHANGELOG Unreleased → Security: certifi line (users: upgrade), aiohttp + opentelemetry-api line (lock only).
- Findings: `agent/findings/aiohttp-alerts.md` § Done.
- live-gap-98-run not touched (behind this card on the board; still waiting on Neil).
- Env note: plain `uv sync` drops the dev extras (pytest-cov); use `uv sync --all-extras`, as CI does.
- Suite: 1382 passed, 89 deselected; ruff check + format clean.

## 2026-10-22 — review fix on 2ca7b62 (steering)

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABeZX1Cg (`[construct review]` of 2ca7b62, two findings). Both verified, both small, fixed in-session:
  1. Wrong PyPI name: the distribution is `outlook-graph-mcp` (`pyproject.toml` `name`, README, ci.yml `published-install`); `outlook-mcp` on PyPI is an unrelated project. Fixed in `agent/findings/aiohttp-alerts.md` (both `pip install` lines, plus a note on the name) and the aiohttp-alerts card. Remaining `outlook-mcp` hits on the card name the GitHub repo/server, correctly.
  2. Alert timing: Dependabot reads `main`, so a lock refresh on `agent/work` closes nothing until #122 merges. VERIFIED `gh api .../dependabot/alerts`: #1 multidict `fixed_at` 2026-10-08T02:11:44Z; #2–4 aiohttp open. aiohttp-lock-refresh card: A now says the alerts close once the PR merges, B says dismissing clears them today, Recommend states the trade. Findings gain a "When the alerts close" paragraph.
- live-gap-98-run skipped: #119 last comments unchanged (Neil 2026-10-20 20:45, Michael 21:27); #98 has 0 comments.
- Queue: no `- [ ]` task a worker can run; two `[?]` wait on Michael.
- Env note: `uv` is at `~/.local/bin/uv` (not on the sandbox PATH); `gh` needs the sandbox off (TLS verify fails inside it).
- Suite: 1379 passed, 89 deselected; ruff check + format clean.

## 2026-10-21 — aiohttp-alerts (found this session; no runnable task was on the board)

- No new steering on PR #122: newest comment is the worker's own reply to IC_kwDOSA_0Lc8AAAABeBmjEQ; no new Michael or `[construct review]` comments, no inline comments.
- live-gap-98-run skipped: #119 still has no result from Neil, #98 has no comments. Worker note on the card updated in place (not appended, to keep the line short).
- pypi-environment-settings is `[?]`, not touched.
- Queued and then ran aiohttp-alerts: the push showed three open Dependabot `aiohttp` alerts (GHSA-2g87-pp6c-29x5, GHSA-5wr8-gcqf-gqf3, GHSA-494j-54xm-rwp3; created 2026-10-20, all patched in 3.14.4). Read-only. VERIFIED: aiohttp never loads (sys.modules check after importing the whole package plus msgraph/kiota); only azure-core's async transport imports it, and `auth.py` uses sync credentials. Fresh installs get 3.14.5 (kiota-auth-azure: `aiohttp>=3.8.0`; PyPI latest 3.14.5). Only `uv.lock` holds 3.14.3.
- `uv lock --upgrade-package aiohttp` diff generated (aiohttp→3.14.5, yarl→1.25.1; `==3.14.4` moves the same two), then the lock was restored byte-for-byte (`git status` clean). Parked `[?]` aiohttp-lock-refresh (A refresh / B dismiss alerts / C leave).
- Findings: `agent/findings/aiohttp-alerts.md`.
- Env note: `uv lock` works in the sandbox with `UV_CACHE_DIR=$TMPDIR/uvcache` and pypi.org + files.pythonhosted.org allowed.
- Suite: 1379 passed, 89 deselected; ruff check + format clean.

## 2026-10-21 — review fix on 51d4226 (steering)

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABeBmjEQ (`[construct review]` of 51d4226: CHANGELOG Unreleased still said RELEASING.md's appendix gives "both settings"; it now has three numbered steps). Verified against RELEASING.md appendix steps 1–3; CHANGELOG now says "the three settings and their order", and the over-long line in that paragraph is rewrapped. Grep for "both settings"/"two settings" in CHANGELOG/RELEASING/README: none left (the dated Done entry in `agent/findings/actions-sha-pin.md` is history, corrected by its own § Correction).
- live-gap-98-run skipped: #119 still has no result from Neil (worker note on the card).
- Queue: no `- [ ]` task a worker can run.
- Suite: 1379 passed, 89 deselected; ruff check + format clean.

## 2026-10-20 — review fixes on pypi-environment (steering)

- Addressed PR comments: IC_kwDOSA_0Lc8AAAABeBdKIw (`[construct review]` of cfbb1d2, three findings). All verified, all small, fixed in-session:
  1. PyPI has no publisher edit (warehouse views: add/delete/constrain only; `_get_publisher_for_environment` falls back to a blank-env publisher). RELEASING.md step now: constrain link from PyPI's email, or delete + re-add; never a second entry beside the blank one. Settings card corrected.
  2. The `v*` environment rule checks only the tag name; "Protect release tags" ruleset has no `creation` rule (VERIFIED `gh api .../rulesets/24471530`). RELEASING.md gains a step: separate tag ruleset with only Restrict creations + Repository admin bypass (or required reviewer on `pypi`). CHANGELOG entry no longer overstates. Settings card now three settings, options A/B/C.
  3. `test_only_the_pypi_job_runs_in_the_pypi_environment` missed quoted values. New `_environment` helper (any `environment:` key counts; quotes, comment, mapping `name:` handled) + `test_the_environment_reader_sees_every_spelling`: red (2 failed, helper missing) → green. Mutation: `"pypi"`, `'pypi'`, `${{ 'pypi' }}` on `registry` each → red.
- live-gap-98-run skipped: #119 still has no result from Neil.
- Findings: `agent/findings/actions-sha-pin.md` § Correction (2026-10-20, from review of cfbb1d2).
- Queue: no `- [ ]` task left that a worker can run (live-gap-98-run waits on Neil; pypi-environment-settings is `[?]`).

## 2026-10-20 — pypi-environment (Michael answered A)

- No new steering on PR #122 (no new Michael or `[construct review]` comments).
- live-gap-98-run (first `[ ]`) skipped again: #119 still has no result from Neil. Worker note added on the card.
- `publish.yml`: `environment: pypi # release tags only; ...` on the `pypi` job only; header and `workflow_dispatch` comments explain it; registry step's retry comment names `--ref vX.Y.Z`.
- RELEASING.md: three-jobs paragraph mentions the environment; retry is `gh workflow run publish.yml --ref vX.Y.Z` with a line on why and `gh run rerun <id> --failed`; appendix table Environment → `pypi` plus a "The `pypi` environment" section (GitHub tag rule `v*` first, PyPI field after one release).
- TDD: `tests/test_workflows_are_pinned.py` test 4 flipped to `test_only_the_pypi_job_runs_in_the_pypi_environment`; new `test_the_documented_publish_retry_names_the_release_ref` (RELEASING.md + publish.yml). Red 2 → green; mutation: env added to `registry` → red, bare retry restored in RELEASING.md → red.
- VERIFIED: PyPI warehouse `_check_environment` returns True for a blank publisher environment (read via `gh api repos/pypi/warehouse/contents/...`), so the workflow-first order is safe. Repo has 0 GitHub environments today.
- Not run: actionlint (binary download blocked in sandbox). PyYAML parse (ephemeral `uv run --no-project --with pyyaml`) shows `environment` only on `pypi`.
- CHANGELOG Unreleased → Security. Findings `agent/findings/actions-sha-pin.md` § Done.
- Parked `[?]` pypi-environment-settings: Michael's two settings (GitHub env before next release, PyPI field after it).
- Suite: 1378 passed, 89 deselected; ruff check + format clean.

## 2026-10-20 — kiota-cap-lift (Michael answered A)

- No new steering on PR #122 (all Michael/`[construct review]` comments already addressed).
- live-gap-98-run (first `[ ]`) skipped: Michael's answer is B, wait for Neil on #119; #119 has no result yet. Worker note added on the card.
- `pyproject.toml`: four `microsoft-kiota-*<1.13` caps and their comment replaced by `msgraph-core>=1.5.2` with a comment naming #80, msgraph-sdk-python-core#1129 and why a floor (msgraph-sdk 1.64.0 still accepts core 1.5.1). `uv lock` → kiota-abstractions/http/auth-azure 1.14.3, core 1.5.2; `uv lock --upgrade-package microsoft-kiota-serialization-json` → 1.14.3 so the lock matches what a fresh install gets. msgraph-sdk stays 1.62.0 in the lock (no unrelated bumps).
- Also: CHANGELOG Unreleased → Changed; README troubleshooting "doesn't yet handle" → "1.5.1 doesn't handle"; wire-test failure text points at the floor; `ci.yml` fresh/published version reports print msgraph-core.
- Evidence: local suite 1377 passed / ruff clean on the new lock. `uv run --no-sync --with msgraph-core==1.5.1 pytest tests/test_me_rewrite_reaches_the_wire.py` → 2 failed ("rewrite did not fire") with kiota 1.14.3, so the floor is load-bearing. CI run 39553812140 (push, 85b8adb): all jobs green; fresh-install log: msgraph-sdk 1.64.0, msgraph-core 1.5.2, kiota-http 1.14.3, `/me` → `https://graph.microsoft.com/v1.0/me`.
- Env note for next worker: `uv` is at `~/.local/bin/uv` (not on PATH) and its cache is outside the sandbox; `.venv/bin/pytest` / `.venv/bin/ruff` run the suite inside it. `gh` needs the sandbox off (TLS).

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
- Guard verified on the host before the first session: with `OUTLOOK_MCP_CONFIG_DIR` pointed at the sealed directory the server's own startup check finds no client id and no token, and even a credential handed the real client id is refused (`AuthenticationRequiredError`) because no account record exists. Offline suite under the guard: 1335 passed, 83 deselected; ruff check and format clean.
- Note for first worker: read MISSION.md's Boundaries in full before anything else. The suite command is the first line of MISSION.md.
