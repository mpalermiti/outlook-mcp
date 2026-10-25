# Mission: outlook-mcp-dev

**Type:** dev · **Started:** 2026-10-12 · **Owner:** Michael

Suite: `uv run pytest && uv run ruff check src/ tests/ scripts/ && uv run ruff format --check src/ tests/ scripts/`

## Why
outlook-mcp is used: a few thousand PyPI installs a month, a ClawHub listing, and an outside contributor (Neil) who files precise, live-verified issues faster than Michael can turn them into fixes. The gap that costs users is between what the server and its docs promise and what it actually does. Most of that gap can be closed and proven offline, with the repo's own payload guards, without ever touching a mailbox.

## Goal
Close the gap between what outlook-mcp promises and what it does, one reviewable change at a time:
1. Fix the user-reported defects that offline tests can prove (the wire payload, the validation, the error text).
2. Keep the docs true to the code: README, SKILL.md, ROADMAP, CHANGELOG, tool counts, parameter names.
3. Harden the release pipeline where it is mechanical.
4. For defects only a live mailbox can show, produce the live test that would catch them, for Michael to run.

## Done when
Open-ended upkeep mission: done is an empty queue. Michael adds issues to the queue as they arrive.

## Boundaries

**The mailbox is off-limits, absolutely.** Every session runs with `OUTLOOK_MCP_CONFIG_DIR` pointed at an empty, read-only directory so the server has nothing to sign in as. Never unset, override or work around that. Never:
- run any credentialed test tier: no `-m live`, `-m integration`, `-m live_write`, no `-m` expression that selects them, never set `OUTLOOK_MCP_LIVE_WRITE`;
- run `scripts/preflight.py`, or `outlook-mcp` / `uv run outlook-mcp` in any form (server, `auth`, `status`, `logout`);
- read, copy or list `~/.outlook-mcp`, or touch the Keychain (`security`, keyring, msal cache files).
If a task can only be verified against a real mailbox, write the live test, leave it unrun, and park a `[?]` asking Michael to run it.

**Nothing ships.** Never bump a version (`pyproject.toml`, `server.json`, CHANGELOG version headings, ROADMAP release lines), tag, create a GitHub release, or run `uv publish`, `mcp-publisher` or `clawhub`. Changelog entries go under `## [Unreleased]`. Never push `main`; all work is on `agent/work`.

**Other people work in this repo.** Before starting any card that names an issue, run `gh issue view <n> --json assignees,comments,closed` and `gh pr list --state all --search "<n>"`. If the issue is closed or assigned, has a PR, or has any comment newer than the date on the card (someone may have said they are taking it), do not start it: check the card off with a note saying why, and take the next one. Never comment on issues, never comment on, review or approve contributor PRs. The rolling PR is the only place you write on GitHub.

**The repo's guards are not yours to loosen.** Never weaken `test_no_dead_parameters.py`, `test_no_dead_modules.py`, `test_sdk_fields_exist.py`, `test_write_payloads_reach_the_wire.py`, `test_error_text_reaches_client.py` or `test_tool_surface_budget.py`. A new parameter that needs a higher `TOOL_SURFACE_CEILING` is a deliberate raise: measure it, record the measured value in the test header the way earlier raises are recorded, and say so in the For Michael note. No dependency bumps, no new dependencies, no cap changes in `pyproject.toml`.

**This repo is public.** Everything under `agent/` and on the rolling PR is visible to anyone. Write nothing private: no local paths outside the repo, no account details, no unreleased security findings.

The live checkout at `~/ClaudeCode/outlook-mcp` is off-limits; work only in this checkout.

## Verification rules
The suite above must be green before and after every session. Every behaviour change starts with a test that fails for the right reason. A write-path change is proven by the serialized payload (`test_write_payloads_reach_the_wire.py` style), not by the model object. Docs changes are checked against `src/` by grep or a test, never from memory.
