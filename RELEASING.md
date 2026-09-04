# Releasing outlook-mcp

Checklist for cutting a new release.

## 1. Smoke-test against the live Graph API

```bash
uv run python scripts/preflight.py
```

Hits every Graph endpoint family the tools depend on with the locally-cached token. Flags any endpoint that returns 403 or 501 — the "not supported for this account type" signal that mocked unit tests can't catch.

Read-only. No writes, no sends, no mailbox state changes.

If the script reports failures, do not tag. Either fix the affected tools or remove them from the release. v1.7.0 shipped four tools backed by `/me/mailboxSettings/*` that Microsoft Graph does not support on personal accounts; v1.7.1 yanked them. This script would have caught it in 30 seconds.

When adding a new tool that hits a Graph endpoint family not yet covered, add a row to `ENDPOINTS` in `scripts/preflight.py`.

## 1b. Live query-shape tests

```bash
uv run pytest -m live -v
```

Preflight answers "does this endpoint exist and respond?" — it treats a 400 as a non-blocking SKIP. This tier answers the different question: **does Graph accept and correctly evaluate the queries we actually build?**

That gap shipped three bugs in 1.12.0, all under a fully green mock suite:

- `$orderby` + any non-date `$filter` → `400 InefficientFilter`, breaking `from_address` and `classification` on every call (#31)
- `list_thread` hit the same rule and 400'd unconditionally — it had never worked in a released version
- `sanitize_kql` stripped `:`, so every documented KQL property restriction returned `200` with **zero results** (#30)

The second failure mode is the dangerous one: a silent 200 with wrong data. Mocks assert what we *send*; only a live call sees what Graph *does*. These tests assert on returned data, not just absence of an exception.

Read-only, and mailbox-independent — they harvest their own fixtures and skip cleanly when the mailbox lacks the needed data. Auto-skipped without a cached token.

If you change how any `$filter`, `$orderby` or `$search` string is built, run this before tagging.

## 1c. Integration smoke tests

```bash
uv run pytest -m integration -v
```

Response-shape checks for each tool family. Read-only.

> These skipped silently for their entire existence — the fixture called a non-existent `AuthManager.login()`, and the `except Exception` swallowed the `AttributeError`. Fixed in 1.13.0. If you see `skipped` here, confirm it's really a missing token and not a broken fixture.

## 2. Tests + lint

```bash
uv run pytest --tb=no -q
uv run ruff check src/ tests/
```

The default run is the offline unit suite only — `addopts` deselects the `integration` and `live` markers, so this needs no network or token. Expect `N passed, 16 deselected` and zero failures.

## 3. Version bump

Update in lockstep:

- `pyproject.toml` — `version = "X.Y.Z"`
- `server.json` — both `version` fields + `description` (tool count if it changed)
- `CHANGELOG.md` — new `## [X.Y.Z] — YYYY-MM-DD` entry
- `SKILL.md` — `## Tools (N)` heading + frontmatter `description` if count changed
- `README.md` — counts and tables if they changed
- `ROADMAP.md` — move shipped items from Near-term to Done
- `CLAUDE.md` — tools listing if you added/removed a module

## 4. PR + merge

```bash
gh pr create --title "vX.Y.Z: <summary>" --body "<changelog excerpt>"
# wait for CI green
gh pr merge <num> --rebase --delete-branch
git checkout main && git pull --ff-only
```

## 5. Tag + GitHub release

```bash
gh release create vX.Y.Z --target main --title "vX.Y.Z" --notes "<changelog body>"
```

## 6. Publish

```bash
uv build
uv publish dist/outlook_graph_mcp-X.Y.Z-py3-none-any.whl dist/outlook_graph_mcp-X.Y.Z.tar.gz

clawhub publish "$(pwd)" --version X.Y.Z --tags latest --changelog "<one-liner>"

mcp-publisher publish   # may need `mcp-publisher login github` if JWT expired
```

## 7. Update GitHub About

If tool count or categories changed:

```bash
gh repo edit mpalermiti/outlook-mcp --description "MCP server for Microsoft Outlook personal accounts via Microsoft Graph API. N tools across K categories — mail, calendar, contacts, tasks, drafts, attachments. Community project, not affiliated with Microsoft."
```

## 8. Verify

```bash
curl -s https://pypi.org/pypi/outlook-graph-mcp/X.Y.Z/json | python3 -c "import json,sys; print(json.load(sys.stdin)['info']['version'])"
```

And confirm the MCP registry shows the new version as `(latest)`:

```bash
curl -s 'https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti&limit=20' | python3 -c "import json,sys; [print(s['server'].get('version'), '(latest)' if s.get('_meta',{}).get('io.modelcontextprotocol.registry/official',{}).get('isLatest') else '') for s in json.load(sys.stdin).get('servers', [])]"
```
