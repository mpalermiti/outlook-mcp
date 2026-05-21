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

## 2. Tests + lint

```bash
uv run pytest --tb=no -q
uv run ruff check src/ tests/
```

A single pre-existing failure (`test_graph.py::test_graph_client_init`, SOCKS env import) is expected and unrelated.

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
