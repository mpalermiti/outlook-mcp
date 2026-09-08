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

## 1d. Write-tier guards (calendar only)

```bash
OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
```

The only tier that writes. It creates short, bounded, attendee-free recurring events on the authenticated calendar and deletes each one in a `finally`.

It exists because recurrence cannot be validated any other way: a `PatternedRecurrence` that is well-formed to the SDK still 400s with `ErrorInvalidRecurrenceRange` if the range disagrees with the series master's start. `outlook_create_event` accepted a `recurrence` argument and silently discarded it for fourteen minor versions (#41) under a fully green mock suite — mocks assert what we *build*.

Double-gated on purpose: the marker is deselected by default **and** the tier skips without `OUTLOOK_MCP_LIVE_WRITE=1`, so a cached token alone can never write to a calendar. Run it if you changed anything under `tools/_recurrence.py` or `calendar_write.py`. See the rules at the top of `tests/conftest.py` before adding to it.

## 2. Tests + lint

```bash
uv run pytest --tb=no -q
uv run ruff check src/ tests/
```

The default run is the offline unit suite only — `addopts` deselects the `integration` and `live` markers, so this needs no network or token. Expect `N passed, 22 deselected` and zero failures.

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

## 5. Tag + GitHub release — this publishes everything

```bash
gh release create vX.Y.Z --target main --title "vX.Y.Z" --notes "<changelog body>"
```

Publishing the release triggers `.github/workflows/publish.yml`, which re-checks the version lockstep, runs tests and lint, builds, publishes to **PyPI**, waits for PyPI's index to catch up, then publishes to the **MCP registry**. Both authenticate through GitHub OIDC — no stored tokens, and no five-minute registry login to race by hand.

Watch it:

```bash
gh run watch
```

If it fails partway, re-run it. Uploads already on PyPI are skipped, so a dispatch retries only what didn't finish:

```bash
gh workflow run publish.yml
```

> The workflow deliberately does **not** run the live tier — those need real credentials. Step 1 is still yours, and still the step that matters: a green offline suite is exactly what shipped 1.13.0 and 1.13.1 broken.

## 6. Publish to ClawHub (the one manual channel)

ClawHub has no OIDC equivalent, so it stays hand-run. Three steps, and the first and last are the ones that matter.

**6a. Check the CLI is current — nothing else will.**

```bash
clawhub --cli-version; npm view clawhub version   # must match
npm i -g clawhub@latest                            # it's npm-global under Homebrew's node, not a brew formula
```

Site discovery advertises `minCliVersion: "0.1.0"`, so an arbitrarily old client is accepted without a warning. On 2026-09-07 a v0.9.0 client (latest was v0.23.3) printed `✔ OK. Published outlook-mcp@1.15.0` for a submission that was actually pending security scans, and two releases were reported as published to users before anyone checked. Current clients print `pending security scans before it becomes public`, which is the truth.

**6b. Publish.**

```bash
clawhub publish "$(pwd)" --version X.Y.Z --tags latest --changelog "<one-liner>"
```

Expect it to take a minute or two and to say **pending security scans**. That is success. The scan has taken ~12 min to ~1 h in practice; the version is not public until it clears.

Note that ClawHub bundles the **whole repo** (everything not in `.gitignore` with a text extension — `tests/` included), not the wheel. Don't leave one-off scripts that touch real data lying in the tree at publish time.

**6c. Verify — the CLI's success message is not verification.**

```bash
curl -s -o /dev/null -w '%{http_code}\n' https://clawhub.ai/api/v1/skills/outlook-mcp/versions/X.Y.Z   # 200 once public
curl -s https://clawhub.ai/api/v1/skills/outlook-mcp | python3 -c "import json,sys; print(json.load(sys.stdin)['latestVersion']['version'])"
```

`404` immediately after publishing is normal (scan pending). `404` an hour later is not — and `clawhub publish` will then refuse the same version number as a duplicate, so don't burn versions probing it; ask on <https://github.com/openclaw/clawhub/issues> (see #3623 for the shape of this).

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


## Appendix — one-time PyPI trusted publisher setup

Step 5 uploads to PyPI without a token by using PyPI's *trusted publishing*: PyPI verifies the workflow through GitHub OIDC rather than a stored API token. It has to be registered once in PyPI's web UI — it cannot be scripted:

1. Go to <https://pypi.org/manage/project/outlook-graph-mcp/settings/publishing/>
2. Add a **GitHub** publisher with exactly:

   | Field | Value |
   | --- | --- |
   | Owner | `mpalermiti` |
   | Repository | `outlook-mcp` |
   | Workflow name | `publish.yml` |
   | Environment | *(leave blank)* |

Until this is saved, the workflow's PyPI step fails with an OIDC/trusted-publishing error. Once saved, `UV_PUBLISH_TOKEN` is no longer needed and can be dropped from the shell environment.
