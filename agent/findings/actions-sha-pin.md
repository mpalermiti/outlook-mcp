# actions-sha-pin — 2026-10-14

## Pins

Every `uses:` in `.github/workflows/ci.yml` and `publish.yml`, resolved with
`gh api repos/<owner>/<action>/git/ref/tags/<tag>` (annotated tags dereferenced
through `git/tags/<sha>`). Each pin stays inside the major the workflow already
used: pinning is not a version bump. Each pin is the commit the moving major tag
pointed at on 2026-10-14, so the code that runs is the same code as before.

| Action | Was | Pinned tag | Major tag → | Commit SHA |
| --- | --- | --- | --- | --- |
| `actions/checkout` | `@v4` | v4.4.0 | `v4` lightweight → same commit | `11d5960a326750d5838078e36cf38b85af677262` |
| `astral-sh/setup-uv` | `@v5` | v5.4.2 | `v5` annotated tag `e58605a9…` → same commit | `d4b2f3b6ecc6e67c4457f6d3e41ec42d3d0fcb86` |
| `actions/upload-artifact` | *(new, publish split)* | v4.6.2 | `v4` → same commit | `ea165f8d65b6e75b540449e92b4886f43607fa02` |
| `actions/download-artifact` | *(new, publish split)* | v4.3.0 | `v4` → same commit | `d3f86a106a0bac45b974a628896c90dbdf5c8093` |

Newer majors exist (checkout v7.0.1, setup-uv v10.3.0, upload-artifact v7.0.2,
download-artifact v8.0.2). Dependabot will open PRs for them. Each one is a
normal review: read the release notes before merging.

## Dependabot

`.github/dependabot.yml`: `github-actions` only, weekly. No `pip`/`uv`
ecosystem. Python caps are hand-set with a comment naming the upstream fix, and
a bot bump would bypass that.

## Publish split

`publish.yml` was one job holding `id-token: write` that also ran `uv sync`,
the full test suite and the build. Every dev dependency and every test ran with
a token that can publish to PyPI.

Now there are three jobs:

- **build** has `contents: read` and no token. It does the checkout, the
  version lockstep check, tests and lint, and `uv build`, then uploads `dist/`
  as an artifact and outputs `version`.
- **pypi** (`needs: build`) has `id-token: write` only. It downloads `dist/`,
  runs setup-uv with `enable-cache: false`, and runs `uv publish`. It has no
  checkout, so no repo code runs in it. The cache is off because a cache
  restored into a job that holds a token is a way to poison the publish.
- **registry** (`needs: [build, pypi]`) has `id-token: write` and
  `contents: read`. It does the checkout (for `server.json` and the
  mcp-publisher script), waits for PyPI, installs the pinned mcp-publisher and
  publishes. It also writes the summary. Version comes from
  `needs.build.outputs.version`.

PyPI trusted publishing matches the repository, the workflow file
(`publish.yml`) and the environment, not the job name. The rename from
`publish` to `pypi` therefore does not touch the PyPI setting (ASSUMED, from
PyPI's publisher form, which asks for no job name). The MCP registry's
`github-oidc` login checks the repository owner.

## Guard

`tests/test_workflows_are_pinned.py` has four tests:

1. Every `uses:` is `owner/action@<40 hex> # vX.Y.Z`.
2. Dependabot covers `github-actions` and nothing else.
3. Only `pypi` runs `uv publish`. Only `pypi` and `registry` hold `id-token: write`. `pypi` has no checkout, test, sync, build or script step, needs `build`, and has the cache off.
4. No publish job names an `environment:`. This stops one from being added without the matching PyPI change.

Red before the change (3 of 4 failed: unpinned uses, no dependabot.yml, uploader was `publish`). Green after. Mutation-checked: unpinning one checkout reddens test 1, and adding `environment: pypi` reddens test 4.

## Verification

- Suite: 1365 passed, 85 deselected. ruff check and format are clean.
- `actionlint` 1.7.12 (run ephemerally via `uvx --from actionlint-py`, not added to the project) is clean on both workflows.
- All three YAML files parse (PyYAML, ephemeral).
- Not verifiable offline: an actual release run. The first release after merge is the real test. If it fails, `gh workflow run publish.yml` re-runs all three jobs idempotently (`--check-url` on PyPI, duplicate-version handling in the registry step).

## Parked: a protected `pypi` environment

The card held this back: do not add `environment:` without Michael.

- **VERIFIED:** the repo has 0 GitHub environments (`gh api repos/mpalermiti/outlook-mcp/environments` → `total_count: 0`). The last five publish runs (v1.22.1 to v1.25.1) all succeeded with no environment, so PyPI's trusted-publisher entry does not demand one. RELEASING.md's appendix says to leave it blank.
- **NOT VERIFIABLE:** PyPI's publisher settings are behind your login, and PyPI serves no provenance for these files (`/integrity/.../provenance` → 404), so the entry can't be read from outside.
- **ASSUMED:** a blank environment on PyPI accepts a job that names one. If so, the safe order is workflow first, PyPI second. PyPI's docs only say the field is optional; they don't say how a blank one matches.

What an environment would buy: today anyone with write access can run `gh workflow run publish.yml` on any branch, and that run can publish. A `pypi` environment limited to `v*` tags (optionally with you as the required reviewer) closes that gap.

Two settings have to change together:

1. **GitHub:** Settings → Environments → New `pypi`. Under deployment branches and tags, choose "Selected" and add the tag rule `v*`. Required reviewers is optional.
2. **PyPI:** Project → Publishing → edit the GitHub publisher and set Environment to `pypi`.

The workflow change is one line under the `pypi` job, `environment: pypi`, plus flipping test 4 to require it. Do the GitHub side first, then the workflow line. Change PyPI last, and only once a release has gone through with the environment named. A mismatch in the other direction fails the upload. A dispatch re-run retries it once fixed.

**Correction (2026-10-14, from review):** two things the paragraphs above left out.

- **The retry has to change with it.** The documented retry, `gh workflow run publish.yml`, dispatches on the default branch. A `pypi` environment limited to `v*` tags refuses that run's `pypi` job, and `registry` is then skipped, so a release that failed after the upload could not be finished the documented way. The same change that adds `environment: pypi` must move the retry (RELEASING.md, the `workflow_dispatch` comment in `publish.yml`, and the Verification section above) to `gh workflow run publish.yml --ref vX.Y.Z`, or to `gh run rerun <id> --failed`, which keeps the release's ref. It is "one line" plus that.
- **What the environment buys is more than the dispatch gap.** With no environment on PyPI's entry, PyPI accepts an OIDC token from any job in `publish.yml`, so `registry`, which holds `id-token: write` and runs the downloaded mcp-publisher, could upload too. The split keeps the token out of `build`; "only `pypi` can publish" holds against PyPI only once PyPI's entry names the environment, since PyPI checks `environment` only when the publisher names one.
