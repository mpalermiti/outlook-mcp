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
- Not verifiable offline: an actual release run. The first release after merge is the real test. If it fails, `gh workflow run publish.yml --ref vX.Y.Z` (the release tag; see below) re-runs all three jobs idempotently (`--check-url` on PyPI, duplicate-version handling in the registry step).

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

## Done (2026-10-20, Michael chose A)

- `publish.yml`: `environment: pypi` on the `pypi` job only. The `workflow_dispatch` comment, the registry step's retry comment and RELEASING.md now give the retry as `gh workflow run publish.yml --ref vX.Y.Z` (RELEASING.md also names `gh run rerun <id> --failed`). RELEASING.md appendix: Environment `pypi`, plus a "The `pypi` environment" section with both settings and their order.
- Guards in `tests/test_workflows_are_pinned.py`: test 4 flipped to `test_only_the_pypi_job_runs_in_the_pypi_environment` (exactly `{pypi: pypi}`), and new `test_the_documented_publish_retry_names_the_release_ref` (no `gh workflow run publish.yml` without `--ref` in RELEASING.md or publish.yml). Both red before, green after; mutation-checked (environment on `registry` → red; bare retry restored in RELEASING.md → red).
- **VERIFIED (2026-10-20, PyPI source):** warehouse `warehouse/oidc/models/github.py::_check_environment` returns `True` when the publisher's environment is `""`, so a blank PyPI entry accepts a job that names `pypi`. The earlier ASSUMED line is settled: the workflow side cannot break a release.
- **VERIFIED (2026-10-20, `gh api .../environments`):** the repo still has 0 environments. ASSUMED (GitHub docs, not exercised): a job naming a missing environment creates it with no protection rules on first run, so merging before Michael creates it publishes as today, just unprotected.
- Not verifiable here: `actionlint` (its binary download is blocked in the sandbox); PyYAML parse is clean and shows `environment` only on `pypi`. A real release is the end-to-end test.
- Still Michael's, in order: (1) GitHub Settings → Environments → `pypi`, tag rule `v*`, before the next release; (2) PyPI publisher Environment → `pypi`, after that release succeeds.

## Correction (2026-10-20, from review of cfbb1d2)

Three things the Done section got wrong, all fixed in the same commit:

- **PyPI cannot edit a publisher.** VERIFIED (warehouse source, `warehouse/manage/views/oidc_publishers.py`): the publisher views are add, delete and `constrain_environment`; constrain adds a publisher with the environment and then removes the blank one. VERIFIED (`warehouse/oidc/models/github.py::_get_publisher_for_environment`): an exact environment match wins, else any publisher with environment `""` is used. So adding a second publisher with `pypi` beside the blank one locks nothing, silently. RELEASING.md step 3 now says: use the constrain link PyPI emails after the first release with the environment, or delete and re-add; never leave the blank one.
- **The tag rule checks only a tag's name.** VERIFIED (`gh api repos/mpalermiti/outlook-mcp/rulesets/24471530`): "Protect release tags" on `refs/tags/v*` has rules `deletion`, `update`, `non_fast_forward`, no `creation`, no bypass actors. Anyone with write access could create a `v*` tag on any commit, and the workflow that runs is the file at that tag. VERIFIED: today the only collaborator is the owner, so nothing is exposed yet. RELEASING.md step 2 now adds a separate tag ruleset with only "Restrict creations" and Repository admin as bypass (separate so the bypass does not also lift the delete/move protections), or a required reviewer on `pypi` instead. CHANGELOG entry says the lock is as strong as that rule. ASSUMED (GitHub docs, not exercised): a ruleset bypass applies to all of its rules; `gh release create` by a bypassing admin can still create the tag.
- **The environment guard missed quoted values.** `test_only_the_pypi_job_runs_in_the_pypi_environment` matched only `[\w-]+`, so `environment: "pypi"` on `registry` passed. Now any `environment:` key on a job counts (`_environment`, which also reads the mapping form's `name:`), pinned by `test_the_environment_reader_sees_every_spelling`. Mutation: `"pypi"`, `'pypi'` and `${{ 'pypi' }}` on `registry` each turn the guard red.

## Re-check (2026-10-28, Michael chose A; both halves done 2026-10-23)

Read-only, `gh api repos/mpalermiti/outlook-mcp/...` on 2026-10-28. Nothing was changed.

**GitHub side — VERIFIED:**

- `environments` → `total_count: 1`, only `pypi`.
- `environments/pypi` → `deployment_branch_policy: {custom_branch_policies: true, protected_branches: false}`; one protection rule, `type: branch_policy`; no required reviewers (A, not B).
- `environments/pypi/deployment-branch-policies` → exactly one policy, `{name: "v*", type: "tag"}`. No branch policy, so a dispatch on `main` is refused.
- `rulesets` → three, all `enforcement: active`: `Protect main` (24471529, branch), `Protect release tags` (24471530, tag), `Release tag creation` (25761696, tag).
- `rulesets/25761696` → `refs/tags/v*`, rules `["creation"]` only, one bypass actor `{actor_type: RepositoryRole, actor_id: 5, bypass_mode: always}`. ASSUMED (GitHub docs, not exercised): repository role id 5 is admin.
- `rulesets/24471530` unchanged from 2026-10-20: `deletion`, `update`, `non_fast_forward`, no bypass actors. The admin bypass lives only on the creation ruleset, so it does not lift delete/move.
- `publish.yml`: `environment: pypi` on the `pypi` job only (guarded by `test_only_the_pypi_job_runs_in_the_pypi_environment`).

**PyPI side — Michael-confirmed (2026-10-23), not checkable from here:** the blank publisher entry was replaced by one bound to environment `pypi`, and the blank one removed. PyPI's publisher settings sit behind his login.

**Consistent with both, not proof of the PyPI half:** every release since has deployed through the `pypi` environment and reached PyPI. `deployments?environment=pypi` shows v1.27.0 → v1.30.1 (five) after both halves were set, each latest status `success`; PyPI's JSON API shows each version uploaded within seconds of its deployment (1.27.0 at 2026-10-24T07:11:17Z). v1.26.0 is evidence for neither half: its run auto-created the environment with no rules (environment `created_at` 2026-10-23T23:47:39Z, the same second as that deployment), the `Release tag creation` ruleset was created 18 minutes after it reached PyPI (2026-10-24T00:05:50Z), and Michael's PyPI change came later still. Corrected 2026-10-28 from "v1.26.0 → v1.30.1 (six)" per the checker's review. A successful upload shows the environment-bound entry accepts the job; it cannot show the blank one is gone, since a blank entry would accept it too.

Closes the card. Nothing further is gated on it.
