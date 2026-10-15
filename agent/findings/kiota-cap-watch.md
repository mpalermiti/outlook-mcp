# kiota cap watch — 2026-10-14

**Status: the upstream fix has shipped. The cap can lift, but not by itself: lifting it safely
also needs a `msgraph-core>=1.5.2` floor.** Nothing in `pyproject.toml` or `uv.lock` was changed
(workers do not change dependencies).

## What the cap is

`pyproject.toml` caps four `microsoft-kiota-*` packages below 1.13 (#80). kiota 1.13.0
(2026-09-18) moved per-request options from `request.options` to
`request.extensions["kiota_request_options"]`; msgraph-core 1.5.1's `AsyncGraphTransport` only
sends a request through the middleware pipeline when `hasattr(request, "options")`, so under
kiota 1.13 it skipped the pipeline entirely, the `/users/me-token-to-replace` → `/me` rewrite
never ran, and every `/me` call failed. The comment says: lift once a msgraph-core release
carries microsoftgraph/msgraph-sdk-python-core#1129.

Locked today: kiota-* 1.12.3, msgraph-core 1.5.1, msgraph-sdk 1.62.0.

## What upstream did (checked 2026-10-14)

| Fact | Source | Label |
|---|---|---|
| #1129 is a PR, "fix: honor Kiota request extensions in Graph transport", **merged 2026-10-06 18:18Z** (merge commit `ca47287`) | `gh pr view 1129 -R microsoftgraph/msgraph-sdk-python-core` | VERIFIED |
| **msgraph-core v1.5.2** released 2026-10-06 18:52Z (GitHub), uploaded to PyPI 19:13Z, not yanked; latest | `gh release list`, `pypi.org/pypi/msgraph-core/json` | VERIFIED |
| v1.5.2 contains the fix: `compare ca47287...v1.5.2` is `ahead` by 2 (release commits); `compare v1.5.1...v1.5.2` lists the three #1129 commits | `gh api .../compare` | VERIFIED |
| v1.5.2's `AsyncGraphTransport` dispatches through the pipeline when `kiota_request_options` is in `request.extensions` **or** `request.options` exists, and reads options from either | released source at tag `v1.5.2` | VERIFIED by reading |
| #1129 added a transport test that routes options through `request.extensions` and asserts the middleware honours them (redirect handler) | PR files | VERIFIED by reading; it does not test the `/me` rewrite specifically |
| msgraph-core 1.5.2 **requires kiota-abstractions/-http/-authentication-azure `>=1.13.0,<2`** | PyPI `requires_dist` | VERIFIED |
| msgraph-sdk latest is **1.64.0** (2026-10-06); it requires only `msgraph_core>=1.5.1` | PyPI `requires_dist` | VERIFIED |
| kiota latest on PyPI is **1.14.3** (2026-10-13); GitHub tagged 1.14.4 today (2026-10-14 17:44Z), not yet on PyPI at check time | PyPI JSON, `gh release list -R microsoft/kiota-python` | VERIFIED |
| Our `/me` rewrite works on msgraph-core 1.5.2 + kiota 1.13/1.14 | — | **ASSUMED** (from the released code). A throwaway-env run of `tests/test_me_rewrite_reaches_the_wire.py` was declined in this session's sandbox; not run |

## Why the cap cannot simply be deleted

Two consequences of the version metadata above:

1. **While the cap stays, users are pinned to msgraph-core 1.5.1.** 1.5.2 needs kiota ≥1.13 and
   our cap forbids it, so any resolver picks 1.5.1. That is safe (1.5.1 + kiota 1.12 is the
   working pair), just old.
2. **Deleting the cap alone reopens #80 for some installs.** msgraph-core 1.5.1 accepts
   kiota `>=1.11.6`, and msgraph-sdk accepts msgraph-core `>=1.5.1`, so "kiota 1.14 + msgraph-core
   1.5.1" is a valid resolution — the exact broken pair. A fresh `uv tool install` picks newest
   (1.5.2) and works; but `pip install` into an environment that already has msgraph-core 1.5.1
   keeps it, and a lowest-direct or constrained resolve can choose it. The lift must add
   `msgraph-core>=1.5.2` as a direct requirement in the same change (a new line in
   `dependencies`, with a comment naming #80 and #1129).

So the change is: drop the four `<1.13` caps, add `msgraph-core>=1.5.2`, re-lock, CHANGELOG
under Unreleased. Whether to take kiota 1.13 **and** 1.14 now (an upper bound `<2` is implied by
msgraph-core) is part of the same decision.

## How a lift would be proven before any release

`ci.yml`'s `fresh-install` job runs on every push (not only releases): it builds the wheel,
resolves without the lock, and runs `tests/test_me_rewrite_reaches_the_wire.py` against what it
resolved — the exact #80 symptom. The `test` job runs the whole suite under the new lock. Both
run on the rolling PR, so a lift is checked on GitHub before it reaches a release. The weekly
`published-install` job then guards the published package after the release.

## Recommendation

Lift it, in one change with the `msgraph-core>=1.5.2` floor, on a push whose CI is read before
merging. Staying capped means every new user runs a kiota two minor versions old and an
msgraph-core missing whatever lands after 1.5.1; the guard that caught #80 now runs on every
push, so the risk of lifting is visible before release. This is a dependency change, which
workers do not make, so it is parked for Michael (queue line `kiota-cap-lift`).

## Re-check commands (no credentials)

```bash
gh pr view 1129 -R microsoftgraph/msgraph-sdk-python-core --json state,mergedAt
gh release list -R microsoftgraph/msgraph-sdk-python-core -L 3
curl -s https://pypi.org/pypi/msgraph-core/json | python3 -c "import json,sys; i=json.load(sys.stdin)['info']; print(i['version'], i['requires_dist'])"
curl -s https://pypi.org/pypi/msgraph-sdk/json  | python3 -c "import json,sys; i=json.load(sys.stdin)['info']; print(i['version'], [r for r in i['requires_dist'] if 'core' in r])"
curl -s https://pypi.org/pypi/microsoft-kiota-http/json | python3 -c "import json,sys; print(json.load(sys.stdin)['info']['version'])"
```
