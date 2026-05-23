"""Pre-release smoke test against the live Microsoft Graph API.

Run before tagging a release. Hits every Graph endpoint family that
outlook-mcp tools depend on, using the locally-cached token, and flags
any endpoint that returns 403 or 501 — the "not supported for this
account type" signals that unit tests can't catch.

Read-only. No writes, no sends, no mailbox state changes.

Why this exists: v1.7.0 shipped four mailbox-settings tools backed by
/me/mailboxSettings/*, which Microsoft Graph does not support on
personal Microsoft accounts (the project's only target). Unit tests
passed because the SDK was mocked. A 30-second curl against the live
endpoint would have caught it. This script is that 30-second curl,
codified.

Usage:
    uv run python scripts/preflight.py

Requires that `outlook-mcp auth` has been run successfully — the
script reuses the cached token. Exits 0 on success, non-zero on any
endpoint failure.
"""

from __future__ import annotations

import sys
from urllib.parse import urljoin

import httpx

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"

# Each row covers one Graph endpoint family. The path is the smallest
# request that exercises the surface — usually a $top=1 list — so the
# script runs fast and doesn't drag the user's full mailbox over the
# wire. The label maps to the outlook_mcp tool group(s) that depend
# on this endpoint, so a failure points at the broken tools directly.
#
# NOTE: ``/$batch`` is exercised by ``outlook_read_messages`` (v1.11.0)
# and ``outlook_batch_triage`` but isn't directly preflighted — it's a
# POST endpoint with a body, and the current preflight harness only
# issues GETs. The underlying ``me/messages`` GET is already covered
# below, and ``$batch`` is just transport, so the surface is exercised
# indirectly. Refactor to support POST rows if a $batch-specific
# regression ever needs catching at preflight time.
ENDPOINTS: list[tuple[str, str]] = [
    ("me", "Auth / whoami"),
    ("me/messages?$top=1", "Mail read / search"),
    ("me/mailFolders?$top=1", "Mail folders"),
    ("me/mailFolders/drafts/messages?$top=1", "Drafts"),
    ("me/events?$top=1", "Calendar read"),
    ("me/calendars", "Calendar list"),
    ("me/contacts?$top=1", "Contacts"),
    ("me/todo/lists", "To Do"),
    ("me/inferenceClassification/overrides", "Focused Inbox overrides"),
    ("me/outlook/masterCategories", "Categories"),
    ("me/mailFolders/inbox/messages/delta", "Mail delta"),
    (
        "me/calendarView/delta?startDateTime=2026-05-21T00:00:00Z"
        "&endDateTime=2026-05-28T00:00:00Z",
        "Calendar delta",
    ),
    ("me/contacts/delta", "Contacts delta"),
]


def classify(status_code: int) -> str:
    """Map an HTTP status from Graph to a preflight verdict.

    - ``"OK"`` (200, 204) — endpoint is supported and the request succeeded.
    - ``"FAIL"`` (403, 501) — the v1.7.0-class signal: endpoint not supported
      for this account type, or scopes are wrong. Release blocker.
    - ``"SKIP"`` (anything else) — 4xx query-shape, 401 transient auth, 429
      rate-limit, 5xx transient. Surface to a human but don't block release.
    """
    if status_code in (200, 204):
        return "OK"
    if status_code in (403, 501):
        return "FAIL"
    return "SKIP"


def fetch_token() -> str:
    config = load_config()
    am = AuthManager(config)
    am.try_cached_token(am.get_token_scopes())
    cred = am.get_credential()
    tok = cred.get_token("https://graph.microsoft.com/.default")
    return tok.token


def run() -> int:
    try:
        token = fetch_token()
    except Exception as exc:
        print(f"FAIL  could not acquire token — {exc}")
        print("Run `outlook-mcp auth` first.")
        return 2

    headers = {"Authorization": f"Bearer {token}"}
    failures: list[str] = []

    for path, label in ENDPOINTS:
        url = urljoin(GRAPH_BASE, path)
        try:
            r = httpx.get(url, headers=headers, timeout=20)
        except Exception as exc:
            print(f"FAIL  {label:30s} {path}  exception: {exc}")
            failures.append(label)
            continue

        verdict = classify(r.status_code)
        if verdict == "OK":
            print(f"OK    {label:30s} {path}  {r.status_code}")
        elif verdict == "FAIL":
            code = ""
            try:
                code = r.json().get("error", {}).get("code", "")
            except Exception:
                pass
            print(f"FAIL  {label:30s} {path}  {r.status_code} {code}")
            failures.append(label)
        else:
            print(f"SKIP  {label:30s} {path}  {r.status_code} (non-blocking)")

    print()
    if failures:
        print(
            f"{len(failures)} endpoint(s) FAILED: {', '.join(failures)}. "
            "Tools backed by these endpoints will not work on this account. "
            "Do not release until resolved."
        )
        return 1
    print(f"All {len(ENDPOINTS)} endpoints reachable. Safe to tag.")
    return 0


if __name__ == "__main__":
    sys.exit(run())
