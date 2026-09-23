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
#
# The same caveat applies to the To Do attachment *write* path: the inline
# base64 POST to ``.../tasks/{id}/attachments`` mutates the task, so it
# stays out of this read-only script; the GET rows that ``_todo_task_rows``
# adds cover the same resource family (a 403/501 "not supported for this
# account type" shows up on the GETs too). The upload-session route
# (``createUploadSession``) is deliberately not probed either: the tools
# don't use it (its upload URL needs per-chunk Authorization/Content-Type
# headers the current inline path doesn't), so its reachability says
# nothing about this release.
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


def _get_json(path_or_url: str, headers: dict[str, str], timeout: int = 20):
    """GET and return (status_code, parsed_json_or_None)."""
    url = path_or_url if path_or_url.startswith("http") else urljoin(GRAPH_BASE, path_or_url)
    try:
        r = httpx.get(url, headers=headers, timeout=timeout)
    except Exception as exc:
        return None, f"exception: {exc}"
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, None


def _default_list_id(headers: dict[str, str]) -> str:
    """The id of the default task list, the same way the To Do tools pick it.

    ``?$top=1`` on ``/me/todo/lists`` is not that: it grabs whichever list
    Graph happens to return first (often enough, none — leaving every task
    family unprobed), and the tools themselves resolve
    ``wellknownListName="defaultList"`` + ``isOwner``, falling back to the
    first list. The preflight harvests tasks from the list the tools would
    actually use.
    """
    path = "me/todo/lists?$top=100"
    pages = 0
    first_id = None
    while path and pages < 20:
        status, body = _get_json(path, headers)
        if status != 200 or not isinstance(body, dict):
            raise RuntimeError(f"listing me/todo/lists -> {status}")
        for lst in body.get("value") or []:
            if first_id is None and lst.get("id"):
                first_id = lst["id"]
            if lst.get("isOwner") and lst.get("wellknownListName") == "defaultList":
                return lst["id"]
        path = body.get("@odata.nextLink")
        pages += 1
    if first_id:
        return first_id
    raise RuntimeError("no task list at all (GET me/todo/lists -> 200, empty)")


def _task_with_attachment(headers: dict[str, str], list_id: str) -> tuple[str, str | None]:
    """A task id from the default list, plus an attachment id if one task has
    one (first task wins for the id; up to five tasks are checked for
    attachments so an attachment on a slightly older task still widens the
    probe)."""
    status, body = _get_json(f"me/todo/lists/{list_id}/tasks?$top=10", headers)
    if status != 200 or not isinstance(body, dict):
        raise RuntimeError(f"listing tasks of the default list -> {status}")
    tasks = body.get("value") or []
    if not tasks:
        raise RuntimeError("default list has no tasks")
    task_id = tasks[0]["id"]
    for task in tasks[:5]:
        tid = task.get("id")
        if not tid:
            continue
        s, b = _get_json(f"me/todo/lists/{list_id}/tasks/{tid}/attachments?$top=1", headers)
        if s == 200 and isinstance(b, dict) and b.get("value"):
            return tid, b["value"][0].get("id")
    return task_id, None


def _todo_task_rows(headers: dict[str, str]) -> tuple[list[tuple[str, str]], list[str], list[str]]:
    """Probe rows for the To Do checklist/attachment families.

    Unlike every static row above, these need a real task id — the endpoints
    live under ``.../tasks/{taskId}/…`` and a made-up id 404s for the wrong
    reason (missing resource, not unsupported family). Read-only throughout.

    Returns ``(rows, not_probed, failures)``: ``rows`` go to the generic
    probe loop; ``not_probed`` names families that could not be provisioned
    and must surface in the summary rather than counting as checked;
    ``failures`` carries 403/501-class verdicts measured here (the entity
    GET is fetched in this function because its verdict is not the status
    code alone — ``contentBytes`` has to be in the body for the download
    tool to have something to read).
    """
    rows: list[tuple[str, str]] = []
    not_probed: list[str] = []
    failures: list[str] = []
    try:
        list_id = _default_list_id(headers)
        task_id, attachment_id = _task_with_attachment(headers, list_id)
    except Exception as exc:
        reason = f"could not provision a live task: {exc}"
        return [], [
            f"To Do checklist items ({reason})",
            f"To Do task attachments ({reason})",
            f"To Do attachment entity GET ({reason})",
        ], []
    task_path = f"me/todo/lists/{list_id}/tasks/{task_id}"

    rows.append((f"{task_path}/checklistItems?$top=1", "To Do checklist items"))
    rows.append((f"{task_path}/attachments?$top=1", "To Do task attachments"))

    if not attachment_id:
        not_probed.append(
            "To Do attachment entity GET + $value (no attachment on the default "
            "list's recent tasks — attach a file to a task to widen this probe)"
        )
        return rows, not_probed, failures

    # The entity GET is what outlook_download_task_attachment actually calls
    # (contentBytes off the attachment entity), so it is probed directly and
    # the body is checked — a 200 without contentBytes is the shape the
    # download tool now refuses, and worth a loud line here.
    status, body = _get_json(
        f"{task_path}/attachments/{attachment_id}", headers, timeout=60
    )
    verdict = classify(status) if status is not None else "FAIL"
    has_bytes = isinstance(body, dict) and "contentBytes" in body
    print(
        f"{verdict:5s} {'To Do attachment entity':30s} "
        f"{task_path}/attachments/{attachment_id}  {status} "
        f"contentBytes={'yes' if has_bytes else 'NO'}"
    )
    if verdict == "FAIL":
        failures.append("To Do attachment entity")
    if status == 200 and not has_bytes:
        print(
            "      NOTE  a 200 without contentBytes is the response shape the "
            "download tool refuses (size cross-check) — if the attachment is "
            "genuinely non-empty, that is a finding, not a pass"
        )
    # Legacy raw route, no longer used by any tool — kept as an extra signal.
    rows.append(
        (f"{task_path}/attachments/{attachment_id}/$value", "To Do attachment $value")
    )
    return rows, not_probed, failures


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
    am.try_cached_token()
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

    todo_rows, not_probed, todo_failures = _todo_task_rows(headers)
    failures.extend(todo_failures)
    endpoints = ENDPOINTS + todo_rows
    reachable = 0

    for path, label in endpoints:
        url = urljoin(GRAPH_BASE, path)
        try:
            r = httpx.get(url, headers=headers, timeout=20)
        except Exception as exc:
            print(f"FAIL  {label:30s} {path}  exception: {exc}")
            failures.append(label)
            continue

        verdict = classify(r.status_code)
        if verdict == "OK":
            reachable += 1
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
    summary = f"{reachable} of {len(endpoints)} probed endpoints reachable."
    if not_probed:
        # Unprobed is not verified — say so instead of printing "Safe to tag"
        # over families the run never touched.
        print(
            f"{summary} {len(not_probed)} NOT probed: "
            + "; ".join(not_probed)
            + ". Not probed is not verified — widen the probe before tagging."
        )
    else:
        print(f"{summary} Safe to tag.")
    return 0


if __name__ == "__main__":
    sys.exit(run())
