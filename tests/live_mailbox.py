"""A test account's mailbox seen from outside the test: find, clean up, answer.

The ``live_send`` tier sends from the main test account to its pair, then needs to
look in the pair's mailbox, and both need cleaning up afterwards. Each account's
sign-in lives in its own settings directory (and, on Linux, its token cache under
``$HOME/.IdentityService``), so an account is reached by starting a process pointed
at its directory — never by switching accounts inside the test process::

    python -m tests.live_mailbox --as pair find --marker "<subject marker>" --wait 180

:func:`run` does that from a test, and ``check-leftovers`` (what
``scripts/live-tests.sh`` runs after a writing tier) does it for both accounts;
``check-leftovers --clean`` (``scripts/live-tests.sh clean``) first deletes what a
crash left, where deleting it sends nothing.

Guarded on its own, whatever the caller asks:

* It reads ``/me`` first, and refuses unless the account is the one ``--as`` names
  (``OUTLOOK_MCP_TEST_ACCOUNT`` or ``OUTLOOK_MCP_TEST_PAIR``).
* It deletes only items that carry a live-test marker (``tests/live_guard.MARKERS``),
  and never a meeting that has attendees: deleting one sends a cancellation, or
  answers the organizer.
* It sends only with ``--allow-send``: ``reply`` and ``accept``, once each, and only
  to the other test account. The calling test counts that send against its run's
  budget before it starts this process.

Output is one JSON object on the last line of stdout. Addresses never appear in
it: a sender is reported as ``account``, ``pair`` or ``other``.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_SCOPE = "https://graph.microsoft.com/.default"
FOLDERS = ("inbox", "junkemail", "sentitems", "deleteditems", "drafts")
POLL_SECONDS = 10


class Refused(SystemExit):
    """The helper will not do what it was asked."""

    def __init__(self, why: str) -> None:
        super().__init__(f"live_mailbox refused: {why}")


# ── From a test: start the helper as one of the two accounts ──────────────


class MailboxError(RuntimeError):
    pass


def run(role: str, *args: str, timeout: float = 420) -> dict:
    """Run the helper as ``role`` ("main" or "pair") and return its JSON answer."""
    from tests.live_guard import settings_from

    settings = settings_from(os.environ)
    env = dict(os.environ)
    if role == "pair":
        if not settings.pair_config_dir:
            raise MailboxError("the pair account's settings directory is not mounted")
        env.update(HOME=settings.pair_config_dir, OUTLOOK_MCP_CONFIG_DIR=settings.pair_config_dir)
    proc = subprocess.run(
        [sys.executable, "-m", "tests.live_mailbox", "--as", role, *args],
        cwd=REPO,
        env=env,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    answer = _last_json(proc.stdout)
    if proc.returncode != 0 or answer is None:
        # The unencrypted-cache notice is expected in the container; leave it out.
        lines = [x for x in (proc.stderr or "").strip().splitlines() if "unencrypted" not in x]
        tail = "\n".join(lines[-6:])
        if "service abuse mode" in (proc.stderr or ""):
            tail = (
                "Microsoft reports the account is in service abuse mode: locked until its "
                "owner signs in on the web and verifies it. " + tail
            )
        raise MailboxError(f"live_mailbox --as {role} {args[0]} failed ({proc.returncode}): {tail}")
    return answer


def _last_json(stdout: str) -> dict | None:
    for line in reversed((stdout or "").strip().splitlines()):
        if line.startswith("{"):
            try:
                return json.loads(line)
            except ValueError:
                return None
    return None


# ── Inside the helper process ──────────────────────────────────────────────


class Mailbox:
    """Raw Graph calls as the signed-in account, after checking who that is."""

    def __init__(self, role: str) -> None:
        import httpx

        from outlook_mcp.auth import AuthManager, _load_auth_record
        from outlook_mcp.config import load_config
        from tests.live_guard import identity_of, settings_from

        settings = settings_from(os.environ)
        self.expected = settings.account if role == "main" else settings.pair
        self.other = settings.pair if role == "main" else settings.account
        if not self.expected or not self.other:
            raise Refused("both test accounts must be named (OUTLOOK_MCP_TEST_ACCOUNT/_PAIR)")
        self.roles = {settings.account: "account", settings.pair: "pair"}

        auth = AuthManager(load_config())
        if not auth.try_cached_token():
            raise Refused(f"the {role} account is not signed in in this settings directory")
        token = auth.get_credential().get_token(TOKEN_SCOPE).token
        self.http = httpx.Client(
            base_url=GRAPH, headers={"Authorization": f"Bearer {token}"}, timeout=30
        )
        me = self.get("/me", {"$select": "userPrincipalName,mail"})
        record = _load_auth_record()
        if identity_of(me, getattr(record, "username", None)) != self.expected:
            raise Refused(f"the signed-in account is not the {role} test account")

    # Plain HTTP, raising on anything but success.
    def get(self, path: str, params: dict | None = None) -> dict:
        response = self.http.get(path, params=params)
        response.raise_for_status()
        return response.json()

    def get_all(self, path: str, params: dict | None = None, pages: int = 20) -> list[dict]:
        found: list[dict] = []
        page = self.get(path, params)
        for _ in range(pages):
            found += page.get("value", [])
            link = page.get("@odata.nextLink")
            if not link:
                break
            page = self.get(link)
        return found

    def role_of(self, address: str | None) -> str:
        return self.roles.get((address or "").strip().casefold(), "other")

    def sender_role(self, message: dict) -> str:
        return self.role_of(((message.get("from") or {}).get("emailAddress") or {}).get("address"))

    # Messages.
    def messages(self, marker: str, folders: list[str]) -> list[dict]:
        found = []
        for folder in folders:
            listing = self.get(
                f"/me/mailFolders/{folder}/messages",
                {
                    "$select": "id,subject,from,replyTo,conversationId,receivedDateTime",
                    "$orderby": "receivedDateTime desc",
                    "$top": "50",
                },
            )
            for m in listing.get("value", []):
                if marker in (m.get("subject") or ""):
                    full = self.get(f"/me/messages/{m['id']}")
                    found.append(
                        {
                            "id": m["id"],
                            "folder": folder,
                            "subject": m.get("subject"),
                            "from": self.sender_role(m),
                            "reply_to": [
                                self.role_of((r.get("emailAddress") or {}).get("address"))
                                for r in m.get("replyTo") or []
                            ],
                            "conversation_id": m.get("conversationId"),
                            "type": full.get("@odata.type", "#microsoft.graph.message"),
                            "meeting_message_type": full.get("meetingMessageType"),
                        }
                    )
        return found

    def permanently_delete(self, message_id: str) -> bool:
        response = self.http.post(f"/me/messages/{message_id}/permanentDelete")
        if response.status_code == 404:
            return False
        if response.is_success:
            return True
        if response.status_code not in (400, 405, 501):
            response.raise_for_status()  # throttled or failing: say so, don't settle for less
        # The account does not support permanentDelete: a plain DELETE instead.
        response = self.http.delete(f"/me/messages/{message_id}")
        if response.status_code == 404:
            return False
        response.raise_for_status()
        return True

    # Events.
    def events(self, marker: str) -> list[dict]:
        found = []
        for e in self.get_all(
            "/me/calendar/events",
            {
                "$select": "id,subject,isCancelled,isOrganizer,organizer,attendees,responseStatus",
                "$top": "100",
            },
        ):
            if marker in (e.get("subject") or ""):
                found.append(
                    {
                        "id": e["id"],
                        "subject": e.get("subject"),
                        "is_cancelled": bool(e.get("isCancelled")),
                        "is_organizer": bool(e.get("isOrganizer")),
                        "organizer": self.role_of(
                            ((e.get("organizer") or {}).get("emailAddress") or {}).get("address")
                        ),
                        "response": (e.get("responseStatus") or {}).get("response"),
                        "attendees": [
                            {
                                "who": self.role_of((a.get("emailAddress") or {}).get("address")),
                                "response": (a.get("status") or {}).get("response"),
                            }
                            for a in e.get("attendees") or []
                        ],
                    }
                )
        return found


def _require_marker(marker: str) -> str:
    from tests.live_guard import MARKERS

    if not any(m in marker for m in MARKERS):
        raise Refused("--marker must contain a live-test marker (tests/live_guard.MARKERS)")
    return marker


def _poll(fetch, done, wait: float):
    deadline = time.monotonic() + wait
    while True:
        got = fetch()
        if done(got) or time.monotonic() >= deadline:
            return got
        time.sleep(POLL_SECONDS)


def cmd_whoami(box: Mailbox, args) -> dict:
    """Only that the account is reachable and is the one expected (checked on start)."""
    return {"reachable": True}


def cmd_find(box: Mailbox, args) -> dict:
    marker = _require_marker(args.marker)

    def fetch():
        found = box.messages(marker, args.folder)
        if args.kind:
            found = [m for m in found if m["meeting_message_type"] == args.kind]
        return found

    found = _poll(fetch, lambda got: len(got) >= args.min, args.wait)
    return {"messages": found}


def cmd_purge(box: Mailbox, args) -> dict:
    marker = _require_marker(args.marker)
    deleted = sum(box.permanently_delete(m["id"]) for m in box.messages(marker, args.folder))
    return {"deleted": deleted}


def cmd_events(box: Mailbox, args) -> dict:
    marker = _require_marker(args.marker)
    return {
        "events": _poll(lambda: box.events(marker), lambda got: len(got) >= args.min, args.wait)
    }


def cmd_remove_events(box: Mailbox, args) -> dict:
    """Delete this account's copies of marked meetings someone else organized.

    Waits for each copy to show as cancelled first: deleting a copy of a live
    meeting could answer the organizer, and the cancellation is what this test bed
    sends anyway.
    """
    marker = _require_marker(args.marker)

    def settled(events):
        return all(e["is_cancelled"] or e["is_organizer"] for e in events)

    events = _poll(lambda: box.events(marker), settled, args.wait)
    removed = kept = 0
    for e in events:
        if e["is_organizer"] or not e["is_cancelled"]:
            kept += 1
            continue
        response = box.http.delete(f"/me/events/{e['id']}")
        if response.status_code != 404:
            response.raise_for_status()
        removed += 1
    return {"removed": removed, "kept": kept}


def cmd_reply(box: Mailbox, args) -> dict:
    marker = _require_marker(args.marker)
    if not args.allow_send:
        raise Refused("reply sends mail; pass --allow-send once the test has counted it")
    found = [
        m for m in box.messages(marker, ["inbox", "junkemail"]) if m["type"].endswith("message")
    ]
    if len(found) != 1:
        raise Refused(f"expected one marked message to reply to, found {len(found)}")
    message = found[0]
    if message["from"] != box.role_of(box.other):
        raise Refused("the message to reply to is not from the other test account")
    if any(r != message["from"] for r in message["reply_to"]):
        raise Refused("the message's Reply-To names someone other than its sender")
    box.http.post(
        f"/me/messages/{message['id']}/reply", json={"comment": args.comment}
    ).raise_for_status()
    return {"replied": True}


def cmd_accept(box: Mailbox, args) -> dict:
    marker = _require_marker(args.marker)
    if not args.allow_send:
        raise Refused("accept sends a response; pass --allow-send once the test has counted it")
    found = [e for e in box.events(marker) if not e["is_organizer"]]
    if len(found) != 1:
        raise Refused(f"expected one marked invitation to accept, found {len(found)}")
    event = found[0]
    if event["organizer"] != box.role_of(box.other):
        raise Refused("the invitation is not from the other test account")
    box.http.post(
        f"/me/events/{event['id']}/accept", json={"sendResponse": True}
    ).raise_for_status()
    return {"accepted": True}


def cmd_leftovers(box: Mailbox, args) -> dict:
    """Every item on this account that still carries a live-test marker. Read-only."""
    from tests.live_guard import MARKERS

    def marked(text) -> bool:
        return any(m in (text or "") for m in MARKERS)

    found: dict[str, list[str]] = {
        "messages": [
            m.get("subject")
            for m in box.get_all("/me/messages", {"$select": "subject", "$top": "250"})
            if marked(m.get("subject"))
        ],
        "calendars": [],
        "events": [],
        "contacts": [
            c.get("displayName")
            for c in box.get_all("/me/contacts", {"$select": "displayName", "$top": "250"})
            if marked(c.get("displayName"))
        ],
        "task_lists": [],
        "tasks": [],
    }
    for cal in box.get_all("/me/calendars", {"$select": "id,name"}):
        if marked(cal.get("name")):
            found["calendars"].append(cal.get("name"))
        found["events"] += [
            e.get("subject")
            for e in box.get_all(
                f"/me/calendars/{cal['id']}/events", {"$select": "subject", "$top": "250"}
            )
            if marked(e.get("subject"))
        ]
    for todo_list in box.get_all("/me/todo/lists"):
        if marked(todo_list.get("displayName")):
            found["task_lists"].append(todo_list.get("displayName"))
        found["tasks"] += [
            t.get("title")
            for t in box.get_all(f"/me/todo/lists/{todo_list['id']}/tasks", {"$top": "100"})
            if marked(t.get("title"))
        ]
    return {"leftovers": found, "total": sum(len(v) for v in found.values())}


def cmd_clean(box: Mailbox, args) -> dict:
    """Delete every marked item on this account whose deletion sends nothing.

    For what a crash or a dropped connection leaves behind: a create whose reply
    never arrived leaves an item the test never learnt the id of. Messages, events
    with no attendees, contacts, To Do tasks and lists. A marked meeting with
    attendees is reported, not deleted: deleting it would send a cancellation.
    """
    from tests.live_guard import MARKERS

    def marked(text) -> bool:
        return any(m in (text or "") for m in MARKERS)

    def delete(path: str) -> None:
        response = box.http.delete(path)
        if response.status_code != 404:
            response.raise_for_status()

    done = {"messages": 0, "events": 0, "contacts": 0, "tasks": 0, "task_lists": 0}
    kept: list[str] = []
    for m in box.get_all("/me/messages", {"$select": "id,subject", "$top": "250"}):
        if marked(m.get("subject")):
            done["messages"] += box.permanently_delete(m["id"])
    for cal in box.get_all("/me/calendars", {"$select": "id"}):
        for e in box.get_all(
            f"/me/calendars/{cal['id']}/events",
            {"$select": "id,subject,attendees,isOrganizer", "$top": "250"},
        ):
            if not marked(e.get("subject")):
                continue
            if e.get("attendees"):
                kept.append(f"meeting with attendees: {e.get('subject')}")
                continue
            delete(f"/me/events/{e['id']}")
            done["events"] += 1
    for c in box.get_all("/me/contacts", {"$select": "id,displayName", "$top": "250"}):
        if marked(c.get("displayName")):
            delete(f"/me/contacts/{c['id']}")
            done["contacts"] += 1
    for todo_list in box.get_all("/me/todo/lists"):
        if marked(todo_list.get("displayName")) and todo_list.get("wellknownListName") not in (
            "defaultList",
            "flaggedEmails",
        ):
            delete(f"/me/todo/lists/{todo_list['id']}")
            done["task_lists"] += 1
            continue
        for t in box.get_all(f"/me/todo/lists/{todo_list['id']}/tasks", {"$top": "100"}):
            if marked(t.get("title")):
                delete(f"/me/todo/lists/{todo_list['id']}/tasks/{t['id']}")
                done["tasks"] += 1
    return {"deleted": done, "kept": kept}


def check_leftovers(clean: bool = False) -> int:
    """Both accounts: print what is left, exit non-zero if anything is.

    The main account must be reachable. The pair is checked when it is mounted; if
    it cannot be reached, that is said, not counted: the send tier skips without it,
    so nothing can have been left there by a run that could not reach it.
    """
    from tests.live_guard import settings_from

    total = 0
    for role in ["main"] + (["pair"] if settings_from(os.environ).pair_config_dir else []):
        try:
            if clean:
                cleaned = run(role, "clean")
                print(f"{role}: deleted {cleaned['deleted']}")
                for item in cleaned["kept"]:
                    print(f"  kept, deleting it would send mail: {item}")
            answer = run(role, "leftovers")
        except MailboxError as exc:
            if role == "main":
                raise
            print(f"{role}: could not check — {exc}")
            continue
        total += answer["total"]
        print(f"{role}: {answer['total']} marked item(s) left")
        for area, items in answer["leftovers"].items():
            for item in items:
                print(f"  {area}: {item}")
    return 1 if total else 0


COMMANDS = {
    "whoami": cmd_whoami,
    "find": cmd_find,
    "purge": cmd_purge,
    "events": cmd_events,
    "remove-events": cmd_remove_events,
    "reply": cmd_reply,
    "accept": cmd_accept,
    "leftovers": cmd_leftovers,
    "clean": cmd_clean,
}


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv[:1] == ["check-leftovers"]:
        return check_leftovers(clean="--clean" in argv[1:])
    parser = argparse.ArgumentParser(prog="python -m tests.live_mailbox")
    parser.add_argument("--as", dest="role", choices=("main", "pair"), required=True)
    parser.add_argument("command", choices=sorted(COMMANDS))
    parser.add_argument("--marker", default="")
    parser.add_argument("--folder", action="append", choices=FOLDERS)
    parser.add_argument("--kind", help="a meetingMessageType, e.g. meetingRequest")
    parser.add_argument("--min", type=int, default=1)
    parser.add_argument("--wait", type=float, default=0)
    parser.add_argument("--comment", default="")
    parser.add_argument("--allow-send", action="store_true")
    args = parser.parse_args(argv)
    args.folder = args.folder or ["inbox", "junkemail"]
    print(json.dumps(COMMANDS[args.command](Mailbox(args.role), args)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
