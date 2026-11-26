"""A recorded Graph fake for the intent tests: one handler for the raw path and the SDK.

``FakeGraph`` answers every request an intent call makes, the SDK's included (a
real ``GraphServiceClient`` built over the same ``httpx.MockTransport``), and
records what was sent. It proves logic, not Graph behaviour: what Graph does with
a query is the live guards' job (``test_live_intent_queries.py``).

Counting follows the contracts: every top-level request is one Graph operation,
except a ``$batch`` envelope, which counts as its sub-requests.
"""

from __future__ import annotations

import asyncio
import json
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any
from urllib.parse import unquote

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"
APPLIED = 'IdType="ImmutableId"'


@dataclass
class Seen:
    """One request as Graph received it."""

    method: str
    path: str
    params: dict[str, str]
    headers: dict[str, str]
    body: Any = None
    in_batch: bool = False

    @property
    def search(self) -> str | None:
        return self.params.get("$search")

    @property
    def filter(self) -> str | None:
        return self.params.get("$filter")


Handler = Callable[[Seen, re.Match], Any]


@dataclass
class Route:
    method: str
    pattern: re.Pattern
    handler: Handler


@dataclass
class FakeGraph:
    """Routes by method and path; a handler returns a body, or ``(status, body)``."""

    routes: list[Route] = field(default_factory=list)
    seen: list[Seen] = field(default_factory=list)
    # What `Preference-Applied` says on id-bearing replies: the id family Graph granted.
    preference_applied: str | None = APPLIED
    # (predicate, seconds): replies to matching requests are held this long.
    delays: list[tuple[Callable[[Seen], bool], float]] = field(default_factory=list)

    def on(self, path: str, handler: Handler, method: str = "GET") -> None:
        self.routes.insert(0, Route(method, re.compile(rf"^{path}$", re.IGNORECASE), handler))

    # ── Counting ────────────────────────────────────────────────────────

    @property
    def operations(self) -> list[Seen]:
        """Every counted Graph operation: top-level requests and batch sub-requests."""
        return [s for s in self.seen if not (s.path.endswith("/$batch") and not s.in_batch)]

    @property
    def counted(self) -> int:
        return len(self.operations)

    def searches(self) -> list[str]:
        return [s.search for s in self.operations if s.search]

    def filters(self) -> list[str]:
        return [s.filter for s in self.operations if s.filter]

    # ── Transport ───────────────────────────────────────────────────────

    async def handle(self, request: httpx.Request) -> httpx.Response:
        path = unquote(request.url.path).removeprefix("/v1.0")
        params = dict(request.url.params)
        headers = {k.lower(): v for k, v in request.headers.items()}
        body = json.loads(request.content) if request.content else None
        if request.method == "POST" and path == "/$batch":
            self.seen.append(Seen("POST", path, params, headers, body))
            replies = [self._sub(sub) for sub in body["requests"]]
            await self._delay(s for s in self.seen[-len(replies) :])
            return httpx.Response(200, json={"responses": replies})
        await self._delay([Seen(request.method, path, params, headers, body)])
        status, payload, reply_headers = self._dispatch(
            Seen(request.method, path, params, headers, body)
        )
        return httpx.Response(status, json=payload, headers=reply_headers)

    def _sub(self, sub: dict) -> dict:
        url = httpx.URL(GRAPH + sub["url"])
        seen = Seen(
            sub["method"],
            unquote(url.path).removeprefix("/v1.0"),
            dict(url.params),
            {k.lower(): v for k, v in (sub.get("headers") or {}).items()},
            sub.get("body"),
            in_batch=True,
        )
        status, payload, headers = self._dispatch(seen)
        return {"id": sub["id"], "status": status, "headers": headers, "body": payload}

    def _dispatch(self, seen: Seen) -> tuple[int, Any, dict]:
        self.seen.append(seen)
        for route in self.routes:
            match = route.pattern.match(seen.path)
            if match and route.method == seen.method:
                answer = route.handler(seen, match)
                status, payload = answer if isinstance(answer, tuple) else (200, answer)
                headers = {}
                if (
                    200 <= status < 300
                    and self.preference_applied
                    and "immutableid" in seen.headers.get("prefer", "").lower()
                ):
                    headers["Preference-Applied"] = self.preference_applied
                return status, payload, headers
        return 404, _error("ErrorItemNotFound", f"no fake route for {seen.method} {seen.path}"), {}

    async def _delay(self, requests) -> None:
        """Hold the reply when a delay rule matches any request in it."""
        for seen in requests:
            for matches, seconds in self.delays:
                if matches(seen):
                    await asyncio.sleep(seconds)
                    return

    # ── Clients ─────────────────────────────────────────────────────────

    def http_client(self) -> httpx.AsyncClient:
        return httpx.AsyncClient(transport=httpx.MockTransport(self.handle))

    def graph_client(self) -> SimpleNamespace:
        """A ``GraphClient`` stand-in: a real SDK client and a credential, both fake."""
        return SimpleNamespace(sdk_client=self._sdk_client(), credential=FakeCredential())

    def _sdk_client(self):
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as gra
        from msgraph_core import GraphClientFactory

        http_client = GraphClientFactory.create_with_default_middleware(
            client=self.http_client(), options=getattr(gra, "options", None)
        )
        auth = AzureIdentityAuthenticationProvider(
            FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=gra.GraphRequestAdapter(auth, http_client))


class FakeCredential:
    def get_token(self, *args, **kwargs):
        from azure.core.credentials import AccessToken

        return AccessToken("not-a-real-token", 9_999_999_999)


def _error(code: str, message: str) -> dict:
    return {"error": {"code": code, "message": message}}


def error(status: int, code: str, message: str = "refused") -> tuple[int, dict]:
    return status, _error(code, message)


def page(items: list[dict], next_link: str | None = None) -> dict:
    body: dict = {"value": items}
    if next_link:
        body["@odata.nextLink"] = next_link
    return body


# ── Graph-shaped rows ─────────────────────────────────────────────────────

ME = "me@example.com"
# The Sent Items folder's id: a message there is one the owner really sent.
SENT_FOLDER = "sent-id"


def addr(address: str, name: str | None = None) -> dict:
    return {"emailAddress": {"address": address, "name": name if name is not None else address}}


def letter(
    to: list,
    *,
    sent: str = "2026-11-01T10:00:00Z",
    cc: list | None = None,
    bcc: list | None = None,
    sender: str = ME,
    odata_type: str | None = None,
    conversation: str = "conv-1",
) -> dict:
    """A Sent Items row as Q1 selects it.

    ``to``, ``cc`` and ``bcc`` take addresses or ``(address, name)`` pairs.
    """

    def people(entries):
        return [addr(*e) if isinstance(e, tuple) else addr(e) for e in entries or []]

    row = {
        "from": addr(sender, "Me"),
        "toRecipients": people(to),
        "ccRecipients": people(cc),
        "bccRecipients": people(bcc),
        "sentDateTime": sent,
        "conversationId": conversation,
        "parentFolderId": SENT_FOLDER,
    }
    if odata_type:
        row["@odata.type"] = odata_type
    return row


def meeting_response(to: str, *, sent: str = "2026-11-02T10:00:00Z") -> dict:
    """An accept sent to an organizer: a meeting message, not a letter."""
    return letter([to], sent=sent, odata_type="#microsoft.graph.eventMessageResponse")


# ── A mailbox behind the fake ─────────────────────────────────────────────

_RECEIVED = re.compile(r"^received(>=|<=)(\d{4}-\d{2}-\d{2})$", re.IGNORECASE)
_FILTER_FROM = re.compile(r"from/emailAddress/address eq '([^']*)'")
_FILTER_CONVERSATION = re.compile(r"conversationId eq '([^']*)'")
_FILTER_MESSAGE_ID = re.compile(r"^internetMessageId eq '(.*)'$")
_FILTER_CONTACT = re.compile(r"emailAddresses/any\(a:a/address eq '([^']*)'\)")


def kql_clauses(value: str) -> list[list[str]]:
    """A `$search` value as AND'd clauses, each a list of OR'd terms."""
    text = value[1:-1] if value.startswith('"') and value.endswith('"') else value
    clauses = [[t for t in g.split() if t != "OR"] for g in re.findall(r"\(([^)]*)\)", text)]
    clauses += [[t] for t in re.sub(r"\([^)]*\)", " ", text).split()]
    return clauses


def _people_text(row: dict, keys: tuple[str, ...]) -> str:
    found = []
    for key in keys:
        entries = row.get(key)
        entries = [entries] if isinstance(entries, dict) else entries or []
        for entry in entries:
            email = entry.get("emailAddress") or {}
            found += [str(email.get("address") or ""), str(email.get("name") or "")]
    return " ".join(found).casefold()


def kql_term(term: str, row: dict) -> bool:
    """A loose stand-in for Graph's KQL: enough to route the queries the resolver sends."""
    received = _RECEIVED.match(term)
    if received:
        day = (row.get("receivedDateTime") or "")[:10]
        op, bound = received.groups()
        return bool(day) and (day >= bound if op == ">=" else day <= bound)
    prop, sep, value = term.partition(":")
    value = value.casefold()
    if sep and prop.lower() == "from":
        return value in _people_text(row, ("from",))
    if sep and prop.lower() == "to":
        return value in _people_text(row, ("toRecipients",))
    if sep and prop.lower() == "recipients":
        return value in _people_text(row, ("toRecipients", "ccRecipients", "bccRecipients"))
    if sep and prop.lower() == "hasattachments":
        return bool(row.get("hasAttachments")) is (value == "true")
    if sep and prop.lower() == "subject":
        return value in str(row.get("subject") or "").casefold()
    if sep and prop.lower() == "attachment":
        return any(value in name.casefold() for name in row.get("_attachment_names", []))
    text = " ".join(
        [
            str(row.get("subject") or ""),
            str(row.get("bodyPreview") or ""),
            _people_text(row, ("from",)),
        ]
    ).casefold()
    return term.casefold() in text


def kql_matches(value: str, row: dict) -> bool:
    return all(any(kql_term(t, row) for t in clause) for clause in kql_clauses(value))


def _select(row: dict, seen: Seen, headers_on_lists: bool = True) -> dict:
    out = dict(row)
    selected = (seen.params.get("$select") or "").split(",")
    if "internetMessageHeaders" not in selected or not headers_on_lists:
        out.pop("internetMessageHeaders", None)
    return out


def _paged(items: list[dict], seen: Seen, path: str, size: int | None = None) -> dict:
    top = int(seen.params.get("$top") or size or 10)
    start = int(seen.params.get("$skiptoken") or 0)
    chunk = items[start : start + top]
    more = start + top < len(items)
    return page(chunk, f"{GRAPH}{path}?$skiptoken={start + top}&$top={top}" if more else None)


@dataclass
class Mailbox:
    """A small mailbox the fake answers from. Every list here is Graph-shaped JSON."""

    sent: list[dict] = field(default_factory=list)
    received: list[dict] = field(default_factory=list)
    contacts: list[dict] = field(default_factory=list)
    calendars: list[dict] = field(default_factory=list)
    events: dict[str, list[dict]] = field(default_factory=dict)
    task_lists: list[dict] = field(default_factory=list)
    tasks: dict[str, list[dict]] = field(default_factory=dict)
    rules: Any = field(default_factory=list)  # a list, or (status, body) to refuse
    folders: list[dict] = field(default_factory=list)
    child_folders: dict[str, list[dict]] = field(default_factory=dict)
    attachments: dict[str, list[dict]] = field(default_factory=dict)
    # P-F5: whether `internetMessageHeaders` comes back on a list's `$select`.
    headers_on_lists: bool = True
    # P-H1a: whether To Do takes `$filter=status ne 'completed'`.
    task_filter_supported: bool = True
    # P-AT1: whether a message search takes KQL `attachment:`.
    attachment_search_supported: bool = True
    # P-SI1: whether a list's $select takes the eventMessage cast.
    cast_select_supported: bool = True

    def install(self, fake: FakeGraph) -> FakeGraph:
        on = fake.on
        on("/me/messages", self._messages)
        on(r"/me/mailFolders/([^/]+)/messages", self._folder_messages)
        on(r"/me/messages/([^/]+)", self._message)
        on(r"/me/messages/([^/]+)/attachments", self._attachments)
        on("/me/contacts", self._contacts)
        on("/me/calendars", lambda seen, m: _paged(self.calendars, seen, "/me/calendars", 100))
        on(r"/me/calendars/([^/]+)/calendarView", self._calendar_view)
        on(r"/me/events/([^/]+)", self._event)
        on("/me/todo/lists", lambda seen, m: page(self.task_lists))
        on(r"/me/todo/lists/([^/]+)/tasks", self._tasks)
        on(r"/me/todo/lists/([^/]+)/tasks/([^/]+)", self._task)
        on("/me/mailFolders/inbox/messageRules", self._rules)
        on("/me/mailFolders", lambda seen, m: page(self.folders))
        on(
            r"/me/mailFolders/([^/]+)/childFolders",
            lambda seen, m: page(self.child_folders.get(m.group(1), [])),
        )
        return fake

    # Mail

    def _sent_items(self, seen: Seen, match) -> Any:
        if "microsoft.graph.eventMessage" in (seen.params.get("$select") or "") and not (
            self.cast_select_supported
        ):
            return error(400, "BadRequest", "Could not find a property named 'eventMessage'")
        if seen.search:
            hits = [r for r in self.sent if kql_matches(seen.search, r)]
            return page(hits[: int(seen.params.get("$top") or 25)])
        newest = sorted(self.sent, key=lambda r: r.get("sentDateTime") or "", reverse=True)
        return _paged(newest, seen, "/me/mailFolders/sentitems/messages", 100)

    def _messages(self, seen: Seen, match) -> Any:
        rows = self.received
        if seen.search:
            if "attachment:" in seen.search and not self.attachment_search_supported:
                return error(400, "BadRequest", "Invalid search query")
            hits = [r for r in rows if kql_matches(seen.search, self._named(r))]
            hits.sort(key=lambda r: r.get("receivedDateTime") or "", reverse=True)
            return page([_select(r, seen) for r in hits[: int(seen.params.get("$top") or 10)]])
        wanted = seen.filter or ""
        if (by_id := _FILTER_MESSAGE_ID.match(wanted)) is not None:
            literal = by_id.group(1).replace("''", "'")
            return page([_select(r, seen) for r in rows if r.get("internetMessageId") == literal])
        if (by_conversation := _FILTER_CONVERSATION.search(wanted)) is not None:
            hits = [r for r in rows if r.get("conversationId") == by_conversation.group(1)]
        elif (by_sender := _FILTER_FROM.search(wanted)) is not None:
            address = by_sender.group(1).replace("''", "'").casefold()
            hits = [r for r in rows if _people_text(r, ("from",)).split(" ")[0] == address]
        else:
            hits = list(rows)
        descending = "desc" in (seen.params.get("$orderby") or "desc")
        hits.sort(key=lambda r: r.get("receivedDateTime") or "", reverse=descending)
        top = int(seen.params.get("$top") or 10)
        return page([_select(r, seen, self.headers_on_lists) for r in hits[:top]])

    def _named(self, row: dict) -> dict:
        """The row with its non-inline attachments' names, for KQL `attachment:`."""
        names = [
            str(a.get("name") or "")
            for a in self.attachments.get(row.get("id"), [])
            if not a.get("isInline")
        ]
        return {**row, "_attachment_names": names}

    def _folder_messages(self, seen: Seen, match) -> Any:
        folder = match.group(1)
        if folder == "sentitems":
            return self._sent_items(seen, match)
        if folder == "drafts":
            rows = [r for r in self.received if r.get("isDraft")]
        else:
            rows = [r for r in self.received if r.get("parentFolderId") == folder]
        hits = [r for r in rows if not seen.search or kql_matches(seen.search, r)]
        return page([_select(r, seen) for r in hits[: int(seen.params.get("$top") or 10)]])

    def _message(self, seen: Seen, match) -> Any:
        for row in self.received:
            if row["id"] == match.group(1):
                return _select(row, seen)
        return error(404, "ErrorItemNotFound", "The specified object was not found in the store.")

    def _attachments(self, seen: Seen, match) -> Any:
        return page(self.attachments.get(match.group(1), []))

    # Contacts

    def _contacts(self, seen: Seen, match) -> Any:
        if seen.search:
            term = seen.search.strip('"').casefold()
            hits = [
                c
                for c in self.contacts
                if any(
                    w.startswith(term) or term in w
                    for w in " ".join(
                        [str(c.get("displayName") or "")]
                        + [str(e.get("address") or "") for e in c.get("emailAddresses") or []]
                    )
                    .casefold()
                    .split()
                )
                or term in str(c.get("displayName") or "").casefold()
            ]
            return page(hits[: int(seen.params.get("$top") or 10)])
        if (exact := _FILTER_CONTACT.search(seen.filter or "")) is not None:
            address = exact.group(1).casefold()
            return page(
                [
                    c
                    for c in self.contacts
                    if any(
                        str(e.get("address") or "").casefold() == address
                        for e in c.get("emailAddresses") or []
                    )
                ]
            )
        return page(self.contacts)

    # Calendar

    def _calendar_view(self, seen: Seen, match) -> Any:
        start = seen.params.get("startDateTime", "")
        end = seen.params.get("endDateTime", "")
        events = [
            e
            for e in self.events.get(match.group(1), [])
            if e["start"]["dateTime"][:19] < end[:19] and e["end"]["dateTime"][:19] > start[:19]
        ]
        return _paged(events, seen, f"/me/calendars/{match.group(1)}/calendarView", 10)

    def _event(self, seen: Seen, match) -> Any:
        for events in self.events.values():
            for event in events:
                if event["id"] == match.group(1):
                    return event
        return error(404, "ErrorItemNotFound")

    # To Do and rules

    def _tasks(self, seen: Seen, match) -> Any:
        tasks = self.tasks.get(match.group(1), [])
        if seen.filter:
            if not self.task_filter_supported:
                return error(400, "BadRequest", "Invalid filter clause")
            tasks = [t for t in tasks if t.get("status") != "completed"]
        return page(tasks[: int(seen.params.get("$top") or 10)])

    def _task(self, seen: Seen, match) -> Any:
        for task in self.tasks.get(match.group(1), []):
            if task["id"] == match.group(2):
                return task
        return error(404, "ErrorItemNotFound")

    def _rules(self, seen: Seen, match) -> Any:
        return self.rules if isinstance(self.rules, tuple) else page(self.rules)


def received(
    msg_id: str,
    sender: str,
    subject: str,
    *,
    name: str | None = None,
    when: str = "2026-11-20T09:00:00Z",
    to: list[str] | None = None,
    cc: list[str] | None = None,
    conversation: str | None = None,
    folder: str = "inbox-id",
    attachments: bool = False,
    headers: list[dict] | None = None,
    draft: bool = False,
    internet_message_id: str | None = None,
) -> dict:
    """A message as the mail queries select it."""
    return {
        "id": msg_id,
        "subject": subject,
        "bodyPreview": f"preview of {subject}",
        "from": addr(sender, name or sender.split("@")[0].title()),
        "toRecipients": [addr(a) for a in (to if to is not None else [ME])],
        "ccRecipients": [addr(a) for a in cc or []],
        "receivedDateTime": when,
        "isRead": False,
        "importance": "normal",
        "hasAttachments": attachments,
        "categories": [],
        "flag": {"flagStatus": "notFlagged"},
        "conversationId": conversation or f"conv-{msg_id}",
        "inferenceClassification": "focused",
        "parentFolderId": folder,
        "isDraft": draft,
        "internetMessageId": internet_message_id or f"<{msg_id}@mail.example.com>",
        "internetMessageHeaders": headers if headers is not None else [],
    }


def auth_header(verdict: str, domain: str) -> list[dict]:
    """Microsoft's Authentication-Results stamp for a From domain."""
    return [
        {
            "name": "Authentication-Results",
            "value": f"spf=pass smtp.mailfrom={domain}; dkim=pass header.d={domain}; "
            f"dmarc={verdict} action=none header.from={domain};compauth={verdict} reason=100",
        }
    ]


def contact(contact_id: str, name: str, *addresses: str) -> dict:
    given, _, surname = name.partition(" ")
    return {
        "id": contact_id,
        "displayName": name,
        "givenName": given,
        "surname": surname,
        "emailAddresses": [{"address": a, "name": name} for a in addresses],
    }


def event(
    event_id: str,
    subject: str,
    start: str,
    end: str,
    *,
    organizer: str = ME,
    organizer_name: str = "Me",
    is_organizer: bool = True,
    response: str = "organizer",
    attendees: list[tuple[str, str]] | None = None,
    show_as: str = "busy",
    all_day: bool = False,
) -> dict:
    return {
        "id": event_id,
        "subject": subject,
        "start": {"dateTime": start, "timeZone": "UTC"},
        "end": {"dateTime": end, "timeZone": "UTC"},
        "isAllDay": all_day,
        "isOrganizer": is_organizer,
        "organizer": addr(organizer, organizer_name),
        "responseStatus": {"response": response},
        "isCancelled": False,
        "showAs": show_as,
        "seriesMasterId": None,
        "type": "singleInstance",
        "attendees": [
            {"emailAddress": {"address": a, "name": n}, "type": "required"}
            for a, n in attendees or []
        ],
    }
