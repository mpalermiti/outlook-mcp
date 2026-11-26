"""Guards for the tiers that reach a real mailbox: whose mailbox, and who may get mail.

The live tiers run against a pair of throwaway test accounts the project owns,
signed in only inside the container ``scripts/live-tests.sh`` starts. Nothing in
this repository names them: their addresses reach the run through the
environment, from a file kept outside the repository.

Three decisions, each pure (no network, no sign-in) so ``tests/test_live_guards.py``
can pin them offline; ``tests/conftest.py`` wires them in:

1. **May this test run at all?** (:func:`refusal`) Decided before any sign-in is
   touched. A live test needs the test accounts configured, or — for the read-only
   ``live`` tier only — the owner's explicit opt-in to run it against their own
   mailbox (``OUTLOOK_MCP_LIVE_OWN_MAILBOX=1``).
2. **Is the signed-in account one of them?** (:func:`account_refusal`) Decided from
   ``/me``, which must answer, and the saved sign-in record, which must agree with it.
   If ``/me`` fails, the test refuses: the record alone is never enough.
3. **May this request go out?** (:class:`ReadOnly`, :class:`NeverSends`,
   :class:`SendGuard`) While a live test runs, every HTTP request the process makes,
   through the SDK, a tool's own client or a raw ``httpx``/``requests`` call, passes
   through :func:`guarding` first. ``live`` and ``integration`` may only read.
   ``live_write`` may change the mailbox but not send: no ``sendMail`` or send,
   reply, forward or RSVP action, no invitation, no change to a meeting with
   attendees, no forwarding rule, automatic reply or calendar sharing.
   ``live_send`` may send, but each recipient must be the paired test account,
   checked on the request itself before it leaves, and each send is counted
   against the run's budget.
"""

from __future__ import annotations

import contextlib
import json
import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from urllib.parse import quote, unquote, urlsplit

import pytest

# ── Settings, from the environment ────────────────────────────────────────

ACCOUNT_ENV = "OUTLOOK_MCP_TEST_ACCOUNT"
PAIR_ENV = "OUTLOOK_MCP_TEST_PAIR"
PAIR_CONFIG_DIR_ENV = "OUTLOOK_MCP_TEST_PAIR_CONFIG_DIR"
PAIR_CAN_SEND_ENV = "OUTLOOK_MCP_TEST_PAIR_CAN_SEND"
OWN_MAILBOX_ENV = "OUTLOOK_MCP_LIVE_OWN_MAILBOX"
SEND_ENV = "OUTLOOK_MCP_LIVE_SEND"
SEND_BUDGET_ENV = "OUTLOOK_MCP_LIVE_SEND_BUDGET"

# The most real sends one run may make. New Outlook.com accounts have very low
# sending limits, and one that sends too much too soon is locked.
MAX_SEND_BUDGET = 10

# Every tier that reaches a real mailbox. tests/conftest.py's offline isolation
# leaves exactly these alone.
LIVE_MARKERS = ("integration", "live", "live_write", "live_send")

# What the owner's own-mailbox opt-in may run: the read-only query-shape tier.
OWN_MAILBOX_TIERS = frozenset({"live"})

# Subjects and names carry these, so anything a crash leaks is easy to find, and
# so the clean-up helpers can never touch an item a test did not create.
LIVE_WRITE_SUBJECT = "[outlook-mcp live-write guard] safe to delete"
LIVE_WRITE_NAME = "LiveWriteGuard-SafeToDelete"
LIVE_SEND_SUBJECT = "[outlook-mcp live-send guard] safe to delete"
MARKERS = (LIVE_WRITE_SUBJECT, LIVE_WRITE_NAME, LIVE_SEND_SUBJECT)


def _address(value: str | None) -> str | None:
    value = (value or "").strip().casefold()
    return value or None


def _on(environ: Mapping[str, str], name: str) -> bool:
    return environ.get(name, "").strip() == "1"


@dataclass(frozen=True)
class LiveSettings:
    """What the environment says about the run. Addresses are casefolded."""

    account: str | None = None
    pair: str | None = None
    own_mailbox: bool = False
    send: bool = False
    send_budget: int | None = None
    pair_config_dir: str | None = None
    pair_can_send: bool = False

    @property
    def test_accounts(self) -> frozenset[str]:
        return frozenset(a for a in (self.account, self.pair) if a)


def settings_from(environ: Mapping[str, str]) -> LiveSettings:
    budget_text = environ.get(SEND_BUDGET_ENV, "").strip()
    try:
        budget = int(budget_text) if budget_text else None
    except ValueError:
        budget = None
    if budget is not None and not 0 <= budget <= MAX_SEND_BUDGET:
        budget = None
    return LiveSettings(
        account=_address(environ.get(ACCOUNT_ENV)),
        pair=_address(environ.get(PAIR_ENV)),
        own_mailbox=_on(environ, OWN_MAILBOX_ENV),
        send=_on(environ, SEND_ENV),
        send_budget=budget,
        pair_config_dir=environ.get(PAIR_CONFIG_DIR_ENV, "").strip() or None,
        pair_can_send=_on(environ, PAIR_CAN_SEND_ENV),
    )


# ── 1 and 2: whose mailbox ────────────────────────────────────────────────

_HOW_TO_RUN = (
    "Live tiers run against the project's test accounts, inside the container "
    "`scripts/live-tests.sh <tier>` starts (RELEASING.md). To run the read-only `live` "
    f"tier against your own mailbox as a smoke check, set {OWN_MAILBOX_ENV}=1."
)


def live_tiers_of(item) -> frozenset[str]:
    """The live markers a test carries, however they were applied."""
    return frozenset(m for m in LIVE_MARKERS if item.get_closest_marker(m))


def own_mailbox_allows(tiers: frozenset[str], settings: LiveSettings) -> bool:
    return settings.own_mailbox and bool(tiers) and tiers <= OWN_MAILBOX_TIERS


def refusal(tiers: frozenset[str], settings: LiveSettings) -> str | None:
    """Why a test in these tiers may not run at all — before any sign-in is read.

    None when it may go on to the account check (or, under the own-mailbox opt-in,
    straight to the test).
    """
    if not tiers:
        return None
    if "live_send" in tiers and not (settings.account and settings.pair):
        return (
            f"Refusing to run: the send tier needs both test accounts named ({ACCOUNT_ENV} "
            f"and {PAIR_ENV}); it mails only between them. " + _HOW_TO_RUN
        )
    if settings.test_accounts:
        return None
    if own_mailbox_allows(tiers, settings):
        return None
    if settings.own_mailbox:
        return (
            f"Refusing to run: {OWN_MAILBOX_ENV}=1 covers only the read-only `live` tier, "
            f"and this test is in {sorted(tiers)}. " + _HOW_TO_RUN
        )
    return f"Refusing to run: no test account is named ({ACCOUNT_ENV}). " + _HOW_TO_RUN


def identity_of(me: Mapping | None, record_username: str | None) -> str | None:
    """The signed-in address, or None: when ``/me`` named no one, or the sources disagree.

    A personal account reports its address as ``userPrincipalName``; ``mail`` is
    often null. ``/me`` must name the account: it is what the token can reach, and
    a failed or empty answer (pass ``{}`` or None) means unknown, never "trust the
    record". The saved sign-in record, which names the account the token cache
    belongs to, must agree when present.
    """
    me = me if isinstance(me, Mapping) else {}
    from_me = {
        value
        for value in (_address(me.get("userPrincipalName")), _address(me.get("mail")))
        if value
    }
    if len(from_me) != 1:
        return None
    who = from_me.pop()
    record = _address(record_username)
    return who if record is None or record == who else None


def account_refusal(identity: str | None, tiers: frozenset[str], settings: LiveSettings) -> str:
    """Why the signed-in account may not run these tiers; "" when it may."""
    if not tiers or own_mailbox_allows(tiers, settings):
        return ""
    if identity is None:
        return (
            "Refusing to run: could not tell which account is signed in (/me did not "
            "answer with one, or it disagrees with the saved sign-in record)."
        )
    if "live_send" in tiers and identity != settings.account:
        return (
            f"Refusing to run: the send tier runs only as the main test account ({ACCOUNT_ENV}); "
            "the signed-in account is not it."
        )
    if identity in settings.test_accounts:
        return ""
    # The address is left out on purpose: this text lands in logs.
    return (
        f"Refusing to run: the signed-in account is not one of the test accounts "
        f"({ACCOUNT_ENV}, {PAIR_ENV}). " + _HOW_TO_RUN
    )


# ── 3: what may leave ─────────────────────────────────────────────────────


class Refused(pytest.fail.Exception):
    """A request the live guards stopped before it left.

    A pytest failure, which is a BaseException, so a tool's ``except Exception``
    cannot turn it into "the message may have been sent".
    """


class SendRefused(Refused):
    """The request would have sent mail it may not send."""


class WriteRefused(Refused):
    """A read-only tier tried to change something."""


@dataclass(frozen=True)
class Send:
    """A request that makes Exchange send mail.

    ``recipients`` is None when they cannot be read off the request itself (a reply
    goes to whoever wrote the original), so nothing can vouch for them.
    """

    what: str
    recipients: tuple[str, ...] | None


# Actions on a message or event whose POST sends mail to people the request does
# not name: the draft's recipients, the original's sender, the meeting's organizer
# or attendees.
_SENDING_ACTIONS = {
    "send": "send a draft",
    "reply": "reply",
    "replyall": "reply all",
    "forward": "forward",
    "cancel": "cancel a meeting",
    "accept": "accept an invitation",
    "decline": "decline an invitation",
    "tentativelyaccept": "tentatively accept an invitation",
}
_RESPONSES = {"accept", "decline", "tentativelyaccept"}
# A rule that forwards or redirects sends every matching message on, for good.
_FORWARDING_ACTIONS = ("forwardto", "redirectto", "forwardasattachmentto")

# Token requests go here, as POSTs; they are not Graph and are never mail.
SIGN_IN_HOSTS = frozenset(
    {"login.microsoftonline.com", "login.live.com", "login.windows.net", "login.microsoft.com"}
)
_READS = frozenset({"GET", "HEAD"})

# Graph accepts an action by its qualified name (`microsoft.graph.sendMail`) and a
# key in parentheses (`events('id')`); both are read as the plain form.
_QUALIFIER = "microsoft.graph."
_KEYED = re.compile(r"^([^()]+)\((.*)\)$")


def _folded(value):
    """``value`` with every object key casefolded, all the way down."""
    if isinstance(value, Mapping):
        return {str(k).casefold(): _folded(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_folded(v) for v in value]
    return value


def _segment(raw: str) -> list[str]:
    part = unquote(raw)
    keyed = _KEYED.match(part)
    if keyed:
        name, key = keyed.group(1), keyed.group(2)
        if len(key) >= 2 and key[0] == key[-1] == "'":
            key = key[1:-1]
        found = [name] + ([key] if key else [])
    else:
        found = [part]
    return [p[len(_QUALIFIER) :] if p.lower().startswith(_QUALIFIER) else p for p in found]


def _segments(url: str) -> list[str]:
    parts = [p for raw in urlsplit(url).path.split("/") if raw for p in _segment(raw)]
    if parts and parts[0].lower() in ("v1.0", "beta"):
        parts = parts[1:]
    return parts


def _host(url: str) -> str:
    return (urlsplit(url).hostname or "").lower()


def _reads_only(url: str, parts: list[str]) -> bool:
    """A POST that reads and changes nothing: ``/me/translateExchangeIds`` on Graph.

    It turns ids from one format into another. ``outlook_find`` needs it because a
    ``$search`` hands back regular ids while lists hand back immutable ones (live
    round 1, 2026-11-26). Only that path, on Graph, through ``/me``.
    """
    return _host(url) == "graph.microsoft.com" and [p.lower() for p in parts] == [
        "me",
        "translateexchangeids",
    ]


def _addresses(entries) -> tuple[str, ...]:
    found = []
    for entry in entries or []:
        email = entry.get("emailaddress") if isinstance(entry, Mapping) else None
        address = email.get("address") if isinstance(email, Mapping) else None
        found.append(_address(address if isinstance(address, str) else None) or "")
    return tuple(found)


def _event_id(parts: list[str]) -> str | None:
    """The id after the last ``events`` segment, if the path names one event."""
    for i in range(len(parts) - 1, -1, -1):
        if parts[i].lower() == "events":
            return parts[i + 1] if i + 1 < len(parts) else None
    return None


def wants_body(method: str, url: str) -> bool:
    """Whether :func:`classify` needs this request's body to decide."""
    if method.upper() not in ("POST", "PATCH", "PUT"):
        return False
    names = [p.lower() for p in _segments(url)]
    return bool(names) and (
        names[-1] in ("sendmail", "$batch", "events", "mailboxsettings")
        or "events" in names
        or "messagerules" in names
        or names[-1] in _RESPONSES
    )


def classify(
    method: str,
    url: str,
    body: Mapping | None,
    meetings: Mapping[str, tuple[str, ...]] | None = None,
) -> Send | None:
    """The mail this request makes Exchange send, or None when it sends none.

    ``meetings`` are meetings that have attendees, by id: ones this run created,
    or ones the guard looked up before letting a change through. Changing or
    deleting one sends them an update or a cancellation.
    """
    method = method.upper()
    meetings = meetings or {}
    parts = _segments(url)
    if not parts:
        return None
    names = [p.lower() for p in parts]
    last = names[-1]
    # Graph reads property names without regard to case, and the SDK writes some
    # capitalised ({"Message": ..., "SaveToSentItems": ...}): compare them folded.
    body = _folded(body) if isinstance(body, Mapping) else {}

    if method == "POST" and last == "$batch":
        found = [
            classify(
                r.get("method", ""),
                "/" + str(r.get("url", "")).lstrip("/"),
                r.get("body"),
                meetings,
            )
            for r in body.get("requests") or []
            if isinstance(r, Mapping)
        ]
        found = [s for s in found if s]
        if not found:
            return None
        if any(s.recipients is None for s in found):
            return Send("; ".join(s.what for s in found), None)
        return Send("; ".join(s.what for s in found), tuple(a for s in found for a in s.recipients))

    if method == "POST" and last == "sendmail":
        message = body.get("message") if isinstance(body.get("message"), Mapping) else {}
        recipients = sum(
            (_addresses(message.get(k)) for k in ("torecipients", "ccrecipients", "bccrecipients")),
            (),
        )
        # Not recipients, but a reply to the message goes there: held to the same rule.
        recipients += _addresses(message.get("replyto"))
        return Send("sendMail", recipients)

    if method in ("POST", "PATCH", "PUT") and "messagerules" in names:
        actions = body.get("actions") if isinstance(body.get("actions"), Mapping) else {}
        if any(actions.get(k) for k in _FORWARDING_ACTIONS):
            return Send("a rule that forwards or redirects mail", None)
    if method in ("PATCH", "PUT", "POST") and last == "mailboxsettings":
        if "automaticrepliessetting" in body:
            return Send("automatic replies", None)
    if method in ("POST", "PATCH") and "calendarpermissions" in names:
        return Send("a calendar sharing invitation", None)

    event_id = _event_id(parts)
    if method == "POST" and last in _SENDING_ACTIONS:
        if last in _RESPONSES and body.get("sendresponse") is False:
            return None  # answered without telling the organizer
        if last == "cancel" and event_id in meetings:
            return Send(_SENDING_ACTIONS[last], meetings[event_id])
        return Send(_SENDING_ACTIONS[last], None)

    if method == "POST" and last == "events":
        attendees = _addresses(body.get("attendees"))
        return Send("an invitation", attendees) if attendees else None

    if event_id is None:
        return None
    if method == "PATCH":
        attendees = _addresses(body.get("attendees"))
        if attendees or event_id in meetings:
            return Send("a meeting update", attendees + meetings.get(event_id, ()))
        return None
    if method == "DELETE" and parts[-1] == event_id and event_id in meetings:
        return Send("a meeting cancellation", meetings[event_id])
    return None


def event_write_target(method: str, url: str) -> str | None:
    """The event a request changes or removes (a PATCH, DELETE or cancel of the
    event itself, not of its attachments), else None."""
    parts = _segments(url)
    event_id = _event_id(parts)
    if event_id is None:
        return None
    method = method.upper()
    if method in ("PATCH", "DELETE") and parts[-1] == event_id:
        return event_id
    if method == "POST" and parts[-1].lower() == "cancel":
        return event_id
    return None


def attendees_url(url: str, event_id: str) -> str:
    """Where to read the attendees of the event ``url`` changes."""
    split = urlsplit(url)
    raw = [r for r in split.path.split("/") if r]
    for i in range(len(raw) - 1, -1, -1):
        name = unquote(raw[i]).split("(", 1)[0].lower()
        if name.removeprefix(_QUALIFIER) == "events":
            path = "/".join(raw[:i] + ["events", quote(event_id, safe="")])
            break
    else:  # pragma: no cover - event_write_target found an events segment
        path = "/".join(raw)
    return f"{split.scheme}://{split.netloc}/{path}?$select=attendees"


class ReadOnly:
    """``live`` and ``integration``, the own-mailbox smoke check included: reads only.

    Anything but a GET (or a ``$batch`` of GETs) to anything but the sign-in
    service is refused before it leaves.
    """

    def __init__(self, tier: str = "this tier") -> None:
        self.tier = tier

    def lookup(self, method: str, url: str) -> str | None:
        return None

    def before(self, method: str, url: str, body: Mapping | None, existing=None) -> None:
        method = method.upper()
        if method in _READS or _host(url) in SIGN_IN_HOSTS:
            return
        parts = _segments(url)
        if method == "POST" and _reads_only(url, parts):
            return
        if method == "POST" and parts and parts[-1] == "$batch":
            requests = (_folded(body) or {}).get("requests") if isinstance(body, Mapping) else None
            if requests and all(
                isinstance(r, Mapping) and str(r.get("method", "")).upper() in _READS
                for r in requests
            ):
                return
        raise WriteRefused(
            f"Refused before it left: {self.tier} is read-only, and this is a {method}. "
            "A test that changes the mailbox belongs in live_write."
        )

    def created(self, event_id: str, attendees: tuple[str, ...]) -> None:
        """Nothing is created in a read-only tier."""


class NeverSends:
    """``live_write``: it may change the mailbox, but no request may send mail.

    An event it did not create this run may be someone's meeting, where a change
    or delete would mail the attendees: before one goes, :func:`guarding` reads
    the event's attendees, and the change is refused if it has any.
    """

    meetings: Mapping[str, tuple[str, ...]] = {}

    def __init__(self, tier: str = "this tier") -> None:
        self.tier = tier
        self.known: set[str] = set()  # attendee-free events this run created

    def lookup(self, method: str, url: str) -> str | None:
        target = event_write_target(method, url)
        return target if target is not None and target not in self.known else None

    def before(self, method: str, url: str, body: Mapping | None, existing=None) -> None:
        if _host(url) in SIGN_IN_HOSTS:
            return
        sent = classify(method, url, body, existing)
        if sent:
            raise SendRefused(
                f"Refused before it left: {self.tier} may not send mail, and this request "
                f"would ({sent.what}). Sending tests belong in the live_send tier."
            )

    def created(self, event_id: str, attendees: tuple[str, ...]) -> None:
        self.known.add(event_id)


@dataclass
class SendGuard:
    """The ``live_send`` tier's rule: mail goes only to the pair, and is counted.

    A test reserves the sends it will make before making any (:meth:`reserve`), so
    a test that cannot finish — an invitation it could not then cancel — never
    starts. Every send draws on the reservation, after each recipient is checked.
    """

    pair: str
    budget: int
    used: int = 0
    reserved: int = 0
    meetings: dict[str, tuple[str, ...]] = field(default_factory=dict)
    known: set[str] = field(default_factory=set)
    log: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        pair = _address(self.pair)
        if not pair or "@" not in pair:
            raise ValueError("SendGuard needs the pair account's address")
        if not 0 <= self.budget <= MAX_SEND_BUDGET:
            raise ValueError(f"the send budget must be 0 to {MAX_SEND_BUDGET}")
        self.pair = pair

    def reserve(self, sends: int) -> str:
        """Hold ``sends`` for the calling test; the reason it cannot, else ""."""
        if self.reserved:
            return "a send reservation is already open; release it first"
        left = self.budget - self.used
        if sends > left:
            return (
                f"Skipped to stay within this run's send budget: the test needs {sends} "
                f"send(s), {left} of {self.budget} left (scripts/live-tests.sh --sends N)."
            )
        self.reserved = sends
        return ""

    def release(self) -> None:
        self.reserved = 0

    def authorize(self, what: str, recipients: Iterable[str] | None, *, to: str | None = None):
        """Check every recipient, then count one send — or refuse before it goes."""
        allowed = _address(to) or self.pair
        if recipients is None:
            raise SendRefused(
                f"Refused before it left: {what} sends to recipients the request does not "
                "name, so they cannot be checked."
            )
        recipients = tuple(_address(r) or "" for r in recipients)
        if not recipients:
            raise SendRefused(f"Refused before it left: {what} names no recipient.")
        strangers = [r for r in recipients if not r or r != allowed]
        if strangers:
            raise SendRefused(
                f"Refused before it left: {what} is addressed to {len(strangers)} "
                "recipient(s) other than the paired test account."
            )
        if self.reserved <= 0:
            raise SendRefused(
                f"Refused before it left: {what} is a send the test did not reserve "
                f"({self.used} of {self.budget} used this run)."
            )
        self.reserved -= 1
        self.used += 1
        self.log.append(what)

    def lookup(self, method: str, url: str) -> str | None:
        return None  # an event it did not create is refused outright

    def before(self, method: str, url: str, body: Mapping | None, existing=None) -> None:
        if _host(url) in SIGN_IN_HOSTS:
            return
        sent = classify(method, url, body, self.meetings)
        if sent is None:
            target = event_write_target(method, url)
            if target is not None and target not in self.meetings and target not in self.known:
                raise SendRefused(
                    "Refused before it left: the send tier changes or deletes only events it "
                    "created this run, and this is not one of them: it could be someone's "
                    "meeting, and the change would mail its attendees."
                )
            return
        self.authorize(sent.what, sent.recipients)

    def created(self, event_id: str, attendees: tuple[str, ...]) -> None:
        if attendees:
            self.meetings[event_id] = attendees
        else:
            self.known.add(event_id)


# ── Holding every request to the guard ────────────────────────────────────


def _json(content) -> Mapping | None:
    if not content:
        return None
    try:
        value = json.loads(content)
    except (ValueError, UnicodeDecodeError):
        return {"<unreadable body>": True}
    return value if isinstance(value, Mapping) else {"<not an object>": True}


def _attendees_found(event_id: str, status: int, content) -> dict[str, tuple[str, ...]]:
    """What a lookup of the event's attendees says, as ``meetings`` for classify."""
    if status == 404:
        return {}  # gone: there is no one to tell
    if not 200 <= status < 300:
        raise SendRefused(
            f"Refused before it left: could not read the attendees of the event it would "
            f"change (HTTP {status}), so it cannot tell whether the change would mail them."
        )
    attendees = _addresses((_folded(_json(content)) or {}).get("attendees"))
    return {event_id: attendees} if attendees else {}


def _created_event(method: str, url: str, body: Mapping | None, status: int):
    """The attendees of an event this request created, or None if it created none."""
    parts = _segments(url)
    if method.upper() != "POST" or not parts or parts[-1].lower() != "events":
        return None
    if not 200 <= status < 300:
        return None
    return _addresses((_folded(body) or {}).get("attendees") if isinstance(body, Mapping) else None)


def _remember(guard, content, attendees) -> None:
    created = _json(content)
    if isinstance(created, Mapping) and isinstance(created.get("id"), str):
        guard.created(created["id"], attendees)


def _authorization(headers) -> dict[str, str]:
    value = headers.get("Authorization") if headers is not None else None
    return {"Authorization": value} if value else {}


@contextlib.contextmanager
def guarding(guard):
    """Hold every HTTP request this process makes to ``guard`` before it leaves.

    Patches the send methods of ``httpx.AsyncClient``, ``httpx.Client`` and
    ``requests.Session``, so it covers the SDK's client, the clients the tools
    open themselves (``$batch``, delta, MIME, upload sessions) and any raw call a
    test makes. A child process is not covered: ``tests/live_mailbox.py``, the
    only one the tiers start, keeps its own rules. A created event is remembered,
    so a later change or delete of it is judged for what it sends.
    """
    import httpx

    patch = pytest.MonkeyPatch()
    async_send = httpx.AsyncClient.send
    sync_send = httpx.Client.send

    async def guarded_async(client, request, *args, **kwargs):
        method, url = request.method, str(request.url)
        body = _json(await request.aread()) if wants_body(method, url) else None
        existing = {}
        target = guard.lookup(method, url)
        if target is not None:
            look = client.build_request(
                "GET", attendees_url(url, target), headers=_authorization(request.headers)
            )
            found = await async_send(client, look)
            existing = _attendees_found(target, found.status_code, found.content)
        guard.before(method, url, body, existing)
        response = await async_send(client, request, *args, **kwargs)
        attendees = _created_event(method, url, body, response.status_code)
        if attendees is not None:
            _remember(guard, await response.aread(), attendees)
        return response

    def guarded_sync(client, request, *args, **kwargs):
        method, url = request.method, str(request.url)
        body = _json(request.read()) if wants_body(method, url) else None
        existing = {}
        target = guard.lookup(method, url)
        if target is not None:
            look = client.build_request(
                "GET", attendees_url(url, target), headers=_authorization(request.headers)
            )
            found = sync_send(client, look)
            existing = _attendees_found(target, found.status_code, found.content)
        guard.before(method, url, body, existing)
        response = sync_send(client, request, *args, **kwargs)
        attendees = _created_event(method, url, body, response.status_code)
        if attendees is not None:
            _remember(guard, response.read(), attendees)
        return response

    patch.setattr(httpx.AsyncClient, "send", guarded_async)
    patch.setattr(httpx.Client, "send", guarded_sync)

    try:
        import requests
    except ImportError:  # pragma: no cover - msal depends on it
        requests = None
    if requests is not None:
        session_send = requests.Session.send

        def guarded_requests(session, request, **kwargs):
            method, url = (request.method or "GET"), request.url
            body = _json(request.body) if wants_body(method, url) else None
            existing = {}
            target = guard.lookup(method, url)
            if target is not None:
                look = requests.Request(
                    "GET", attendees_url(url, target), headers=_authorization(request.headers)
                ).prepare()
                found = session_send(session, look, timeout=30)
                existing = _attendees_found(target, found.status_code, found.content)
            guard.before(method, url, body, existing)
            response = session_send(session, request, **kwargs)
            attendees = _created_event(method, url, body, response.status_code)
            if attendees is not None:
                _remember(guard, response.content, attendees)
            return response

        patch.setattr(requests.Session, "send", guarded_requests)
    try:
        yield guard
    finally:
        patch.undo()
