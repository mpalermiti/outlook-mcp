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
   ``/me`` and the saved sign-in record, which must agree.
3. **May this request go out?** (:class:`NeverSends`, :class:`SendGuard`) Every
   request the live tiers' Graph client makes passes through :func:`install`. A
   request that makes Exchange send mail — ``sendMail``, a send or reply action, an
   invitation, an update or cancellation of a meeting — is refused in every tier but
   ``live_send``. There, each recipient must be the paired test account, checked on
   the request itself before it leaves, and each send is counted against the run's
   budget.
"""

from __future__ import annotations

import json
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from urllib.parse import unquote, urlsplit

import pytest

# ── Settings, from the environment ────────────────────────────────────────

ACCOUNT_ENV = "OUTLOOK_MCP_TEST_ACCOUNT"
PAIR_ENV = "OUTLOOK_MCP_TEST_PAIR"
PAIR_CONFIG_DIR_ENV = "OUTLOOK_MCP_TEST_PAIR_CONFIG_DIR"
PAIR_CAN_SEND_ENV = "OUTLOOK_MCP_TEST_PAIR_CAN_SEND"
OWN_MAILBOX_ENV = "OUTLOOK_MCP_LIVE_OWN_MAILBOX"
SEND_ENV = "OUTLOOK_MCP_LIVE_SEND"
SEND_BUDGET_ENV = "OUTLOOK_MCP_LIVE_SEND_BUDGET"

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
    if budget is not None and budget < 0:
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


def identity_of(me: Mapping, record_username: str | None) -> str | None:
    """The signed-in address, or None when the sources disagree or say nothing.

    A personal account reports its address as ``userPrincipalName``; ``mail`` is
    often null. The saved sign-in record names the account the token cache
    belongs to. Every one of them that is present must agree.
    """
    found = {
        value
        for value in (
            _address(me.get("userPrincipalName")),
            _address(me.get("mail")),
            _address(record_username),
        )
        if value
    }
    return found.pop() if len(found) == 1 else None


def account_refusal(identity: str | None, tiers: frozenset[str], settings: LiveSettings) -> str:
    """Why the signed-in account may not run these tiers; "" when it may."""
    if not tiers or own_mailbox_allows(tiers, settings):
        return ""
    if identity is None:
        return (
            "Refusing to run: could not tell which account is signed in (/me and the "
            "saved sign-in record are missing or disagree)."
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


# ── 3: who may get mail ───────────────────────────────────────────────────


class SendRefused(pytest.fail.Exception):
    """A request would have sent mail it may not send.

    A pytest failure, which is a BaseException, so a tool's ``except Exception``
    cannot turn it into "the message may have been sent".
    """


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


def _folded(value):
    """``value`` with every object key casefolded, all the way down."""
    if isinstance(value, Mapping):
        return {str(k).casefold(): _folded(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_folded(v) for v in value]
    return value


def _segments(url: str) -> list[str]:
    path = urlsplit(url).path
    parts = [unquote(p) for p in path.split("/") if p]
    if parts and parts[0].lower() in ("v1.0", "beta"):
        parts = parts[1:]
    return parts


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
    if method.upper() not in ("POST", "PATCH"):
        return False
    parts = _segments(url)
    return bool(parts) and (
        parts[-1].lower() in ("sendmail", "$batch", "events")
        or _event_id(parts) is not None
        or parts[-1].lower() in _RESPONSES
    )


def classify(
    method: str,
    url: str,
    body: Mapping | None,
    meetings: Mapping[str, tuple[str, ...]] | None = None,
) -> Send | None:
    """The mail this request makes Exchange send, or None when it sends none.

    ``meetings`` are the meetings this run created, by id, with their attendees:
    changing or deleting one sends them an update or a cancellation.
    """
    method = method.upper()
    meetings = meetings or {}
    parts = _segments(url)
    if not parts:
        return None
    last = parts[-1].lower()
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
    if method == "DELETE" and event_id in meetings:
        return Send("a meeting cancellation", meetings[event_id])
    return None


def deletes_an_event(method: str, url: str) -> str | None:
    """The event id a DELETE (or cancel) removes, else None."""
    parts = _segments(url)
    event_id = _event_id(parts)
    if event_id is None:
        return None
    if method.upper() == "DELETE" and parts[-1] == event_id:
        return event_id
    if method.upper() == "POST" and parts[-1].lower() == "cancel":
        return event_id
    return None


class NeverSends:
    """Every tier but ``live_send``: no request may send mail."""

    meetings: Mapping[str, tuple[str, ...]] = {}

    def __init__(self, tier: str = "this tier") -> None:
        self.tier = tier

    def before(self, method: str, url: str, body: Mapping | None) -> None:
        sent = classify(method, url, body)
        if sent:
            raise SendRefused(
                f"Refused before it left: {self.tier} may not send mail, and this request "
                f"would ({sent.what}). Sending tests belong in the live_send tier."
            )

    def created(self, event_id: str, attendees: tuple[str, ...]) -> None:
        """Nothing to remember: an invitation never gets this far."""


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
    log: list[str] = field(default_factory=list)

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
        strangers = [r for r in recipients if r != allowed]
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

    def before(self, method: str, url: str, body: Mapping | None) -> None:
        sent = classify(method, url, body, self.meetings)
        if sent is None:
            event_id = deletes_an_event(method, url)
            if event_id is not None and event_id not in self.meetings:
                # On the test account only this tier organizes meetings, but an
                # event it did not create could still be someone's meeting.
                raise SendRefused(
                    "Refused before it left: the send tier deletes only meetings it created "
                    "this run, and this event is not one of them."
                )
            return
        self.authorize(sent.what, sent.recipients)

    def created(self, event_id: str, attendees: tuple[str, ...]) -> None:
        self.meetings[event_id] = attendees


def install(graph_client, guard) -> None:
    """Hold every request ``graph_client`` makes to ``guard`` before it leaves.

    Hooks on the SDK's own httpx client, so a tool cannot route around it. The
    response hook remembers each meeting the tier created, so changing or deleting
    it later is counted as the update or cancellation it sends.
    """
    http = graph_client.sdk_client.request_adapter._http_client

    async def body_of(message) -> Mapping | None:
        content = await message.aread()
        if not content:
            return None
        try:
            return json.loads(content)
        except ValueError:
            return {"<unreadable body>": True}

    async def on_request(request) -> None:
        url = str(request.url)
        body = await body_of(request) if wants_body(request.method, url) else None
        guard.before(request.method, url, body)
        request.extensions["live_guard_body"] = body

    async def on_response(response) -> None:
        request = response.request
        body = request.extensions.get("live_guard_body")
        if request.method != "POST" or not body or not response.is_success:
            return
        url = str(request.url)
        if _segments(url)[-1].lower() != "events":
            return
        attendees = _addresses(_folded(body).get("attendees"))
        if not attendees:
            return
        created = await body_of(response)
        if isinstance(created, Mapping) and isinstance(created.get("id"), str):
            guard.created(created["id"], attendees)

    http.event_hooks["request"].append(on_request)
    http.event_hooks["response"].append(on_response)
