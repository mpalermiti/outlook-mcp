"""The live tiers' guards, tested offline with stand-in accounts and fakes.

``tests/live_guard.py`` decides whose mailbox a live test may touch and who may
get mail from it. Those decisions only run for real inside the live tiers, which
CI never selects, so a guard that quietly stopped refusing would not fail anything
until it was needed. They are pinned here instead, with nothing live involved:
the addresses below are made up (example.com/example.org), and the Graph client
is the real SDK over a transport that is a function.
"""

from __future__ import annotations

import json

import pytest

from outlook_mcp.config import Config
from outlook_mcp.tools.calendar_write import create_event, delete_event
from outlook_mcp.tools.mail_write import send_message
from tests.live_guard import (
    ACCOUNT_ENV,
    OWN_MAILBOX_ENV,
    PAIR_ENV,
    SEND_BUDGET_ENV,
    SEND_ENV,
    LiveSettings,
    NeverSends,
    SendGuard,
    SendRefused,
    account_refusal,
    classify,
    identity_of,
    install,
    refusal,
    settings_from,
)
from tests.test_batch import _graph_client_over

MAIN = "main-test@example.com"
PAIR = "pair-test@example.com"
OWNER = "someone@example.org"
STRANGER = "stranger@example.org"
GRAPH = "https://graph.microsoft.com/v1.0"

TEST_PAIR = LiveSettings(account=MAIN, pair=PAIR)
NOTHING = LiveSettings()
OWN = LiveSettings(own_mailbox=True)

LIVE = frozenset({"live"})
INTEGRATION = frozenset({"integration"})
WRITE = frozenset({"live_write"})
SEND = frozenset({"live_send"})


def _who(*addresses):
    return [{"emailAddress": {"address": a}} for a in addresses]


# ── Settings ──────────────────────────────────────────────────────────────


def test_settings_come_from_the_environment_casefolded():
    settings = settings_from(
        {
            ACCOUNT_ENV: " Main-Test@Example.com ",
            PAIR_ENV: PAIR,
            SEND_ENV: "1",
            SEND_BUDGET_ENV: "3",
        }
    )
    assert settings.account == MAIN
    assert settings.test_accounts == {MAIN, PAIR}
    assert settings.send and settings.send_budget == 3
    assert not settings.own_mailbox


@pytest.mark.parametrize("budget", ["", "three", "-1"])
def test_a_budget_that_is_not_a_count_is_no_budget(budget):
    assert settings_from({SEND_BUDGET_ENV: budget}).send_budget is None


def test_an_opt_in_is_exactly_1():
    assert not settings_from({OWN_MAILBOX_ENV: "true"}).own_mailbox
    assert settings_from({OWN_MAILBOX_ENV: "1"}).own_mailbox


# ── Whose mailbox: before any sign-in ─────────────────────────────────────


@pytest.mark.parametrize("tiers", [LIVE, INTEGRATION, WRITE])
def test_with_no_test_account_and_no_opt_in_every_live_tier_is_refused(tiers):
    assert "no test account is named" in refusal(tiers, NOTHING)


def test_the_own_mailbox_opt_in_covers_the_read_only_live_tier_only():
    assert refusal(LIVE, OWN) is None
    for tiers in (INTEGRATION, WRITE, LIVE | WRITE):
        assert "covers only the read-only `live` tier" in refusal(tiers, OWN)
    assert "needs both test accounts" in refusal(SEND, OWN)


def test_the_send_tier_needs_both_accounts_named():
    assert "needs both test accounts" in refusal(SEND, NOTHING)
    assert "needs both test accounts" in refusal(SEND, LiveSettings(account=MAIN))
    assert refusal(SEND, TEST_PAIR) is None


def test_an_offline_test_is_never_refused():
    assert refusal(frozenset(), NOTHING) is None
    assert account_refusal(None, frozenset(), NOTHING) == ""


# ── Whose mailbox: the signed-in account ──────────────────────────────────


@pytest.mark.parametrize("tiers", [LIVE, INTEGRATION, WRITE, SEND])
def test_a_wrong_account_is_refused_in_every_tier(tiers):
    reason = account_refusal(OWNER, tiers, TEST_PAIR)
    assert reason.startswith("Refusing to run")
    assert OWNER not in reason  # the text lands in logs


@pytest.mark.parametrize("tiers", [LIVE, INTEGRATION, WRITE])
def test_either_test_account_may_run_the_other_tiers(tiers):
    assert account_refusal(MAIN, tiers, TEST_PAIR) == ""
    assert account_refusal(PAIR, tiers, TEST_PAIR) == ""


def test_the_send_tier_runs_only_as_the_main_account():
    assert account_refusal(MAIN, SEND, TEST_PAIR) == ""
    assert "only as the main test account" in account_refusal(PAIR, SEND, TEST_PAIR)


def test_an_unknown_account_is_refused():
    assert "could not tell which account" in account_refusal(None, WRITE, TEST_PAIR)


def test_the_opt_in_lets_any_account_run_live_but_nothing_else():
    both = LiveSettings(account=MAIN, pair=PAIR, own_mailbox=True)
    assert account_refusal(OWNER, LIVE, both) == ""
    assert account_refusal(OWNER, WRITE, both).startswith("Refusing to run")


def test_identity_needs_every_source_to_agree():
    assert identity_of({"userPrincipalName": "Main-Test@example.com", "mail": None}, None) == MAIN
    assert identity_of({"userPrincipalName": MAIN}, MAIN.upper()) == MAIN
    assert identity_of({"userPrincipalName": MAIN, "mail": PAIR}, None) is None
    assert identity_of({"userPrincipalName": MAIN}, PAIR) is None
    assert identity_of({}, None) is None


# ── Who may get mail: what counts as a send ───────────────────────────────


def test_send_mail_names_every_recipient():
    body = {
        "Message": {  # capitalised, as the SDK writes it
            "toRecipients": _who(PAIR),
            "ccRecipients": _who(MAIN),
            "bccRecipients": _who(STRANGER),
            "replyTo": _who(OWNER),
        }
    }
    sent = classify("POST", f"{GRAPH}/me/sendMail", body)
    assert sent.what == "sendMail"
    assert sent.recipients == (PAIR, MAIN, STRANGER, OWNER)


@pytest.mark.parametrize(
    "path", ["messages/m1/send", "messages/m1/reply", "messages/m1/replyAll", "events/e1/forward"]
)
def test_a_send_whose_recipients_the_request_does_not_name(path):
    sent = classify("POST", f"{GRAPH}/me/{path}", {})
    assert sent is not None and sent.recipients is None


def test_an_event_is_an_invitation_only_with_attendees():
    assert classify("POST", f"{GRAPH}/me/events", {"subject": "x"}) is None
    sent = classify("POST", f"{GRAPH}/me/calendars/c1/events", {"attendees": _who(PAIR)})
    assert sent == classify("POST", f"{GRAPH}/me/events", {"attendees": _who(PAIR)})
    assert sent.recipients == (PAIR,)


def test_changing_or_deleting_a_meeting_this_run_made_sends_to_its_attendees():
    meetings = {"AAMk=": (PAIR,)}
    patch = classify("PATCH", f"{GRAPH}/me/events/AAMk%3D", {"subject": "y"}, meetings)
    assert patch.recipients == (PAIR,)
    cancel = classify("DELETE", f"{GRAPH}/me/events/AAMk%3D", None, meetings)
    assert cancel.what == "a meeting cancellation" and cancel.recipients == (PAIR,)
    assert classify("DELETE", f"{GRAPH}/me/events/other", None, meetings) is None
    assert classify("PATCH", f"{GRAPH}/me/events/other", {"subject": "y"}, meetings) is None


def test_answering_an_invitation_sends_unless_told_not_to():
    assert classify("POST", f"{GRAPH}/me/events/e1/accept", {"sendResponse": True})
    assert classify("POST", f"{GRAPH}/me/events/e1/decline", {})
    assert classify("POST", f"{GRAPH}/me/events/e1/accept", {"sendResponse": False}) is None


def test_a_batch_is_looked_inside():
    reads = {"requests": [{"id": "1", "method": "GET", "url": "/me/messages/m1"}]}
    assert classify("POST", f"{GRAPH}/$batch", reads) is None
    sends = {
        "requests": [
            {"id": "1", "method": "GET", "url": "/me/messages/m1"},
            {"id": "2", "method": "POST", "url": "/me/messages/m1/reply", "body": {}},
        ]
    }
    assert classify("POST", f"{GRAPH}/$batch", sends).recipients is None


def test_reads_and_unsent_drafts_are_not_sends():
    draft = {"subject": "x", "toRecipients": _who(STRANGER)}
    assert classify("POST", f"{GRAPH}/me/messages", draft) is None
    assert classify("POST", f"{GRAPH}/me/messages/m1/createReply", {}) is None
    assert classify("GET", f"{GRAPH}/me/events", None) is None
    assert classify("DELETE", f"{GRAPH}/me/messages/m1", None) is None


# ── Who may get mail: the two guards ──────────────────────────────────────


def test_never_sends_refuses_any_send_and_allows_the_write_tier():
    guard = NeverSends(tier="live_write")
    with pytest.raises(SendRefused, match="live_write may not send"):
        guard.before("POST", f"{GRAPH}/me/sendMail", {"message": {"toRecipients": _who(PAIR)}})
    with pytest.raises(SendRefused):
        guard.before("POST", f"{GRAPH}/me/events", {"attendees": _who(PAIR)})
    guard.before("POST", f"{GRAPH}/me/events", {"subject": "no attendees"})
    guard.before("DELETE", f"{GRAPH}/me/events/e1", None)
    guard.before("POST", f"{GRAPH}/me/messages", {"toRecipients": _who(STRANGER)})


def test_a_send_to_anyone_but_the_pair_is_refused_and_not_counted():
    guard = SendGuard(pair=PAIR, budget=3)
    assert guard.reserve(1) == ""
    for recipients in ([STRANGER], [PAIR, STRANGER], [MAIN], []):
        with pytest.raises(SendRefused):
            guard.authorize("sendMail", recipients)
    assert guard.used == 0 and guard.reserved == 1


def test_a_send_with_recipients_it_cannot_read_is_refused():
    guard = SendGuard(pair=PAIR, budget=3)
    guard.reserve(1)
    with pytest.raises(SendRefused, match="cannot be checked"):
        guard.before("POST", f"{GRAPH}/me/messages/m1/reply", {})
    assert guard.used == 0


def test_the_pair_is_matched_whatever_its_case():
    guard = SendGuard(pair=PAIR, budget=1)
    guard.reserve(1)
    guard.authorize("sendMail", [PAIR.upper()])
    assert guard.used == 1


def test_the_budget_is_enforced():
    guard = SendGuard(pair=PAIR, budget=2)
    assert "needs 3 send(s), 2 of 2 left" in guard.reserve(3)
    assert guard.reserve(2) == ""
    guard.authorize("one", [PAIR])
    guard.authorize("two", [PAIR])
    with pytest.raises(SendRefused, match="did not reserve"):
        guard.authorize("three", [PAIR])
    guard.release()
    assert "0 of 2 left" in guard.reserve(1)
    assert guard.used == 2 and guard.log == ["one", "two"]


def test_an_unreserved_send_is_refused():
    guard = SendGuard(pair=PAIR, budget=5)
    with pytest.raises(SendRefused, match="did not reserve"):
        guard.authorize("sendMail", [PAIR])


def test_the_send_tier_deletes_only_meetings_it_made():
    guard = SendGuard(pair=PAIR, budget=1)
    with pytest.raises(SendRefused, match="deletes only meetings it created"):
        guard.before("DELETE", f"{GRAPH}/me/events/someone-elses", None)
    guard.created("mine", (PAIR,))
    guard.reserve(1)
    guard.before("DELETE", f"{GRAPH}/me/events/mine", None)
    assert guard.log == ["a meeting cancellation"]


# ── On the wire: the real tools through the real SDK ──────────────────────


class _Recorder:
    """A Graph endpoint that records what reached it and answers like Graph."""

    def __init__(self) -> None:
        self.seen: list[tuple[str, str]] = []

    def __call__(self, request):
        import httpx

        self.seen.append((request.method, request.url.path))
        if request.method == "POST" and request.url.path.endswith("/events"):
            sent = json.loads(request.content)
            return httpx.Response(201, json={"id": "meeting-1", "subject": sent.get("subject")})
        return httpx.Response(202 if request.method == "POST" else 204)


class _Wrapper:
    """Shaped like outlook_mcp.graph.GraphClient: what `install` hooks into."""

    def __init__(self, sdk) -> None:
        self.sdk_client = sdk


def _client(guard):
    graph = _Recorder()
    wrapper = _Wrapper(_graph_client_over(graph))
    install(wrapper, guard)
    return wrapper.sdk_client, graph


_CONFIG = Config(timezone="UTC")


async def test_the_send_tool_cannot_reach_a_stranger():
    guard = SendGuard(pair=PAIR, budget=1)
    guard.reserve(1)
    sdk, graph = _client(guard)
    with pytest.raises(SendRefused):
        await send_message(sdk, to=[PAIR], cc=[STRANGER], subject="s", body="b", config=_CONFIG)
    assert graph.seen == [] and guard.used == 0

    await send_message(sdk, to=[PAIR], subject="s", body="b", config=_CONFIG)
    assert graph.seen == [("POST", "/v1.0/me/sendMail")] and guard.used == 1


async def test_an_invitation_and_its_cancellation_are_both_counted():
    guard = SendGuard(pair=PAIR, budget=2)
    guard.reserve(2)
    sdk, graph = _client(guard)
    created = await create_event(
        sdk,
        subject="s",
        start="2030-01-07T17:00:00",
        end="2030-01-07T17:30:00",
        timezone="UTC",
        attendees=[PAIR],
        config=_CONFIG,
    )
    assert guard.meetings == {created["event_id"]: (PAIR,)}
    await delete_event(sdk, created["event_id"], config=_CONFIG)
    assert guard.log == ["an invitation", "a meeting cancellation"]
    assert guard.used == 2


async def test_an_invitation_beyond_the_budget_never_leaves():
    guard = SendGuard(pair=PAIR, budget=0)
    sdk, graph = _client(guard)
    with pytest.raises(SendRefused):
        await create_event(
            sdk,
            subject="s",
            start="2030-01-07T17:00:00",
            end="2030-01-07T17:30:00",
            timezone="UTC",
            attendees=[PAIR],
            config=_CONFIG,
        )
    assert graph.seen == []


async def test_the_shared_live_client_refuses_any_send():
    sdk, graph = _client(NeverSends(tier="live_write"))
    with pytest.raises(SendRefused):
        await send_message(sdk, to=[PAIR], subject="s", body="b", config=_CONFIG)
    assert graph.seen == []
