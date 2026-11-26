"""The send tier: real mail and invitations between the two test accounts.

Mocks show what we *build*; only Exchange shows what arrives. These tests send
from the main test account to its pair — the only recipient the tier's Graph
client lets through (``tests/live_guard.SendGuard``) — look for the result in the
pair's mailbox, then delete everything they sent or received from both
mailboxes (``tests/live_mailbox.py``, run as each account).

Each test reserves its sends first and skips when the run's budget cannot cover
them (``scripts/live-tests.sh live_send --sends N``). Tests that need the pair to
answer skip until it can send (``OUTLOOK_MCP_TEST_PAIR_CAN_SEND=1``): Outlook.com
blocks a new account from sending until its owner verifies it.

Delivery can take a minute or two, and a new account's mail often lands in Junk,
so the pair's Inbox and Junk Email both count as delivered.
"""

from __future__ import annotations

import asyncio
import os
import uuid
from datetime import datetime, timedelta, timezone

import pytest

from outlook_mcp.tools.calendar_read import get_event
from outlook_mcp.tools.calendar_write import create_event, delete_event
from outlook_mcp.tools.mail_thread import list_thread
from outlook_mcp.tools.mail_write import send_message
from tests.conftest import LIVE_SEND_SUBJECT
from tests.live_guard import PAIR_CAN_SEND_ENV, settings_from

pytestmark = pytest.mark.live_send

# Seconds to wait for mail to reach the other mailbox.
DELIVERY_WAIT = "240"
DELIVERED = ("--folder", "inbox", "--folder", "junkemail")
EVERYWHERE = (*DELIVERED, "--folder", "sentitems", "--folder", "deleteditems")
BODY = "Sent by outlook-mcp's live_send tier between its two test accounts. Safe to delete."

needs_the_pair_to_send = pytest.mark.skipif(
    not settings_from(os.environ).pair_can_send,
    reason=(
        "Needs the pair account to send (a reply or an acceptance), and it cannot yet: "
        "Outlook.com blocks a new account from sending until its owner verifies it. "
        f"Set {PAIR_CAN_SEND_ENV}=1 in the live-tests env file once it can."
    ),
)


def _accounts() -> tuple[str, str]:
    settings = settings_from(os.environ)
    return settings.account, settings.pair


def _marker(what: str) -> str:
    return f"{LIVE_SEND_SUBJECT} {what} {uuid.uuid4().hex[:8]}"


def _reserve(sends, count: int) -> None:
    reason = sends.reserve(count)
    if reason:
        pytest.skip(reason)


def _a_slot() -> tuple[str, str]:
    """A half hour three weeks out, as wall-clock UTC with no offset.

    No offset on purpose: a typed instant makes `create_event` read the event back
    and patch it if Graph moved it, and patching a meeting sends an update.
    """
    start = (datetime.now(timezone.utc) + timedelta(days=21)).replace(
        hour=17, minute=0, second=0, microsecond=0
    )
    end = start + timedelta(minutes=30)
    return start.strftime("%Y-%m-%dT%H:%M:%S"), end.strftime("%Y-%m-%dT%H:%M:%S")


def _clean_up(mailbox, marker: str) -> None:
    """Delete every marked item from both mailboxes: messages, and the pair's
    cancelled copies of a meeting. Only items carrying this test's marker."""
    mailbox.run("pair", "remove-events", "--marker", marker, "--wait", "180")
    mailbox.run("pair", "purge", "--marker", marker, *EVERYWHERE)
    mailbox.run("main", "purge", "--marker", marker, *EVERYWHERE, "--folder", "drafts")


async def test_mail_to_the_pair_arrives(send_graph_client, send_config, sends, pair_mailbox):
    """`outlook_send_message` delivers, and keeps a copy in Sent Items."""
    _reserve(sends, 1)
    _, pair = _accounts()
    marker = _marker("mail")
    try:
        result = await send_message(
            send_graph_client.sdk_client, to=[pair], subject=marker, body=BODY, config=send_config
        )
        assert result == {"status": "sent"}
        assert sends.reserved == 0, "the guard did not count the send"

        got = pair_mailbox.run(
            "pair", "find", "--marker", marker, *DELIVERED, "--wait", DELIVERY_WAIT
        )["messages"]
        assert len(got) == 1, "the message did not reach the pair's Inbox or Junk Email"
        assert got[0]["from"] == "account"
        assert got[0]["subject"] == marker

        kept = pair_mailbox.run(
            "main", "find", "--marker", marker, "--folder", "sentitems", "--wait", "60"
        )["messages"]
        assert len(kept) == 1, "no copy in the sender's Sent Items"
    finally:
        _clean_up(pair_mailbox, marker)


async def test_an_invitation_reaches_the_pair_and_its_cancellation_follows(
    send_graph_client, send_config, sends, pair_mailbox
):
    """`outlook_create_event` with an attendee invites them; `outlook_delete_event`
    on the organizer's copy cancels for them. Two sends."""
    _reserve(sends, 2)
    _, pair = _accounts()
    marker = _marker("invite")
    start, end = _a_slot()
    sdk = send_graph_client.sdk_client
    event_id = None
    try:
        created = await create_event(
            sdk,
            subject=marker,
            start=start,
            end=end,
            timezone="UTC",
            body=BODY,
            attendees=[pair],
            config=send_config,
        )
        event_id = created["event_id"]
        assert event_id in sends.meetings, "the guard did not see the invitation go out"

        request = pair_mailbox.run(
            "pair",
            "find",
            "--marker",
            marker,
            *DELIVERED,
            "--kind",
            "meetingRequest",
            "--wait",
            DELIVERY_WAIT,
        )["messages"]
        assert len(request) == 1, "the invitation did not reach the pair's Inbox or Junk Email"
        assert request[0]["from"] == "account"

        await delete_event(sdk, event_id, config=send_config)
        event_id = None
        cancellation = pair_mailbox.run(
            "pair",
            "find",
            "--marker",
            marker,
            *DELIVERED,
            "--folder",
            "deleteditems",
            "--kind",
            "meetingCancelled",
            "--wait",
            DELIVERY_WAIT,
        )["messages"]
        assert len(cancellation) == 1, "the cancellation did not reach the pair"
        assert cancellation[0]["from"] == "account"
    finally:
        if event_id is not None:
            await delete_event(sdk, event_id, config=send_config)  # the reserved cancellation
        _clean_up(pair_mailbox, marker)


@needs_the_pair_to_send
async def test_a_reply_from_the_pair_joins_the_thread(
    send_graph_client, send_config, sends, pair_mailbox
):
    """The pair's reply comes back into the same conversation, and
    `outlook_list_thread` returns both messages. Two sends: ours and the reply."""
    _reserve(sends, 2)
    account, pair = _accounts()
    marker = _marker("reply")
    sdk = send_graph_client.sdk_client
    try:
        await send_message(sdk, to=[pair], subject=marker, body=BODY, config=send_config)
        got = pair_mailbox.run(
            "pair", "find", "--marker", marker, *DELIVERED, "--wait", DELIVERY_WAIT
        )["messages"]
        assert len(got) == 1, "the message did not reach the pair"

        sends.authorize("the pair's reply", [account], to=account)
        pair_mailbox.run("pair", "reply", "--marker", marker, "--comment", BODY, "--allow-send")

        replies = pair_mailbox.run(
            "main", "find", "--marker", marker, *DELIVERED, "--wait", DELIVERY_WAIT
        )["messages"]
        assert len(replies) == 1, "the pair's reply did not come back"
        assert replies[0]["from"] == "pair"
        original = pair_mailbox.run("main", "find", "--marker", marker, "--folder", "sentitems")[
            "messages"
        ]
        assert replies[0]["conversation_id"] == original[0]["conversation_id"]

        thread = await list_thread(sdk, replies[0]["conversation_id"])
        assert thread["count"] == 2, thread
    finally:
        _clean_up(pair_mailbox, marker)


@needs_the_pair_to_send
async def test_the_pairs_acceptance_reaches_the_organizer(
    send_graph_client, send_config, sends, pair_mailbox
):
    """The pair accepts; `outlook_get_event` on the organizer's copy shows it.
    Three sends: the invitation, the acceptance and the cancellation."""
    _reserve(sends, 3)
    account, pair = _accounts()
    marker = _marker("accept")
    start, end = _a_slot()
    sdk = send_graph_client.sdk_client
    event_id = None
    try:
        created = await create_event(
            sdk,
            subject=marker,
            start=start,
            end=end,
            timezone="UTC",
            body=BODY,
            attendees=[pair],
            config=send_config,
        )
        event_id = created["event_id"]
        on_their_calendar = pair_mailbox.run(
            "pair", "events", "--marker", marker, "--wait", DELIVERY_WAIT
        )["events"]
        assert on_their_calendar, (
            "the invitation made no event on the pair's calendar: Outlook.com files a new "
            "sender's invitation in Junk Email without one. Mark the account's mail as not "
            "junk in the pair's mailbox, then run again."
        )

        sends.authorize("the pair's acceptance", [account], to=account)
        pair_mailbox.run("pair", "accept", "--marker", marker, "--allow-send")

        response = None
        for _ in range(int(DELIVERY_WAIT) // 10):
            detail = await get_event(sdk, event_id)
            response = next(
                (
                    a.get("response")
                    for a in detail["attendees"]
                    if (a.get("email") or "").casefold() == pair
                ),
                None,
            )
            if response == "accepted":
                break
            await asyncio.sleep(10)
        assert response == "accepted", f"the organizer's copy shows {response!r}"
    finally:
        if event_id is not None:
            await delete_event(sdk, event_id, config=send_config)  # the reserved cancellation
        _clean_up(pair_mailbox, marker)
