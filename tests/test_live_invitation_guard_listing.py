"""Offline guard for how the live invitation guards list messages.

`tests/test_live_meeting_invitation.py` never runs offline, so a change to its
listing would go unnoticed until a credentialed run. These drive its listing
and the invitation guard through a real GraphServiceClient over a mock
transport: the filter on the message class, the page cap, and the loud failure
when Graph lists an invitation without its `@odata.type`.
"""

from __future__ import annotations

from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import httpx
import pytest

import tests.test_live_meeting_invitation as live
from tests.test_batch import _graph_client_over

LISTING_PATH = "/v1.0/me/messages"
INVITATION_TYPE = "#microsoft.graph.eventMessageRequest"


class _Mailbox:
    """Serves listing pages in order, and invitations without a calendar event."""

    def __init__(self, *pages: list[dict], next_link: bool = False):
        self.pages = list(pages)
        self.next_link = next_link
        self.listings: list[httpx.URL] = []
        self.expanded: list[str] = []

    def __call__(self, request: httpx.Request) -> httpx.Response:
        path = request.url.path
        if path == LISTING_PATH:
            self.listings.append(request.url)
            index = len(self.listings) - 1
            body: dict = {"value": self.pages[min(index, len(self.pages) - 1)]}
            if self.next_link:
                body["@odata.nextLink"] = (
                    f"https://graph.microsoft.com{LISTING_PATH}?$skiptoken=page{index + 1}"
                )
            return httpx.Response(200, json=body)
        message_id = path.rsplit("/", 1)[-1]
        self.expanded.append(message_id)
        # A cancelled or removed meeting: the invitation expands to no event.
        return httpx.Response(200, json={"@odata.type": INVITATION_TYPE, "id": message_id})

    def query(self, index: int = 0) -> dict[str, list[str]]:
        return parse_qs(urlsplit(str(self.listings[index])).query)


async def _outcome(coro) -> tuple[str, str]:
    """Run a live test body and say whether it passed, skipped or failed."""
    try:
        await coro
    except pytest.skip.Exception as exc:
        return "skipped", str(exc)
    except pytest.fail.Exception as exc:
        return "failed", str(exc)
    return "passed", ""


def _invitation_guard(mailbox: _Mailbox):
    client = SimpleNamespace(sdk_client=_graph_client_over(mailbox))
    return live.test_invitation_resolves_to_the_mailbox_event(client)


async def test_invitation_guard_lists_meeting_requests_by_message_class():
    mailbox = _Mailbox([])

    status, reason = await _outcome(_invitation_guard(mailbox))

    assert status == "skipped"
    assert "0 IPM.Schedule.Meeting.Request messages" in reason
    assert len(mailbox.listings) == 1
    assert mailbox.query() == {
        "$filter": [
            "singleValueExtendedProperties/Any(ep: ep/id eq 'String 0x001A' "
            "and ep/value eq 'IPM.Schedule.Meeting.Request')"
        ],
        "$select": ["id"],
        "$top": ["50"],
    }


async def test_plain_mail_sample_lists_notes_by_message_class():
    mailbox = _Mailbox([{"id": "AAMk-note"}])

    pages = live.message_pages(_graph_client_over(mailbox), live.PLAIN_MAIL)
    sample = await anext(pages)
    await pages.aclose()

    assert [m.id for m in sample] == ["AAMk-note"]
    assert mailbox.query()["$filter"] == [
        "singleValueExtendedProperties/Any(ep: ep/id eq 'String 0x001A' and ep/value eq 'IPM.Note')"
    ]


async def test_invitation_listed_without_its_odata_type_fails_loudly():
    mailbox = _Mailbox([{"id": "AAMk-untyped"}])

    status, reason = await _outcome(_invitation_guard(mailbox))

    assert status == "failed"
    assert "@odata.type #microsoft.graph.eventMessageRequest" in reason
    assert "AAMk-untyped" in reason
    assert mailbox.expanded == []


async def test_invitation_guard_follows_next_link_for_at_most_20_pages():
    mailbox = _Mailbox(
        *[[{"@odata.type": INVITATION_TYPE, "id": f"AAMk-{n}"}] for n in range(25)],
        next_link=True,
    )

    status, reason = await _outcome(_invitation_guard(mailbox))

    assert status == "skipped"
    assert "among 20 IPM.Schedule.Meeting.Request messages (20 of at most 20 pages read)" in reason
    assert len(mailbox.listings) == 20
    assert mailbox.expanded == [f"AAMk-{n}" for n in range(20)]
    assert [mailbox.query(i).get("$skiptoken") for i in (1, 19)] == [["page1"], ["page19"]]


async def test_invitation_guard_caps_expansions_within_a_page():
    mailbox = _Mailbox(
        [{"@odata.type": INVITATION_TYPE, "id": f"AAMk-{n}"} for n in range(75)],
        next_link=True,
    )

    status, reason = await _outcome(_invitation_guard(mailbox))

    assert status == "skipped"
    assert len(mailbox.expanded) == 50
    assert mailbox.expanded == [f"AAMk-{n}" for n in range(50)]
    assert len(mailbox.listings) == 1
    assert "50" in reason


async def test_invitation_guard_caps_expansions_across_pages():
    mailbox = _Mailbox(
        *[
            [{"@odata.type": INVITATION_TYPE, "id": f"AAMk-{page}-{n}"} for n in range(30)]
            for page in range(3)
        ],
        next_link=True,
    )

    status, _ = await _outcome(_invitation_guard(mailbox))

    assert status == "skipped"
    assert len(mailbox.expanded) == 50
    assert mailbox.expanded[-1] == "AAMk-1-19"
    assert len(mailbox.listings) == 2
