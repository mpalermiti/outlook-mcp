"""Does the converted-time check read the event back with a GET on the wire?

A release run against real Graph on 2026-11-25 showed the check as first built
never ran: it asked for the write's reply in UTC with `Prefer:
outlook.timezone="UTC"`, and Graph answered the POST in the event's own zone
(`2026-12-27T18:00:00`, `America/Los_Angeles`). The check could not read that,
failed open, and five live tests failed. It now reads the event back with
`GET .../events/{id}?$select=start,end` and that header, which Graph honours on
a GET, and the write carries no `Prefer` at all.

A mock of `.get()` sees a query-parameters object, not the URL it becomes, and
the read uses the `/me/events` builder's parameters class on whichever builder
the event was written to. So these tests build the client the way `msgraph`
does, hand it a transport that answers as Graph did live, and assert on what
the real request pipeline sends, as `test_search_cursor_reaches_the_wire.py`
does.
"""

from __future__ import annotations

import inspect
import json
from urllib.parse import parse_qs, urlsplit

import httpx
import pytest
from azure.core.credentials import AccessToken
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph import graph_request_adapter as _gra
from msgraph_core import GraphClientFactory

from outlook_mcp.config import Config
from outlook_mcp.tools.calendar_write import create_event, update_event

pytestmark = pytest.mark.asyncio

ME = "https://graph.microsoft.com/v1.0/me"
EVENT_ID = "AAMkAGnew="
_CFG = Config(client_id="test")
_UTC = 'outlook.timezone="UTC"'


class _FakeCredential:
    def get_token(self, *args, **kwargs):
        return AccessToken("not-a-real-token", 9_999_999_999)


def _times(start: str, end: str, zone: str) -> dict:
    """An event's `start` and `end` as Graph spells them, seven fractional digits."""
    return {
        "start": {"dateTime": f"{start}.0000000", "timeZone": zone},
        "end": {"dateTime": f"{end}.0000000", "timeZone": zone},
    }


#: The POST's reply as Graph sent it live, whatever `Prefer` said: in the
#: event's own zone. 18:00 Los Angeles on the 12th is 02:00Z on the 13th.
_WRITE_REPLY = {
    "id": EVENT_ID,
    "subject": "Dinner",
    "type": "singleInstance",
    "originalStartTimeZone": "America/Los_Angeles",
    "originalEndTimeZone": "America/Los_Angeles",
    **_times("2026-12-12T18:00:00", "2026-12-12T19:00:00", "America/Los_Angeles"),
}
_AS_TYPED = {"id": EVENT_ID, **_times("2026-12-13T02:00:00", "2026-12-13T03:00:00", "UTC")}
_AN_HOUR_LATE = {"id": EVENT_ID, **_times("2026-12-13T03:00:00", "2026-12-13T04:00:00", "UTC")}


def _graph(read_backs: list[dict], **routes: dict) -> tuple[GraphServiceClient, list]:
    """A real SDK client over a transport that answers as Graph did live.

    A POST or PATCH gets the write's reply; a GET with `$select=start,end` gets
    the next of ``read_backs`` (a body, or a whole `httpx.Response`); any other
    GET gets ``routes[path]``. Returns the client and each request sent, in order.
    """
    sent: list[httpx.Request] = []
    queue = list(read_backs)

    async def transport(request: httpx.Request) -> httpx.Response:
        sent.append(request)
        if request.method in ("POST", "PATCH"):
            return httpx.Response(201 if request.method == "POST" else 200, json=_WRITE_REPLY)
        if _query(request).get("$select") == ["start,end"]:
            answer = queue.pop(0)
            return (
                answer if isinstance(answer, httpx.Response) else httpx.Response(200, json=answer)
            )
        return httpx.Response(200, json=routes[request.url.path])

    factory = GraphClientFactory.create_with_default_middleware
    kwargs = {}
    params = inspect.signature(factory).parameters
    if "options" in params:
        kwargs["options"] = getattr(_gra, "options", None)
    if "client" in params:
        kwargs["client"] = httpx.AsyncClient(transport=httpx.MockTransport(transport))
    auth = AzureIdentityAuthenticationProvider(
        _FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
    )
    client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, factory(**kwargs)))
    return client, sent


def _query(request: httpx.Request) -> dict[str, list[str]]:
    return parse_qs(urlsplit(str(request.url)).query)


def _shape(request: httpx.Request) -> tuple[str, str, dict, str | None]:
    """A request as `(method, path, query, Prefer)`."""
    return (
        request.method,
        request.url.path,
        _query(request),
        request.headers.get("Prefer"),
    )


def _read_back(path: str) -> tuple[str, str, dict, str]:
    return ("GET", path, {"$select": ["start,end"]}, _UTC)


async def _create(client, **kwargs) -> dict:
    return await create_event(
        client,
        subject="Dinner",
        start="2026-12-13T02:00:00Z",
        end="2026-12-13T03:00:00Z",
        timezone="America/Los_Angeles",
        config=_CFG,
        **kwargs,
    )


async def test_a_converted_create_is_read_back_with_a_get():
    client, sent = _graph([_AS_TYPED])

    result = await _create(client)

    assert [_shape(r) for r in sent] == [
        # No `Prefer` on the write: Graph ignored it there.
        ("POST", "/v1.0/me/events", {}, None),
        _read_back(f"/v1.0/me/events/{EVENT_ID}"),
    ]
    posted = json.loads(sent[0].content)
    assert posted["start"] == {"dateTime": "2026-12-12T18:00:00", "timeZone": "America/Los_Angeles"}
    assert "note" not in result


async def test_a_moved_create_is_put_back_and_read_again():
    client, sent = _graph([_AN_HOUR_LATE, _AS_TYPED])

    result = await _create(client)

    path = f"/v1.0/me/events/{EVENT_ID}"
    assert [_shape(r) for r in sent] == [
        ("POST", "/v1.0/me/events", {}, None),
        _read_back(path),
        ("PATCH", path, {}, None),
        _read_back(path),
    ]
    assert json.loads(sent[2].content) == {
        "@odata.type": "#microsoft.graph.event",
        "start": {"dateTime": "2026-12-13T02:00:00Z", "timeZone": "America/Los_Angeles"},
        "end": {"dateTime": "2026-12-13T03:00:00Z", "timeZone": "America/Los_Angeles"},
    }
    assert "so the time is right" in result["note"]


async def test_a_series_create_reads_nothing_back():
    client, sent = _graph([])

    await _create(
        client,
        recurrence={
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["saturday"]},
            "range": {"type": "numbered", "numberOfOccurrences": 2},
        },
    )

    assert [_shape(r) for r in sent] == [("POST", "/v1.0/me/events", {}, None)]


async def test_a_create_in_a_named_calendar_is_read_back_there():
    """The read uses the `/me/events` builder's query class on a calendar's builder."""
    calendars = {
        "value": [
            {
                "id": "DEFAULT1=",
                "name": "Calendar",
                "canEdit": True,
                "isDefaultCalendar": True,
                "owner": {"address": "me@outlook.com"},
            },
            {
                "id": "KIDS77=",
                "name": "Kids + School",
                "canEdit": True,
                "isDefaultCalendar": False,
                "owner": {"address": "me@outlook.com"},
            },
        ]
    }
    client, sent = _graph([_AS_TYPED], **{"/v1.0/me/calendars": calendars})

    await _create(client, calendar="Kids + School")

    assert [_shape(r)[:2] for r in sent] == [
        ("GET", "/v1.0/me/calendars"),
        ("POST", "/v1.0/me/calendars/KIDS77=/events"),
        ("GET", f"/v1.0/me/calendars/KIDS77=/events/{EVENT_ID}"),
    ]
    assert _shape(sent[2]) == _read_back(f"/v1.0/me/calendars/KIDS77=/events/{EVENT_ID}")


#: A single Los Angeles event as a plain GET returns it: 01:00Z to 03:00Z.
_STORED = {
    **_WRITE_REPLY,
    "isAllDay": False,
    **_times("2026-12-13T01:00:00", "2026-12-13T03:00:00", "UTC"),
}
_PATH = f"/v1.0/me/events/{EVENT_ID}"


async def test_a_throttled_read_back_is_sent_once():
    """kiota's retry is off for the read: one 429, no wait, and the event is created.

    With it on, a `Retry-After` of minutes holds back the answer to a create
    that already succeeded, long enough for a client to give up and send it again.
    """
    throttled = httpx.Response(429, headers={"Retry-After": "1"}, json={"error": {}})
    client, sent = _graph([throttled])

    result = await _create(client)

    assert [_shape(r) for r in sent] == [
        ("POST", "/v1.0/me/events", {}, None),
        _read_back(f"/v1.0/me/events/{EVENT_ID}"),
    ]
    assert result == {"status": "created", "event_id": EVENT_ID, "subject": "Dinner"}


async def test_a_converted_update_is_read_back_with_a_get():
    path = _PATH
    client, sent = _graph([_AS_TYPED], **{path: _STORED})

    result = await update_event(
        client, event_id=EVENT_ID, start="2026-12-13T02:00:00Z", config=_CFG
    )

    assert [_shape(r) for r in sent] == [
        # The read for the zone the event is anchored in, as before.
        ("GET", path, {}, None),
        ("PATCH", path, {}, None),
        _read_back(path),
    ]
    assert result == {"status": "updated", "event_id": EVENT_ID}


async def test_a_moved_update_is_put_back_and_read_again():
    client, sent = _graph([_AN_HOUR_LATE, _AS_TYPED], **{_PATH: _STORED})

    result = await update_event(
        client, event_id=EVENT_ID, start="2026-12-13T02:00:00Z", config=_CFG
    )

    assert [_shape(r) for r in sent] == [
        ("GET", _PATH, {}, None),
        ("PATCH", _PATH, {}, None),
        _read_back(_PATH),
        ("PATCH", _PATH, {}, None),
        _read_back(_PATH),
    ]
    assert json.loads(sent[3].content) == {
        "@odata.type": "#microsoft.graph.event",
        "start": {"dateTime": "2026-12-13T02:00:00Z", "timeZone": "America/Los_Angeles"},
    }
    assert "so the time is right" in result["note"]


async def test_an_update_read_back_at_its_old_time_is_left_alone():
    """The read shows 01:00Z, where the event stood before the move to 02:00Z (S1).

    That hour is the move asked for, not a zone rule's, so no second PATCH goes
    out and nothing is said.
    """
    as_it_was = {"id": EVENT_ID, **_times("2026-12-13T01:00:00", "2026-12-13T03:00:00", "UTC")}
    client, sent = _graph([as_it_was], **{_PATH: _STORED})

    result = await update_event(
        client, event_id=EVENT_ID, start="2026-12-13T02:00:00Z", config=_CFG
    )

    assert [_shape(r) for r in sent] == [
        ("GET", _PATH, {}, None),
        ("PATCH", _PATH, {}, None),
        _read_back(_PATH),
    ]
    assert result == {"status": "updated", "event_id": EVENT_ID}
