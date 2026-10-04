"""The delta cursor is attacker-reachable input. Prove the token can't follow it.

``fetch_delta_pages`` attaches ``Authorization: Bearer <Graph token>`` and then
GETs whatever URL it was handed. Two of those URLs come from outside:

- ``delta_token`` — the caller's stored cursor. ``_delta.py`` deliberately does
  not persist it ("that's the caller's job"), so it is agent-held state, and an
  agent takes instructions from mail it reads.
- ``@odata.nextLink`` — read back out of a response body mid-walk.

If either can name a host, the bearer token for the whole mailbox goes there.
These tests assert on the header that actually reached the wire, not on the
exception type, because the leak is a sent request and nothing else.
"""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.tools._delta import fetch_delta_pages, require_graph_url

GRAPH_DELTA = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta"

# Each of these is a plausible cursor string that must never receive the token.
OFF_GRAPH_URLS = [
    pytest.param("https://evil.example/v1.0/me/messages/delta", id="foreign-host"),
    pytest.param("http://graph.microsoft.com/v1.0/me/messages/delta", id="http-scheme"),
    pytest.param(
        "https://graph.microsoft.com.evil.example/v1.0/me/messages/delta",
        id="suffix-lookalike",
    ),
    pytest.param(
        "https://graph.microsoft.com@evil.example/v1.0/me/messages/delta",
        id="userinfo-prefix",
    ),
    pytest.param(
        "https://evilgraph.microsoft.com/v1.0/me/messages/delta",
        id="prefix-lookalike",
    ),
    pytest.param("file:///etc/passwd", id="file-scheme"),
    pytest.param("//evil.example/v1.0/me/messages/delta", id="scheme-relative"),
]


# ── Helpers ──────────────────────────────────────────────────────────


def _credential():
    cred = MagicMock()
    cred.get_token = MagicMock(return_value=MagicMock(token="SECRET-GRAPH-TOKEN"))
    return cred


def _http_response(body: dict):
    r = MagicMock()
    r.status_code = 200
    r.json = MagicMock(return_value=body)
    r.raise_for_status = MagicMock()
    return r


def _recording_client(responses):
    """Patch httpx.AsyncClient and record every (url, headers) pair sent."""
    responses = list(responses)
    sent: list[tuple[str, dict]] = []

    async def fake_get(url, headers=None):
        sent.append((url, dict(headers or {})))
        return responses.pop(0)

    fake_client = MagicMock()
    fake_client.get = AsyncMock(side_effect=fake_get)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    return patch(
        "outlook_mcp.tools._delta.httpx.AsyncClient",
        return_value=fake_client,
    ), sent


# ── The validator itself ─────────────────────────────────────────────


class TestRequireGraphURL:
    @pytest.mark.parametrize("url", OFF_GRAPH_URLS)
    def test_refuses_any_url_that_is_not_graph_over_https(self, url):
        with pytest.raises(OutlookMCPError) as exc:
            require_graph_url(url, source="delta_token")
        # The model has to be able to act on this, so the offending value and
        # the reason both belong in the text the client receives.
        assert "graph.microsoft.com" in str(exc.value)

    def test_accepts_a_real_graph_url(self):
        assert require_graph_url(GRAPH_DELTA, source="delta_token") == GRAPH_DELTA

    def test_accepts_graph_host_regardless_of_case(self):
        url = "https://GRAPH.microsoft.com/v1.0/me/messages/delta"
        assert require_graph_url(url, source="delta_token") == url

    def test_refuses_empty_and_junk_cursors(self):
        for junk in ["", "   ", "not-a-url", "https://"]:
            with pytest.raises(OutlookMCPError):
                require_graph_url(junk, source="delta_token")

    def test_error_names_where_the_bad_url_came_from(self):
        with pytest.raises(OutlookMCPError) as exc:
            require_graph_url("https://evil.example/x", source="@odata.nextLink")
        assert "@odata.nextLink" in str(exc.value)


# ── The token must not reach the wire ────────────────────────────────


@pytest.mark.parametrize("url", OFF_GRAPH_URLS)
@pytest.mark.asyncio
async def test_poisoned_caller_cursor_sends_no_request_at_all(url):
    patcher, sent = _recording_client([])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url=GRAPH_DELTA,
                delta_token=url,
                page_size=10,
                resource="mail",
            )
    assert sent == [], f"bearer token was sent to {sent[0][0] if sent else ''}"


@pytest.mark.asyncio
async def test_poisoned_nextlink_stops_the_walk_before_the_second_request():
    """Page one is real Graph; its nextLink is not. The walk must stop there."""
    page_one = _http_response(
        {
            "value": [{"id": "m1"}],
            "@odata.nextLink": "https://evil.example/v1.0/me/messages/delta",
        }
    )
    patcher, sent = _recording_client([page_one, _http_response({"value": []})])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url=GRAPH_DELTA,
                delta_token=None,
                page_size=10,
                resource="mail",
            )

    assert len(sent) == 1, "followed the poisoned nextLink"
    assert sent[0][0] == GRAPH_DELTA


@pytest.mark.asyncio
async def test_poisoned_deltalink_is_not_handed_back_as_a_cursor():
    """A bad deltaLink must not be stored by the caller and replayed later."""
    body = {
        "value": [{"id": "m1"}],
        "@odata.deltaLink": "https://evil.example/v1.0/me/messages/delta",
    }
    patcher, _sent = _recording_client([_http_response(body)])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url=GRAPH_DELTA,
                delta_token=None,
                page_size=10,
                resource="mail",
            )


@pytest.mark.asyncio
async def test_no_request_in_the_suite_carries_the_bearer_off_graph():
    """The property that matters, stated once over a full multi-page walk."""
    pages = [
        _http_response(
            {
                "value": [{"id": "m1"}],
                "@odata.nextLink": f"{GRAPH_DELTA}?$skiptoken=abc",
            }
        ),
        _http_response({"value": [{"id": "m2"}], "@odata.deltaLink": GRAPH_DELTA}),
    ]
    patcher, sent = _recording_client(pages)
    with patcher:
        await fetch_delta_pages(
            _credential(),
            initial_url=GRAPH_DELTA,
            delta_token=None,
            page_size=10,
            resource="mail",
        )

    assert len(sent) == 2
    for url, headers in sent:
        assert url.startswith("https://graph.microsoft.com/")
        assert headers["Authorization"] == "Bearer SECRET-GRAPH-TOKEN"


# ── The happy path still works ───────────────────────────────────────


@pytest.mark.asyncio
async def test_a_legitimate_graph_cursor_is_still_followed():
    resume = f"{GRAPH_DELTA}?$deltatoken=xyz"
    patcher, sent = _recording_client(
        [_http_response({"value": [{"id": "m1"}], "@odata.deltaLink": GRAPH_DELTA})]
    )
    with patcher:
        items, token, has_more = await fetch_delta_pages(
            _credential(),
            initial_url="",
            delta_token=resume,
            page_size=10,
            resource="mail",
        )

    assert [i["id"] for i in items] == ["m1"]
    assert token == GRAPH_DELTA
    assert has_more is False
    assert sent[0][0] == resume


class TestParserDifferentials:
    """Agreeing with ``urlsplit`` is not the same as agreeing with httpx.

    ``urlsplit`` silently deletes tab, CR and LF before parsing, so a string it
    reads as Graph can be read as a different host by an HTTP client. Rather
    than try to match every parser, refuse anything with a control or space
    character in it and compare the whole netloc.
    """

    @pytest.mark.parametrize(
        "url",
        [
            pytest.param(
                "https://evil.example\t@graph.microsoft.com/x", id="tab-smuggled"
            ),
            pytest.param(
                "https://evil.example\n@graph.microsoft.com/x", id="lf-smuggled"
            ),
            pytest.param(
                "https://evil.example\r@graph.microsoft.com/x", id="cr-smuggled"
            ),
            pytest.param("\x01https://graph.microsoft.com/x", id="leading-control"),
            pytest.param("https://graph.microsoft.com\x00.evil/x", id="null-byte"),
            pytest.param("https://graph.microsoft.com/x\ty", id="tab-in-path"),
        ],
    )
    def test_control_characters_are_refused_outright(self, url):
        with pytest.raises(OutlookMCPError):
            require_graph_url(url, source="delta_token")

    @pytest.mark.parametrize(
        "url",
        [
            pytest.param("https://graph.microsoft.com:evil/x", id="junk-port"),
            pytest.param("https://graph.microsoft.com:443/x", id="explicit-port"),
            pytest.param("https://user:pw@graph.microsoft.com/x", id="userinfo"),
            pytest.param(
                "https://evil.example[graph.microsoft.com]/x", id="bracketed-host"
            ),
        ],
    )
    def test_netloc_must_match_exactly(self, url):
        """Graph emits neither userinfo nor a port, so equality is safe here."""
        with pytest.raises(OutlookMCPError):
            require_graph_url(url, source="delta_token")

    def test_returns_the_string_it_actually_validated(self):
        """Validating one string and sending another is how checks get bypassed."""
        padded = f"  {GRAPH_DELTA}  "
        assert require_graph_url(padded, source="delta_token") == GRAPH_DELTA


# ── A cursor is bound to the tool that issued it ─────────────────────
# Pinning the host keeps the token on Graph. It does not keep a tool on its own
# data: a cursor is a whole URL, so with only the host checked, the calendar
# delta tool would GET `/v1.0/me/messages` when handed that as its cursor and
# return what came back. Each tool now accepts its own delta endpoint and
# nothing else — for the caller's cursor, every nextLink, and the deltaLink it
# hands back.

G = "https://graph.microsoft.com"

OWN_CURSORS = [
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/inbox/messages/delta?$top=5", id="mail-first"),
    pytest.param(
        "mail",
        f"{G}/v1.0/me/mailFolders('AQMkADAwATM0MDAAMS1iNTcAZC04MgBlLTAwAi0wMAoALgAAA')"
        "/messages/delta?$deltatoken=abc",
        id="mail-deltalink-key-form",
    ),
    pytest.param(
        "mail",
        f"{G}/v1.0/me/mailFolders/AQMkADAw-_%3D/messages/delta?$skiptoken=abc",
        id="mail-encoded-id",
    ),
    pytest.param(
        "calendar",
        f"{G}/v1.0/me/calendarView/delta"
        "?startDateTime=2026-05-21T00%3A00%3A00Z&endDateTime=2026-05-28T00%3A00%3A00Z",
        id="calendar-first",
    ),
    pytest.param("calendar", f"{G}/v1.0/me/calendarView/delta?$deltatoken=abc", id="calendar"),
    pytest.param("contacts", f"{G}/v1.0/me/contacts/delta", id="contacts-first"),
    pytest.param("contacts", f"{G}/v1.0/me/contacts/delta?$skiptoken=abc", id="contacts"),
]

FOREIGN_CURSORS = [
    pytest.param(
        "calendar", f"{G}/v1.0/me/messages?$select=subject&$top=100", id="mail-via-calendar"
    ),
    pytest.param("contacts", f"{G}/v1.0/me/todo/lists", id="todo-via-contacts"),
    pytest.param("mail", f"{G}/v1.0/me/events?$top=50", id="events-via-mail"),
    pytest.param(
        "mail", f"{G}/v1.0/me/calendarView/delta?$deltatoken=abc", id="calendar-cursor-in-mail"
    ),
    pytest.param(
        "calendar", f"{G}/v1.0/me/mailFolders/inbox/messages/delta", id="mail-cursor-in-calendar"
    ),
    pytest.param("contacts", f"{G}/v1.0/me/calendarView/delta", id="calendar-cursor-in-contacts"),
    pytest.param("mail", f"{G}/beta/me/mailFolders/inbox/messages/delta", id="beta"),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/inbox/messages", id="not-delta"),
    pytest.param(
        "mail", f"{G}/v1.0/me/mailFolders/inbox/messages/delta/extra", id="trailing-segment"
    ),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/../messages/delta", id="dot-dot-as-the-id"),
    pytest.param(
        "mail",
        f"{G}/v1.0/me/mailFolders/inbox/messages/delta/../../../../events",
        id="dot-dot-after-a-good-prefix",
    ),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/%2e%2e/messages/delta", id="encoded-dot-dot"),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/a%2Fb/messages/delta", id="encoded-slash"),
    pytest.param("contacts", f"{G}/v1.0/me/contacts", id="contacts-list-not-delta"),
    pytest.param("contacts", f"{G}/", id="no-path"),
    # This server only ever asks as `/me`, and Graph answers in kind, so no
    # other spelling of a mailbox is a cursor it issued.
    pytest.param("contacts", f"{G}/v1.0/users/someone/contacts/delta", id="another-user"),
    pytest.param("contacts", f"{G}/v1.0/users('someone')/contacts/delta", id="another-user-key"),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/a%5Cb/messages/delta", id="encoded-backslash"),
    pytest.param("mail", f"{G}/v1.0/me/mailFolders/a\\b/messages/delta", id="raw-backslash"),
    # `re.IGNORECASE` on a str pattern folds a few non-ASCII letters onto ASCII
    # ones (long s, dotless i, the Kelvin sign); the paths are ASCII.
    pytest.param("contacts", f"{G}/v1.0/me/contact\u017f/delta", id="unicode-case-fold"),
]


class TestCursorIsBoundToItsTool:
    @pytest.mark.parametrize(("resource", "url"), OWN_CURSORS)
    def test_a_tools_own_endpoint_is_accepted(self, resource, url):
        assert require_graph_url(url, source="delta_token", resource=resource) == url

    @pytest.mark.parametrize(("resource", "url"), FOREIGN_CURSORS)
    def test_any_other_graph_path_is_refused(self, resource, url):
        with pytest.raises(OutlookMCPError) as exc:
            require_graph_url(url, source="delta_token", resource=resource)
        assert "delta_token" in str(exc.value)

    def test_the_host_is_still_checked_first(self):
        with pytest.raises(OutlookMCPError) as exc:
            require_graph_url(
                "https://evil.example/v1.0/me/contacts/delta",
                source="delta_token",
                resource="contacts",
            )
        assert "non-Graph" in str(exc.value)

    def test_segment_names_match_in_any_ascii_case(self):
        """Graph ignores case in these segment names, so a link in another case is the same one."""
        url = f"{G}/v1.0/me/mailfolders('AQMkADNkNAAAgEMAAAA')/messages/delta?$skiptoken=abc"
        assert require_graph_url(url, source="@odata.nextLink", resource="mail") == url

    def test_a_first_url_that_cannot_be_built_does_not_blame_a_cursor(self):
        """`initial_url` is ours, not the caller's: the refusal must not say to discard a cursor."""
        with pytest.raises(OutlookMCPError) as exc:
            require_graph_url(
                f"{G}/v1.0/me/mailFolders/a%2Fb/messages/delta",
                source="initial_url",
                resource="mail",
            )
        message = str(exc.value)
        assert "Discard this cursor" not in message
        assert "folder" in message

    def test_an_unknown_resource_is_a_programming_error_not_a_pass(self):
        with pytest.raises(KeyError):
            require_graph_url(f"{G}/v1.0/me/contacts/delta", source="x", resource="notes")


@pytest.mark.asyncio
async def test_another_tools_cursor_sends_no_request_at_all():
    patcher, sent = _recording_client([])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url="",
                delta_token=f"{G}/v1.0/me/messages?$select=subject&$top=100",
                page_size=10,
                resource="calendar",
            )
    assert sent == []


@pytest.mark.asyncio
async def test_a_nextlink_to_another_resource_stops_the_walk():
    page_one = _http_response(
        {"value": [{"id": "m1"}], "@odata.nextLink": f"{G}/v1.0/me/events?$top=50"}
    )
    patcher, sent = _recording_client([page_one, _http_response({"value": []})])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url=GRAPH_DELTA,
                delta_token=None,
                page_size=10,
                resource="mail",
            )
    assert len(sent) == 1


@pytest.mark.asyncio
async def test_a_deltalink_to_another_resource_is_not_handed_back():
    body = {"value": [{"id": "m1"}], "@odata.deltaLink": f"{G}/v1.0/me/contacts/delta"}
    patcher, _sent = _recording_client([_http_response(body)])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await fetch_delta_pages(
                _credential(),
                initial_url=GRAPH_DELTA,
                delta_token=None,
                page_size=10,
                resource="mail",
            )


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("tool", "kwargs"),
    [
        pytest.param("mail", {}, id="mail"),
        pytest.param("calendar", {"start": None, "end": None}, id="calendar"),
        pytest.param("contacts", {}, id="contacts"),
    ],
)
async def test_each_delta_tool_refuses_a_cursor_that_is_not_its_own(tool, kwargs):
    """Through the tools themselves: a path on Graph that no delta tool owns."""
    from outlook_mcp.tools.calendar_delta import list_events_delta
    from outlook_mcp.tools.contacts_delta import list_contacts_delta
    from outlook_mcp.tools.mail_delta import list_inbox_delta

    tools = {
        "mail": list_inbox_delta,
        "calendar": list_events_delta,
        "contacts": list_contacts_delta,
    }
    graph_client = MagicMock()
    graph_client.credential = _credential()

    patcher, sent = _recording_client([])
    with patcher:
        with pytest.raises(OutlookMCPError):
            await tools[tool](
                graph_client, delta_token=f"{G}/v1.0/me/todo/lists/AAMk/tasks", **kwargs
            )
    assert sent == []
