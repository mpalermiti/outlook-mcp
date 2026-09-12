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
