"""Tests for calendar delta tool."""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.tools.calendar_delta import _format_event_delta, list_events_delta

# ── Helpers ──────────────────────────────────────────────────────────


def _raw_event(**overrides) -> dict:
    """Build a raw Graph JSON event dict (the wire shape)."""
    base = {
        "id": "EVT_AAA",
        "subject": "Standup",
        "start": {"dateTime": "2026-05-22T15:00:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2026-05-22T15:30:00.0000000", "timeZone": "UTC"},
        "location": {"displayName": "Online"},
        "isAllDay": False,
        "organizer": {"emailAddress": {"address": "lead@test.com", "name": "Lead"}},
        "responseStatus": {"response": "accepted"},
        "isOnlineMeeting": True,
        "showAs": "busy",
    }
    base.update(overrides)
    return base


def _mock_graph_client():
    client = MagicMock()
    client.credential = MagicMock()
    return client


def _http_response(body: dict, status: int = 200):
    r = MagicMock()
    r.status_code = status
    r.json = MagicMock(return_value=body)
    r.raise_for_status = MagicMock()
    return r


def _async_client_with(responses):
    responses = list(responses)

    async def fake_get(url, headers=None):
        # Stash the call for later assertion
        fake_get.last_call = {"url": url, "headers": headers}
        return responses.pop(0)

    fake_get.last_call = None

    fake_client = MagicMock()
    fake_client.get = AsyncMock(side_effect=fake_get)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    return (
        patch(
            "outlook_mcp.tools._delta.httpx.AsyncClient",
            return_value=fake_client,
        ),
        fake_client,
        fake_get,
    )


# ── Formatter ────────────────────────────────────────────────────────


class TestFormatEventDelta:
    def test_maps_basic_fields(self):
        out = _format_event_delta(_raw_event())
        assert out["id"] == "EVT_AAA"
        assert out["subject"] == "Standup"
        assert out["organizer"] == "Lead"
        assert out["response_status"] == "accepted"
        assert out["is_online"] is True
        assert out["location"] == "Online"

    def test_missing_subject_and_organizer(self):
        out = _format_event_delta(
            _raw_event(subject=None, organizer={"emailAddress": None})
        )
        assert out["subject"] == "(no subject)"
        assert out["organizer"] == ""

    def test_carries_show_as(self):
        """Delta sends no $select, so showAs is already in the raw JSON.

        Confirmed against a live mailbox: `showAs` is present on every item
        /me/calendarView/delta returns. The formatter dropping it was the only
        thing standing between the caller and the field.
        """
        out = _format_event_delta(_raw_event(showAs="workingElsewhere"))
        assert out["show_as"] == "workingElsewhere"

    def test_absent_show_as_is_empty_string(self):
        """Matches calendar_read's convention, so one response has one shape."""
        raw = _raw_event()
        del raw["showAs"]

        assert _format_event_delta(raw)["show_as"] == ""

    def test_show_as_matches_the_listing_formatter(self):
        """A parity claim should be a test, not a docstring sentence (#63).

        Scoped to the field this change adds rather than the whole key set:
        the two shapes also differ on `type`, which is issue #69 and not this
        change's to fix. When #69 lands, this becomes a full key-set
        comparison and the scoping note goes away.
        """
        from outlook_mcp.tools.calendar_read import _format_event_summary

        sdk_event = MagicMock()
        sdk_event.subject = "Standup"
        sdk_event.location = MagicMock(display_name="Online")
        sdk_event.organizer = None
        sdk_event.response_status = None
        sdk_event.start = None
        sdk_event.end = None
        sdk_event.type = None
        sdk_event.show_as = MagicMock(value="oof")

        summary = _format_event_summary(sdk_event)
        delta = _format_event_delta(_raw_event(showAs="oof"))

        assert summary["show_as"] == delta["show_as"] == "oof"


# ── First call ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_first_call_uses_prefer_header_and_window():
    body = {
        "value": [_raw_event(id="e1"), _raw_event(id="e2")],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta",
    }
    patch_client, fake_client, fake_get = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start="2026-05-21T00:00:00Z",
            end="2026-05-28T00:00:00Z",
            page_size=25,
        )

    assert len(result["events"]) == 2
    assert result["delta_token"] == "https://graph.microsoft.com/v1.0/cal-delta"
    assert result["has_more"] is False

    # Prefer header should carry the maxpagesize
    headers = fake_get.last_call["headers"]
    assert headers.get("Prefer") == "odata.maxpagesize=25"

    # URL has the window encoded but no $top
    url = fake_get.last_call["url"]
    assert "startDateTime=" in url
    assert "endDateTime=" in url
    assert "$top" not in url


# ── Subsequent call with delta_token ─────────────────────────────────


@pytest.mark.asyncio
async def test_subsequent_call_uses_delta_token_url_verbatim():
    body = {
        "value": [_raw_event(id="changed")],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta-new",
    }
    prior = "https://graph.microsoft.com/v1.0/cal-delta-prior"
    patch_client, fake_client, fake_get = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start=None,
            end=None,
            delta_token=prior,
        )

    assert fake_get.last_call["url"] == prior
    assert result["delta_token"] == "https://graph.microsoft.com/v1.0/cal-delta-new"


# ── Missing window on first call ─────────────────────────────────────


@pytest.mark.asyncio
async def test_missing_start_raises_value_error():
    with pytest.raises(ValueError, match="start"):
        await list_events_delta(_mock_graph_client(), start=None, end="2026-05-28T00:00:00Z")


@pytest.mark.asyncio
async def test_missing_end_raises_value_error():
    with pytest.raises(ValueError, match="end"):
        await list_events_delta(_mock_graph_client(), start="2026-05-21T00:00:00Z", end=None)


# ── Safety cap ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_cap_reached_returns_nextlink_and_has_more():
    page = lambda i, link: {  # noqa: E731
        "value": [_raw_event(id=f"e{i*50 + n}") for n in range(50)],
        "@odata.nextLink": link,
    }
    responses = [
        _http_response(page(0, "https://graph.microsoft.com/v1.0/p2")),
        _http_response(page(1, "https://graph.microsoft.com/v1.0/p3")),
        _http_response(page(2, "https://graph.microsoft.com/v1.0/p4")),
        _http_response(page(3, "https://graph.microsoft.com/v1.0/p5")),
    ]
    patch_client, fake_client, _ = _async_client_with(responses)
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start="2026-05-21T00:00:00Z",
            end="2026-05-28T00:00:00Z",
            page_size=50,
        )

    assert len(result["events"]) == 200
    assert result["has_more"] is True
    assert result["delta_token"] == "https://graph.microsoft.com/v1.0/p5"


# ── Final page reached ───────────────────────────────────────────────


@pytest.mark.asyncio
async def test_follows_nextlink_until_deltalink():
    responses = [
        _http_response({
            "value": [_raw_event(id="e1")],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/p2",
        }),
        _http_response({
            "value": [_raw_event(id="e2")],
            "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta-final",
        }),
    ]
    patch_client, fake_client, _ = _async_client_with(responses)
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start="2026-05-21T00:00:00Z",
            end="2026-05-28T00:00:00Z",
        )

    assert len(result["events"]) == 2
    assert result["delta_token"] == "https://graph.microsoft.com/v1.0/cal-delta-final"
    assert result["has_more"] is False


# ── Tombstones ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_removed_event_collapses_to_id_only():
    body = {
        "value": [
            _raw_event(id="e1"),
            {"id": "e2-deleted", "@removed": {"reason": "deleted"}},
        ],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta",
    }
    patch_client, _, _ = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start="2026-05-21T00:00:00Z",
            end="2026-05-28T00:00:00Z",
        )

    tomb = result["events"][1]
    assert tomb == {"id": "e2-deleted", "is_deleted": True}
    assert set(tomb.keys()) == {"id", "is_deleted"}
    assert result["events"][0]["is_deleted"] is False


# ── No-changes case ─────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_no_changes_returns_empty_list_and_delta_token():
    body = {"value": [], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta-same"}
    patch_client, _, _ = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_events_delta(
            _mock_graph_client(),
            start=None,
            end=None,
            delta_token="https://graph.microsoft.com/v1.0/cal-delta-prior",
        )

    assert result["events"] == []
    assert result["delta_token"] == "https://graph.microsoft.com/v1.0/cal-delta-same"
    assert result["has_more"] is False
