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
        "type": "singleInstance",
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

    def test_carries_type(self):
        """#69. Same reasoning as `show_as`, one field over.

        The delta endpoint takes no `$select`, so `type` is in the raw JSON
        whatever the listing asked for — the formatter simply never read it,
        which is why an agent seeding from `outlook_list_events` and refreshing
        from this tool saw the key disappear.
        """
        out = _format_event_delta(_raw_event(type="seriesMaster"))
        assert out["type"] == "seriesMaster"

    def test_absent_type_is_an_empty_string_not_a_missing_key(self):
        """The distinction the caller has to be able to make.

        A missing key is a `KeyError` in a caller that read `type` off the
        listing; an empty string is the listing's own convention for "Graph
        didn't send it". Both formatters answer the same way.
        """
        raw = _raw_event()
        del raw["type"]

        assert _format_event_delta(raw)["type"] == ""


class TestTheDeltaSummaryMirrorsTheListingSummary:
    """The parity claim in `_format_event_delta`'s docstring, as a test.

    `SKILL.md` and `outlook_changes_since` both steer recurring work to the
    delta tool, so an agent seeds from `outlook_list_events` and refreshes from
    this one. A key on one side and not the other either raises `KeyError` in
    the caller or silently downgrades the field — which is exactly what #69
    reported for `type`, and #63 reported for contacts one module over.

    The docstring said "field-for-field" and nothing pinned it. #73 could only
    scope its version of this to `show_as`, because `type` was still missing
    and the full comparison would have failed for a reason that was not #73's
    to fix. This is that test unscoped.

    Parity is asserted between the two *formatters*. The tool's own output
    carries one key more — `format_delta_item` appends `is_deleted` to every
    live item, and collapses a tombstone to `{id, is_deleted: True}` without
    calling the formatter at all. That envelope is shared with mail and
    contacts and is not what drifted.
    """

    @staticmethod
    def _listing_summary():
        from outlook_mcp.tools.calendar_read import _format_event_summary

        # `spec` so the formatter cannot silently read an attribute this
        # fixture never set: a plain MagicMock auto-creates one, and a field
        # added to the summary would then land here as a truthy stub instead
        # of the AttributeError that tells us to update both sides.
        event = MagicMock(
            spec=[
                "id",
                "subject",
                "start",
                "end",
                "location",
                "is_all_day",
                "organizer",
                "response_status",
                "is_online_meeting",
                "type",
                "show_as",
            ]
        )
        event.id = "EVT_AAA"
        event.subject = "Standup"
        event.start = MagicMock(date_time="2026-05-22T15:00:00.0000000", time_zone="UTC")
        event.end = MagicMock(date_time="2026-05-22T15:30:00.0000000", time_zone="UTC")
        event.location = MagicMock(display_name="Online")
        event.is_all_day = False
        # `name` is a MagicMock *constructor* keyword, not an attribute, so it
        # has to be assigned after the fact or the formatter reads a mock repr.
        event.organizer = MagicMock(email_address=MagicMock())
        event.organizer.email_address.name = "Lead"
        event.response_status = MagicMock(response=MagicMock(value="accepted"))
        event.is_online_meeting = True
        event.type = MagicMock(value="singleInstance")
        event.show_as = MagicMock(value="busy")
        return _format_event_summary(event)

    def test_both_summaries_carry_the_same_keys(self):
        assert set(_format_event_delta(_raw_event())) == set(self._listing_summary())

    def test_type_agrees_on_both_sides(self):
        """Key-set equality alone would pass with both sides empty.

        #69's defect was a key that was *present* and always `""`, so the
        key-set test above would have been green throughout it. The value has
        to be asserted too, from a source that carries a real one.
        """
        delta = _format_event_delta(_raw_event(type="singleInstance"))

        assert delta["type"] == self._listing_summary()["type"] == "singleInstance"


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
        "value": [_raw_event(id=f"e{i * 50 + n}") for n in range(50)],
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
        _http_response(
            {
                "value": [_raw_event(id="e1")],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/p2",
            }
        ),
        _http_response(
            {
                "value": [_raw_event(id="e2")],
                "@odata.deltaLink": "https://graph.microsoft.com/v1.0/cal-delta-final",
            }
        ),
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
