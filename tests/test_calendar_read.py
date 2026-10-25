"""Tests for calendar read tools."""

from datetime import datetime, timedelta
from datetime import timezone as dt_timezone
from unittest.mock import AsyncMock, MagicMock
from zoneinfo import ZoneInfo

import pytest

from outlook_mcp.pagination import decode_cursor_payload, encode_cursor, encode_cursor_payload
from outlook_mcp.tools.calendar_read import (
    _compute_calendar_range,
    _format_event_summary,
    get_event,
    list_events,
)


def _make_mock_event(**overrides):
    """Factory for mock Graph SDK event objects."""
    event = MagicMock()
    event.id = overrides.get("id", "AAMkAG123=")
    event.subject = overrides.get("subject", "Team Meeting")
    event.start = MagicMock(
        date_time=overrides.get("start_dt", "2026-04-15T10:00:00"),
        time_zone=overrides.get("start_tz", "UTC"),
    )
    event.end = MagicMock(
        date_time=overrides.get("end_dt", "2026-04-15T11:00:00"),
        time_zone=overrides.get("end_tz", "UTC"),
    )
    event.location = MagicMock(display_name=overrides.get("location", "Room 101"))
    event.is_all_day = overrides.get("is_all_day", False)
    event.organizer = MagicMock()
    event.organizer.email_address = MagicMock()
    event.organizer.email_address.name = overrides.get("organizer_name", "Boss")
    event.organizer.email_address.address = overrides.get("organizer_email", "boss@test.com")
    event.response_status = MagicMock(
        response=MagicMock(value=overrides.get("response_status", "accepted"))
    )
    event.is_online_meeting = overrides.get("is_online", False)
    event.categories = overrides.get("categories", [])
    # Detail fields
    event.body = MagicMock(content=overrides.get("body_content", "<p>Agenda here</p>"))
    event.online_meeting = MagicMock(join_url=overrides.get("join_url", None))
    event.recurrence = overrides.get("recurrence", None)
    event.type = overrides.get("type", None)
    # Set explicitly, not left to MagicMock's auto-attribute: an auto-created
    # child is truthy and stringifies to something no Graph response contains,
    # so a formatter reading the wrong field would still look like it worked.
    event.original_start_time_zone = overrides.get("original_start_tz", "UTC")
    event.original_end_time_zone = overrides.get("original_end_tz", "UTC")
    event.show_as = overrides.get("show_as", None)
    attendee_data = overrides.get("attendees", [])
    attendees = []
    for a in attendee_data:
        att = MagicMock()
        att.email_address = MagicMock()
        att.email_address.name = a.get("name", "")
        att.email_address.address = a.get("email", "")
        att.status = MagicMock(response=MagicMock(value=a.get("response", "none")))
        attendees.append(att)
    event.attendees = attendees
    return event


class TestFormatEventSummary:
    def test_formats_basic_event(self):
        """_format_event_summary extracts all expected fields."""
        event = _make_mock_event()
        result = _format_event_summary(event)
        assert result["id"] == "AAMkAG123="
        assert result["subject"] == "Team Meeting"
        assert result["start"] == "2026-04-15T10:00:00 (UTC)"
        assert result["end"] == "2026-04-15T11:00:00 (UTC)"
        assert result["location"] == "Room 101"
        assert result["is_all_day"] is False
        assert result["organizer"] == "Boss"
        assert result["response_status"] == "accepted"
        assert result["is_online"] is False


class TestListEvents:
    async def test_list_events_returns_events(self):
        """list_events returns structured event list."""
        mock_event = _make_mock_event()
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC")
        assert result["count"] == 1
        assert result["events"][0]["subject"] == "Team Meeting"
        assert result["has_more"] is False

    async def test_list_events_days_computes_range(self):
        """days param computes start/end relative to now."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        result = await list_events(mock_client, days=3, timezone="UTC")
        assert result["count"] == 0
        # Verify calendarView was called
        mock_client.me.calendar_view.get.assert_called_once()

    async def test_list_events_with_explicit_dates(self):
        """list_events with after/before validates and uses them."""
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        result = await list_events(
            mock_client,
            after="2026-04-15T00:00:00Z",
            before="2026-04-22T00:00:00Z",
            timezone="UTC",
        )
        assert result["count"] == 0

    async def test_list_events_rejects_invalid_dates(self):
        """list_events rejects bad date strings."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError, match="Invalid datetime"):
            await list_events(mock_client, after="not-a-date", timezone="UTC")


class TestListEventsConcise:
    async def test_list_events_concise_drops_body_and_attendees(self):
        """concise=True returns is_organizer + attendees_count, no body/organizer/categories."""
        mock_event = _make_mock_event(
            attendees=[
                {"name": "Alice", "email": "alice@test.com", "response": "accepted"},
                {"name": "Bob", "email": "bob@test.com", "response": "none"},
            ],
            categories=["Blue Category"],
        )
        # is_organizer is read via getattr, set it explicitly on the mock.
        mock_event.is_organizer = True

        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC", concise=True)
        event = result["events"][0]
        assert event["id"] == "AAMkAG123="
        assert event["subject"] == "Team Meeting"
        # Concise-mode-only fields.
        assert event["is_organizer"] is True
        assert event["is_online_meeting"] is False
        assert event["attendees_count"] == 2
        # Fields that must be dropped.
        assert "organizer" not in event
        assert "response_status" not in event
        assert "categories" not in event
        assert "body" not in event
        assert "attendees" not in event

    async def test_list_events_default_keeps_full_shape(self):
        """concise=False (default) preserves the existing event shape."""
        mock_event = _make_mock_event()
        mock_client = AsyncMock()
        mock_client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[mock_event], odata_next_link=None)
        )

        result = await list_events(mock_client, days=7, timezone="UTC")
        event = result["events"][0]
        # Existing fields.
        assert "organizer" in event
        assert "response_status" in event
        assert "is_online" in event
        # New concise-only fields must NOT bleed into default mode.
        assert "is_organizer" not in event
        assert "attendees_count" not in event


def _make_calendars(*specs):
    cals = []
    for name, cal_id in specs:
        cal = MagicMock()
        cal.name = name
        cal.id = cal_id
        cals.append(cal)
    return cals


def _client_with_calendars(cals):
    """A client whose /me/calendars lists ``cals`` (one page) and whose default
    and per-calendar calendarView both answer."""
    client = MagicMock()
    client.me.calendars.get = AsyncMock(return_value=MagicMock(value=cals, odata_next_link=None))
    client.me.calendar_view.get = AsyncMock(return_value=MagicMock(value=[], odata_next_link=None))
    builder = MagicMock()
    builder.calendar_view.get = AsyncMock(
        return_value=MagicMock(value=[_make_mock_event()], odata_next_link=None)
    )
    client.me.calendars.by_calendar_id = MagicMock(return_value=builder)
    return client, builder


def _query_params(async_mock):
    """The typed query-parameters object the last call sent — what Graph sees."""
    return async_mock.call_args.kwargs["request_configuration"].query_parameters


TWO_CALENDARS = (("Calendar", "DEFAULT1="), ("Work", "WORK456="))


class TestCalendarSelection:
    """`calendar` is the only way to reach a secondary calendar.

    Before it, list_events always hit ``/me/calendarView`` — the default
    calendar — so a user whose events live in a second calendar (a class
    schedule, a shared team calendar) got an empty listing with no hint why.
    """

    async def test_omitted_calendar_keeps_the_default_path(self):
        """No selector, no extra roundtrip: /me/calendars is never listed."""
        client = AsyncMock()
        client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=[], odata_next_link=None)
        )

        await list_events(client, days=7, timezone="UTC")

        client.me.calendar_view.get.assert_called_once()
        client.me.calendars.get.assert_not_called()

    @pytest.mark.parametrize("alias", ["Primary", "PRIMARY", "", "   "])
    async def test_primary_and_blank_mean_the_default_calendar(self, alias):
        """ "primary" (any case), empty and whitespace all mean the default — one rule, not two."""
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        await list_events(client, days=7, timezone="UTC", calendar=alias)

        client.me.calendar_view.get.assert_called_once()
        client.me.calendars.get.assert_not_called()

    async def test_display_name_resolves_and_queries_that_calendar(self):
        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        result = await list_events(client, days=7, timezone="UTC", calendar="work")

        client.me.calendars.by_calendar_id.assert_called_once_with("WORK456=")
        builder.calendar_view.get.assert_called_once()
        assert result["count"] == 1
        assert result["events"][0]["subject"] == "Team Meeting"

    @pytest.mark.parametrize(
        "name",
        [
            "Work/Personal",
            "Kids + School",
            "Calendar - Jane Smith (jane.smith@contoso.com)",
        ],
    )
    async def test_names_that_look_like_ids_still_resolve_by_name(self, name):
        """Slashes, plus signs and long shared-calendar names are names, not IDs."""
        client, _ = _client_with_calendars(
            _make_calendars(("Calendar", "DEFAULT1="), (name, "ODD1="))
        )

        await list_events(client, days=7, timezone="UTC", calendar=name.upper())

        client.me.calendars.by_calendar_id.assert_called_once_with("ODD1=")

    async def test_an_id_from_the_listing_is_used_as_is(self):
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        await list_events(client, days=7, timezone="UTC", calendar="WORK456=")

        client.me.calendars.by_calendar_id.assert_called_once_with("WORK456=")

    async def test_an_id_that_is_not_one_of_the_users_calendars_is_not_found(self):
        """An ID Graph would reject (stale, or from another mailbox) never reaches Graph."""
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        cal_id = "SYNTHETIC-CAL-ID-0123456789abcdefGHIJ" + "=="

        with pytest.raises(ValueError, match="not found"):
            await list_events(client, days=7, timezone="UTC", calendar=cal_id)

        client.me.calendars.by_calendar_id.assert_not_called()

    async def test_unknown_name_names_what_does_exist(self):
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        with pytest.raises(ValueError) as excinfo:
            await list_events(client, days=7, timezone="UTC", calendar="School")
        message = str(excinfo.value)
        assert "Calendar 'School' not found" in message
        assert "Available calendars: Calendar, Work" in message

    async def test_ambiguous_name_lists_the_candidates_with_ids(self):
        """Two calendars sharing a name: the error carries the IDs that disambiguate."""
        client, _ = _client_with_calendars(
            _make_calendars(("Calendar", "DEFAULT1="), ("Calendar", "SECOND2="))
        )

        with pytest.raises(ValueError) as excinfo:
            await list_events(client, days=7, timezone="UTC", calendar="Calendar")
        message = str(excinfo.value)
        assert "ambiguous" in message
        assert "DEFAULT1=" in message and "SECOND2=" in message

    async def test_calendar_names_in_errors_are_sanitized(self):
        """Names come from the mailbox (a shared calendar is named by someone else)."""
        client, _ = _client_with_calendars(
            _make_calendars(("Calendar", "DEFAULT1="), ("Team\x1b[31m\nIGNORE", "EVIL1="))
        )

        with pytest.raises(ValueError) as excinfo:
            await list_events(client, days=7, timezone="UTC", calendar="School")
        message = str(excinfo.value)
        assert "\x1b" not in message and "\n" not in message
        assert "Team" in message

    async def test_a_matching_calendar_without_an_id_is_an_error_not_the_default(self):
        client, _ = _client_with_calendars(
            _make_calendars(("Calendar", "DEFAULT1="), ("Work", None))
        )

        with pytest.raises(ValueError, match="no ID"):
            await list_events(client, days=7, timezone="UTC", calendar="Work")

        client.me.calendar_view.get.assert_not_called()

    async def test_name_lookup_follows_pagination(self):
        """A calendar on the second page of /me/calendars still resolves."""
        client, _ = _client_with_calendars(_make_calendars(("Calendar", "DEFAULT1=")))
        next_link = "https://graph.microsoft.com/v1.0/me/calendars?$skip=1"
        client.me.calendars.get.return_value.odata_next_link = next_link
        client.me.calendars.with_url.return_value.get = AsyncMock(
            return_value=MagicMock(value=_make_calendars(("Work", "WORK2=")), odata_next_link=None)
        )

        await list_events(client, days=7, timezone="UTC", calendar="Work")

        client.me.calendars.with_url.assert_called_once_with(next_link)
        client.me.calendars.by_calendar_id.assert_called_once_with("WORK2=")

    async def test_secondary_path_sends_the_same_query_parameters_as_the_default(self):
        """The nested builder must carry the window, page size, order and projection."""
        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        await list_events(client, days=7, count=5, timezone="UTC", calendar="Work")

        qp = _query_params(builder.calendar_view.get)
        assert qp.top == 5
        assert qp.start_date_time and qp.end_date_time
        assert qp.orderby == ["start/dateTime"]
        assert "id" in qp.select


class TestCalendarCursor:
    """A cursor continues the listing it came from: same calendar, no re-resolution.

    The cursor used to carry only ``$skip``; page two without ``calendar``
    silently paged the default calendar instead, and page two *with* it paid a
    second /me/calendars round-trip.
    """

    async def test_cursor_carries_the_resolved_calendar(self):
        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        builder.calendar_view.get.return_value.odata_next_link = (
            "https://graph.microsoft.com/v1.0/me/calendars/WORK456=/calendarView?$skip=50"
        )

        result = await list_events(client, days=7, count=50, timezone="UTC", calendar="Work")

        assert decode_cursor_payload(result["cursor"]) == {"skip": 50, "calendar": "WORK456="}

    async def test_default_calendar_cursor_records_the_default(self):
        client = AsyncMock()
        client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(
                value=[],
                odata_next_link="https://graph.microsoft.com/v1.0/me/calendarView?$skip=50",
            )
        )

        result = await list_events(client, days=7, count=50, timezone="UTC")

        assert decode_cursor_payload(result["cursor"]) == {"skip": 50, "calendar": None}

    async def test_a_cursor_continues_its_calendar_without_re_listing(self):
        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        cursor = encode_cursor_payload({"skip": 50, "calendar": "WORK456="})

        await list_events(client, days=7, timezone="UTC", cursor=cursor)

        client.me.calendars.get.assert_not_called()
        client.me.calendars.by_calendar_id.assert_called_once_with("WORK456=")
        assert _query_params(builder.calendar_view.get).skip == 50

    async def test_a_cursor_wins_over_a_conflicting_calendar_argument(self):
        """Page two of the default calendar stays on the default calendar."""
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        cursor = encode_cursor_payload({"skip": 50, "calendar": None})

        await list_events(client, days=7, timezone="UTC", cursor=cursor, calendar="Work")

        client.me.calendar_view.get.assert_called_once()
        client.me.calendars.get.assert_not_called()

    async def test_a_legacy_cursor_without_a_calendar_resolves_the_argument(self):
        """Cursors minted before the calendar key still work with an explicit calendar."""
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        await list_events(client, days=7, timezone="UTC", cursor=encode_cursor(50), calendar="Work")

        client.me.calendars.by_calendar_id.assert_called_once_with("WORK456=")

    async def test_a_malformed_cursor_fails_before_any_network_call(self):
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        with pytest.raises(ValueError, match="Invalid pagination cursor"):
            await list_events(client, days=7, timezone="UTC", cursor="garbage", calendar="Work")

        client.me.calendars.get.assert_not_called()

    async def test_the_calendar_id_inside_a_cursor_is_validated(self):
        """The cursor is caller-held state, and its calendar goes into a URL path."""
        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        cursor = encode_cursor_payload({"skip": 0, "calendar": "../users/x@y.com?x"})

        with pytest.raises(ValueError, match="invalid characters"):
            await list_events(client, days=7, timezone="UTC", cursor=cursor)

        client.me.calendars.by_calendar_id.assert_not_called()


class TestGetEvent:
    async def test_get_event_validates_id(self):
        """get_event rejects invalid event IDs."""
        mock_client = AsyncMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await get_event(mock_client, event_id="bad id with spaces!")

    async def test_get_event_returns_full_detail(self):
        """get_event returns full event detail including body and attendees."""
        mock_event = _make_mock_event(
            body_content="<p>Full agenda</p>",
            join_url="https://teams.microsoft.com/join/123",
            attendees=[
                {"name": "Alice", "email": "alice@test.com", "response": "accepted"},
                {"name": "Bob", "email": "bob@test.com", "response": "tentativelyAccepted"},
            ],
            categories=["Blue Category"],
        )
        builder = MagicMock()
        builder.get = AsyncMock(return_value=mock_event)
        mock_client = MagicMock()
        mock_client.me.events.by_event_id = MagicMock(return_value=builder)

        result = await get_event(mock_client, event_id="AAMkAG123=")
        assert result["id"] == "AAMkAG123="
        assert result["body"] == "<p>Full agenda</p>"
        assert result["online_meeting_url"] == "https://teams.microsoft.com/join/123"
        assert len(result["attendees"]) == 2
        assert result["attendees"][0]["name"] == "Alice"
        assert result["attendees"][0]["response"] == "accepted"
        assert result["categories"] == ["Blue Category"]


class TestEventDetailAnchorZone:
    """The detail payload has to say which zone the event is anchored in.

    ``start``/``end`` come back in UTC because that is what Graph returns with
    no ``Prefer: outlook.timezone`` header, and what every other datetime in
    this server's responses is. That makes the UTC in ``start`` say nothing
    about the anchor — an event anchored in America/Los_Angeles and one
    anchored in UTC are indistinguishable in it, while behaving differently the
    week daylight saving ends. ``original_start_time_zone`` is the difference.

    ``get_event`` sends no ``$select`` at all, so there is no
    ``_PAIRS`` row to add here: Graph returns every property and the
    formatter-narrower-than-select bug class cannot arise. Adding a select
    purely so a row became possible would *create* the failure mode the row
    would then guard.
    """

    def _detail(self, **overrides):
        from outlook_mcp.tools.calendar_read import _format_event_detail

        return _format_event_detail(_make_mock_event(**overrides))

    def test_the_anchor_zone_is_reported(self):
        detail = self._detail(
            original_start_tz="America/Los_Angeles",
            original_end_tz="America/Los_Angeles",
        )

        assert detail["original_start_time_zone"] == "America/Los_Angeles"
        assert detail["original_end_time_zone"] == "America/Los_Angeles"

    def test_a_utc_anchor_is_distinguishable_from_a_named_one(self):
        """The assertion that fails if the field is hardcoded or read off start.

        Both events below report the same UTC ``start``; only the anchor tells
        them apart, which is the entire reason the field is here.
        """
        named = self._detail(start_dt="2026-10-28T16:00:00", original_start_tz="Europe/London")
        utc = self._detail(start_dt="2026-10-28T16:00:00", original_start_tz="UTC")

        assert named["start"] == utc["start"]
        assert named["original_start_time_zone"] != utc["original_start_time_zone"]

    def test_the_summary_does_not_carry_it(self):
        """Listings are unchanged; the anchor zone is detail-only, deliberately.

        When this was written the summary/delta parity claim was prose only and
        already false, so keeping the anchor zone off the summary avoided
        leaning on a claim nothing enforced. #69 closed that gap: the two
        formatters now agree key-for-key and
        `test_calendar_delta.py::…::test_both_summaries_carry_the_same_keys`
        enforces it. So this assertion has grown a second job — add the anchor
        zone to the summary and that parity test fails too, unless the delta
        formatter grows it in the same commit.
        """
        assert "original_start_time_zone" not in _format_event_summary(_make_mock_event())


class TestEventDetailRecurrence:
    """#41 follow-on: recurrence must come back as the same JSON shape create accepts."""

    def _detail(self, **overrides):
        from outlook_mcp.tools.calendar_read import _format_event_detail

        return _format_event_detail(_make_mock_event(**overrides))

    def test_no_recurrence_is_none(self):
        assert self._detail()["recurrence"] is None

    def test_recurrence_is_a_dict_not_a_repr(self):
        """str(PatternedRecurrence) leaked a ~500-char Python repr into the payload."""
        from outlook_mcp.tools._recurrence import build_event_recurrence

        pr = build_event_recurrence(
            {
                "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
                "range": {"type": "noEnd", "startDate": "2026-09-07"},
            },
            start="2026-09-07T12:30:00Z",
        )

        assert self._detail(recurrence=pr)["recurrence"] == {
            "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
            "range": {"type": "noEnd", "startDate": "2026-09-07"},
        }

    def test_series_master_type_is_surfaced(self):
        """`type` is what tells a client the series actually took — see #41's repro."""
        assert self._detail(type=MagicMock(value="seriesMaster"))["type"] == "seriesMaster"

    def test_missing_type_is_empty_string(self):
        assert self._detail()["type"] == ""

    def test_summary_carries_type_so_listings_can_tell_an_occurrence_apart(self):
        """`occurrence`, not `seriesMaster` — which is what a listing sees.

        This formatter feeds both shapes, so a master is a real value *for
        `outlook_get_event`* (asserted above). `list_events` reads
        `calendarView`, which returns expanded instances, so what tells
        recurring work apart there is `occurrence`/`exception` against
        `singleInstance`. The old name and fixture said otherwise, and an
        agent filtering a listing for `seriesMaster` matches nothing.
        """
        summary = _format_event_summary(_make_mock_event(type=MagicMock(value="occurrence")))

        assert summary["type"] == "occurrence"

    def test_summary_still_omits_the_recurrence_object(self):
        """`type` is one short string; the full pattern stays detail-only."""
        assert "recurrence" not in _format_event_summary(_make_mock_event())

    def test_concise_mode_omits_type(self):
        from outlook_mcp.tools.calendar_read import _format_event_concise

        # `occurrence`, like its sibling above: this formatter only ever sees a
        # calendarView response, which never carries a master.
        concise = _format_event_concise(_make_mock_event(type=MagicMock(value="occurrence")))

        assert "type" not in concise


class TestShowAs:
    """Outlook's "Show as" (Graph `showAs`) on the three read shapes."""

    def test_summary_carries_show_as(self):
        summary = _format_event_summary(_make_mock_event(show_as=MagicMock(value="tentative")))

        assert summary["show_as"] == "tentative"

    def test_detail_carries_show_as(self):
        from outlook_mcp.tools.calendar_read import _format_event_detail

        detail = _format_event_detail(_make_mock_event(show_as=MagicMock(value="oof")))

        assert detail["show_as"] == "oof"

    def test_missing_show_as_is_empty_string(self):
        """Follows `type`'s convention on this shape, so one response has one."""
        assert _format_event_summary(_make_mock_event())["show_as"] == ""

    def test_a_plain_string_show_as_survives_the_enum_fallback(self):
        """The `else str(...)` half of the extraction, which the SDK path never hits.

        `event.show_as` is a `FreeBusyStatus` today, so every other test here
        exercises the `.value` branch and the fallback is dead weight nothing
        would notice breaking. It is not speculative: the same two-branch idiom
        guards `type` and `response_status` in this formatter because kiota has
        handed back raw strings before. Pinning it means a refactor that drops
        the fallback fails here rather than in a caller's listing.
        """
        assert _format_event_summary(_make_mock_event(show_as="tentative"))["show_as"] == (
            "tentative"
        )

    def test_concise_mode_omits_show_as(self):
        """A deliberate trade, not an oversight — see _format_event_concise."""
        from outlook_mcp.tools.calendar_read import _format_event_concise

        concise = _format_event_concise(_make_mock_event(show_as=MagicMock(value="free")))

        assert "show_as" not in concise


class TestTheListingSendsTheSelectTheGuardChecks:
    """The seam `test_select_covers_the_formatter.py` cannot see.

    That guard reads `_SUMMARY_SELECT` / `_CONCISE_SELECT` off the module and
    compares them to what the formatters read. It never calls `list_events`, so
    it would stay green if the listing stopped sending those constants — an
    inline literal beside them, a field appended at the call site, a branch
    picking the wrong one. The constants would be provably consistent with the
    formatters and provably not what Graph was asked for, which is #69 again
    with an extra step.

    Two assertions, one per mode, on the typed query-parameters object the SDK
    actually receives. `.select` is a list because `build_request_config`
    splits the comma form, so compare against the same split rather than
    re-typing the field names — a hand-typed copy beside a derived one is what
    drifts.

    This replaces #73's `test_list_events_asks_graph_for_show_as` and
    `test_concise_listing_does_not_pay_for_show_as`, which asserted one field
    in each direction. The `_PAIRS` rows now cover every field in both
    directions; what was left over is only this — that the constants are what
    we send.
    """

    @staticmethod
    async def _sent_select(**kwargs):
        mock_client = MagicMock()
        response = MagicMock(value=[], odata_next_link=None)
        mock_client.me.calendar_view.get = AsyncMock(return_value=response)
        mock_client.me.calendars = MagicMock()

        await list_events(mock_client, days=1, **kwargs)
        return _query_params(mock_client.me.calendar_view.get).select

    async def test_the_normal_listing_sends_the_summary_select(self):
        from outlook_mcp.tools.calendar_read import _SUMMARY_SELECT

        assert await self._sent_select() == _SUMMARY_SELECT.split(",")

    async def test_the_concise_listing_sends_the_concise_select(self):
        from outlook_mcp.tools.calendar_read import _CONCISE_SELECT

        assert await self._sent_select(concise=True) == _CONCISE_SELECT.split(",")


class TestTimezoneResolutionReachesTheCaller:
    """The resolver itself is tested in `test_validation.py`; this is the wiring.

    What it pins is that a bad `config.timezone` still reaches the model through
    the calendar path rather than dying as a message-free crash — the whole
    reason the resolver raises `ValueError` instead of letting
    `ZoneInfoNotFoundError` out.
    """

    def test_the_message_survives_the_wrapper_the_model_sees(self):
        with pytest.raises(ValueError) as excinfo:
            _compute_calendar_range(7, None, None, "Not/AZone")
        assert "Not/AZone" in str(excinfo.value)


class TestCalendarRangeWithoutBounds:
    def test_days_widens_the_window_from_now(self):
        start, end = _compute_calendar_range(7, None, None, "America/Los_Angeles")
        assert start < end
        assert start.endswith("Z") and end.endswith("Z")

    def test_explicit_bounds_win_over_days(self):
        start, end = _compute_calendar_range(
            7, "2026-09-11T00:00:00", "2026-09-12T00:00:00", "America/Los_Angeles"
        )
        # Midnight Pacific on a DST date is 07:00Z.
        assert start == "2026-09-11T07:00:00Z"
        assert end == "2026-09-12T07:00:00Z"


class TestAmbiguousHourKeepsItsFold:
    """The defaulted `start` must survive a DST fall-back unchanged.

    PEP 495: arithmetic on an aware datetime ignores `fold` and yields fold=0.
    `datetime.now(tz) + timedelta(days=0)` therefore preserves every visible
    field and resets the invisible one, and `.astimezone()` reads exactly that
    field to choose the offset — so during the repeated hour the window opened
    an hour early and reported ended events as upcoming.

    Shape assertions cannot see this: `start < end` stays true, and the two
    datetimes even compare EQUAL to each other, because PEP 495 has intra-zone
    comparison ignore fold as well. Only the UTC value pins it.
    """

    ZONE = "America/Los_Angeles"
    # 09:30Z on 2026-11-01 is the SECOND pass through 01:30 local (fold=1).
    AMBIGUOUS_UTC = datetime(2026, 11, 1, 9, 30, tzinfo=dt_timezone.utc)

    @pytest.fixture
    def frozen_now(self, monkeypatch):
        """Freeze `datetime.now(tz)` inside the repeated hour, with fold intact."""
        real = datetime

        class FrozenDatetime(real):
            @classmethod
            def now(cls, tz=None):
                if tz is None:
                    return real.fromtimestamp(self.AMBIGUOUS_UTC.timestamp())
                return real.fromtimestamp(self.AMBIGUOUS_UTC.timestamp(), tz)

        monkeypatch.setattr("outlook_mcp.tools.calendar_read.datetime", FrozenDatetime)

    def test_the_frozen_moment_really_is_ambiguous(self):
        """Guard the premise: if fold were 0 the test below would prove nothing."""
        now = datetime.fromtimestamp(self.AMBIGUOUS_UTC.timestamp(), ZoneInfo(self.ZONE))
        assert now.fold == 1, "this instant is no longer in the repeated hour"
        assert now.hour == 1 and now.minute == 30

    def test_start_is_the_second_pass_not_the_first(self, frozen_now):
        start, _ = _compute_calendar_range(7, None, None, self.ZONE)
        assert start == "2026-11-01T09:30:00Z"  # 08:30Z would be the first pass

    def test_window_is_still_the_requested_length(self, frozen_now):
        """Fixing start must not shorten the window; end keeps local-day arithmetic."""
        start, end = _compute_calendar_range(7, None, None, self.ZONE)
        assert start < end
        # 7 local days across a fall-back is 7*24 + 1 hours.
        delta = datetime.strptime(end, "%Y-%m-%dT%H:%M:%SZ") - datetime.strptime(
            start, "%Y-%m-%dT%H:%M:%SZ"
        )
        assert delta == timedelta(days=7)

    def test_explicit_after_is_unaffected(self, frozen_now):
        """Only the defaulted path reads the clock at all."""
        start, _ = _compute_calendar_range(7, "2026-11-01T00:00:00", None, self.ZONE)
        assert start == "2026-11-01T07:00:00Z"


# ── event keys ────────────────────────────────────────────────────────

_KEY_PROP = "String {51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey"
_KEY_EXPAND = f"singleValueExtendedProperties($filter=id eq '{_KEY_PROP}')"


def _keyed(event, value, prop_id=_KEY_PROP):
    prop = MagicMock()
    prop.id = prop_id
    prop.value = value
    event.single_value_extended_properties = [prop]
    return event


def _unkeyed(event):
    event.single_value_extended_properties = None
    return event


class TestTheKeyPropertyIsFixed:
    def test_the_property_id_never_changes(self):
        """Events already stamped become unfindable if this string moves."""
        from outlook_mcp.tools.calendar_read import EVENT_KEY_PROPERTY_ID

        assert EVENT_KEY_PROPERTY_ID == _KEY_PROP


class TestGetEventReturnsItsKey:
    @staticmethod
    def _client(event):
        builder = MagicMock()
        builder.get = AsyncMock(return_value=event)
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)
        return client, builder

    async def test_the_read_expands_only_the_key_property(self):
        client, builder = self._client(_unkeyed(_make_mock_event()))

        await get_event(client, event_id="AAMkAG123=")

        assert _query_params(builder.get).expand == [_KEY_EXPAND]
        # Still no $select: every other property comes back as before.
        assert _query_params(builder.get).select is None

    async def test_a_keyed_event_reports_its_key(self):
        client, _ = self._client(_keyed(_make_mock_event(), "booking:ABC123"))

        assert (await get_event(client, event_id="AAMkAG123="))["key"] == "booking:ABC123"

    async def test_graph_may_spell_the_property_id_in_another_case(self):
        client, _ = self._client(_keyed(_make_mock_event(), "booking:ABC123", _KEY_PROP.upper()))

        assert (await get_event(client, event_id="AAMkAG123="))["key"] == "booking:ABC123"

    async def test_an_unkeyed_event_reports_none(self):
        client, _ = self._client(_unkeyed(_make_mock_event()))

        assert (await get_event(client, event_id="AAMkAG123="))["key"] is None

    async def test_another_property_is_not_a_key(self):
        other = "String {00000000-0000-0000-0000-000000000000} Name EventKey"
        client, _ = self._client(_keyed(_make_mock_event(), "someone-else", other))

        assert (await get_event(client, event_id="AAMkAG123="))["key"] is None


class TestListEventsReturnsKeys:
    @staticmethod
    def _client(events):
        client = MagicMock()
        client.me.calendar_view.get = AsyncMock(
            return_value=MagicMock(value=events, odata_next_link=None)
        )
        return client

    async def test_the_listing_expands_the_key_and_shows_it_where_set(self):
        client = self._client(
            [
                _keyed(_make_mock_event(id="K1="), "booking:ABC123"),
                _unkeyed(_make_mock_event(id="U1=")),
            ]
        )

        result = await list_events(client, days=7, timezone="UTC")

        assert _query_params(client.me.calendar_view.get).expand == [_KEY_EXPAND]
        keyed, plain = result["events"]
        assert keyed["key"] == "booking:ABC123"
        # Absent rather than null: most events have no key, and a null on every
        # row of a page is paid for on every listing.
        assert "key" not in plain

    async def test_concise_sends_no_expand_and_returns_no_key(self):
        client = self._client([_keyed(_make_mock_event(id="K1="), "booking:ABC123")])

        result = await list_events(client, days=7, timezone="UTC", concise=True)

        assert _query_params(client.me.calendar_view.get).expand is None
        assert "key" not in result["events"][0]

    async def test_a_secondary_calendar_listing_expands_it_too(self):
        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        builder.calendar_view.get.return_value.value = [
            _keyed(_make_mock_event(id="K1="), "booking:ABC123")
        ]

        result = await list_events(client, days=7, timezone="UTC", calendar="Work")

        assert _query_params(builder.calendar_view.get).expand == [_KEY_EXPAND]
        assert result["events"][0]["key"] == "booking:ABC123"


class TestFindEventsByKey:
    @staticmethod
    def _client(events, next_link=None):
        client = MagicMock()
        client.me.events.get = AsyncMock(
            return_value=MagicMock(value=events, odata_next_link=next_link)
        )
        return client

    async def test_the_filter_matches_this_property_and_this_value(self):
        from outlook_mcp.tools.calendar_read import _SUMMARY_SELECT, find_events_by_key

        client = self._client([])

        await find_events_by_key(client, "booking:ABC123")

        params = _query_params(client.me.events.get)
        assert params.filter == (
            f"singleValueExtendedProperties/Any(ep: ep/id eq '{_KEY_PROP}' "
            "and ep/value eq 'booking:ABC123')"
        )
        assert params.expand == [_KEY_EXPAND]
        assert params.select == [f.strip() for f in _SUMMARY_SELECT.split(",")]
        assert params.top == 25

    async def test_a_quote_in_the_key_is_doubled_in_the_literal(self, monkeypatch):
        """The validator refuses a quote today; the filter must not rely on it."""
        from outlook_mcp.tools import calendar_read

        monkeypatch.setattr(calendar_read, "validate_event_key", lambda key: key)
        client = self._client([])

        await calendar_read.find_events_by_key(client, "o'brien:1")

        assert _query_params(client.me.events.get).filter.endswith("and ep/value eq 'o''brien:1')")

    async def test_matches_come_back_in_the_listing_shape_with_their_key(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        master = _keyed(_make_mock_event(id="M1=", type=MagicMock(value="seriesMaster")), "k:1")
        client = self._client([master])

        result = await find_events_by_key(client, "k:1")

        assert result["key"] == "k:1"
        assert result["count"] == 1
        assert result["has_more"] is False
        found = result["events"][0]
        assert found["id"] == "M1="
        assert found["type"] == "seriesMaster"
        assert found["key"] == "k:1"
        assert {"subject", "start", "end", "show_as"} <= set(found)

    async def test_nothing_found_is_an_empty_answer_not_an_error(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        result = await find_events_by_key(self._client([]), "k:1")

        assert result == {"key": "k:1", "events": [], "count": 0, "has_more": False}

    async def test_more_than_a_page_says_so(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        events = [_keyed(_make_mock_event(id=f"E{i}="), "k:1") for i in range(25)]
        client = self._client(
            events, next_link="https://graph.microsoft.com/v1.0/me/events?$skip=25"
        )

        result = await find_events_by_key(client, "k:1")

        assert result["count"] == 25
        assert result["has_more"] is True

    async def test_an_invalid_key_is_refused_before_any_request(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        client = self._client([])
        client.me.calendars.get = AsyncMock()

        with pytest.raises(ValueError, match="Invalid event key"):
            await find_events_by_key(client, "it's", calendar="Work")

        client.me.events.get.assert_not_called()
        client.me.calendars.get.assert_not_called()

    async def test_a_named_calendar_is_searched_instead_of_the_default(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        client, builder = _client_with_calendars(_make_calendars(*TWO_CALENDARS))
        builder.events.get = AsyncMock(return_value=MagicMock(value=[], odata_next_link=None))
        client.me.events.get = AsyncMock()

        await find_events_by_key(client, "k:1", calendar="work")

        client.me.calendars.by_calendar_id.assert_called_once_with("WORK456=")
        builder.events.get.assert_called_once()
        client.me.events.get.assert_not_called()
        assert "ep/value eq 'k:1'" in _query_params(builder.events.get).filter

    async def test_an_unknown_calendar_is_refused_naming_what_exists(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        client, _ = _client_with_calendars(_make_calendars(*TWO_CALENDARS))

        with pytest.raises(ValueError, match="not found"):
            await find_events_by_key(client, "k:1", calendar="Holidays")


class TestEventKeyQueriesOnTheWire:
    """The URLs the real SDK pipeline sends, decoded the way Graph decodes them.

    `mail_read` builds its one extended-property `$expand` as a raw URL because
    kiota once dropped the inner `$filter`. These go through the typed request
    builders, so the test that matters is what reaches the transport.
    """

    @staticmethod
    def _client(answer: dict):
        import httpx
        from azure.core.credentials import AccessToken
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as _gra
        from msgraph_core import GraphClientFactory

        class _Credential:
            def get_token(self, *args, **kwargs):
                return AccessToken("not-a-real-token", 9_999_999_999)

        sent: list[httpx.Request] = []

        def handle(request: httpx.Request) -> httpx.Response:
            sent.append(request)
            return httpx.Response(200, json=answer)

        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options,
            client=httpx.AsyncClient(transport=httpx.MockTransport(handle)),
        )
        auth = AzureIdentityAuthenticationProvider(
            _Credential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client)), sent

    @staticmethod
    def _query(request) -> dict[str, str]:
        from urllib.parse import parse_qs

        return {k: v[0] for k, v in parse_qs(request.url.query.decode()).items()}

    async def test_the_lookup_filter_reaches_graph_intact(self):
        from outlook_mcp.tools.calendar_read import find_events_by_key

        client, sent = self._client({"value": []})

        await find_events_by_key(client, "flight:ab12cd:LHR-JFK")

        (request,) = sent
        assert request.url.path == "/v1.0/me/events"
        query = self._query(request)
        assert query["$filter"] == (
            f"singleValueExtendedProperties/Any(ep: ep/id eq '{_KEY_PROP}' "
            "and ep/value eq 'flight:ab12cd:LHR-JFK')"
        )
        assert query["$expand"] == _KEY_EXPAND
        assert query["$top"] == "25"

    async def test_the_key_expand_reaches_graph_on_a_read_and_a_listing(self):
        client, sent = self._client(
            {"id": "AAMkAG123=", "subject": "s", "value": [], "singleValueExtendedProperties": []}
        )

        await get_event(client, event_id="AAMkAG123=")
        await list_events(client, days=7, timezone="UTC")

        read, listing = sent
        assert read.url.path == "/v1.0/me/events/AAMkAG123="
        assert self._query(read) == {"$expand": _KEY_EXPAND}
        assert listing.url.path == "/v1.0/me/calendarView"
        assert self._query(listing)["$expand"] == _KEY_EXPAND

    async def test_a_key_graph_returns_is_read_back(self):
        client, _ = self._client(
            {
                "id": "AAMkAG123=",
                "subject": "s",
                "singleValueExtendedProperties": [
                    {"id": _KEY_PROP.lower(), "value": "booking:ABC123"}
                ],
            }
        )

        assert (await get_event(client, event_id="AAMkAG123="))["key"] == "booking:ABC123"
