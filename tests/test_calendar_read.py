"""Tests for calendar read tools."""

from datetime import datetime, timedelta
from datetime import timezone as dt_timezone
from unittest.mock import AsyncMock, MagicMock
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import pytest

from outlook_mcp.pagination import decode_cursor_payload, encode_cursor, encode_cursor_payload
from outlook_mcp.tools.calendar_read import (
    _compute_calendar_range,
    _format_event_summary,
    _has_time_zone_database,
    _resolve_timezone,
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

    def test_summary_carries_type_so_listings_can_tell_series_apart(self):
        summary = _format_event_summary(_make_mock_event(type=MagicMock(value="seriesMaster")))

        assert summary["type"] == "seriesMaster"

    def test_summary_still_omits_the_recurrence_object(self):
        """`type` is one short string; the full pattern stays detail-only."""
        assert "recurrence" not in _format_event_summary(_make_mock_event())

    def test_concise_mode_omits_type(self):
        from outlook_mcp.tools.calendar_read import _format_event_concise

        concise = _format_event_concise(_make_mock_event(type=MagicMock(value="seriesMaster")))

        assert "type" not in concise


class TestTimezoneResolution:
    """A calendar range needs a time zone database; say so when there isn't one.

    The regression these guard: on a host with no IANA database (Windows, or a
    slim Linux image) every calendar tool failed with a bare `Error executing
    tool outlook_list_events` and no text, because `_wrap_tool_errors` holds an
    unexpected exception's message server-side.
    """

    def test_a_real_zone_resolves(self):
        assert _resolve_timezone("America/Los_Angeles").key == "America/Los_Angeles"

    def test_a_typo_names_the_zone_and_where_to_fix_it(self):
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("America/Los_Angelez")
        message = str(excinfo.value)
        assert "America/Los_Angelez" in message
        assert "config.json" in message

    def test_a_missing_database_blames_the_install_not_the_config(self, monkeypatch):
        """No database at all: the fix is the install, not the config value."""
        monkeypatch.setattr(
            "outlook_mcp.tools.calendar_read.ZoneInfo",
            MagicMock(side_effect=ZoneInfoNotFoundError("no such key")),
        )
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("America/Los_Angeles")
        message = str(excinfo.value)
        assert "tzdata" in message
        # The installable is outlook-graph-mcp; outlook-mcp is only the command.
        assert "outlook-graph-mcp" in message
        assert "config.json" not in message

    def test_a_path_shaped_key_is_refused_like_any_other_bad_zone(self):
        """zoneinfo raises plain ValueError, not the subclass, for these."""
        for key in ("/etc/localtime", "../../etc/passwd"):
            with pytest.raises(ValueError) as excinfo:
                _resolve_timezone(key)
            assert "Invalid timezone" in str(excinfo.value)

    def test_the_database_probe_sees_a_real_database(self):
        assert _has_time_zone_database() is True

    def test_the_message_survives_the_wrapper_the_model_sees(self):
        """ValueError is the anticipated-failure channel, so the text reaches the caller."""
        with pytest.raises(ValueError) as excinfo:
            _compute_calendar_range(7, None, None, "Not/AZone")
        assert "Not/AZone" in str(excinfo.value)

    def test_a_long_zone_name_is_truncated_like_every_other_echo(self):
        with pytest.raises(ValueError) as excinfo:
            _resolve_timezone("X" * 200)
        assert "X" * 51 not in str(excinfo.value)


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
