"""Writing into a calendar or To Do list someone else owns needs `mail_send`.

A calendar or list another person shared with this account is theirs to read:
whatever is written there — an event, a file on one, a task, a checklist step —
is shown to them, the same as an invitation would be. An invitation already
needs `mail_send` under a policy that withholds it, so content added to
someone else's calendar or list does too.

Ownership is looked up only under a policy that withholds `mail_send`; with
`mail_send` allowed (or no policy at all) the write costs no extra request.
What cannot be shown to be the user's own counts as someone else's, including
a lookup Graph answers with a 403. Any other failed lookup (a 404, a lapsed
sign-in, the network) is reported as it is, and nothing is written.

Deleting, completing, ticking a step and RSVPs add no content and are left
alone.
"""

from __future__ import annotations

import json
from unittest.mock import AsyncMock, MagicMock

import pytest
from kiota_abstractions.api_error import APIError
from msgraph.generated.models.calendar import Calendar
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.todo_task_list import TodoTaskList
from msgraph.generated.models.wellknown_list_name import WellknownListName

from outlook_mcp.config import Config
from outlook_mcp.errors import LEAVE_SETTINGS_TO_THE_USER, PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.calendar_attachments import attach_to_event, remove_event_attachment
from outlook_mcp.tools.calendar_write import create_event, delete_event, rsvp, update_event
from outlook_mcp.tools.todo import (
    add_checklist_item,
    complete_task,
    create_task,
    delete_checklist_item,
    delete_task,
    list_task_lists,
    update_checklist_item,
    update_task,
)
from outlook_mcp.tools.todo_attachments import delete_task_attachment, upload_task_attachment

_ME = "Me@Outlook.com"
_JANE = "jane@example.com"


def _cfg(tmp_path, *categories, **kwargs):
    return Config(
        client_id="test",
        allow_categories=list(categories),
        attachments_dir=str(tmp_path),
        **kwargs,
    )


def _api_error(status):
    """What the SDK raises for a Graph error status (``ODataError`` is a subclass)."""
    return APIError(f"Graph answered {status}", response_status_code=status)


# A lookup failure Graph did not answer with a 403: reported as it is, never as
# a `mail_send` refusal. Factories, so each test raises its own exception.
_OTHER_FAILURES = [
    pytest.param(lambda: _api_error(404), id="404"),
    pytest.param(lambda: _api_error(401), id="401"),
    pytest.param(lambda: RuntimeError("connection reset"), id="not-a-graph-error"),
]

_ACCESS_DENIED = {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}}
_NOT_FOUND = {"error": {"code": "ErrorItemNotFound", "message": "Not found."}}


def _wire_client(answers):
    """The real SDK pipeline over a mock transport, answering by (method, path)."""
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
        key = (request.method, request.url.path)
        if key in answers:
            status, body = answers[key]
            return httpx.Response(status, json=body) if body is not None else httpx.Response(status)
        raise AssertionError(f"unexpected request: {request.method} {request.url}")

    http_client = GraphClientFactory.create_with_default_middleware(
        options=_gra.options,
        client=httpx.AsyncClient(transport=httpx.MockTransport(handle)),
    )
    auth = AzureIdentityAuthenticationProvider(
        _Credential(), scopes=["https://graph.microsoft.com/.default"]
    )
    return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client)), sent


# ── calendars: create_event(calendar=…) ──────────────────────────────────────


def _calendar(cal_id, name, owner, *, is_default=False):
    """A calendar as the `/me/calendars` listing returns it; `owner=None` is one
    listed without an owner at all."""
    cal = Calendar(id=cal_id, name=name, can_edit=True, is_default_calendar=is_default)
    if owner is not None:
        cal.owner = EmailAddress(name="Owner", address=owner)
    return cal


def _listing():
    return [
        _calendar("DEF=", "Calendar", _ME, is_default=True),
        # Same owner as the default, in another case: still the user's own.
        _calendar("KIDS=", "Kids", _ME.lower()),
        _calendar("JANE=", "Jane", _JANE),
    ]


def _create_client(calendars):
    client = MagicMock()
    client.me.calendars.get = AsyncMock(
        return_value=MagicMock(value=list(calendars), odata_next_link=None)
    )
    client.me.events.post = AsyncMock(return_value=MagicMock(id="AAMkNEW=", subject="Dentist"))
    secondary = MagicMock()
    secondary.events.post = AsyncMock(return_value=MagicMock(id="AAMkNEW=", subject="Dentist"))
    client.me.calendars.by_calendar_id = MagicMock(return_value=secondary)
    client.me.calendar.get = AsyncMock()
    return client, secondary


async def _create(client, config, calendar):
    return await create_event(
        client,
        subject="Dentist",
        start="2026-11-03T09:00:00",
        end="2026-11-03T09:30:00",
        calendar=calendar,
        config=config,
    )


class TestCreatingInAChosenCalendar:
    async def test_an_own_calendar_besides_the_default_is_calendar_only(self, tmp_path):
        client, secondary = _create_client(_listing())

        result = await _create(client, _cfg(tmp_path, "calendar_write"), "Kids")

        assert result["status"] == "created"
        client.me.calendars.by_calendar_id.assert_called_once_with("KIDS=")
        secondary.events.post.assert_called_once()
        # Decided from the listing the resolver already made, nothing more.
        client.me.calendars.get.assert_called_once()
        client.me.calendar.get.assert_not_called()

    async def test_the_default_calendar_by_name_is_always_own(self, tmp_path):
        """Even listed with no owner: the default calendar is the user's."""
        client, secondary = _create_client(
            [_calendar("DEF=", "Calendar", None, is_default=True), _calendar("J=", "J", _JANE)]
        )

        await _create(client, _cfg(tmp_path, "calendar_write"), "Calendar")

        secondary.events.post.assert_called_once()

    async def test_a_calendar_someone_else_owns_needs_mail_send(self, tmp_path):
        client, secondary = _create_client(_listing())

        with pytest.raises(PermissionDeniedError) as exc:
            await _create(client, _cfg(tmp_path, "calendar_write"), "Jane")

        message = str(exc.value)
        assert "mail_send" in message
        assert "owner will see" in message
        assert "sends email" not in message
        assert LEAVE_SETTINGS_TO_THE_USER in message
        secondary.events.post.assert_not_called()
        client.me.events.post.assert_not_called()

    async def test_it_passes_once_mail_send_is_allowed(self, tmp_path):
        client, secondary = _create_client(_listing())

        await _create(client, _cfg(tmp_path, "calendar_write", "mail_send"), "Jane")

        secondary.events.post.assert_called_once()
        client.me.calendars.get.assert_called_once()
        client.me.calendar.get.assert_not_called()

    async def test_an_unrestricted_server_writes_there_as_before(self, tmp_path):
        client, secondary = _create_client(_listing())

        await _create(client, _cfg(tmp_path), "Jane")

        secondary.events.post.assert_called_once()
        client.me.calendars.get.assert_called_once()

    @pytest.mark.parametrize(
        "owner",
        [None, EmailAddress(name="Someone"), EmailAddress(address="")],
        ids=["no-owner", "owner-without-address", "empty-address"],
    )
    async def test_a_calendar_whose_owner_is_unknown_is_refused(self, tmp_path, owner):
        mystery = _calendar("MYST=", "Mystery", None)
        mystery.owner = owner
        client, secondary = _create_client([*_listing(), mystery])

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await _create(client, _cfg(tmp_path, "calendar_write"), "Mystery")

        secondary.events.post.assert_not_called()

    async def test_with_no_default_calendar_listed_nothing_counts_as_own(self, tmp_path):
        client, secondary = _create_client([_calendar("KIDS=", "Kids", _ME)])

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await _create(client, _cfg(tmp_path, "calendar_write"), "Kids")

        secondary.events.post.assert_not_called()

    async def test_a_default_calendar_without_an_owner_confirms_nothing(self, tmp_path):
        client, secondary = _create_client(
            [_calendar("DEF=", "Calendar", None, is_default=True), _calendar("K=", "Kids", _ME)]
        )

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await _create(client, _cfg(tmp_path, "calendar_write"), "Kids")

        secondary.events.post.assert_not_called()


class TestCreatingInAChosenCalendarOnTheWire:
    _LISTED = [
        {
            "id": "AAMkDEF=",
            "name": "Calendar",
            "canEdit": True,
            "isDefaultCalendar": True,
            "owner": {"name": "Me", "address": _ME},
        },
        {
            "id": "AAMkJANE=",
            "name": "Jane",
            "canEdit": True,
            "isDefaultCalendar": False,
            "owner": {"name": "Jane", "address": _JANE},
        },
        {
            "id": "AAMkKIDS=",
            "name": "Kids",
            "canEdit": True,
            "isDefaultCalendar": False,
            "owner": {"name": "Me", "address": _ME.upper()},
        },
    ]

    def _client(self):
        return _wire_client(
            {
                ("GET", "/v1.0/me/calendars"): (200, {"value": self._LISTED}),
                ("POST", "/v1.0/me/calendars/AAMkKIDS=/events"): (
                    201,
                    {"id": "AAMkNEW=", "subject": "Dentist"},
                ),
            }
        )

    async def test_someone_elses_calendar_is_refused_after_the_one_listing(self, tmp_path):
        client, sent = self._client()

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await _create(client, _cfg(tmp_path, "calendar_write"), "Jane")

        assert [(r.method, r.url.path) for r in sent] == [("GET", "/v1.0/me/calendars")]
        # Owner and isDefaultCalendar come back only while nothing narrows the
        # listing to other fields.
        assert "$select" not in sent[0].url.params

    async def test_an_own_calendar_is_the_listing_and_the_post(self, tmp_path):
        client, sent = self._client()

        await _create(client, _cfg(tmp_path, "calendar_write"), "Kids")

        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/calendars"),
            ("POST", "/v1.0/me/calendars/AAMkKIDS=/events"),
        ]


# ── calendars: writes addressed by event id ──────────────────────────────────


def _event_client(event_calendar, *, default_owner=_ME, lookup_error=None, default_error=None):
    """A client whose event `AAMkEVT=` lives on ``event_calendar``."""
    client = MagicMock()
    builder = client.me.events.by_event_id.return_value
    builder.get = AsyncMock(return_value=MagicMock(attendees=[]))
    builder.patch = AsyncMock(return_value=MagicMock(id="AAMkEVT="))
    builder.delete = AsyncMock()
    builder.accept.post = AsyncMock()
    builder.attachments.post = AsyncMock(return_value=MagicMock(id="ATT1="))
    builder.attachments.by_attachment_id.return_value.delete = AsyncMock()
    builder.calendar.get = AsyncMock(return_value=event_calendar, side_effect=lookup_error)
    client.me.calendar.get = AsyncMock(
        return_value=Calendar(owner=EmailAddress(address=default_owner)),
        side_effect=default_error,
    )
    return client, builder


def _on_default():
    return Calendar(id="DEF=", is_default_calendar=True, owner=EmailAddress(address=_ME))


def _on(owner):
    cal = Calendar(id="OTHER=", is_default_calendar=False)
    if owner is not None:
        cal.owner = EmailAddress(address=owner)
    return cal


async def _update(client, config):
    return await update_event(client, event_id="AAMkEVT=", subject="Renamed", config=config)


async def _attach(client, config):
    return await attach_to_event(client, "AAMkEVT=", ["ticket.txt"], config=config)


_EVENT_WRITES = [pytest.param(_update, id="update_event"), pytest.param(_attach, id="attach")]


def _sent(builder):
    return builder.patch.await_count + builder.attachments.post.await_count


class TestWritesToAnEventOnSomeoneElsesCalendar:
    @pytest.fixture(autouse=True)
    def _file(self, tmp_path):
        (tmp_path / "ticket.txt").write_bytes(b"seat 12A")

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_an_event_on_the_default_calendar_is_calendar_only(self, tmp_path, write):
        client, builder = _event_client(_on_default())

        await write(client, _cfg(tmp_path, "calendar_write"))

        assert _sent(builder) == 1
        builder.calendar.get.assert_awaited_once()
        config = builder.calendar.get.call_args.kwargs["request_configuration"]
        assert config.query_parameters.select == ["id", "isDefaultCalendar", "owner"]
        # The default calendar is the user's own: no second lookup.
        client.me.calendar.get.assert_not_called()

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_an_event_on_an_own_calendar_is_calendar_only(self, tmp_path, write):
        client, builder = _event_client(_on(_ME.lower()))

        await write(client, _cfg(tmp_path, "calendar_write"))

        assert _sent(builder) == 1
        config = client.me.calendar.get.call_args.kwargs["request_configuration"]
        assert config.query_parameters.select == ["owner"]

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_an_event_on_someone_elses_calendar_needs_mail_send(self, tmp_path, write):
        client, builder = _event_client(_on(_JANE))

        with pytest.raises(PermissionDeniedError) as exc:
            await write(client, _cfg(tmp_path, "calendar_write"))

        assert "mail_send" in str(exc.value)
        assert "owner will see" in str(exc.value)
        assert "sends email" not in str(exc.value)
        assert LEAVE_SETTINGS_TO_THE_USER in str(exc.value)
        assert _sent(builder) == 0

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_it_passes_once_mail_send_is_allowed_with_no_lookup(self, tmp_path, write):
        client, builder = _event_client(_on(_JANE))

        await write(client, _cfg(tmp_path, "calendar_write", "mail_send"))

        assert _sent(builder) == 1
        builder.calendar.get.assert_not_called()
        client.me.calendar.get.assert_not_called()

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_an_unrestricted_server_makes_no_lookup(self, tmp_path, write):
        client, builder = _event_client(_on(_JANE))

        await write(client, _cfg(tmp_path))

        assert _sent(builder) == 1
        builder.calendar.get.assert_not_called()
        client.me.calendar.get.assert_not_called()

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    @pytest.mark.parametrize(
        "case",
        [
            pytest.param({"event_calendar": _on(None)}, id="no-owner"),
            pytest.param(
                {"event_calendar": _on_default(), "lookup_error": _api_error(403)},
                id="calendar-lookup-403",
            ),
            pytest.param(
                {"event_calendar": _on(_ME), "default_error": _api_error(403)},
                id="default-lookup-403",
            ),
            pytest.param({"event_calendar": _on(_ME), "default_owner": None}, id="default-unowned"),
        ],
    )
    async def test_an_owner_that_cannot_be_confirmed_is_refused(self, tmp_path, write, case):
        client, builder = _event_client(**case)

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await write(client, _cfg(tmp_path, "calendar_write"))

        assert _sent(builder) == 0

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    @pytest.mark.parametrize("failure", _OTHER_FAILURES)
    @pytest.mark.parametrize("which", ["event-calendar", "default-calendar"])
    async def test_any_other_failed_lookup_is_reported_as_it_is(
        self, tmp_path, write, failure, which
    ):
        """A stale id or a lapsed sign-in is not a reason to ask for `mail_send`."""
        error = failure()
        if which == "event-calendar":
            client, builder = _event_client(_on(_ME), lookup_error=error)
        else:
            client, builder = _event_client(_on(_ME), default_error=error)

        with pytest.raises(type(error)) as exc:
            await write(client, _cfg(tmp_path, "calendar_write"))

        assert exc.value is error
        assert _sent(builder) == 0

    @pytest.mark.parametrize("write", _EVENT_WRITES)
    async def test_read_only_and_the_category_are_checked_first(self, tmp_path, write):
        client, builder = _event_client(_on(_JANE))

        with pytest.raises(ReadOnlyError):
            await write(client, _cfg(tmp_path, "calendar_write", read_only=True))
        with pytest.raises(PermissionDeniedError, match="calendar_write"):
            await write(client, _cfg(tmp_path, "mail_drafts"))

        builder.calendar.get.assert_not_called()
        assert _sent(builder) == 0

    async def test_writes_that_add_nothing_are_left_alone(self, tmp_path):
        """Deleting an event, a bare RSVP and removing a file put no content in
        front of the calendar's owner."""
        client, builder = _event_client(_on(_JANE))
        config = _cfg(tmp_path, "calendar_write")

        await delete_event(client, "AAMkEVT=", config=config)
        await rsvp(client, "AAMkEVT=", "accept", config=config)
        await remove_event_attachment(client, "AAMkEVT=", "ATT1=", config=config)

        builder.delete.assert_awaited_once()
        builder.accept.post.assert_awaited_once()
        builder.attachments.by_attachment_id.return_value.delete.assert_awaited_once()
        builder.calendar.get.assert_not_called()


class TestEventCalendarLookupOnTheWire:
    @staticmethod
    def _answers(event_calendar):
        return {
            ("GET", "/v1.0/me/events/AAMkEVT=/calendar"): (200, event_calendar),
            ("GET", "/v1.0/me/calendar"): (200, {"owner": {"name": "Me", "address": _ME}}),
            ("PATCH", "/v1.0/me/events/AAMkEVT="): (200, {"id": "AAMkEVT="}),
        }

    async def test_someone_elses_calendar_is_two_reads_and_no_patch(self, tmp_path):
        client, sent = _wire_client(
            self._answers(
                {
                    "id": "AAMkJANE=",
                    "isDefaultCalendar": False,
                    "owner": {"name": "Jane", "address": _JANE},
                }
            )
        )

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await update_event(
                client, event_id="AAMkEVT=", show_as="busy", config=_cfg(tmp_path, "calendar_write")
            )

        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/events/AAMkEVT=/calendar"),
            ("GET", "/v1.0/me/calendar"),
        ]
        assert sent[0].url.params["$select"] == "id,isDefaultCalendar,owner"
        assert sent[1].url.params["$select"] == "owner"

    async def test_the_default_calendar_is_one_read_then_the_patch(self, tmp_path):
        client, sent = _wire_client(
            self._answers(
                {
                    "id": "AAMkDEF=",
                    "isDefaultCalendar": True,
                    "owner": {"name": "Me", "address": _ME},
                }
            )
        )

        await update_event(
            client, event_id="AAMkEVT=", show_as="busy", config=_cfg(tmp_path, "calendar_write")
        )

        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/events/AAMkEVT=/calendar"),
            ("PATCH", "/v1.0/me/events/AAMkEVT="),
        ]
        assert json.loads(sent[-1].content)["showAs"] == "busy"

    async def test_a_403_on_the_lookup_is_refused_as_unconfirmed(self, tmp_path):
        client, sent = _wire_client(
            {("GET", "/v1.0/me/events/AAMkEVT=/calendar"): (403, _ACCESS_DENIED)}
        )

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await update_event(
                client, event_id="AAMkEVT=", show_as="busy", config=_cfg(tmp_path, "calendar_write")
            )

        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/events/AAMkEVT=/calendar")
        ]

    async def test_a_404_on_the_lookup_is_graphs_own_error(self, tmp_path):
        client, sent = _wire_client(
            {("GET", "/v1.0/me/events/AAMkEVT=/calendar"): (404, _NOT_FOUND)}
        )

        with pytest.raises(APIError) as exc:
            await update_event(
                client, event_id="AAMkEVT=", show_as="busy", config=_cfg(tmp_path, "calendar_write")
            )

        assert exc.value.response_status_code == 404
        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/events/AAMkEVT=/calendar")
        ]


# ── To Do lists ──────────────────────────────────────────────────────────────


def _default_list():
    return TodoTaskList(
        id="DEFAULT=",
        display_name="Tasks",
        is_owner=True,
        is_shared=False,
        wellknown_list_name=WellknownListName.DefaultList,
    )


def _todo_client(task_list=None, *, lookup_error=None, lists=None):
    """A client whose own default list is `DEFAULT=`; any list's own GET
    answers ``task_list``."""
    client = MagicMock()
    client.me.todo.lists.get = AsyncMock(
        return_value=MagicMock(
            value=list(lists) if lists is not None else [_default_list()], odata_next_link=None
        )
    )
    list_item = client.me.todo.lists.by_todo_task_list_id.return_value
    list_item.get = AsyncMock(return_value=task_list, side_effect=lookup_error)
    list_item.tasks.post = AsyncMock(return_value=MagicMock(id="T1=", title="Buy milk"))
    task_item = list_item.tasks.by_todo_task_id.return_value
    task_item.patch = AsyncMock()
    task_item.delete = AsyncMock()
    task_item.checklist_items.post = AsyncMock(return_value=MagicMock(id="C1=", display_name="x"))
    task_item.checklist_items.by_checklist_item_id.return_value.patch = AsyncMock()
    task_item.checklist_items.by_checklist_item_id.return_value.delete = AsyncMock()
    uploaded = MagicMock(id="A1=")
    uploaded.name = "note.txt"
    task_item.attachments.post = AsyncMock(return_value=uploaded)
    task_item.attachments.by_attachment_base_id.return_value.delete = AsyncMock()
    return client, list_item, task_item


def _todo_writes_sent(list_item, task_item):
    return (
        list_item.tasks.post.await_count
        + task_item.patch.await_count
        + task_item.checklist_items.post.await_count
        + task_item.checklist_items.by_checklist_item_id.return_value.patch.await_count
        + task_item.attachments.post.await_count
    )


def _shared(is_owner=False):
    return TodoTaskList(id="SHARED=", display_name="Family", is_owner=is_owner, is_shared=True)


_TODO_WRITES = [
    pytest.param(
        lambda c, cfg, lid: create_task(c, "Buy milk", list_id=lid, config=cfg), id="create_task"
    ),
    pytest.param(
        lambda c, cfg, lid: update_task(c, "T1=", list_id=lid, title="Buy oat milk", config=cfg),
        id="update_task",
    ),
    pytest.param(
        lambda c, cfg, lid: add_checklist_item(c, "T1=", "Step", list_id=lid, config=cfg),
        id="add_checklist_item",
    ),
    pytest.param(
        lambda c, cfg, lid: update_checklist_item(
            c, "T1=", "C1=", display_name="Step", list_id=lid, config=cfg
        ),
        id="rename_checklist_item",
    ),
    pytest.param(
        lambda c, cfg, lid: upload_task_attachment(c, "T1=", "note.txt", list_id=lid, config=cfg),
        id="upload_task_attachment",
    ),
]


class TestWritesToSomeoneElsesList:
    @pytest.fixture(autouse=True)
    def _file(self, tmp_path):
        (tmp_path / "note.txt").write_bytes(b"milk, eggs")

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_the_default_list_needs_no_lookup(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared())

        await write(client, _cfg(tmp_path, "todo_write"), None)

        assert _todo_writes_sent(list_item, task_item) == 1
        client.me.todo.lists.by_todo_task_list_id.assert_called_with("DEFAULT=")
        list_item.get.assert_not_called()

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_a_list_someone_else_owns_needs_mail_send(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared())

        with pytest.raises(PermissionDeniedError) as exc:
            await write(client, _cfg(tmp_path, "todo_write"), "SHARED=")

        assert "mail_send" in str(exc.value)
        assert "owner will see" in str(exc.value)
        assert "sends email" not in str(exc.value)
        assert LEAVE_SETTINGS_TO_THE_USER in str(exc.value)
        assert _todo_writes_sent(list_item, task_item) == 0
        list_item.get.assert_awaited_once_with()

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_it_passes_once_mail_send_is_allowed_with_no_lookup(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared())

        await write(client, _cfg(tmp_path, "todo_write", "mail_send"), "SHARED=")

        assert _todo_writes_sent(list_item, task_item) == 1
        list_item.get.assert_not_called()

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_an_unrestricted_server_makes_no_lookup(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared())

        await write(client, _cfg(tmp_path), "SHARED=")

        assert _todo_writes_sent(list_item, task_item) == 1
        list_item.get.assert_not_called()

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_an_own_list_is_todo_only(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared(is_owner=True))

        await write(client, _cfg(tmp_path, "todo_write"), "SHARED=")

        assert _todo_writes_sent(list_item, task_item) == 1

    @pytest.mark.parametrize("write", _TODO_WRITES)
    @pytest.mark.parametrize(
        "case",
        [
            pytest.param({"task_list": _shared(is_owner=None)}, id="no-isOwner"),
            pytest.param({"task_list": None}, id="empty-answer"),
            pytest.param(
                {"task_list": _shared(is_owner=True), "lookup_error": _api_error(403)},
                id="lookup-403",
            ),
        ],
    )
    async def test_an_owner_that_cannot_be_confirmed_is_refused(self, tmp_path, write, case):
        client, list_item, task_item = _todo_client(**case)

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await write(client, _cfg(tmp_path, "todo_write"), "SHARED=")

        assert _todo_writes_sent(list_item, task_item) == 0

    @pytest.mark.parametrize("write", _TODO_WRITES)
    @pytest.mark.parametrize("failure", _OTHER_FAILURES)
    async def test_any_other_failed_lookup_is_reported_as_it_is(self, tmp_path, write, failure):
        """A mistyped list id or a lapsed sign-in is not a reason to ask for `mail_send`."""
        error = failure()
        client, list_item, task_item = _todo_client(_shared(is_owner=True), lookup_error=error)

        with pytest.raises(type(error)) as exc:
            await write(client, _cfg(tmp_path, "todo_write"), "SHARED=")

        assert exc.value is error
        assert _todo_writes_sent(list_item, task_item) == 0

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_a_fallback_first_list_is_not_taken_as_the_default(self, tmp_path, write):
        """With no `defaultList` found, `_resolve_list_id` falls back to the
        first list, which may be a shared one: it is looked up like any other."""
        client, list_item, task_item = _todo_client(_shared(), lists=[_shared()])

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await write(client, _cfg(tmp_path, "todo_write"), None)

        assert _todo_writes_sent(list_item, task_item) == 0

    @pytest.mark.parametrize("write", _TODO_WRITES)
    async def test_read_only_and_the_category_are_checked_first(self, tmp_path, write):
        client, list_item, task_item = _todo_client(_shared())

        with pytest.raises(ReadOnlyError):
            await write(client, _cfg(tmp_path, "todo_write", read_only=True), "SHARED=")
        with pytest.raises(PermissionDeniedError, match="todo_write"):
            await write(client, _cfg(tmp_path, "calendar_write"), "SHARED=")

        list_item.get.assert_not_called()

    async def test_writes_that_add_nothing_are_left_alone(self, tmp_path):
        """Completing, deleting and ticking a step put no new content in the list."""
        client, list_item, task_item = _todo_client(_shared())
        config = _cfg(tmp_path, "todo_write")
        steps = task_item.checklist_items.by_checklist_item_id.return_value

        await complete_task(client, "T1=", list_id="SHARED=", config=config)
        await delete_task(client, "T1=", list_id="SHARED=", config=config)
        await update_checklist_item(
            client, "T1=", "C1=", is_checked=True, list_id="SHARED=", config=config
        )
        await delete_checklist_item(client, "T1=", "C1=", list_id="SHARED=", config=config)
        await delete_task_attachment(client, "T1=", "A1=", list_id="SHARED=", config=config)

        assert task_item.patch.await_count == 1
        task_item.delete.assert_awaited_once()
        steps.patch.assert_awaited_once()
        steps.delete.assert_awaited_once()
        task_item.attachments.by_attachment_base_id.return_value.delete.assert_awaited_once()
        list_item.get.assert_not_called()


class TestListLookupOnTheWire:
    async def test_the_lookup_is_a_plain_get_of_the_list(self, tmp_path):
        """No `$select`: To Do lists answer one with a 400 on consumer accounts."""
        client, sent = _wire_client(
            {
                ("GET", "/v1.0/me/todo/lists/SHARED="): (
                    200,
                    {
                        "id": "SHARED=",
                        "displayName": "Family",
                        "isOwner": False,
                        "isShared": True,
                        "wellknownListName": "none",
                    },
                ),
            }
        )

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await create_task(
                client, "Buy milk", list_id="SHARED=", config=_cfg(tmp_path, "todo_write")
            )

        assert [(r.method, r.url.path) for r in sent] == [("GET", "/v1.0/me/todo/lists/SHARED=")]
        assert sent[0].url.query == b""

    async def test_an_own_list_is_the_lookup_and_the_post(self, tmp_path):
        client, sent = _wire_client(
            {
                ("GET", "/v1.0/me/todo/lists/OWN="): (
                    200,
                    {"id": "OWN=", "displayName": "Mine", "isOwner": True, "isShared": False},
                ),
                ("POST", "/v1.0/me/todo/lists/OWN=/tasks"): (201, {"id": "T1=", "title": "x"}),
            }
        )

        await create_task(client, "Buy milk", list_id="OWN=", config=_cfg(tmp_path, "todo_write"))

        assert [(r.method, r.url.path) for r in sent] == [
            ("GET", "/v1.0/me/todo/lists/OWN="),
            ("POST", "/v1.0/me/todo/lists/OWN=/tasks"),
        ]

    async def test_a_403_on_the_lookup_is_refused_as_unconfirmed(self, tmp_path):
        client, sent = _wire_client({("GET", "/v1.0/me/todo/lists/SHARED="): (403, _ACCESS_DENIED)})

        with pytest.raises(PermissionDeniedError, match="mail_send"):
            await create_task(
                client, "Buy milk", list_id="SHARED=", config=_cfg(tmp_path, "todo_write")
            )

        assert [(r.method, r.url.path) for r in sent] == [("GET", "/v1.0/me/todo/lists/SHARED=")]

    async def test_a_404_on_the_lookup_is_graphs_own_error(self, tmp_path):
        client, sent = _wire_client({("GET", "/v1.0/me/todo/lists/GONE="): (404, _NOT_FOUND)})

        with pytest.raises(APIError) as exc:
            await create_task(
                client, "Buy milk", list_id="GONE=", config=_cfg(tmp_path, "todo_write")
            )

        assert exc.value.response_status_code == 404
        assert [(r.method, r.url.path) for r in sent] == [("GET", "/v1.0/me/todo/lists/GONE=")]


class TestTaskListsSayWhoOwnsThem:
    async def test_the_listing_reports_is_owner_and_is_shared(self):
        client, _ = _wire_client(
            {
                ("GET", "/v1.0/me/todo/lists"): (
                    200,
                    {
                        "value": [
                            {
                                "id": "DEFAULT=",
                                "displayName": "Tasks",
                                "isOwner": True,
                                "isShared": False,
                                "wellknownListName": "defaultList",
                            },
                            {
                                "id": "SHARED=",
                                "displayName": "Family",
                                "isOwner": False,
                                "isShared": True,
                                "wellknownListName": "none",
                            },
                        ]
                    },
                ),
            }
        )

        result = await list_task_lists(client)

        assert [
            (lst["id"], lst["is_default"], lst["is_owner"], lst["is_shared"])
            for lst in result["task_lists"]
        ] == [("DEFAULT=", True, True, False), ("SHARED=", False, False, True)]

    async def test_a_list_reported_without_them_says_so(self):
        """Missing is reported as null, not as a guess either way."""
        client, _, _ = _todo_client(lists=[TodoTaskList(id="L=", display_name="Mine")])

        (listed,) = (await list_task_lists(client))["task_lists"]

        assert listed["is_owner"] is None
        assert listed["is_shared"] is None
