"""Tier 2 of the silent-no-op audit: does each write argument reach the wire?

Two of the three no-op shapes found in 1.15–1.18 were invisible to
attribute-level assertions:

- ``outlook_create_event`` accepted ``recurrence`` and never set it (#41).
- ``update_event(remove_recurrence=True)`` set ``event.recurrence = None`` —
  which the SDK **omits from the payload entirely**. ``patched.recurrence is
  None`` would have passed while the PATCH went out empty.

So these tests do not look at the model. They serialize the object handed to
``.post()``/``.patch()`` exactly as kiota would send it, and assert each
argument's value is present in that JSON. A parameter that is dropped by the
handler *or* by the SDK fails here.

``exactly as kiota would send it`` was not true until #63. This file used a
bare ``JsonSerializationWriter``; the real client enables the backing store
unconditionally, so every write actually goes out through
``BackingStoreSerializationWriterProxyFactory``. The two disagree on one
thing, and it is the thing that bites: a field explicitly assigned ``None``
is silently omitted by the bare writer and *emitted* by the adapter — under
its **Python** name, and for a nested model onto the **parent** object. A
partial contact address went out carrying ``"country_or_region": null`` at
the top level and Graph answered 400. Nothing here could see it.

So ``wire()`` now serializes the way the adapter does, and ``assert_on_wire``
asserts in both directions: every argument reached the payload, and nothing
reached it that we never assigned.

Every string argument gets a distinctive sentinel so a match is unambiguous.
Booleans and enums are asserted by the wire key/value they must produce.

What this cannot see: Graph accepting a field and ignoring it (``is_online``
on personal accounts). That needs the live tier.
"""

from __future__ import annotations

import json
from unittest.mock import AsyncMock, MagicMock

import pytest
from kiota_abstractions.store import BackingStoreSerializationWriterProxyFactory
from kiota_serialization_json.json_serialization_writer_factory import (
    JsonSerializationWriterFactory,
)

from outlook_mcp.config import Config
from outlook_mcp.tools import calendar_write, contacts, mail_drafts, mail_write, todo

_CFG = Config(client_id="test")


def wire(model) -> str:
    """The JSON the Graph request adapter would actually send for this model.

    Deliberately not a bare ``JsonSerializationWriter``. ``BaseGraphServiceClient``
    routes the adapter's writer factory through
    ``enable_backing_store_for_serialization_writer_factory``, so the real client
    always serializes through the backing-store proxy — and only that path emits
    a field that was explicitly assigned ``None``. The bare writer drops those,
    which is the whole difference between this guard seeing a leak and not.

    **Destructive: build a fresh model for every call.** Serializing marks the
    entire object graph clean — the proxy's ``on_after`` sets
    ``is_initialization_completed``, whose setter rewrites every entry as
    unchanged and recurses into nested models. A second call on the same object
    reports no nulls no matter what was assigned to it.
    """
    factory = BackingStoreSerializationWriterProxyFactory(JsonSerializationWriterFactory())
    writer = factory.get_serialization_writer("application/json")
    writer.write_object_value(None, model)
    try:
        return writer.get_serialized_content().decode()
    except ValueError as exc:
        if "Invalid Json output" not in str(exc):
            raise
        # kiota writes a top-level null beside the object body rather than
        # inside it, and then refuses to serialize the mixed document at all.
        # Reported as an assertion because it is a bug in the caller, not here.
        raise AssertionError(
            "A top-level field was explicitly assigned None, so this payload "
            "cannot be serialized and the request would fail before leaving the "
            "process. Leave the field unset instead — or, if Graph genuinely "
            "needs an explicit null, send it through additional_data (see "
            "calendar_write.update_event's remove_recurrence)."
        ) from exc


def _walk_keys(body: str):
    """Every (key, value) pair in the payload, nested objects and arrays included."""

    def walk(node):
        if isinstance(node, dict):
            for key, value in node.items():
                yield key, value
                yield from walk(value)
        elif isinstance(node, list):
            for item in node:
                yield from walk(item)

    return list(walk(json.loads(body)))


def assert_on_wire(model, *needles: str, allow_null: bool = False) -> None:
    """Assert every needle reached the payload, and that nothing else did.

    ``allow_null`` opts out of the null check for the one shape that wants an
    explicit null on the wire — ``remove_recurrence``, which sends it through
    ``additional_data`` on purpose.
    """
    body = wire(model)

    missing = [n for n in needles if n not in body]
    assert not missing, (
        f"Argument value(s) never reached the serialized payload: {missing}\n"
        f"Wire JSON was:\n{body}"
    )

    pairs = _walk_keys(body)

    # Graph properties are camelCase. A snake_case key can only be a Python
    # attribute name the backing store emitted for a field assigned None —
    # always a bug, and one Graph answers with a 400 or, when the name happens
    # to collide with a real property, by silently clearing it.
    leaked = sorted({key for key, _ in pairs if "_" in key and not key.startswith("@")})
    assert not leaked, (
        f"Python attribute name(s) reached the wire: {leaked}. A field was "
        f"assigned None somewhere; leave it unset instead.\n"
        f"Wire JSON was:\n{body}"
    )

    if not allow_null:
        nulls = sorted({key for key, value in pairs if value is None})
        assert not nulls, (
            f"Field(s) we never assigned reached the wire as null: {nulls}. "
            f"Graph treats an explicit null as 'clear this', so this would "
            f"destroy data the caller did not ask to change. Pass "
            f"allow_null=True if the null is deliberate.\n"
            f"Wire JSON was:\n{body}"
        )


# ── calendar ──────────────────────────────────────────────────────────


class TestCalendarWrite:
    async def test_create_event_every_argument_reaches_the_wire(self):
        client = AsyncMock()
        client.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="s"))

        await calendar_write.create_event(
            client,
            subject="SENTINEL-SUBJECT-4a1",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:30:00Z",
            location="SENTINEL-LOCATION-4a2",
            body="SENTINEL-BODY-4a3",
            attendees=["sentinel.attendee@example.com"],
            is_all_day=True,
            recurrence={
                "pattern": {"type": "weekly", "interval": 2, "daysOfWeek": ["monday"]},
                "range": {"type": "numbered", "numberOfOccurrences": 3},
            },
            config=_CFG,
        )

        assert_on_wire(
            client.me.events.post.call_args[0][0],
            "SENTINEL-SUBJECT-4a1",
            "2026-09-07T12:30:00",
            "2026-09-07T13:30:00",
            "SENTINEL-LOCATION-4a2",
            "SENTINEL-BODY-4a3",
            "sentinel.attendee@example.com",
            '"isAllDay": true',
            '"type": "weekly"',
            '"interval": 2',
            '"numberOfOccurrences": 3',
        )

    async def test_create_event_puts_the_anchor_zone_on_the_wire(self):
        """The zone is a value Graph reads, so assert it in the payload.

        ``event.start.time_zone`` was the literal ``"UTC"`` until the timezone
        fix, and a model-level assertion on it would have passed just as
        happily then — the string was always *there*. What changed is which
        string, and the only way to state that is to pin the one that ships.
        """
        client = MagicMock()
        client.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="s"))

        await calendar_write.create_event(
            client,
            subject="SENTINEL-SUBJECT-4c1",
            start="2026-10-28T09:00:00",
            end="2026-10-28T10:00:00",
            timezone="America/New_York",
            config=_CFG,
        )

        assert_on_wire(
            client.me.events.post.call_args[0][0],
            "SENTINEL-SUBJECT-4c1",
            '"timeZone": "America/New_York"',
            '"dateTime": "2026-10-28T09:00:00"',
        )

    async def test_update_event_every_argument_reaches_the_wire(self):
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        # A start/end patch reads the event first for the zone it must carry.
        # The mock answers with a real string, not a MagicMock, because a
        # MagicMock is truthy and would sail through the fallback while
        # serializing to something no Graph response ever contains.
        current = MagicMock(type=MagicMock(value="singleInstance"))
        current.start = MagicMock(date_time="2026-10-22T00:00:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "UTC"
        current.original_end_time_zone = "UTC"
        builder.get = AsyncMock(return_value=current)
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client,
            event_id="AAMkAG123=",
            subject="SENTINEL-SUBJECT-5b1",
            start="2026-10-22T00:00:00Z",
            end="2026-10-23T00:00:00Z",
            location="SENTINEL-LOCATION-5b2",
            body="SENTINEL-BODY-5b3",
            recurrence="weekly",
            attendees=["sentinel.guest@example.com"],
            is_all_day=True,
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SENTINEL-SUBJECT-5b1",
            "2026-10-22T00:00:00",
            "SENTINEL-LOCATION-5b2",
            "SENTINEL-BODY-5b3",
            "sentinel.guest@example.com",
            '"isAllDay": true',
            '"daysOfWeek": ["thursday"]',  # 2026-10-22 is a Thursday
        )

    async def test_update_event_remove_recurrence_is_an_explicit_null(self):
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG
        )

        assert_on_wire(
            builder.patch.call_args[0][0], '"recurrence": null', allow_null=True
        )


# ── mail ──────────────────────────────────────────────────────────────


class TestMailWrite:
    async def test_send_message_every_argument_reaches_the_wire(self):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        await mail_write.send_message(
            client,
            to=["sentinel.to@example.com"],
            subject="SENTINEL-SUBJECT-6c1",
            body="<b>SENTINEL-BODY-6c2</b>",
            cc=["sentinel.cc@example.com"],
            bcc=["sentinel.bcc@example.com"],
            is_html=True,
            importance="high",
            request_read_receipt=True,
            reply_to=["sentinel.replyto@example.com"],
            config=_CFG,
        )

        assert_on_wire(
            client.me.send_mail.post.call_args[0][0],
            "sentinel.to@example.com",
            "SENTINEL-SUBJECT-6c1",
            "SENTINEL-BODY-6c2",
            "sentinel.cc@example.com",
            "sentinel.bcc@example.com",
            '"contentType": "html"',
            '"importance": "high"',
            '"isReadReceiptRequested": true',
            "sentinel.replyto@example.com",
        )

    async def test_reply_html_body_reaches_the_wire(self):
        builder = MagicMock()
        builder.reply.post = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_write.reply(
            client,
            message_id="AAMkAG123=",
            body="<i>SENTINEL-REPLY-7d1</i>",
            is_html=True,
            config=_CFG,
        )

        assert_on_wire(
            builder.reply.post.call_args[0][0],
            "SENTINEL-REPLY-7d1",
            '"contentType": "html"',
        )

    async def test_reply_plain_body_reaches_the_wire(self):
        builder = MagicMock()
        builder.reply.post = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_write.reply(
            client, message_id="AAMkAG123=", body="SENTINEL-REPLY-7d2", config=_CFG
        )

        # kiota emits action-parameter keys in PascalCase ("Comment", "Message",
        # "SaveToSentItems"); Graph accepts them — reply has always worked this way.
        assert_on_wire(builder.reply.post.call_args[0][0], '"Comment": "SENTINEL-REPLY-7d2"')

    async def test_forward_every_argument_reaches_the_wire(self):
        builder = MagicMock()
        builder.forward.post = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_write.forward(
            client,
            message_id="AAMkAG123=",
            to=["sentinel.fwd@example.com"],
            comment="SENTINEL-COMMENT-8e1",
            config=_CFG,
        )

        assert_on_wire(
            builder.forward.post.call_args[0][0],
            "sentinel.fwd@example.com",
            "SENTINEL-COMMENT-8e1",
        )


class TestMailDrafts:
    async def test_create_draft_every_argument_reaches_the_wire(self):
        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=MagicMock(id="D1"))

        await mail_drafts.create_draft(
            client,
            to=["sentinel.to@example.com"],
            subject="SENTINEL-SUBJECT-9f1",
            body="<p>SENTINEL-BODY-9f2</p>",
            cc=["sentinel.cc@example.com"],
            bcc=["sentinel.bcc@example.com"],
            is_html=True,
            importance="low",
            reply_to=["sentinel.replyto@example.com"],
            deferred_send_datetime="2026-12-01T09:00:00Z",
            config=_CFG,
        )

        assert_on_wire(
            client.me.messages.post.call_args[0][0],
            "sentinel.to@example.com",
            "SENTINEL-SUBJECT-9f1",
            "SENTINEL-BODY-9f2",
            "sentinel.cc@example.com",
            "sentinel.bcc@example.com",
            '"contentType": "html"',
            '"importance": "low"',
            "sentinel.replyto@example.com",
            "2026-12-01T09:00:00",
        )

    async def test_update_draft_every_argument_reaches_the_wire(self):
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_drafts.update_draft(
            client,
            draft_id="AAMkAG123=",
            subject="SENTINEL-SUBJECT-a01",
            body="SENTINEL-BODY-a02",
            to=["sentinel.to2@example.com"],
            cc=["sentinel.cc2@example.com"],
            reply_to=["sentinel.rt2@example.com"],
            deferred_send_datetime="2026-12-02T09:00:00Z",
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SENTINEL-SUBJECT-a01",
            "SENTINEL-BODY-a02",
            "sentinel.to2@example.com",
            "sentinel.cc2@example.com",
            "sentinel.rt2@example.com",
            "2026-12-02T09:00:00",
        )


# ── contacts ──────────────────────────────────────────────────────────


class TestContacts:
    async def test_create_contact_every_argument_reaches_the_wire(self):
        client = MagicMock()
        client.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))

        await contacts.create_contact(
            client,
            first_name="SentinelFirst",
            last_name="SentinelLast",
            email="sentinel.contact@example.com",
            phone="+15555550100",
            company="SENTINEL-COMPANY-b11",
            title="SENTINEL-TITLE-b12",
            config=_CFG,
        )

        assert_on_wire(
            client.me.contacts.post.call_args[0][0],
            "SentinelFirst",
            "SentinelLast",
            "sentinel.contact@example.com",
            "+15555550100",
            "SENTINEL-COMPANY-b11",
            "SENTINEL-TITLE-b12",
        )

    async def test_update_contact_every_argument_reaches_the_wire(self):
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client,
            contact_id="AAMkAG123=",
            first_name="SentinelFirst2",
            last_name="SentinelLast2",
            email="sentinel.contact2@example.com",
            phone="+15555550101",
            home_address={
                "street": "SentinelStreet2",
                "city": "SentinelCity2",
                "state": "SentinelState2",
                "postal_code": "SentinelZip2",
                "country_or_region": "SentinelCountry2",
            },
            business_address={"city": "SentinelBusinessCity2"},
            other_address={"city": "SentinelOtherCity2"},
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SentinelFirst2",
            "SentinelLast2",
            "sentinel.contact2@example.com",
            "+15555550101",
            # An address is a nested model, so "reaches the wire" also means
            # kiota serialized the child object under the right camelCase keys —
            # and that each of the three slots landed in its own.
            '"street": "SentinelStreet2"',
            '"city": "SentinelCity2"',
            '"state": "SentinelState2"',
            '"postalCode": "SentinelZip2"',
            '"countryOrRegion": "SentinelCountry2"',
            '"businessAddress": {"city": "SentinelBusinessCity2"}',
            '"otherAddress": {"city": "SentinelOtherCity2"}',
        )


# ── to do ─────────────────────────────────────────────────────────────


class TestTodo:
    async def test_create_task_every_argument_reaches_the_wire(self):
        from tests.test_todo import _build_mock_client

        client = _build_mock_client()

        await todo.create_task(
            client,
            title="SENTINEL-TITLE-c21",
            due="2026-11-05T17:00:00Z",
            importance="high",
            body="SENTINEL-BODY-c22",
            reminder=True,
            recurrence={
                "pattern": {"type": "daily", "interval": 3},
                "range": {"type": "noEnd", "startDate": "2026-11-05"},
            },
            config=_CFG,
        )

        post = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.post
        assert_on_wire(
            post.call_args.args[0],
            "SENTINEL-TITLE-c21",
            "2026-11-05T17:00:00",
            '"importance": "high"',
            "SENTINEL-BODY-c22",
            '"isReminderOn": true',
            '"type": "daily"',
            '"interval": 3',
        )

    async def test_update_task_every_argument_reaches_the_wire(self):
        from tests.test_todo import _build_mock_client

        client = _build_mock_client()

        await todo.update_task(
            client,
            task_id="AAMkAG123=",
            title="SENTINEL-TITLE-d31",
            due="2026-11-06T17:00:00Z",
            body="SENTINEL-BODY-d32",
            importance="low",
            config=_CFG,
        )

        patch = (
            client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id.return_value.patch
        )
        assert_on_wire(
            patch.call_args.args[0],
            "SENTINEL-TITLE-d31",
            "2026-11-06T17:00:00",
            "SENTINEL-BODY-d32",
            '"importance": "low"',
        )


# ── the helper itself ─────────────────────────────────────────────────


def test_wire_helper_detects_a_dropped_field():
    """The original purpose: an argument the handler never set must fail here."""
    from msgraph.generated.models.event import Event

    event = Event()
    event.subject = "kept"

    assert '"subject": "kept"' in wire(event)

    with pytest.raises(AssertionError, match="never reached"):
        assert_on_wire(Event(subject="kept"), '"location"')


def test_a_nested_none_is_emitted_onto_the_parent_under_its_python_name():
    """The mechanism this file was blind to until #63.

    A field assigned ``None`` on a *nested* model does not vanish and does not
    stay nested — the backing store emits it on the **parent**, keyed by the
    Python attribute name. ``country_or_region`` is not a Graph property, so
    Graph answered 400; had the name been single-word it would have matched a
    real property and Graph would have cleared it instead.

    If kiota ever stops doing this, this test fails and the adapter-fidelity
    serializer in ``wire()`` can go back to being a bare writer.
    """
    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.physical_address import PhysicalAddress

    def a_contact_with_a_half_assigned_address() -> Contact:
        address = PhysicalAddress()
        address.city = "Bothell"
        address.country_or_region = None
        contact = Contact()
        contact.home_address = address
        return contact

    body = wire(a_contact_with_a_half_assigned_address())
    assert '"country_or_region": null' in body, body

    with pytest.raises(AssertionError, match="Python attribute name"):
        assert_on_wire(a_contact_with_a_half_assigned_address(), '"city": "Bothell"')


def test_a_top_level_none_is_reported_as_an_assertion_not_a_kiota_error():
    """``event.recurrence = None`` cannot be serialized at all.

    kiota writes the null beside the object body rather than inside it, then
    refuses the mixed document with ``ValueError("Invalid Json output")``. That
    is why ``update_event(remove_recurrence=True)`` routes the null through
    ``additional_data`` — the plain assignment does not produce a null, it
    produces an unsendable request. The bare writer showed neither, which is
    how the workaround's comment came to describe the wrong mechanism.
    """
    from msgraph.generated.models.event import Event

    event = Event()
    event.subject = "kept"
    event.recurrence = None

    with pytest.raises(AssertionError, match="top-level field"):
        wire(event)


def test_serializing_twice_would_report_a_clean_payload():
    """Pin the sharp edge documented on ``wire()``: the call mutates its model.

    ``on_after`` completes initialization on the backing store, which marks
    every entry — nested models included — unchanged. A guard that serialized
    the same object twice would pass vacuously on the second look, so every
    call has to build its own model.
    """
    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.physical_address import PhysicalAddress

    address = PhysicalAddress()
    address.city = "Bothell"
    address.country_or_region = None
    contact = Contact()
    contact.home_address = address

    assert '"country_or_region": null' in wire(contact)
    assert "country_or_region" not in wire(contact)
