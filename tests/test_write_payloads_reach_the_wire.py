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

Every string argument gets a distinctive sentinel so a match is unambiguous.
Booleans and enums are asserted by the wire key/value they must produce.

What this cannot see: Graph accepting a field and ignoring it (``is_online``
on personal accounts). That needs the live tier.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest
from kiota_serialization_json.json_serialization_writer import JsonSerializationWriter

from outlook_mcp.config import Config
from outlook_mcp.tools import calendar_write, contacts, mail_drafts, mail_write, todo

_CFG = Config(client_id="test")


def wire(model) -> str:
    """The JSON kiota would put on the wire for this SDK model."""
    writer = JsonSerializationWriter()
    model.serialize(writer)
    return writer.get_serialized_content().decode()


def assert_on_wire(model, *needles: str) -> None:
    body = wire(model)
    missing = [n for n in needles if n not in body]
    assert not missing, (
        f"Argument value(s) never reached the serialized payload: {missing}\n"
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

    async def test_update_event_every_argument_reaches_the_wire(self):
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
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

        assert_on_wire(builder.patch.call_args[0][0], '"recurrence": null')


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
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SentinelFirst2",
            "SentinelLast2",
            "sentinel.contact2@example.com",
            "+15555550101",
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
    """Pin the reason this file exists: the SDK omits None-valued fields."""
    from msgraph.generated.models.event import Event

    event = Event()
    event.subject = "kept"
    event.recurrence = None

    body = wire(event)
    assert '"subject": "kept"' in body
    assert "recurrence" not in body

    with pytest.raises(AssertionError, match="never reached"):
        assert_on_wire(event, '"recurrence"')
