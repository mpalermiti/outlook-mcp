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
from outlook_mcp.tools import (
    calendar_write,
    contacts,
    mail_attachments,
    mail_drafts,
    mail_folders,
    mail_thread,
    mail_triage,
    mail_write,
    todo,
    todo_attachments,
)

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
        f"Argument value(s) never reached the serialized payload: {missing}\nWire JSON was:\n{body}"
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
            show_as="tentative",
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
            # camelCase `showAs`, not the Python `show_as`. An enum assigned to
            # an SDK field is only as good as what kiota emits for it.
            '"showAs": "tentative"',
        )

    async def test_an_event_key_reaches_the_wire_on_create_and_update(self):
        """The key lives in a hidden extended property, under Graph's camelCase name."""
        client = AsyncMock()
        client.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="s"))
        await calendar_write.create_event(
            client,
            subject="s",
            start="2026-09-07T12:30:00Z",
            end="2026-09-07T13:30:00Z",
            key="SENTINEL-KEY-4b1",
            config=_CFG,
        )
        assert_on_wire(
            client.me.events.post.call_args[0][0],
            '"singleValueExtendedProperties"',
            '"String {51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey"',
            '"value": "SENTINEL-KEY-4b1"',
        )

        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)
        await calendar_write.update_event(
            client, event_id="AAMkAG123=", key="SENTINEL-KEY-4b3", config=_CFG
        )
        assert_on_wire(
            builder.patch.call_args[0][0],
            '"singleValueExtendedProperties"',
            '"String {51f1f359-76ee-427e-ac52-22691bf5b2f8} Name EventKey"',
            '"value": "SENTINEL-KEY-4b3"',
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
        # Anchored in a NON-UTC zone deliberately. With a UTC fixture the
        # preserved anchor and the pre-fix hardcoded `"UTC"` serialize
        # byte-identically, so reverting the update-side fix left this green.
        current = MagicMock(type=MagicMock(value="singleInstance"))
        current.start = MagicMock(date_time="2026-10-22T00:00:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "America/New_York"
        current.original_end_time_zone = "America/New_York"
        current.is_all_day = False
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
            show_as="workingElsewhere",
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SENTINEL-SUBJECT-5b1",
            "2026-10-22T00:00:00",
            "SENTINEL-LOCATION-5b2",
            "SENTINEL-BODY-5b3",
            "sentinel.guest@example.com",
            # UTC, not the New York anchor above: an all-day event is stored
            # anchored in UTC whatever zone it is sent, so that is what goes on
            # the wire. The preserved anchor has its own test below, because
            # the two are mutually exclusive by design.
            '"timeZone": "UTC"',
            '"isAllDay": true',
            '"daysOfWeek": ["thursday"]',  # 2026-10-22 is a Thursday
            '"showAs": "workingElsewhere"',
        )

    async def test_update_event_puts_the_preserved_anchor_on_the_wire(self):
        """The anchor is a value Graph reads, so assert it in the payload.

        Anchored in a non-UTC zone deliberately. The sibling test above must
        use an all-day event, which is always UTC-anchored — and with a UTC
        fixture the preserved anchor and the pre-fix hardcoded `"UTC"`
        serialize byte-identically, so reverting the update-side fix left it
        green. This is the one that can tell them apart.
        """
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        current = MagicMock(type=MagicMock(value="singleInstance"))
        current.start = MagicMock(date_time="2026-10-22T13:00:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "America/New_York"
        current.original_end_time_zone = "America/New_York"
        current.is_all_day = False
        builder.get = AsyncMock(return_value=current)
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client,
            event_id="AAMkAG123=",
            subject="SENTINEL-SUBJECT-5c1",
            start="2026-10-22T09:00:00",
            end="2026-10-22T10:00:00",
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SENTINEL-SUBJECT-5c1",
            '"timeZone": "America/New_York"',
            '"dateTime": "2026-10-22T09:00:00"',
        )

    async def test_update_event_show_as_alone_patches_nothing_else(self):
        """The shape a partial patch has to have: one field in, one field out.

        The null check is the half that matters here — a `show_as`-only PATCH
        that also carried ``"start": null`` would clear the event's start, and
        the model-level assertion (``patched.start is None``) passes either way
        because that is exactly what the leak looks like on the model.
        """
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(client, event_id="AAMkAG123=", show_as="oof", config=_CFG)

        # One serialization, both assertions: wire() is destructive, so a second
        # call on the same model reports a clean payload no matter what leaked.
        body = json.loads(wire(builder.patch.call_args[0][0]))

        assert body["showAs"] == "oof"
        # Key-set equality rather than assert_on_wire: it subsumes both of that
        # helper's checks here — a snake_case leak or an unasked-for null would
        # each show up as an extra key — and additionally pins that no other
        # field was sent at all.
        assert set(body) == {"@odata.type", "showAs"}

    async def test_update_event_puts_a_requested_zone_on_the_wire(self):
        """An explicit `timezone` is a value Graph reads, so assert it in the payload.

        The fixture is anchored in `America/New_York` and the argument asks for
        `Europe/London`, so neither the stored anchor nor a hardcoded `"UTC"`
        can satisfy this. A model-level assertion on `start.time_zone` would
        pass for a version that silently kept the stored zone.
        """
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        current = MagicMock(type=MagicMock(value="singleInstance"))
        current.start = MagicMock(date_time="2026-10-22T13:00:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "America/New_York"
        current.original_end_time_zone = "America/New_York"
        current.is_all_day = False
        builder.get = AsyncMock(return_value=current)
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client,
            event_id="AAMkAG123=",
            subject="SENTINEL-SUBJECT-5d1",
            start="2026-10-22T09:00:00",
            end="2026-10-22T10:00:00",
            timezone="Europe/London",
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SENTINEL-SUBJECT-5d1",
            '"timeZone": "Europe/London"',
            '"dateTime": "2026-10-22T09:00:00"',
        )

    async def test_a_re_anchored_series_puts_its_moved_days_on_the_wire(self):
        """The moved pattern is a value Graph reads, so assert it in the payload.

        A fortnightly Thursday 02:00Z series re-anchored to Wednesday 18:00 in
        Los Angeles has to go out as a Wednesday series, its week boundary moved
        with it. A model-level assertion on
        `days_of_week` cannot see what the backing store serializes; this does,
        and `assert_on_wire` also refuses any stray null or snake_case key the
        rebuilt recurrence might leak.
        """
        from outlook_mcp.tools._recurrence import build_event_recurrence

        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        current = MagicMock(type=MagicMock(value="seriesMaster"))
        current.start = MagicMock(date_time="2026-11-05T02:00:00.0000000", time_zone="UTC")
        current.original_start_time_zone = "UTC"
        current.original_end_time_zone = "UTC"
        current.is_all_day = False
        current.additional_data = {"@odata.etag": 'W/"key"'}
        current.cancelled_occurrences = []
        current.exception_occurrences = []
        current.recurrence = build_event_recurrence(
            {
                "pattern": {
                    "type": "weekly",
                    "interval": 2,
                    "daysOfWeek": ["thursday"],
                    "firstDayOfWeek": "sunday",
                },
                "range": {"type": "numbered", "numberOfOccurrences": 3},
            },
            start="2026-11-05T02:00:00Z",
            zone="UTC",
        )
        builder.get = AsyncMock(return_value=current)
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client,
            event_id="AAMkAG123=",
            start="2026-11-04T18:00:00",
            end="2026-11-04T18:30:00",
            timezone="America/Los_Angeles",
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            '"daysOfWeek": ["wednesday"]',
            '"firstDayOfWeek": "saturday"',
            '"startDate": "2026-11-04"',
            '"timeZone": "America/Los_Angeles"',
        )

    async def test_update_event_remove_recurrence_is_an_explicit_null(self):
        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="E1"))
        client = MagicMock()
        client.me.events.by_event_id = MagicMock(return_value=builder)

        await calendar_write.update_event(
            client, event_id="AAMkAG123=", remove_recurrence=True, config=_CFG
        )

        assert_on_wire(builder.patch.call_args[0][0], '"recurrence": null', allow_null=True)


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

    # A reply, and a forward with a comment, go through a draft: the create call,
    # the PATCH with the text and the send are separate payloads (#143 follow-up).

    @staticmethod
    def _thread_client():
        from msgraph.generated.models.body_type import BodyType

        created = MagicMock(id="R1")
        created.body = MagicMock(
            content="<html><body><div>QUOTED-ORIGINAL</div></body></html>",
            content_type=BodyType.Html,
        )
        source = MagicMock()
        for action in ("create_reply", "create_reply_all", "create_forward"):
            getattr(source, action).post = AsyncMock(return_value=created)
        draft = MagicMock()
        draft.patch = AsyncMock()
        draft.send.post = AsyncMock()
        builders = {"AAMkAG123=": source, "R1": draft}
        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(side_effect=lambda i: builders[i])
        return client, source, draft

    async def test_reply_html_body_reaches_the_wire(self):
        client, source, draft = self._thread_client()

        await mail_write.reply(
            client,
            message_id="AAMkAG123=",
            body="<i>SENTINEL-REPLY-7d1</i>",
            is_html=True,
            config=_CFG,
        )

        assert "SENTINEL-REPLY-7d1" not in wire(source.create_reply.post.call_args[0][0])
        assert_on_wire(
            draft.patch.call_args[0][0],
            "<html><body><i>SENTINEL-REPLY-7d1</i><div>QUOTED-ORIGINAL</div></body></html>",
            '"contentType": "html"',
        )

    async def test_reply_plain_body_reaches_the_wire(self):
        client, source, draft = self._thread_client()

        await mail_write.reply(
            client, message_id="AAMkAG123=", body="SENTINEL-REPLY-7d2\nline two", config=_CFG
        )

        # Not as "Comment": Graph joins its lines with spaces on an HTML message.
        assert '"Comment"' not in wire(source.create_reply.post.call_args[0][0])
        assert_on_wire(
            draft.patch.call_args[0][0],
            "<html><body><p>SENTINEL-REPLY-7d2<br>line two</p><div>QUOTED-ORIGINAL</div>",
            '"contentType": "html"',
        )

    async def test_forward_every_argument_reaches_the_wire(self):
        client, source, draft = self._thread_client()

        await mail_write.forward(
            client,
            message_id="AAMkAG123=",
            to=["sentinel.fwd@example.com"],
            comment="SENTINEL-COMMENT-8e1",
            config=_CFG,
        )

        create = wire(source.create_forward.post.call_args[0][0])
        assert (
            '"toRecipients": [{"emailAddress": {"address": "sentinel.fwd@example.com"}}]' in create
        )
        # On `message` only, not also the action's own ToRecipients: Graph 400s on both.
        assert '"ToRecipients"' not in create
        assert_on_wire(draft.patch.call_args[0][0], "<p>SENTINEL-COMMENT-8e1</p>")

    async def test_forward_without_a_comment_reaches_the_wire(self):
        builder = MagicMock()
        builder.forward.post = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_write.forward(
            client, message_id="AAMkAG123=", to=["sentinel.fwd2@example.com"], config=_CFG
        )

        assert_on_wire(builder.forward.post.call_args[0][0], "sentinel.fwd2@example.com")


class TestInlineMailAttachments:
    """Inline attachments (#129): the decoded file goes out as ``contentBytes``.

    The needle is the file's exact base64, so a match is byte-for-byte fidelity on the
    wire — the upload half of the round trip `test_attachments_as_content.py` checks for
    downloads, with a PDF and a PNG whose bytes are not valid UTF-8.
    """

    async def test_attach_to_draft_inline_file_reaches_the_wire(self, tmp_path):
        import base64 as _b64

        from tests.test_attachments_as_content import PDF

        client = MagicMock()
        draft = client.me.messages.by_message_id.return_value
        draft.get = AsyncMock(return_value=MagicMock(is_draft=True))
        draft.attachments.post = AsyncMock(return_value=MagicMock(id="A1"))

        await mail_attachments.attach_to_draft(
            client,
            draft_id="AAMkAG123=",
            attachments=[
                {
                    "name": "SENTINEL-INLINE-b71.pdf",
                    "content_base64": _b64.b64encode(PDF).decode(),
                    "content_type": "application/x-sentinel-b72",
                }
            ],
            config=Config(client_id="test", attachments_dir=str(tmp_path / "att")),
        )

        assert_on_wire(
            draft.attachments.post.call_args.args[0],
            '"@odata.type": "#microsoft.graph.fileAttachment"',
            '"name": "SENTINEL-INLINE-b71.pdf"',
            '"contentType": "application/x-sentinel-b72"',
            f'"contentBytes": "{_b64.b64encode(PDF).decode()}"',
        )

    async def test_send_with_attachments_inline_file_reaches_the_wire(self, tmp_path):
        import base64 as _b64

        from tests.test_attachments_as_content import PNG

        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        await mail_attachments.send_with_attachments(
            client,
            to=["sentinel.to@example.com"],
            subject="SENTINEL-SUBJECT-c71",
            body="SENTINEL-BODY-c72",
            attachments=[
                {"name": "SENTINEL-INLINE-c73.png", "content_base64": _b64.b64encode(PNG).decode()}
            ],
            config=Config(client_id="test", attachments_dir=str(tmp_path / "att")),
        )

        assert_on_wire(
            client.me.send_mail.post.call_args.args[0],
            "sentinel.to@example.com",
            "SENTINEL-SUBJECT-c71",
            "SENTINEL-BODY-c72",
            '"@odata.type": "#microsoft.graph.fileAttachment"',
            '"name": "SENTINEL-INLINE-c73.png"',
            '"contentType": "image/png"',
            f'"contentBytes": "{_b64.b64encode(PNG).decode()}"',
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
        builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
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

    # Drafts in a thread (#143). The create call and the PATCH are separate
    # payloads, so each is serialized on its own.

    @staticmethod
    def _thread_client(content="<html><body><div>QUOTED-ORIGINAL</div></body></html>", html=True):
        from msgraph.generated.models.body_type import BodyType

        draft = MagicMock(id="AAMkThreadDraft=", to_recipients=[], cc_recipients=[])
        draft.body = MagicMock(
            content=content, content_type=BodyType.Html if html else BodyType.Text
        )
        source = MagicMock()
        for action in ("create_reply", "create_reply_all", "create_forward"):
            getattr(source, action).post = AsyncMock(return_value=draft)
        new_draft = MagicMock()
        new_draft.patch = AsyncMock()
        builders = {"AAMkOriginal=": source, "AAMkThreadDraft=": new_draft}
        client = MagicMock()
        client.me.messages.by_message_id = MagicMock(side_effect=lambda i: builders[i])
        return client, source, new_draft

    async def test_reply_draft_every_argument_reaches_the_wire(self):
        client, source, new_draft = self._thread_client()

        await mail_drafts.create_draft(
            client,
            to=["sentinel.to3@example.com"],
            subject="SENTINEL-SUBJECT-a11",
            body="SENTINEL-BODY-a12",
            cc=["sentinel.cc3@example.com"],
            bcc=["sentinel.bcc3@example.com"],
            importance="high",
            reply_to=["sentinel.rt3@example.com"],
            deferred_send_datetime="2026-12-03T09:00:00Z",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        create = source.create_reply.post.call_args[0][0]
        assert_on_wire(
            create,
            "SENTINEL-SUBJECT-a11",
            "sentinel.to3@example.com",
            "sentinel.cc3@example.com",
            "sentinel.bcc3@example.com",
            '"importance": "high"',
            "sentinel.rt3@example.com",
        )
        # The body is not on the create call: Graph's `comment` would flatten it.
        assert "SENTINEL-BODY-a12" not in wire(source.create_reply.post.call_args[0][0])
        assert_on_wire(
            new_draft.patch.call_args[0][0],
            "<html><body><p>SENTINEL-BODY-a12</p><div>QUOTED-ORIGINAL</div></body></html>",
            '"contentType": "html"',
            "2026-12-03T09:00:00",
            "0x3FEF",
        )

    async def test_reply_all_draft_reaches_the_wire(self):
        client, source, new_draft = self._thread_client()

        await mail_drafts.create_draft(
            client,
            to=None,
            subject=None,
            body="SENTINEL-BODY-a21",
            reply_to_message_id="AAMkOriginal=",
            reply_all=True,
            config=_CFG,
        )

        assert '"Comment"' not in wire(source.create_reply_all.post.call_args[0][0])
        assert_on_wire(
            new_draft.patch.call_args[0][0], "<p>SENTINEL-BODY-a21</p>", '"contentType": "html"'
        )

    async def test_forward_draft_reaches_the_wire(self):
        client, source, new_draft = self._thread_client()

        await mail_drafts.create_draft(
            client,
            to=["sentinel.fwd2@example.com"],
            subject=None,
            body="SENTINEL-BODY-a31",
            forward_message_id="AAMkOriginal=",
            config=_CFG,
        )

        body = wire(source.create_forward.post.call_args[0][0])
        assert '"Comment"' not in body
        assert_on_wire(
            new_draft.patch.call_args[0][0], "<p>SENTINEL-BODY-a31</p>", '"contentType": "html"'
        )
        # On `message`, not also the action's own ToRecipients: Graph 400s on both.
        assert (
            '"toRecipients": [{"emailAddress": {"address": "sentinel.fwd2@example.com"}}]' in body
        )
        assert '"ToRecipients"' not in body

    async def test_html_reply_draft_body_reaches_the_wire_above_the_quote(self):
        client, source, new_draft = self._thread_client()

        await mail_drafts.create_draft(
            client,
            to=None,
            subject=None,
            body="<b>SENTINEL-BODY-a41</b>",
            is_html=True,
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        assert '"Comment"' not in wire(source.create_reply.post.call_args[0][0])
        assert_on_wire(
            new_draft.patch.call_args[0][0],
            '"contentType": "html"',
            "<html><body><b>SENTINEL-BODY-a41</b><div>QUOTED-ORIGINAL</div></body></html>",
        )

    async def test_plain_reply_draft_reaches_the_wire_with_its_layout(self):
        """Graph's `comment` joined these lines with spaces; the PATCH keeps them."""
        client, _, new_draft = self._thread_client()

        await mail_drafts.create_draft(
            client,
            to=None,
            subject=None,
            body="SENTINEL-a51 one\nSENTINEL-a52 two\n\nSENTINEL-a53\n\n- SENTINEL-a54",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        assert_on_wire(
            new_draft.patch.call_args[0][0],
            '"contentType": "html"',
            "<html><body><p>SENTINEL-a51 one<br>SENTINEL-a52 two</p><p>SENTINEL-a53</p>"
            "<ul><li>SENTINEL-a54</li></ul><div>QUOTED-ORIGINAL</div></body></html>",
        )

    async def test_plain_reply_draft_on_a_plain_text_original_reaches_the_wire_as_text(self):
        client, _, new_draft = self._thread_client(content="\r\nQUOTED-ORIGINAL\r\n", html=False)

        await mail_drafts.create_draft(
            client,
            to=None,
            subject=None,
            body="SENTINEL-a61\nSENTINEL-a62",
            reply_to_message_id="AAMkOriginal=",
            config=_CFG,
        )

        assert_on_wire(
            new_draft.patch.call_args[0][0],
            '"contentType": "text"',
            '"content": "SENTINEL-a61\\r\\nSENTINEL-a62\\r\\n\\r\\nQUOTED-ORIGINAL\\r\\n"',
        )


class TestMarkdownBodies:
    """`is_markdown` (#168): the body reaches the wire as HTML converted from Markdown."""

    _MD = "**SENTINEL-MD-b1**\n\n- [SENTINEL-MD-b2](https://example.com/b3)"
    _HTML = (
        "<p><strong>SENTINEL-MD-b1</strong></p>"
        '<ul><li><a href=\\"https://example.com/b3\\">SENTINEL-MD-b2</a></li></ul>'
    )

    async def test_send_message(self):
        client = MagicMock()
        client.me.send_mail.post = AsyncMock()

        await mail_write.send_message(
            client, ["a@example.com"], "s", self._MD, is_markdown=True, config=_CFG
        )

        assert_on_wire(
            client.me.send_mail.post.call_args[0][0], self._HTML, '"contentType": "html"'
        )

    async def test_create_draft(self):
        client = MagicMock()
        client.me.messages.post = AsyncMock(return_value=MagicMock(id="D1"))

        await mail_drafts.create_draft(
            client, ["a@example.com"], "s", self._MD, is_markdown=True, config=_CFG
        )

        assert_on_wire(client.me.messages.post.call_args[0][0], self._HTML, '"contentType": "html"')

    async def test_update_draft(self):
        builder = MagicMock()
        builder.get = AsyncMock(return_value=MagicMock(is_draft=True))
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.messages.by_message_id.return_value = builder

        await mail_drafts.update_draft(
            client, "AAMkAG123=", body=self._MD, is_markdown=True, config=_CFG
        )

        assert_on_wire(builder.patch.call_args[0][0], self._HTML, '"contentType": "html"')

    async def test_reply_and_forward(self):
        for call in (
            lambda c: mail_write.reply(c, "AAMkAG123=", self._MD, is_markdown=True, config=_CFG),
            lambda c: mail_write.forward(
                c, "AAMkAG123=", ["a@example.com"], comment=self._MD, is_markdown=True, config=_CFG
            ),
        ):
            client, _, draft = TestMailWrite._thread_client()

            await call(client)

            assert_on_wire(draft.patch.call_args[0][0], self._HTML, '"contentType": "html"')


# ── folder paths (#144) ───────────────────────────────────────────────


def _folder_tree_client():
    """Inbox/Travel/Flights, listed level by level; the move/copy/create builders
    are shared so each test can read what was posted."""
    travel = MagicMock(id="AAMkTravelFolder=", display_name="Travel", child_folder_count=1)
    flights = MagicMock(id="AAMkFlightsFolder=", display_name="Flights", child_folder_count=0)
    children = {"inbox": [travel], "AAMkTravelFolder=": [flights]}
    builders = {}

    def by_folder(folder_id):
        if folder_id not in builders:
            b = MagicMock()
            b.child_folders.get = AsyncMock(
                return_value=MagicMock(value=children.get(folder_id, []), odata_next_link=None)
            )
            b.child_folders.post = AsyncMock(return_value=MagicMock(id="N1", display_name="n"))
            builders[folder_id] = b
        return builders[folder_id]

    client = MagicMock()
    client.me.mail_folders.by_mail_folder_id = MagicMock(side_effect=by_folder)
    message = MagicMock()
    message.move.post = AsyncMock()
    message.copy.post = AsyncMock()
    client.me.messages.by_message_id.return_value = message
    return client, message, builders


class TestFolderPaths:
    async def test_move_to_a_path_puts_that_folder_on_the_wire(self):
        client, message, _ = _folder_tree_client()

        await mail_triage.move_message(
            client, message_id="AAMkAG123=", folder="Inbox/Travel/Flights", config=_CFG
        )

        # Action-parameter keys are PascalCase on the wire, as for reply.
        assert_on_wire(message.move.post.call_args[0][0], '"DestinationId": "AAMkFlightsFolder="')

    async def test_copy_to_a_path_puts_that_folder_on_the_wire(self):
        client, message, _ = _folder_tree_client()

        await mail_thread.copy_message(
            client, message_id="AAMkAG123=", folder="Inbox/Travel", config=_CFG
        )

        assert_on_wire(message.copy.post.call_args[0][0], '"DestinationId": "AAMkTravelFolder="')

    async def test_create_folder_under_a_path_posts_to_that_folder(self):
        client, _, builders = _folder_tree_client()

        await mail_folders.create_folder(
            client, name="SENTINEL-FOLDER-c01", parent_folder="Inbox/Travel", config=_CFG
        )

        posted = builders["AAMkTravelFolder="].child_folders.post
        posted.assert_awaited_once()
        assert_on_wire(posted.call_args[0][0], "SENTINEL-FOLDER-c01")


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
            job_title="SENTINEL-JOBTITLE-b13",
            birthday="1987-06-05",
            middle_name="SentinelMiddle",
            personal_notes="SENTINEL-NOTES-b17",
            categories=["SENTINEL-CATEGORY-b18"],
            config=_CFG,
        )

        assert_on_wire(
            client.me.contacts.post.call_args[0][0],
            "SentinelFirst",
            "SentinelLast",
            "sentinel.contact@example.com",
            "+15555550100",
            "SENTINEL-COMPANY-b11",
            # Two fields a model can confuse, so each is pinned to its own key:
            # `title` is Graph's name prefix ("Dr."), `jobTitle` the job (#119).
            '"title": "SENTINEL-TITLE-b12"',
            '"jobTitle": "SENTINEL-JOBTITLE-b13"',
            # A date in, an instant out: Graph's `birthday` is a DateTimeOffset.
            '"birthday": "1987-06-05T11:59:00+00:00"',
            '"middleName": "SentinelMiddle"',
            '"personalNotes": "SENTINEL-NOTES-b17"',
            '"categories": ["SENTINEL-CATEGORY-b18"]',
        )

    async def test_create_contact_job_title_never_becomes_the_name_prefix(self):
        """#119: a job written to `title` became part of the display name."""
        client = MagicMock()
        client.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))

        await contacts.create_contact(client, first_name="Jane", job_title="Engineer", config=_CFG)

        body = wire(client.me.contacts.post.call_args[0][0])
        assert '"jobTitle": "Engineer"' in body
        assert '"title"' not in body

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
            company="SENTINEL-COMPANY-b16",
            job_title="SENTINEL-JOBTITLE-b14",
            home_address={
                "street": "SentinelStreet2",
                "city": "SentinelCity2",
                "state": "SentinelState2",
                "postal_code": "SentinelZip2",
                "country_or_region": "SentinelCountry2",
            },
            business_address={"city": "SentinelBusinessCity2"},
            other_address={"city": "SentinelOtherCity2"},
            birthday="1988-07-06",
            middle_name="SentinelMiddle2",
            personal_notes="SENTINEL-NOTES-b19",
            categories=["SENTINEL-CATEGORY-b20", "SENTINEL-CATEGORY-b21"],
            config=_CFG,
        )

        assert_on_wire(
            builder.patch.call_args[0][0],
            "SentinelFirst2",
            "SentinelLast2",
            "sentinel.contact2@example.com",
            "+15555550101",
            '"companyName": "SENTINEL-COMPANY-b16"',
            '"jobTitle": "SENTINEL-JOBTITLE-b14"',
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
            '"birthday": "1988-07-06T11:59:00+00:00"',
            '"middleName": "SentinelMiddle2"',
            '"personalNotes": "SENTINEL-NOTES-b19"',
            '"categories": ["SENTINEL-CATEGORY-b20", "SENTINEL-CATEGORY-b21"]',
        )

    async def test_update_contact_can_change_the_company(self):
        """Create sets `company`; update could not, so a job change left the old employer."""
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client, contact_id="AAMkAG123=", company="Contoso Ltd", config=_CFG
        )

        body = wire(builder.patch.call_args[0][0])
        assert '"companyName": "Contoso Ltd"' in body
        assert '"jobTitle"' not in body

    async def test_update_contact_categories_replace_the_list_and_empty_clears_it(self):
        """Graph replaces a contact's categories with the list sent; [] removes them all."""
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(client, contact_id="AAMkAG123=", categories=[], config=_CFG)

        body = wire(builder.patch.call_args[0][0])
        assert '"categories": []' in body

    async def test_update_contact_leaves_categories_alone_unless_given(self):
        """A notes edit must not touch the categories (or anything else)."""
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client, contact_id="AAMkAG123=", personal_notes="Met at the fair", config=_CFG
        )

        body = wire(builder.patch.call_args[0][0])
        assert '"personalNotes": "Met at the fair"' in body
        assert '"categories"' not in body
        assert '"middleName"' not in body

    async def test_update_contact_can_correct_the_name_prefix(self):
        """#119 follow-up: a job stuck in the prefix was only fixable by recreating."""
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client, contact_id="AAMkAG123=", title="SENTINEL-TITLE-b15", config=_CFG
        )

        assert_on_wire(builder.patch.call_args[0][0], '"title": "SENTINEL-TITLE-b15"')

    # #120: `phone` only ever wrote `mobilePhone`, so an office line from a
    # business card was saved as a mobile. Each type lands in its own Graph
    # field and nowhere else; the list fields carry exactly `[phone]`.
    @pytest.mark.parametrize(
        ("phone_type", "needle", "absent"),
        [
            ("mobile", '"mobilePhone": "+15555550102"', ("businessPhones", "homePhones")),
            ("business", '"businessPhones": ["+15555550102"]', ("mobilePhone", "homePhones")),
            ("home", '"homePhones": ["+15555550102"]', ("mobilePhone", "businessPhones")),
        ],
    )
    async def test_create_contact_phone_lands_in_its_type(self, phone_type, needle, absent):
        client = MagicMock()
        client.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))

        await contacts.create_contact(
            client, first_name="Jane", phone="+15555550102", phone_type=phone_type, config=_CFG
        )

        model = client.me.contacts.post.call_args[0][0]
        body = wire(model)
        assert needle in body, body
        for key in absent:
            assert f'"{key}"' not in body, body

    @pytest.mark.parametrize(
        ("phone_type", "needle", "absent"),
        [
            ("mobile", '"mobilePhone": "+15555550103"', ("businessPhones", "homePhones")),
            ("business", '"businessPhones": ["+15555550103"]', ("mobilePhone", "homePhones")),
            ("home", '"homePhones": ["+15555550103"]', ("mobilePhone", "businessPhones")),
        ],
    )
    async def test_update_contact_phone_lands_in_its_type(self, phone_type, needle, absent):
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client,
            contact_id="AAMkAG123=",
            phone="+15555550103",
            phone_type=phone_type,
            config=_CFG,
        )

        body = wire(builder.patch.call_args[0][0])
        assert needle in body, body
        for key in absent:
            assert f'"{key}"' not in body, body

    async def test_update_contact_moves_a_number_out_of_mobile_in_one_patch(self):
        """#120 follow-up: `phone_type` files a number but cannot empty the old slot.

        The issue's own live transcript is the target shape — one PATCH carrying
        `"mobilePhone": null` beside the new list. `mobile_phone = None` would
        not get there: the backing store emits it under its Python name.
        """
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client,
            contact_id="AAMkAG123=",
            phone="+15555550105",
            phone_type="business",
            clear_phones=["mobile"],
            config=_CFG,
        )

        body = wire(builder.patch.call_args[0][0])
        assert '"businessPhones": ["+15555550105"]' in body, body
        assert '"mobilePhone": null' in body, body
        assert '"homePhones"' not in body, body
        # The one null is the one asked for; anything else would clear data.
        assert_on_wire(builder.patch.call_args[0][0], '"mobilePhone": null', allow_null=True)
        nulls = sorted(k for k, v in json.loads(body).items() if v is None)
        assert nulls == ["mobilePhone"], body

    async def test_update_contact_empties_the_number_lists_it_is_told_to(self):
        builder = MagicMock()
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client, contact_id="AAMkAG123=", clear_phones=["business", "home"], config=_CFG
        )

        assert_on_wire(builder.patch.call_args[0][0], '"businessPhones": []', '"homePhones": []')
        assert '"mobilePhone"' not in wire(builder.patch.call_args[0][0])

    async def test_update_contact_merged_addresses_reach_the_wire(self):
        """`add_emails` / `remove_emails` send the whole merged list, and nothing else.

        The kept address goes back with the `name` it was read with, the added one
        with none, and the removed one not at all. A PATCH replaces the collection,
        so an address missing here is an address deleted.
        """
        stored = MagicMock(spec=["address", "name"])
        stored.address = "sentinel.kept@example.com"
        stored.name = "SENTINEL-NAME-b16"
        dropped = MagicMock(spec=["address", "name"])
        dropped.address = "sentinel.dropped@example.com"
        dropped.name = None
        current = MagicMock(spec=["email_addresses", "additional_data"])
        current.email_addresses = [stored, dropped]
        current.additional_data = {"@odata.etag": 'W/"sentinel"'}

        builder = MagicMock()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client,
            contact_id="AAMkAG123=",
            add_emails=["sentinel.added@example.com"],
            remove_emails=["sentinel.dropped@example.com"],
            config=_CFG,
        )

        model = builder.patch.call_args[0][0]
        assert_on_wire(
            model,
            '"emailAddresses": [{"address": "sentinel.kept@example.com", '
            '"name": "SENTINEL-NAME-b16"}, {"address": "sentinel.added@example.com"}]',
        )
        assert "sentinel.dropped@example.com" not in wire(builder.patch.call_args[0][0])

    async def test_update_contact_removing_every_address_sends_an_empty_list(self):
        current = MagicMock(spec=["email_addresses", "additional_data"])
        only = MagicMock(spec=["address", "name"])
        only.address = "sentinel.only@example.com"
        only.name = "Only"
        current.email_addresses = [only]
        current.additional_data = {}

        builder = MagicMock()
        builder.get = AsyncMock(return_value=current)
        builder.patch = AsyncMock()
        client = MagicMock()
        client.me.contacts.by_contact_id.return_value = builder

        await contacts.update_contact(
            client,
            contact_id="AAMkAG123=",
            remove_emails=["sentinel.only@example.com"],
            config=_CFG,
        )

        assert_on_wire(builder.patch.call_args[0][0], '"emailAddresses": []')

    async def test_phone_type_defaults_to_mobile(self):
        """Today's behaviour stays the default: existing callers are unaffected."""
        client = MagicMock()
        client.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))

        await contacts.create_contact(client, first_name="Jane", phone="+15555550104", config=_CFG)

        assert_on_wire(client.me.contacts.post.call_args[0][0], '"mobilePhone": "+15555550104"')


# ── to do ─────────────────────────────────────────────────────────────


def _todo_task_request(client):
    """The mocked request builder for one task: what `tasks.by_todo_task_id(...)` returns.

    Named once because the chain is wider than a line and has no call to break
    at: ruff format joins it back onto one line, which E501 then rejects.
    """
    return client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id.return_value


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

        patch = _todo_task_request(client).patch
        assert_on_wire(
            patch.call_args.args[0],
            "SENTINEL-TITLE-d31",
            "2026-11-06T17:00:00",
            "SENTINEL-BODY-d32",
            '"importance": "low"',
        )

    async def test_add_checklist_item_argument_reaches_the_wire(self):
        from tests.test_todo import _build_mock_client

        client = _build_mock_client()

        await todo.add_checklist_item(
            client,
            task_id="task1",
            display_name="SENTINEL-STEP-e41",
            config=_CFG,
        )

        post = _todo_task_request(client).checklist_items.post
        assert_on_wire(post.call_args.args[0], '"displayName": "SENTINEL-STEP-e41"')

    async def test_update_checklist_item_partial_patch_sends_only_what_was_asked(self):
        """The #63 leak shape: a partial PATCH where the caller picks the
        fields. A rename must not carry isChecked or checkedDateTime — an
        explicit null would clear server-side state the caller kept."""
        from tests.test_todo import _build_mock_client

        client = _build_mock_client()

        await todo.update_checklist_item(
            client,
            task_id="task1",
            checklist_item_id="ci1",
            display_name="SENTINEL-RENAME-e42",
            config=_CFG,
        )

        item = _todo_task_request(client).checklist_items.by_checklist_item_id.return_value
        payload = item.patch.call_args.args[0]
        assert_on_wire(payload, '"displayName": "SENTINEL-RENAME-e42"')

        body = wire(payload)
        assert '"isChecked"' not in body, body
        assert "checkedDateTime" not in body, body

    async def test_update_checklist_item_is_checked_reaches_the_wire(self):
        """The positive half, which the wire suite never asserted: a check
        must serialize as isChecked:true on the PATCH — the model attribute
        existing proves nothing about the JSON kiota sends."""
        from tests.test_todo import _build_mock_client

        client = _build_mock_client()

        await todo.update_checklist_item(
            client,
            task_id="task1",
            checklist_item_id="ci1",
            is_checked=True,
            config=_CFG,
        )

        item = _todo_task_request(client).checklist_items.by_checklist_item_id.return_value
        payload = item.patch.call_args.args[0]
        assert_on_wire(payload, '"isChecked": true')

    async def test_upload_task_attachment_argument_reaches_the_wire(self, tmp_path):
        """The inline-POST payload: base64 contentBytes (not a filename, not
        a session reference) plus the metadata Graph needs to rebuild it."""
        import base64 as _b64

        from tests.test_todo_attachments import _build_mock_client, _cfg

        content = b"SENTINEL-UPLOAD-BYTES-e43"
        src = tmp_path / "att" / "SENTINEL-UPLOAD-e43.bin"
        src.parent.mkdir(parents=True)
        src.write_bytes(content)

        client = _build_mock_client()

        await todo_attachments.upload_task_attachment(
            client,
            task_id="task1",
            file_path=str(src),
            config=_cfg(tmp_path),
        )

        post = _todo_task_request(client).attachments.post
        assert_on_wire(
            post.call_args.args[0],
            '"@odata.type": "#microsoft.graph.taskFileAttachment"',
            '"name": "SENTINEL-UPLOAD-e43.bin"',
            f'"size": {len(content)}',
            f'"contentBytes": "{_b64.b64encode(content).decode()}"',
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
