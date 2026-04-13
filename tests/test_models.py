"""Tests for Pydantic models."""

import pytest

from outlook_mcp.models.calendar import CreateEventInput, RsvpInput
from outlook_mcp.models.common import ListResponse
from outlook_mcp.models.mail import (
    DeleteInput,
    ForwardInput,
    ReplyInput,
    SendMessageInput,
    TriageInput,
)


class TestMailModels:
    def test_send_message_validates_emails(self):
        msg = SendMessageInput(
            to=["valid@outlook.com"],
            subject="Test",
            body="Hello",
        )
        assert msg.to == ["valid@outlook.com"]

    def test_send_message_rejects_empty_to(self):
        with pytest.raises(ValueError):
            SendMessageInput(to=[], subject="Test", body="Hello")

    def test_send_message_defaults(self):
        msg = SendMessageInput(
            to=["a@b.com"], subject="Test", body="Hello"
        )
        assert msg.is_html is False
        assert msg.importance == "normal"
        assert msg.cc is None
        assert msg.bcc is None

    def test_send_message_rejects_bad_importance(self):
        with pytest.raises(ValueError):
            SendMessageInput(
                to=["a@b.com"], subject="Test", body="Hello", importance="urgent"
            )

    def test_triage_input(self):
        t = TriageInput(message_id="AAMkAG123=", action="flag", value="flagged")
        assert t.value == "flagged"

    def test_delete_input_defaults(self):
        d = DeleteInput(message_id="AAMkAG123=")
        assert d.permanent is False

    def test_reply_input(self):
        r = ReplyInput(message_id="AAMkAG123=", body="Thanks!", reply_all=False)
        assert r.reply_all is False

    def test_forward_rejects_empty_to(self):
        with pytest.raises(ValueError):
            ForwardInput(message_id="AAMkAG123=", to=[])


class TestCalendarModels:
    def test_create_event_required_fields(self):
        e = CreateEventInput(
            subject="Meeting",
            start="2026-04-15T10:00:00",
            end="2026-04-15T11:00:00",
        )
        assert e.subject == "Meeting"
        assert e.is_all_day is False
        assert e.is_online is False

    def test_rsvp_validates_response(self):
        r = RsvpInput(event_id="AAMkAG123=", response="accept")
        assert r.response == "accept"

    def test_rsvp_rejects_invalid_response(self):
        with pytest.raises(ValueError):
            RsvpInput(event_id="AAMkAG123=", response="maybe")

    def test_create_event_rejects_bad_recurrence(self):
        with pytest.raises(ValueError):
            CreateEventInput(
                subject="Meeting",
                start="2026-04-15T10:00:00",
                end="2026-04-15T11:00:00",
                recurrence="biweekly",
            )


class TestCommonModels:
    def test_list_response(self):
        resp = ListResponse(items=[{"id": "1"}], count=1, has_more=False)
        assert resp.count == 1
