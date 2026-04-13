"""Tests for Tier 2 Pydantic models (Contacts + To Do)."""

import pytest

from outlook_mcp.models.contacts import (
    ContactSummary,
    CreateContactInput,
    UpdateContactInput,
)
from outlook_mcp.models.todo import (
    CreateTaskInput,
    TaskDetail,
    TaskListSummary,
    TaskSummary,
    UpdateTaskInput,
)


class TestContactModels:
    def test_contact_summary(self):
        c = ContactSummary(id="abc123", display_name="John Doe", email="john@test.com")
        assert c.display_name == "John Doe"

    def test_create_contact_requires_first_name(self):
        c = CreateContactInput(first_name="John")
        assert c.first_name == "John"
        assert c.last_name is None

    def test_create_contact_all_fields(self):
        c = CreateContactInput(
            first_name="John",
            last_name="Doe",
            email="john@test.com",
            phone="+1234567890",
            company="Acme",
            title="Engineer",
        )
        assert c.company == "Acme"

    def test_update_contact_partial(self):
        u = UpdateContactInput(contact_id="abc123", email="new@test.com")
        assert u.first_name is None
        assert u.email == "new@test.com"


class TestTodoModels:
    def test_task_list_summary(self):
        tl = TaskListSummary(id="list1", display_name="My Tasks", is_default=True)
        assert tl.is_default is True

    def test_task_summary_defaults(self):
        t = TaskSummary(id="task1", title="Buy milk")
        assert t.status == "notStarted"
        assert t.importance == "normal"

    def test_create_task_minimal(self):
        t = CreateTaskInput(title="Buy milk")
        assert t.list_id is None
        assert t.due is None

    def test_create_task_validates_importance(self):
        with pytest.raises(ValueError):
            CreateTaskInput(title="Test", importance="critical")

    def test_create_task_validates_recurrence(self):
        with pytest.raises(ValueError):
            CreateTaskInput(title="Test", recurrence="biweekly")

    def test_create_task_valid_recurrence(self):
        t = CreateTaskInput(title="Standup", recurrence="weekdays")
        assert t.recurrence == "weekdays"

    def test_task_detail_has_body(self):
        t = TaskDetail(id="task1", title="Test", body="Details here")
        assert t.body == "Details here"

    def test_update_task_partial(self):
        u = UpdateTaskInput(task_id="task1", title="Updated title")
        assert u.due is None
