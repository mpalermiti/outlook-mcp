"""Tests for To Do tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.tools.todo import (
    complete_task,
    create_task,
    delete_task,
    list_task_lists,
    list_tasks,
    update_task,
)


def _mock_task_list(
    list_id="list1",
    display_name="Tasks",
    is_owner=True,
    wellknown="defaultList",
):
    """Helper to build a mock task list."""
    mock = MagicMock()
    mock.id = list_id
    mock.display_name = display_name
    mock.is_owner = is_owner
    mock.wellknown_list_name = MagicMock(value=wellknown)
    return mock


def _mock_task(
    task_id="task1",
    title="Buy groceries",
    status="notStarted",
    importance="normal",
    due=None,
    reminder=False,
    created="2026-04-12T10:00:00Z",
    completed=None,
    body_content="",
    body_type="text",
    recurrence=None,
):
    """Helper to build a mock task."""
    mock = MagicMock()
    mock.id = task_id
    mock.title = title
    mock.status = MagicMock(value=status)
    mock.importance = MagicMock(value=importance)
    mock.due_date_time = due
    mock.is_reminder_on = reminder
    mock.created_date_time = created
    mock.completed_date_time = completed
    mock.body = MagicMock(content=body_content, content_type=MagicMock(value=body_type))
    mock.recurrence = recurrence
    return mock


def _build_mock_client(lists=None, tasks=None, odata_next_link=None):
    """Build a fully-wired mock Graph client for To Do operations."""
    if lists is None:
        lists = [_mock_task_list()]
    if tasks is None:
        tasks = [_mock_task()]

    mock_client = MagicMock()

    # GET /me/todo/lists
    mock_client.me.todo.lists.get = AsyncMock(return_value=MagicMock(value=lists))

    # Task-level mocks
    mock_task_item = MagicMock()
    mock_task_item.patch = AsyncMock()
    mock_task_item.delete = AsyncMock()

    mock_tasks = MagicMock()
    mock_tasks.get = AsyncMock(return_value=MagicMock(value=tasks, odata_next_link=odata_next_link))
    mock_tasks.post = AsyncMock(return_value=tasks[0] if tasks else _mock_task())
    mock_tasks.by_todo_task_id = MagicMock(return_value=mock_task_item)

    mock_list_item = MagicMock()
    mock_list_item.tasks = mock_tasks

    mock_client.me.todo.lists.by_todo_task_list_id = MagicMock(return_value=mock_list_item)

    return mock_client


# --- list_task_lists ---


class TestListTaskLists:
    async def test_returns_task_lists(self):
        """list_task_lists returns formatted list of task lists."""
        lists = [
            _mock_task_list("list1", "Tasks", True, "defaultList"),
            _mock_task_list("list2", "Shopping", False, "none"),
        ]
        client = _build_mock_client(lists=lists)

        result = await list_task_lists(client)

        assert result["count"] == 2
        assert result["task_lists"][0]["id"] == "list1"
        assert result["task_lists"][0]["display_name"] == "Tasks"
        assert result["task_lists"][0]["is_default"] is True
        assert result["task_lists"][1]["is_default"] is False
        client.me.todo.lists.get.assert_called_once()

    async def test_empty_lists(self):
        """list_task_lists handles no task lists."""
        client = _build_mock_client(lists=[])

        result = await list_task_lists(client)

        assert result["count"] == 0
        assert result["task_lists"] == []


# --- list_tasks ---


class TestListTasks:
    async def test_list_tasks_default_list(self):
        """list_tasks resolves default list when list_id is None."""
        client = _build_mock_client()

        result = await list_tasks(client)

        assert result["count"] == 1
        assert result["tasks"][0]["id"] == "task1"
        assert result["tasks"][0]["title"] == "Buy groceries"
        assert result["tasks"][0]["status"] == "notStarted"
        # Should have resolved default list
        client.me.todo.lists.get.assert_called_once()
        client.me.todo.lists.by_todo_task_list_id.assert_called_with("list1")

    async def test_list_tasks_explicit_list(self):
        """list_tasks uses provided list_id without resolving default."""
        client = _build_mock_client()

        result = await list_tasks(client, list_id="mylist123")

        assert result["count"] == 1
        # Should NOT have called lists.get to resolve default
        client.me.todo.lists.get.assert_not_called()
        client.me.todo.lists.by_todo_task_list_id.assert_called_with("mylist123")

    async def test_list_tasks_with_status_filter(self):
        """list_tasks passes status filter to query params."""
        client = _build_mock_client()

        result = await list_tasks(client, status="completed")

        assert result["count"] == 1
        # Verify filter was passed via request_configuration
        call_kwargs = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.get.call_args
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert qp.filter is not None
        assert "completed" in qp.filter

    async def test_list_tasks_pagination(self):
        """list_tasks returns next_cursor when odata_next_link present."""
        client = _build_mock_client(
            odata_next_link="https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks?$skip=25"
        )

        result = await list_tasks(client)

        assert result["has_more"] is True
        assert result["next_cursor"] is not None


# --- create_task ---


class TestCreateTask:
    async def test_create_basic_task(self):
        """create_task creates a task with title on default list."""
        client = _build_mock_client()

        result = await create_task(client, title="Buy milk")

        assert result["status"] == "created"
        assert result["task_id"] == "task1"
        client.me.todo.lists.by_todo_task_list_id.return_value.tasks.post.assert_called_once()

    async def test_create_task_with_due_date(self):
        """create_task validates and sets due date."""
        client = _build_mock_client()

        result = await create_task(client, title="Report", due="2026-04-15")

        assert result["status"] == "created"

    async def test_create_task_invalid_due_date(self):
        """create_task rejects invalid due date."""
        client = _build_mock_client()

        with pytest.raises(ValueError):
            await create_task(client, title="Bad date", due="not-a-date")

    async def test_create_task_read_only(self):
        """create_task raises ReadOnlyError in read-only mode."""
        client = _build_mock_client()

        with pytest.raises(ReadOnlyError):
            await create_task(client, title="Blocked", read_only=True)


# --- update_task ---


class TestUpdateTask:
    async def test_update_task_title(self):
        """update_task patches task with new title."""
        client = _build_mock_client()

        result = await update_task(client, task_id="task1", title="Updated title")

        assert result["status"] == "updated"
        mock_item = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id
        mock_item.assert_called_with("task1")
        mock_item.return_value.patch.assert_called_once()

    async def test_update_task_read_only(self):
        """update_task raises ReadOnlyError in read-only mode."""
        client = _build_mock_client()

        with pytest.raises(ReadOnlyError):
            await update_task(client, task_id="task1", title="Nope", read_only=True)

    async def test_update_task_validates_id(self):
        """update_task validates the task_id."""
        client = _build_mock_client()

        with pytest.raises(ValueError):
            await update_task(client, task_id="", title="Bad")


# --- complete_task ---


class TestCompleteTask:
    async def test_complete_task(self):
        """complete_task patches task with completed status."""
        client = _build_mock_client()

        result = await complete_task(client, task_id="task1")

        assert result["status"] == "completed"
        mock_item = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id
        mock_item.assert_called_with("task1")
        mock_item.return_value.patch.assert_called_once()

    async def test_complete_task_read_only(self):
        """complete_task raises ReadOnlyError in read-only mode."""
        client = _build_mock_client()

        with pytest.raises(ReadOnlyError):
            await complete_task(client, task_id="task1", read_only=True)


# --- delete_task ---


class TestDeleteTask:
    async def test_delete_task(self):
        """delete_task calls DELETE on the task."""
        client = _build_mock_client()

        result = await delete_task(client, task_id="task1")

        assert result["status"] == "deleted"
        mock_item = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id
        mock_item.assert_called_with("task1")
        mock_item.return_value.delete.assert_called_once()

    async def test_delete_task_read_only(self):
        """delete_task raises ReadOnlyError in read-only mode."""
        client = _build_mock_client()

        with pytest.raises(ReadOnlyError):
            await delete_task(client, task_id="task1", read_only=True)
