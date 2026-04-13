"""To Do tools: task lists, tasks — CRUD + complete."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import apply_pagination, wrap_nextlink
from outlook_mcp.validation import sanitize_output, validate_datetime, validate_graph_id


def _check_read_only(read_only: bool, tool_name: str) -> None:
    """Raise ReadOnlyError if server is in read-only mode."""
    if read_only:
        raise ReadOnlyError(tool_name)


async def _resolve_list_id(graph_client: Any, list_id: str | None) -> str:
    """Resolve list_id — use provided value or find the default list.

    Default list: isOwner=True and wellknownListName="defaultList".
    Falls back to the first list if no explicit default is found.
    """
    if list_id:
        return list_id

    response = await graph_client.me.todo.lists.get()
    lists = response.value or []

    if not lists:
        raise ValueError("No task lists found. Create a list in Microsoft To Do first.")

    # Prefer the default list
    for lst in lists:
        wellknown = ""
        if lst.wellknown_list_name:
            wellknown = (
                lst.wellknown_list_name.value
                if hasattr(lst.wellknown_list_name, "value")
                else str(lst.wellknown_list_name)
            )
        if lst.is_owner and wellknown == "defaultList":
            return lst.id

    # Fallback: first list
    return lists[0].id


def _format_task(task: Any) -> dict:
    """Convert a Graph SDK TodoTask to a clean dict."""
    status = "notStarted"
    if task.status:
        status = task.status.value if hasattr(task.status, "value") else str(task.status)

    importance = "normal"
    if task.importance:
        importance = (
            task.importance.value if hasattr(task.importance, "value") else str(task.importance)
        )

    body_content = ""
    if task.body and task.body.content:
        body_content = sanitize_output(task.body.content, multiline=True)

    due = None
    if task.due_date_time:
        # Graph returns DateTimeTimeZone object for due
        if hasattr(task.due_date_time, "date_time"):
            due = task.due_date_time.date_time
        else:
            due = str(task.due_date_time)

    completed = None
    if task.completed_date_time:
        if hasattr(task.completed_date_time, "date_time"):
            completed = task.completed_date_time.date_time
        else:
            completed = str(task.completed_date_time)

    return {
        "id": task.id,
        "title": sanitize_output(task.title or ""),
        "status": status,
        "importance": importance,
        "due": due,
        "completed": completed,
        "created": str(task.created_date_time or ""),
        "is_reminder_on": bool(task.is_reminder_on),
        "body": body_content,
        "has_recurrence": task.recurrence is not None,
    }


async def list_task_lists(graph_client: Any) -> dict:
    """List all To Do task lists.

    GET /me/todo/lists
    Returns {task_lists: [{id, display_name, is_default}], count}.
    """
    response = await graph_client.me.todo.lists.get()
    lists = response.value or []

    task_lists = []
    for lst in lists:
        wellknown = ""
        if lst.wellknown_list_name:
            wellknown = (
                lst.wellknown_list_name.value
                if hasattr(lst.wellknown_list_name, "value")
                else str(lst.wellknown_list_name)
            )
        is_default = bool(lst.is_owner and wellknown == "defaultList")

        task_lists.append(
            {
                "id": lst.id,
                "display_name": sanitize_output(lst.display_name or ""),
                "is_default": is_default,
            }
        )

    return {
        "task_lists": task_lists,
        "count": len(task_lists),
    }


async def list_tasks(
    graph_client: Any,
    list_id: str | None = None,
    status: str | None = None,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List tasks in a To Do list.

    GET /me/todo/lists/{id}/tasks
    If list_id is None, uses the default list.
    Filter by status: notStarted, inProgress, completed.
    """
    resolved_id = await _resolve_list_id(graph_client, list_id)

    query_params: dict[str, Any] = {
        "$orderby": "createdDateTime desc",
    }

    # Status filter
    if status:
        valid_statuses = {"notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"}
        if status not in valid_statuses:
            raise ValueError(f"Invalid status '{status}'. Must be one of: {valid_statuses}")
        query_params["$filter"] = f"status eq '{status}'"

    # Pagination
    query_params = apply_pagination(query_params, count, cursor)

    response = await graph_client.me.todo.lists.by_todo_task_list_id(resolved_id).tasks.get(
        query_params=query_params
    )

    tasks = [_format_task(t) for t in (response.value or [])]
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "tasks": tasks,
        "count": len(tasks),
        "has_more": next_cursor is not None,
        "next_cursor": next_cursor,
    }


async def create_task(
    graph_client: Any,
    title: str,
    list_id: str | None = None,
    due: str | None = None,
    importance: str | None = None,
    body: str | None = None,
    reminder: bool | None = None,
    recurrence: dict | None = None,
    read_only: bool = False,
) -> dict:
    """Create a task in a To Do list.

    POST /me/todo/lists/{id}/tasks
    Validates due date if provided.
    """
    _check_read_only(read_only, "outlook_create_task")

    resolved_id = await _resolve_list_id(graph_client, list_id)

    # Build the task body as a dict for the SDK
    task_body: dict[str, Any] = {"title": title}

    if due:
        validate_datetime(due)
        # Graph To Do uses DateTimeTimeZone for dueDateTime
        task_body["dueDateTime"] = {
            "dateTime": due,
            "timeZone": "UTC",
        }

    if importance:
        valid_importances = {"low", "normal", "high"}
        if importance not in valid_importances:
            raise ValueError(
                f"Invalid importance '{importance}'. Must be one of: {valid_importances}"
            )
        task_body["importance"] = importance

    if body:
        task_body["body"] = {
            "content": body,
            "contentType": "text",
        }

    if reminder is not None:
        task_body["isReminderOn"] = reminder

    if recurrence:
        task_body["recurrence"] = recurrence

    response = await graph_client.me.todo.lists.by_todo_task_list_id(resolved_id).tasks.post(
        task_body
    )

    return {
        "status": "created",
        "task_id": response.id,
        "title": sanitize_output(response.title or ""),
    }


async def update_task(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
    title: str | None = None,
    due: str | None = None,
    body: str | None = None,
    importance: str | None = None,
    read_only: bool = False,
) -> dict:
    """Update a task in a To Do list.

    PATCH /me/todo/lists/{id}/tasks/{taskId}
    Only patches provided fields.
    """
    _check_read_only(read_only, "outlook_update_task")
    task_id = validate_graph_id(task_id)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    patch_body: dict[str, Any] = {}

    if title is not None:
        patch_body["title"] = title

    if due is not None:
        validate_datetime(due)
        patch_body["dueDateTime"] = {
            "dateTime": due,
            "timeZone": "UTC",
        }

    if body is not None:
        patch_body["body"] = {
            "content": body,
            "contentType": "text",
        }

    if importance is not None:
        valid_importances = {"low", "normal", "high"}
        if importance not in valid_importances:
            raise ValueError(
                f"Invalid importance '{importance}'. Must be one of: {valid_importances}"
            )
        patch_body["importance"] = importance

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .patch(patch_body)
    )

    return {
        "status": "updated",
        "task_id": task_id,
    }


async def complete_task(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
    read_only: bool = False,
) -> dict:
    """Mark a task as completed.

    PATCH with status="completed" and completedDateTime set to now (UTC).
    """
    _check_read_only(read_only, "outlook_complete_task")
    task_id = validate_graph_id(task_id)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.0000000Z")

    patch_body = {
        "status": "completed",
        "completedDateTime": {
            "dateTime": now_utc,
            "timeZone": "UTC",
        },
    }

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .patch(patch_body)
    )

    return {
        "status": "completed",
        "task_id": task_id,
    }


async def delete_task(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
    read_only: bool = False,
) -> dict:
    """Delete a task from a To Do list.

    DELETE /me/todo/lists/{id}/tasks/{taskId}
    """
    _check_read_only(read_only, "outlook_delete_task")
    task_id = validate_graph_id(task_id)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .delete()
    )

    return {
        "status": "deleted",
        "task_id": task_id,
    }
