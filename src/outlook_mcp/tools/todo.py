"""To Do tools: task lists, tasks — CRUD + complete."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any
from weakref import WeakKeyDictionary

from outlook_mcp.config import Config
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_TODO_WRITE, check_permission
from outlook_mcp.tools._recurrence import build_patterned_recurrence as _build_recurrence
from outlook_mcp.validation import sanitize_output, validate_datetime, validate_graph_id

_VALID_IMPORTANCES = {"low", "normal", "high"}

# Backstop for the nextLink walks (see fetch_all_task_lists): 20 pages of
# $top=100 is 2,000 lists/checklist pages — far past any real mailbox — so a
# server that keeps handing us a nextLink is broken, not big.
_PAGE_CAP = 20


def _importance_enum(value: str) -> Any:
    """Map a string importance to the SDK Importance enum."""
    from msgraph.generated.models.importance import Importance

    if value not in _VALID_IMPORTANCES:
        raise ValueError(
            f"Invalid importance '{value}'. Must be one of: {sorted(_VALID_IMPORTANCES)}"
        )
    return {
        "low": Importance.Low,
        "normal": Importance.Normal,
        "high": Importance.High,
    }[value]


def _datetime_timezone(iso_dt: str, tz: str = "UTC") -> Any:
    """Wrap an ISO datetime string in a Graph DateTimeTimeZone typed model."""
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone

    dtz = DateTimeTimeZone()
    dtz.date_time = iso_dt
    dtz.time_zone = tz
    return dtz


def _text_body(content: str) -> Any:
    """Wrap a text string in a Graph ItemBody typed model."""
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.item_body import ItemBody

    ib = ItemBody()
    ib.content = content
    ib.content_type = BodyType.Text
    return ib


def _iso_datetime(value: Any) -> str:
    """Normalize an SDK datetime field to ISO 8601, or ``""`` when absent.

    kiota deserializes ``createdDateTime`` / ``checkedDateTime`` into real
    ``datetime`` objects, and ``str(datetime)`` is ``"2026-09-15 09:00:00+00:00"``
    — space separator, no ``T``, no ``Z`` — while Graph's own DateTimeTimeZone
    strings (``due``) are already ISO. One response, two incompatible formats.
    Strings pass through untouched, so a mock (or a Graph field that stays a
    string) is not reformatted.
    """
    if isinstance(value, datetime):
        return value.isoformat()
    return str(value or "")


# Default-list resolution per GraphServiceClient — see _resolve_list_id. A
# WeakKeyDictionary so a retired client's entry goes with it, and so tests that
# build a fresh mock client per case never share one.
_DEFAULT_LIST_BY_CLIENT: WeakKeyDictionary = WeakKeyDictionary()


def _lists_first_page_config() -> Any:
    """RequestConfiguration for the first ``/me/todo/lists`` page."""
    from msgraph.generated.users.item.todo.lists.lists_request_builder import (
        ListsRequestBuilder,
    )

    return build_request_config(
        ListsRequestBuilder.ListsRequestBuilderGetQueryParameters, {"$top": 100}
    )


async def fetch_all_task_lists(graph_client: Any, top: int = 100) -> list[Any]:
    """Every task list in ``/me/todo/lists``, following ``@odata.nextLink``.

    The pages are followed with ``with_url`` on the raw link, the same way
    ``calendar_resolver.fetch_all_calendars`` walks calendars. That matters
    here because To Do lists page with ``$skiptoken``, and the SDK's typed
    ``ListsRequestBuilderGetQueryParameters`` has no ``skiptoken`` field —
    rebuilding query parameters for page two silently drops the token and
    re-fetches page one forever. The raw link is followed verbatim instead,
    with a page cap so a misbehaving server cannot loop us.
    """
    response = await graph_client.me.todo.lists.get(
        request_configuration=_lists_first_page_config()
    )
    collected = list(response.value) if response and response.value else []
    pages = 1
    while True:
        next_link = getattr(response, "odata_next_link", None)
        if not isinstance(next_link, str) or not next_link:
            break
        pages += 1
        if pages > _PAGE_CAP:
            raise ValueError(
                f"/me/todo/lists kept returning a nextLink for {pages} pages "
                f"({len(collected)} lists so far) — refusing to walk further; "
                "this looks like a broken paging response, not a big mailbox"
            )
        response = await graph_client.me.todo.lists.with_url(next_link).get()
        collected.extend(list(response.value) if response and response.value else [])
    return collected


async def _resolve_list_id(graph_client: Any, list_id: str | None) -> str:
    """Resolve list_id — use provided value or find the default list.

    Default list: isOwner=True and wellknownListName="defaultList". Falls back
    to the first list if no explicit default is found.

    ``None`` means "the default list"; anything else — including the empty
    string, which optional-string schemas hand us — is validated as a Graph id
    rather than silently retargeting the default list.

    The default-list lookup is cached per GraphServiceClient: every one of the
    To Do tools resolves on each call, which doubled the request count of a
    normal checklist flow. The cache lives as long as the client does — the
    server rebuilds the client on an account switch, so a different mailbox
    starts cold. The trade-off is a stale id if the user deletes their default
    list mid-session; Graph answers 404 with a "re-list" hint, and nothing in
    this surface creates or deletes lists.

    Only a *found* defaultList is cached. The first-list fallback is not: a
    mailbox whose ``defaultList`` entry is missing (renamed away, or hidden by
    paging trouble) would otherwise pin what may be a shared list for the life
    of the process — the next call re-resolves and can find the real default.

    The walk follows the raw ``@odata.nextLink`` with ``with_url`` (see
    ``fetch_all_task_lists``): rebuilding typed query parameters drops
    ``$skiptoken`` — the SDK class has no such field — so a rebuilt request
    re-fetches page one and the walk never advances.
    """
    if list_id is not None:
        if not list_id.strip():
            raise ValueError(
                "list_id must be a task-list id, or omitted (None) for the "
                "default list — an empty string would silently target the "
                "default list"
            )
        return validate_graph_id(list_id)

    cached = _DEFAULT_LIST_BY_CLIENT.get(graph_client)
    if cached is not None:
        return cached

    first_list_id: str | None = None
    response = await graph_client.me.todo.lists.get(
        request_configuration=_lists_first_page_config()
    )
    pages = 1
    while True:
        for lst in (response.value if response else None) or []:
            if first_list_id is None:
                first_list_id = lst.id
            wellknown = ""
            if lst.wellknown_list_name:
                wellknown = (
                    lst.wellknown_list_name.value
                    if hasattr(lst.wellknown_list_name, "value")
                    else str(lst.wellknown_list_name)
                )
            if lst.is_owner and wellknown == "defaultList":
                _DEFAULT_LIST_BY_CLIENT[graph_client] = lst.id
                return lst.id
        # The default list can sit past the first page; keep walking until
        # Graph stops handing us a nextLink — following the raw link, not
        # rebuilt parameters (see the docstring).
        next_link = getattr(response, "odata_next_link", None)
        if not isinstance(next_link, str) or not next_link:
            break
        pages += 1
        if pages > _PAGE_CAP:
            raise ValueError(
                f"/me/todo/lists kept returning a nextLink for {pages} pages "
                "while looking for the default list — refusing to walk further"
            )
        response = await graph_client.me.todo.lists.with_url(next_link).get()

    if first_list_id is None:
        raise ValueError("No task lists found. Create a list in Microsoft To Do first.")

    # Fallback: first list. Not cached — see the docstring.
    return first_list_id


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
        "created": _iso_datetime(task.created_date_time),
        "is_reminder_on": bool(task.is_reminder_on),
        "body": body_content,
        "has_recurrence": task.recurrence is not None,
    }


def _format_checklist_item(item: Any) -> dict:
    """Convert a Graph SDK ChecklistItem to a clean dict."""
    return {
        "id": item.id,
        "display_name": sanitize_output(item.display_name or ""),
        "is_checked": bool(item.is_checked),
        "checked_at": _iso_datetime(item.checked_date_time),
        "created": _iso_datetime(item.created_date_time),
    }


def _checked_last(items: list[dict]) -> list[dict]:
    """Order checklist items unchecked-first, then by creation time.

    Graph returns expanded checklistItems in no guaranteed order (it hands
    back whatever the backing store produced), while every To Do client
    surface shows open steps above completed ones. An agent reporting task
    progress reads the first unchecked item as "the next step", so the order
    we emit is load-bearing, not cosmetic — and it has to be *deterministic*,
    not just unchecked-first: sorting on the boolean alone leaves the order
    inside each group as whatever Graph returned, so two identical calls on an
    unchanged task could name a different "next step". ``created`` (emitted by
    ``_format_checklist_item``, ISO so it sorts chronologically) is the
    tiebreak. ISO strings sort correctly; the ``""`` for a missing created
    sorts first, harmlessly.
    """
    return sorted(items, key=lambda i: (i["is_checked"], i["created"]))


async def list_task_lists(graph_client: Any) -> dict:
    """List all To Do task lists.

    GET /me/todo/lists
    Returns {task_lists: [{id, display_name, is_default}], count, has_more,
    next_cursor}. The walk follows every ``@odata.nextLink`` server-side (see
    ``fetch_all_task_lists``), so the listing is always complete: ``has_more``
    is False and ``next_cursor`` is None on every return — the pair is emitted
    anyway so the shape matches the other list tools, and a caller that polls
    on it simply stops.
    """
    lists = await fetch_all_task_lists(graph_client)

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
        "has_more": False,
        "next_cursor": None,
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

    from msgraph.generated.users.item.todo.lists.item.tasks.tasks_request_builder import (
        TasksRequestBuilder,
    )

    req_config = build_request_config(
        TasksRequestBuilder.TasksRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.todo.lists.by_todo_task_list_id(resolved_id).tasks.get(
        request_configuration=req_config
    )

    tasks = [_format_task(t) for t in (response.value or [])]
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "tasks": tasks,
        "count": len(tasks),
        "has_more": next_cursor is not None,
        "next_cursor": next_cursor,
    }


async def get_task(
    graph_client: Any,
    task_id: str,
    list_id: str | None = None,
) -> dict:
    """Get full task details: notes (body), checklist items, recurrence flag.

    GET /me/todo/lists/{id}/tasks/{taskId}?$expand=checklistItems
    """
    task_id = validate_graph_id(task_id)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.users.item.todo.lists.item.tasks.item.todo_task_item_request_builder import (  # noqa: E501
        TodoTaskItemRequestBuilder,
    )

    req_config = build_request_config(
        TodoTaskItemRequestBuilder.TodoTaskItemRequestBuilderGetQueryParameters,
        {"$expand": "checklistItems"},
    )
    task = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .get(request_configuration=req_config)
    )
    if task is None or task.id is None:
        # Optional[Task] on the SDK side: an empty 200/204 used to crash in
        # _format_task with an AttributeError whose text never reached the model.
        raise ValueError(
            f"Graph returned no task for id {task_id} — the id may be stale; "
            "re-list with outlook_list_tasks for current ids"
        )

    result = _format_task(task)
    # The SDK model types checklist_items as a plain list[ChecklistItem] — with
    # $expand Graph fills that list directly, and an empty expansion is [].
    # There is no collection response to unwrap (verified live: every task
    # carrying sub-steps crashed on a phantom `.value`).
    raw_items = list(getattr(task, "checklist_items", None) or [])
    # Defensive: an $expand whose result exceeds Graph's page limit comes back
    # truncated with `checklistItems@odata.nextLink` in the task's
    # additional_data. Whether To Do ever actually pages this expansion is
    # unverified (forcing it live would take a checklist past the page size,
    # and this is a read path) — so this follows the link if it ever appears
    # rather than assuming it cannot, the same with_url walk the lists use.
    next_link = (getattr(task, "additional_data", None) or {}).get(
        "checklistItems@odata.nextLink"
    )
    pages = 1
    while isinstance(next_link, str) and next_link:
        pages += 1
        if pages > _PAGE_CAP:
            raise ValueError(
                f"checklistItems expansion kept returning a nextLink for "
                f"{pages} pages on task {task_id} — refusing to walk further"
            )
        page = await (
            graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
            .tasks.by_todo_task_id(task_id)
            .checklist_items.with_url(next_link)
            .get()
        )
        raw_items.extend(list(page.value) if page and page.value else [])
        next_link = getattr(page, "odata_next_link", None)
    items = [_format_checklist_item(i) for i in raw_items]
    result["checklist_items"] = _checked_last(items)
    result["checklist_count"] = len(items)
    return result


async def create_task(
    graph_client: Any,
    title: str,
    list_id: str | None = None,
    due: str | None = None,
    importance: str | None = None,
    body: str | None = None,
    reminder: bool | None = None,
    recurrence: dict | None = None,
    *,
    config: Config,
) -> dict:
    """Create a task in a To Do list.

    POST /me/todo/lists/{id}/tasks
    Validates due date if provided.
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_create_task")

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.todo_task import TodoTask

    task_body = TodoTask()
    task_body.title = title

    if due:
        validate_datetime(due)
        task_body.due_date_time = _datetime_timezone(due)

    if importance:
        task_body.importance = _importance_enum(importance)

    if body:
        task_body.body = _text_body(body)

    if reminder is not None:
        # Graph silently stores isReminderOn=false unless reminderDateTime is
        # also set (verified live). Anchor the reminder on the due time; without
        # a due there is nothing to anchor it on, so say so instead of no-oping.
        if reminder and not due:
            raise ValueError(
                "reminder=True needs a `due` datetime to anchor the reminder on — "
                "Graph ignores isReminderOn without reminderDateTime"
            )
        task_body.is_reminder_on = reminder
        if reminder:
            task_body.reminder_date_time = _datetime_timezone(due)

    if recurrence:
        task_body.recurrence = _build_recurrence(recurrence)

    response = await graph_client.me.todo.lists.by_todo_task_list_id(resolved_id).tasks.post(
        task_body
    )

    if response is None or response.id is None:
        # Optional[TodoTask] on the SDK side; an empty 201/204 used to raise
        # AttributeError here. The task exists server-side at this point, so
        # say that rather than let an agent's retry create a duplicate.
        raise ValueError(
            "Task was created but Graph returned no id for it — do not retry "
            "blindly (that would create a duplicate); find it with "
            "outlook_list_tasks instead"
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
    *,
    config: Config,
) -> dict:
    """Update a task in a To Do list.

    PATCH /me/todo/lists/{id}/tasks/{taskId}
    Only patches provided fields.
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_update_task")
    task_id = validate_graph_id(task_id)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.todo_task import TodoTask

    patch_body = TodoTask()

    if title is not None:
        patch_body.title = title

    if due is not None:
        validate_datetime(due)
        patch_body.due_date_time = _datetime_timezone(due)

    if body is not None:
        patch_body.body = _text_body(body)

    if importance is not None:
        patch_body.importance = _importance_enum(importance)

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
    *,
    config: Config,
) -> dict:
    """Mark a task as completed.

    PATCH with status="completed" and completedDateTime set to now (UTC).
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_complete_task")
    task_id = validate_graph_id(task_id)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.0000000Z")

    from msgraph.generated.models.task_status import TaskStatus
    from msgraph.generated.models.todo_task import TodoTask

    patch_body = TodoTask()
    patch_body.status = TaskStatus.Completed
    patch_body.completed_date_time = _datetime_timezone(now_utc)

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
    *,
    config: Config,
) -> dict:
    """Delete a task from a To Do list.

    DELETE /me/todo/lists/{id}/tasks/{taskId}
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_delete_task")
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


def _validated_display_name(display_name: str) -> str:
    """Reject an empty/whitespace checklist label before it reaches Graph."""
    name = (display_name or "").strip()
    if not name:
        raise ValueError("display_name must be a non-empty string")
    return name


async def add_checklist_item(
    graph_client: Any,
    task_id: str,
    display_name: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Add a checklist item (sub-step) to a task.

    POST /me/todo/lists/{id}/tasks/{taskId}/checklistItems
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_add_checklist_item")
    task_id = validate_graph_id(task_id)
    name = _validated_display_name(display_name)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.checklist_item import ChecklistItem

    item = ChecklistItem()
    item.display_name = name

    response = await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .checklist_items.post(item)
    )

    if response is None or response.id is None:
        # Optional[ChecklistItem] on the SDK side; an empty 201/204 used to
        # raise AttributeError *after* the sub-step was created, and the
        # agent's natural retry doubled it.
        raise ValueError(
            "Checklist item was added but Graph returned no id for it — do "
            "not retry blindly (that would create a duplicate); read the "
            "task back with outlook_get_task instead"
        )

    return {
        "status": "added",
        "task_id": task_id,
        "checklist_item_id": response.id,
        "display_name": sanitize_output(response.display_name or ""),
    }


async def update_checklist_item(
    graph_client: Any,
    task_id: str,
    checklist_item_id: str,
    display_name: str | None = None,
    is_checked: bool | None = None,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Update a checklist item (partial patch — only provided fields change).

    PATCH /me/todo/lists/{id}/tasks/{taskId}/checklistItems/{checklistItemId}
    checkedDateTime is maintained server-side from isChecked; sending it
    ourselves would race the server's own bookkeeping.
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_update_checklist_item")
    task_id = validate_graph_id(task_id)
    checklist_item_id = validate_graph_id(checklist_item_id)

    if display_name is None and is_checked is None:
        raise ValueError("Provide at least one of display_name or is_checked")

    # Validate every caller input before the network round trip, matching
    # add_checklist_item — a blank rename is a caller fix, not a Graph call.
    checked_name: str | None = None
    if display_name is not None:
        checked_name = _validated_display_name(display_name)

    resolved_id = await _resolve_list_id(graph_client, list_id)

    from msgraph.generated.models.checklist_item import ChecklistItem

    patch_body = ChecklistItem()
    if checked_name is not None:
        patch_body.display_name = checked_name
    if is_checked is not None:
        patch_body.is_checked = is_checked

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .checklist_items.by_checklist_item_id(checklist_item_id)
        .patch(patch_body)
    )

    return {
        "status": "updated",
        "task_id": task_id,
        "checklist_item_id": checklist_item_id,
    }


async def delete_checklist_item(
    graph_client: Any,
    task_id: str,
    checklist_item_id: str,
    list_id: str | None = None,
    *,
    config: Config,
) -> dict:
    """Delete a checklist item from a task.

    DELETE /me/todo/lists/{id}/tasks/{taskId}/checklistItems/{checklistItemId}
    """
    check_permission(config, CATEGORY_TODO_WRITE, "outlook_delete_checklist_item")
    task_id = validate_graph_id(task_id)
    checklist_item_id = validate_graph_id(checklist_item_id)
    resolved_id = await _resolve_list_id(graph_client, list_id)

    await (
        graph_client.me.todo.lists.by_todo_task_list_id(resolved_id)
        .tasks.by_todo_task_id(task_id)
        .checklist_items.by_checklist_item_id(checklist_item_id)
        .delete()
    )

    return {
        "status": "deleted",
        "task_id": task_id,
        "checklist_item_id": checklist_item_id,
    }
