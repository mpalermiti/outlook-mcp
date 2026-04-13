"""To Do task-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, field_validator


class TaskListSummary(BaseModel):
    """To Do list summary."""

    id: str
    display_name: str
    is_default: bool = False


class TaskSummary(BaseModel):
    """Compact task representation for list results."""

    id: str
    title: str
    status: str = "notStarted"
    due: str | None = None
    importance: str = "normal"
    is_reminder_on: bool = False
    created: str = ""


class TaskDetail(BaseModel):
    """Full task representation."""

    id: str
    title: str
    status: str = "notStarted"
    due: str | None = None
    importance: str = "normal"
    body: str = ""
    is_reminder_on: bool = False
    reminder_date_time: str | None = None
    created: str = ""
    completed: str | None = None
    recurrence: dict | None = None


class CreateTaskInput(BaseModel):
    """Input for creating a To Do task."""

    title: str
    list_id: str | None = None
    due: str | None = None
    importance: str = "normal"
    body: str | None = None
    reminder: str | None = None
    recurrence: str | None = None

    @field_validator("importance")
    @classmethod
    def validate_importance(cls, v: str) -> str:
        if v not in ("low", "normal", "high"):
            raise ValueError(f"importance must be low, normal, or high; got {v}")
        return v

    @field_validator("recurrence")
    @classmethod
    def validate_recurrence(cls, v: str | None) -> str | None:
        if v is not None and v not in (
            "daily",
            "weekdays",
            "weekly",
            "monthly",
            "yearly",
        ):
            raise ValueError(
                f"recurrence must be daily/weekdays/weekly/monthly/yearly; got {v}"
            )
        return v


class UpdateTaskInput(BaseModel):
    """Input for updating a task."""

    task_id: str
    list_id: str | None = None
    title: str | None = None
    due: str | None = None
    body: str | None = None
    importance: str | None = None
