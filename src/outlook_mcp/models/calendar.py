"""Calendar-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator


class EventSummary(BaseModel):
    """Compact event representation for list results."""

    id: str
    subject: str
    start: str
    end: str
    location: str = ""
    is_all_day: bool = False
    organizer: str = ""
    response_status: str = ""
    is_online: bool = False


class EventDetail(BaseModel):
    """Full event representation."""

    id: str
    subject: str
    start: str
    end: str
    location: str = ""
    body: str = ""
    is_all_day: bool = False
    organizer: dict[str, str] = Field(default_factory=dict)
    attendees: list[dict[str, str]] = Field(default_factory=list)
    response_status: str = ""
    is_online: bool = False
    online_meeting_url: str | None = None
    recurrence: dict | None = None
    type: str = ""
    categories: list[str] = Field(default_factory=list)


class CreateEventInput(BaseModel):
    """Input for creating a calendar event."""

    subject: str
    start: str
    end: str
    location: str | None = None
    body: str | None = None
    attendees: list[str] | None = None
    is_all_day: bool = False
    is_online: bool = False
    recurrence: dict | str | None = None

    @field_validator("recurrence")
    @classmethod
    def validate_recurrence(cls, v: dict | str | None) -> dict | str | None:
        """Accept a Graph recurrence object, a JSON string of one, or a shorthand."""
        if v is None or isinstance(v, dict):
            return v
        if v.strip().startswith(("{", "[")):
            return v
        if v not in ("daily", "weekdays", "weekly", "monthly", "yearly"):
            raise ValueError(f"recurrence must be daily/weekdays/weekly/monthly/yearly; got {v}")
        return v


class UpdateEventInput(BaseModel):
    """Input for updating an event."""

    event_id: str
    subject: str | None = None
    start: str | None = None
    end: str | None = None
    location: str | None = None
    body: str | None = None


class RsvpInput(BaseModel):
    """Input for RSVPing to an event."""

    event_id: str
    response: str
    message: str | None = None

    @field_validator("response")
    @classmethod
    def validate_response(cls, v: str) -> str:
        if v not in ("accept", "decline", "tentative"):
            raise ValueError(f"response must be accept/decline/tentative; got {v}")
        return v
