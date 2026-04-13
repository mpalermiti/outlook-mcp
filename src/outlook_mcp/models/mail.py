"""Mail-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator


class MessageSummary(BaseModel):
    """Compact message representation for list results."""

    id: str
    subject: str
    from_email: str
    from_name: str = ""
    received: str
    is_read: bool
    importance: str = "normal"
    preview: str = ""
    has_attachments: bool = False
    categories: list[str] = Field(default_factory=list)
    flag: str = "notFlagged"
    conversation_id: str = ""


class MessageDetail(BaseModel):
    """Full message representation."""

    id: str
    subject: str
    from_email: str
    from_name: str = ""
    to: list[dict[str, str]] = Field(default_factory=list)
    cc: list[dict[str, str]] = Field(default_factory=list)
    received: str
    body: str = ""
    body_html: str | None = None
    is_read: bool
    importance: str = "normal"
    has_attachments: bool = False
    attachments: list[dict[str, str]] = Field(default_factory=list)
    categories: list[str] = Field(default_factory=list)
    flag: str = "notFlagged"
    conversation_id: str = ""


class SendMessageInput(BaseModel):
    """Input for sending a message."""

    to: list[str] = Field(min_length=1)
    subject: str
    body: str
    cc: list[str] | None = None
    bcc: list[str] | None = None
    is_html: bool = False
    importance: str = "normal"
    sensitivity: str = "normal"
    request_read_receipt: bool = False

    @field_validator("importance")
    @classmethod
    def validate_importance(cls, v: str) -> str:
        if v not in ("low", "normal", "high"):
            raise ValueError(f"importance must be low, normal, or high; got {v}")
        return v

    @field_validator("sensitivity")
    @classmethod
    def validate_sensitivity(cls, v: str) -> str:
        if v not in ("normal", "personal", "private", "confidential"):
            raise ValueError(f"sensitivity must be normal/personal/private/confidential; got {v}")
        return v


class ReplyInput(BaseModel):
    """Input for replying to a message."""

    message_id: str
    body: str
    reply_all: bool = False
    is_html: bool = False


class ForwardInput(BaseModel):
    """Input for forwarding a message."""

    message_id: str
    to: list[str] = Field(min_length=1)
    comment: str | None = None


class TriageInput(BaseModel):
    """Input for triage actions (move, flag, categorize, mark read)."""

    message_id: str
    action: str
    value: str


class DeleteInput(BaseModel):
    """Input for deleting a message (soft or hard)."""

    message_id: str
    permanent: bool = False
