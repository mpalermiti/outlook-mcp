"""Contact-related Pydantic models."""

from __future__ import annotations

from pydantic import BaseModel, Field


class ContactSummary(BaseModel):
    """Compact contact representation for list results."""

    id: str
    display_name: str = ""
    email: str = ""
    phone: str = ""
    company: str = ""


class ContactDetail(BaseModel):
    """Full contact representation."""

    id: str
    first_name: str = ""
    last_name: str = ""
    display_name: str = ""
    email_addresses: list[dict[str, str]] = Field(default_factory=list)
    phones: list[dict[str, str]] = Field(default_factory=list)
    company: str = ""
    title: str = ""
    department: str = ""
    birthday: str | None = None


class CreateContactInput(BaseModel):
    """Input for creating a contact."""

    first_name: str
    last_name: str | None = None
    email: str | None = None
    phone: str | None = None
    company: str | None = None
    title: str | None = None


class UpdateContactInput(BaseModel):
    """Input for updating a contact."""

    contact_id: str
    first_name: str | None = None
    last_name: str | None = None
    email: str | None = None
    phone: str | None = None
    company: str | None = None
    title: str | None = None
