"""Shared models."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel


class ListResponse(BaseModel):
    """Standard list response wrapper."""

    items: list[Any]
    count: int
    has_more: bool = False
    cursor: str | None = None
