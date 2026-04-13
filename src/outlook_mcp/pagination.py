"""Cursor-based pagination helpers for Tier 2."""

from __future__ import annotations

import base64
import json
from typing import Any
from urllib.parse import parse_qs, urlparse

from kiota_abstractions.base_request_configuration import RequestConfiguration


def encode_cursor(skip: int) -> str:
    """Encode a skip value into an opaque base64 cursor."""
    payload = json.dumps({"skip": skip})
    return base64.urlsafe_b64encode(payload.encode()).decode()


def decode_cursor(cursor: str) -> int:
    """Decode a cursor back to a skip value.

    Raises ValueError if cursor is invalid or tampered.
    """
    try:
        payload = base64.urlsafe_b64decode(cursor.encode()).decode()
        data = json.loads(payload)
        skip = data["skip"]
        if not isinstance(skip, int) or skip < 0:
            raise ValueError("Invalid skip value in cursor")
        return skip
    except (KeyError, json.JSONDecodeError, UnicodeDecodeError) as e:
        raise ValueError(f"Invalid pagination cursor: {cursor}") from e


def wrap_nextlink(odata_next_link: str | None) -> str | None:
    """Extract skip from an @odata.nextLink URL and encode as cursor.

    Returns None if there is no next link.
    """
    if not odata_next_link:
        return None

    # Graph API nextLink contains $skip parameter
    parsed = urlparse(odata_next_link)
    params = parse_qs(parsed.query)

    # Try $skip first
    skip_values = params.get("$skip", [])
    if skip_values:
        try:
            return encode_cursor(int(skip_values[0]))
        except (ValueError, IndexError):
            pass

    # Fallback: some Graph endpoints use $skiptoken instead of $skip
    skiptoken = params.get("$skiptoken", [])
    if skiptoken:
        payload = json.dumps({"skiptoken": skiptoken[0]})
        return base64.urlsafe_b64encode(payload.encode()).decode()

    return None


def apply_pagination(
    query_params: dict,
    count: int,
    cursor: str | None = None,
) -> dict:
    """Add $top and $skip to query params from cursor.

    Args:
        query_params: Existing OData query parameters.
        count: Number of items to request (clamped to 1-100).
        cursor: Opaque cursor from previous response, or None for first page.

    Returns:
        Updated query_params dict with $top and optionally $skip.
    """
    count = max(1, min(100, count))
    query_params["$top"] = count

    if cursor:
        try:
            payload = base64.urlsafe_b64decode(cursor.encode()).decode()
            data = json.loads(payload)
            if "skip" in data:
                query_params["$skip"] = data["skip"]
            elif "skiptoken" in data:
                query_params["$skiptoken"] = data["skiptoken"]
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            raise ValueError("Invalid pagination cursor") from e

    return query_params


# OData key → QueryParameters attribute name
_PARAM_MAP = {
    "$top": "top",
    "$skip": "skip",
    "$filter": "filter",
    "$search": "search",
    "$orderby": "orderby",
    "$select": "select",
    "$skiptoken": "skiptoken",
    "$expand": "expand",
    "start_date_time": "start_date_time",
    "end_date_time": "end_date_time",
}

# Attributes that take list[str] rather than a scalar
_LIST_ATTRS = {"orderby", "select", "expand"}


def build_request_config(
    query_params_class: type,
    params: dict[str, Any],
) -> RequestConfiguration:
    """Build a typed RequestConfiguration from an OData params dict.

    Converts the dict-based query_params used throughout the codebase
    into the typed RequestConfiguration that msgraph-sdk v1.55+ requires.
    """
    qp = query_params_class()

    for key, value in params.items():
        attr = _PARAM_MAP.get(key, key)
        if not hasattr(qp, attr):
            continue
        # Convert comma-separated strings to lists for list-typed attrs
        if attr in _LIST_ATTRS and isinstance(value, str):
            value = [v.strip() for v in value.split(",")]
        setattr(qp, attr, value)

    return RequestConfiguration(query_parameters=qp)
