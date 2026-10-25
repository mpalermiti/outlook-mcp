"""Cursor-based pagination helpers for Tier 2."""

from __future__ import annotations

import base64
import json
from typing import Any
from urllib.parse import parse_qs, quote, urlparse

from kiota_abstractions.base_request_configuration import RequestConfiguration


def encode_cursor_payload(data: dict[str, Any]) -> str:
    """Encode a cursor payload dict as an opaque base64 cursor.

    Paging owns ``skip`` / ``skiptoken``; a tool may add its own keys to pin
    the scope a listing was made in (``calendar_read`` stores the calendar), so
    a follow-up page cannot drift to a different scope.
    """
    return base64.urlsafe_b64encode(json.dumps(data).encode()).decode()


def decode_cursor_payload(cursor: str) -> dict[str, Any]:
    """Decode an opaque cursor back to its payload dict.

    Raises ValueError("Invalid pagination cursor") for anything that is not a
    base64-encoded JSON object — including base64's own padding errors, which
    are ValueErrors with a message ("Incorrect padding") that tells an agent
    nothing about which argument was wrong.
    """
    try:
        data = json.loads(base64.urlsafe_b64decode(cursor.encode()).decode())
    except (ValueError, UnicodeDecodeError) as e:  # binascii.Error is a ValueError
        raise ValueError("Invalid pagination cursor") from e
    if not isinstance(data, dict):
        raise ValueError("Invalid pagination cursor")
    return data


def encode_cursor(skip: int) -> str:
    """Encode a skip value into an opaque base64 cursor."""
    return encode_cursor_payload({"skip": skip})


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
        data = decode_cursor_payload(cursor)
        if "skip" in data:
            query_params["$skip"] = data["skip"]
        elif "skiptoken" in data:
            token = data["skiptoken"]
            # Caller-held state that goes into a URL: anything but the
            # non-empty string Graph issued is refused here, before any request.
            if not isinstance(token, str) or not token:
                raise ValueError("Invalid pagination cursor")
            query_params["$skiptoken"] = token

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


async def get_page(builder: Any, query_params_class: type, params: dict[str, Any]) -> Any:
    """GET one page from ``builder``, including a ``$skiptoken`` from a cursor.

    ``build_request_config`` cannot carry ``$skiptoken``: no typed ``users/*``
    query-parameters class has the field, and their URL templates stop at
    ``$top``, so the token is dropped and a cursor from a listing Graph pages
    by ``$skiptoken`` fetches page one again (a message ``$search`` pages no
    other way). Graph's ``@odata.nextLink`` for such a page is the request it
    answered plus ``$skiptoken``, so that is what is sent: the URL the builder
    makes from the typed parameters, with the token appended as one
    percent-encoded query value. The token can change nothing else about the
    request: not the host, the path or another parameter. ``params`` comes
    from ``apply_pagination``, which has already refused a token that is not a
    non-empty string.
    """
    params = dict(params)
    skiptoken = params.pop("$skiptoken", None)
    config = build_request_config(query_params_class, params)
    if skiptoken is None:
        return await builder.get(request_configuration=config)
    info = builder.to_get_request_information(config)
    info.path_parameters["baseurl"] = builder.request_adapter.base_url
    url = info.url
    separator = "&" if "?" in url else "?"
    return await builder.with_url(f"{url}{separator}%24skiptoken={quote(skiptoken, safe='')}").get()
