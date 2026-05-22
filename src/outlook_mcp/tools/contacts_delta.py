"""Contacts delta tool: ``list_contacts_delta``.

Wraps ``GET /me/contacts/delta``.

Graph constraints worth knowing before reading the code:

- This endpoint accepts **no query parameters at all** on the first
  call — confirmed by a curl 400 against the live API. Page sizing is
  via the ``Prefer: odata.maxpagesize=N`` header only.
- Subsequent calls use the returned deltaLink / nextLink URL verbatim
  (Graph's encoded URL is fine; we don't try to strip or augment query
  params on it).

To keep the surface predictable we just don't take any first-call query
arguments here. There's nothing to strip and no client-side validation
needed beyond ``page_size``.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import urljoin

from outlook_mcp.tools._delta import GRAPH_BASE, fetch_delta_pages, format_delta_item
from outlook_mcp.validation import sanitize_output


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _primary_phone_from_dict(raw: dict) -> str:
    """Pick the most representative phone from a raw Graph contact dict.

    Mirrors ``contacts._primary_phone`` — consumer Outlook accounts
    expose ``mobilePhone`` (single), ``homePhones`` (list), and
    ``businessPhones`` (list). Prefer mobile, then first home, then
    first business.
    """
    mobile = raw.get("mobilePhone") or ""
    if mobile:
        return mobile
    for key in ("homePhones", "businessPhones"):
        values = raw.get(key) or []
        if values:
            return values[0]
    return ""


def _format_contact_delta(raw: dict) -> dict:
    """Build the wire-shape contact summary from a raw Graph JSON contact.

    Mirrors ``contacts._format_contact_summary`` field-for-field so
    callers don't have to branch on whether a given dict came from a
    delta call or a regular list call.
    """
    email = ""
    emails = raw.get("emailAddresses") or []
    if emails:
        email = emails[0].get("address") or ""

    return {
        "id": raw.get("id"),
        "display_name": sanitize_output(raw.get("displayName") or ""),
        "email": email,
        "phone": _primary_phone_from_dict(raw),
        "company": sanitize_output(raw.get("companyName") or ""),
    }


async def list_contacts_delta(
    graph_client: Any,
    page_size: int = 50,
    delta_token: str | None = None,
) -> dict:
    """List contact changes since the last delta token.

    Args:
        graph_client: A ``GraphClient`` (delta tools need the
            ``credential`` attribute to mint raw bearer tokens).
        page_size: Items per Graph page (1-100). Mapped to
            ``Prefer: odata.maxpagesize=N`` — Graph rejects every other
            query parameter on this endpoint.
        delta_token: Opaque cursor from a previous call.

    Returns:
        ``{contacts, delta_token, has_more}`` — see ``mail_delta`` for
        full semantics. Tombstones come back as
        ``{id, is_deleted: True}``.
    """
    page_size = _clamp(page_size, 1, 100)
    credential = graph_client.credential

    initial_url = ""
    if not delta_token:
        # /me/contacts/delta accepts no query params — Graph 400s if you
        # even pass a $select. The page size header below is the only
        # knob we get.
        initial_url = urljoin(GRAPH_BASE, "me/contacts/delta")

    raw_items, next_token, has_more = await fetch_delta_pages(
        credential,
        initial_url=initial_url,
        delta_token=delta_token,
        page_size=page_size,
        headers={"Prefer": f"odata.maxpagesize={page_size}"},
    )

    contacts = [format_delta_item(item, _format_contact_delta) for item in raw_items]

    return {
        "contacts": contacts,
        "delta_token": next_token,
        "has_more": has_more,
    }
