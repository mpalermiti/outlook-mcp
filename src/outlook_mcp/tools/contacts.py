"""Contact tools: list, search, get, create, update, delete."""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import apply_pagination, wrap_nextlink
from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_email,
    validate_graph_id,
    validate_phone,
)


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _check_read_only(read_only: bool, tool_name: str) -> None:
    """Raise ReadOnlyError if server is in read-only mode."""
    if read_only:
        raise ReadOnlyError(tool_name)


def _format_contact_summary(contact: Any) -> dict:
    """Convert Graph SDK contact to summary dict."""
    email = ""
    if contact.email_addresses:
        first_email = contact.email_addresses[0]
        email = getattr(first_email, "address", "") or ""

    phone = ""
    if contact.phones:
        first_phone = contact.phones[0]
        phone = getattr(first_phone, "number", "") or ""

    return {
        "id": contact.id,
        "display_name": sanitize_output(contact.display_name or ""),
        "email": email,
        "phone": phone,
        "company": sanitize_output(contact.company_name or ""),
    }


def _format_contact_detail(contact: Any) -> dict:
    """Convert Graph SDK contact to full detail dict."""
    email_addresses = []
    for e in contact.email_addresses or []:
        email_addresses.append(
            {
                "address": getattr(e, "address", "") or "",
                "name": sanitize_output(getattr(e, "name", "") or ""),
            }
        )

    phones = []
    for p in contact.phones or []:
        phones.append(
            {
                "number": getattr(p, "number", "") or "",
                "type": getattr(p, "type", "") or "",
            }
        )

    return {
        "id": contact.id,
        "first_name": sanitize_output(contact.given_name or ""),
        "last_name": sanitize_output(contact.surname or ""),
        "display_name": sanitize_output(contact.display_name or ""),
        "email_addresses": email_addresses,
        "phones": phones,
        "company": sanitize_output(contact.company_name or ""),
        "title": sanitize_output(contact.title or ""),
        "department": sanitize_output(getattr(contact, "department", "") or ""),
        "birthday": str(contact.birthday) if contact.birthday else None,
    }


async def list_contacts(
    graph_client: Any,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List contacts with pagination."""
    query_params: dict[str, Any] = {
        "$orderby": "displayName",
        "$select": ("id,displayName,givenName,surname,emailAddresses,phones,companyName,title"),
    }
    query_params = apply_pagination(query_params, count, cursor)

    response = await graph_client.me.contacts.get(query_params=query_params)

    contacts = [_format_contact_summary(c) for c in (response.value or [])]
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "contacts": contacts,
        "count": len(contacts),
        "has_more": next_cursor is not None,
        "cursor": next_cursor,
    }


async def search_contacts(
    graph_client: Any,
    query: str,
    count: int = 25,
) -> dict:
    """Search contacts using KQL."""
    count = _clamp(count, 1, 100)
    safe_query = sanitize_kql(query)

    query_params: dict[str, Any] = {
        "$top": count,
        "$search": safe_query,
        "$select": ("id,displayName,givenName,surname,emailAddresses,phones,companyName,title"),
    }

    response = await graph_client.me.contacts.get(query_params=query_params)

    contacts = [_format_contact_summary(c) for c in (response.value or [])]

    return {
        "contacts": contacts,
        "count": len(contacts),
        "has_more": response.odata_next_link is not None,
        "cursor": wrap_nextlink(response.odata_next_link),
    }


async def get_contact(
    graph_client: Any,
    contact_id: str,
) -> dict:
    """Get a single contact by ID."""
    contact_id = validate_graph_id(contact_id)

    contact = await graph_client.me.contacts.by_contact_id(contact_id).get()

    return _format_contact_detail(contact)


async def create_contact(
    graph_client: Any,
    first_name: str,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    company: str | None = None,
    title: str | None = None,
    read_only: bool = False,
) -> dict:
    """Create a new contact."""
    _check_read_only(read_only, "outlook_create_contact")

    if email:
        validate_email(email)
    if phone:
        validate_phone(phone)

    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.phone import Phone

    new_contact = Contact()
    new_contact.given_name = first_name
    if last_name:
        new_contact.surname = last_name
    if email:
        ea = EmailAddress()
        ea.address = email
        ea.name = first_name
        new_contact.email_addresses = [ea]
    if phone:
        p = Phone()
        p.number = phone
        new_contact.phones = [p]
    if company:
        new_contact.company_name = company
    if title:
        new_contact.title = title

    created = await graph_client.me.contacts.post(new_contact)

    return {
        "status": "created",
        "id": created.id,
    }


async def update_contact(
    graph_client: Any,
    contact_id: str,
    first_name: str | None = None,
    last_name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    read_only: bool = False,
) -> dict:
    """Update an existing contact (partial patch)."""
    _check_read_only(read_only, "outlook_update_contact")
    contact_id = validate_graph_id(contact_id)

    if email:
        validate_email(email)
    if phone:
        validate_phone(phone)

    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.phone import Phone

    patch_body = Contact()
    if first_name is not None:
        patch_body.given_name = first_name
    if last_name is not None:
        patch_body.surname = last_name
    if email is not None:
        ea = EmailAddress()
        ea.address = email
        patch_body.email_addresses = [ea]
    if phone is not None:
        p = Phone()
        p.number = phone
        patch_body.phones = [p]

    await graph_client.me.contacts.by_contact_id(contact_id).patch(patch_body)

    return {"status": "updated"}


async def delete_contact(
    graph_client: Any,
    contact_id: str,
    read_only: bool = False,
) -> dict:
    """Delete a contact by ID."""
    _check_read_only(read_only, "outlook_delete_contact")
    contact_id = validate_graph_id(contact_id)

    await graph_client.me.contacts.by_contact_id(contact_id).delete()

    return {"status": "deleted"}
