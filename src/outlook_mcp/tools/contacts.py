"""Contact tools: list, search, get, create, update, delete."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.pagination import apply_pagination, build_request_config, wrap_nextlink
from outlook_mcp.permissions import CATEGORY_CONTACTS_WRITE, check_permission
from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_birthday,
    validate_email,
    validate_graph_id,
    validate_phone,
)


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


# Every field _format_contact_summary reads, and nothing it does not. A $select
# narrower than the formatter silently blanks whatever it leaves out, and a
# wider one pays for bytes nobody reads. (givenName, surname and title were the
# latter: selected since forever, never read by the summary shape.)
_SUMMARY_SELECT = "id,displayName,emailAddresses,mobilePhone,homePhones,businessPhones,companyName"

# Only the listing asks for categories. Graph's $search on contacts does not
# return them — verified against live Graph, empty under a narrow $select, a
# wide one, and no $select at all — so a search result carrying the key would
# report every contact as uncategorised, which is exactly the "empty means
# absent" confusion this module is trying to stop telling. The listing gets the
# field; the search omits the key rather than lying about it.
_LIST_SELECT = _SUMMARY_SELECT + ",categories"

# get_contact sends no $select at all, so Graph returns the whole contact and
# the detail formatter can read whatever it likes: the data was never the
# problem there, the formatter simply did not read these five.
#
# One tuple, read and write. The read path returns these keys and the write path
# accepts them, so handing an address straight back — the documented way to keep
# the parts you are not changing — is mechanical rather than a renaming exercise.
_ADDRESS_FIELDS = ("street", "city", "state", "postal_code", "country_or_region")


def _format_address(address: Any) -> dict | None:
    """Flatten a Graph ``physicalAddress``, or ``None`` when it holds nothing.

    Graph hands back an empty address object rather than null for addresses the
    contact does not have, so an emptiness check is what distinguishes "no home
    address" from "a home address whose street we failed to read". Whitespace is
    not content — :func:`_build_address` refuses to write a blank part, so an
    address of spaces reported here as real would be one this tool tells you to
    hand back and then rejects.
    """
    parts = {
        field: sanitize_output(getattr(address, field, "") or "", multiline=True)
        for field in _ADDRESS_FIELDS
    }
    return parts if any(value.strip() for value in parts.values()) else None


def _build_address(slot: str, parts: Any) -> Any:
    """A Graph ``physicalAddress`` from the shape :func:`_format_address` returns.

    The mirror of the read path, field for field: a write path that cannot
    express what the read path returns is how an address ends up readable but
    not correctable.

    An unknown key is an error rather than a silent drop. The caller is a model
    reading a docstring, and a misspelt part that patched nothing would be the
    #41 shape again — a call that succeeds and does nothing. Raised as a plain
    ``ValueError``, like every other tool module: ``_wrap_tool_errors`` turns it
    into ``ToolInputError`` so the text reaches the model as an anticipated
    failure rather than being withheld as a crash.

    A part that was not supplied is **left unset**, never assigned ``None``.
    Graph's request adapter serializes through the backing store, which emits a
    field that was explicitly set to ``None`` — and for a nested model it emits
    it onto the *parent*, under its Python name. A city-only address therefore
    went out as ``{"country_or_region": null, ..., "homeAddress": {"city": …}}``
    and Graph answered ``400 The property 'country_or_region' does not exist on
    type 'microsoft.graph.contact'``. Assigning only what we were given is the
    whole fix; ``TestPartialAddressDoesNotLeakNulls`` pins it.
    """
    from msgraph.generated.models.physical_address import PhysicalAddress

    if not isinstance(parts, dict):
        raise ValueError(
            f"{slot}_address takes an object like "
            f"{{'street': '…', 'city': '…'}}, not {type(parts).__name__}. "
            f"Valid parts: {', '.join(_ADDRESS_FIELDS)}."
        )
    unknown = [key for key in parts if key not in _ADDRESS_FIELDS]
    if unknown:
        raise ValueError(
            f"{slot}_address has unknown part(s): {', '.join(sorted(unknown))}. "
            f"Valid parts: {', '.join(_ADDRESS_FIELDS)} — the same keys "
            f"outlook_get_contact returns."
        )

    present = {}
    for field, value in parts.items():
        if value is None:
            continue
        if not isinstance(value, str):
            raise ValueError(
                f"{slot}_address['{field}'] must be a string, not {type(value).__name__}."
            )
        if value.strip():
            present[field] = value.strip()

    if not present:
        # Reporting "updated" for a patch that carried nothing is the failure
        # this module exists to stop telling. Clearing an address is a separate
        # capability nobody has asked for yet; until then, say so.
        raise ValueError(
            f"{slot}_address has no content. Omit it to leave the stored address "
            f"unchanged — this tool cannot clear an address."
        )

    # Assigned field by field rather than through a `setattr` loop: an attribute
    # that is not a real SDK field is a silent no-op (#41's shape), and the
    # static guard that catches those — test_sdk_fields_exist — reads
    # `<var>.<attr> = ...` and cannot see `setattr`.
    address = PhysicalAddress()
    if "street" in present:
        address.street = present["street"]
    if "city" in present:
        address.city = present["city"]
    if "state" in present:
        address.state = present["state"]
    if "postal_code" in present:
        address.postal_code = present["postal_code"]
    if "country_or_region" in present:
        address.country_or_region = present["country_or_region"]
    return address


_PHONE_TYPES = ("mobile", "business", "home")


def _check_phone_type(phone: str | None, phone_type: str) -> None:
    """Refuse a type that would quietly file the number somewhere else, or nowhere.

    An unknown type falling back to mobile is #120 again with extra steps, and a
    type with no number is a patch that carries nothing and reports "updated".
    """
    if phone_type not in _PHONE_TYPES:
        raise ValueError(
            f"Unknown phone_type {phone_type[:20]!r}. Use one of: {', '.join(_PHONE_TYPES)}."
        )
    if not phone and phone_type != "mobile":
        raise ValueError("phone_type needs phone: it says which kind of number phone is.")


def _assign_phone(contact: Any, phone: str, phone_type: str) -> None:
    """File ``phone`` under the Graph field its type names (#120).

    Consumer contacts have one ``mobilePhone`` but *lists* of business and home
    numbers, so a business or home write replaces that list with ``[phone]`` —
    the same replace-not-merge contract the addresses have. The other two fields
    are left unset, never assigned ``None``: the backing store would emit them.
    """
    if phone_type == "business":
        contact.business_phones = [phone]
    elif phone_type == "home":
        contact.home_phones = [phone]
    else:
        contact.mobile_phone = phone


def _check_clear_phones(clear_phones: list[str], phone: str | None, phone_type: str) -> None:
    """Refuse a clear that would do nothing, or undo the number being written."""
    unknown = [t for t in clear_phones if t not in _PHONE_TYPES]
    if unknown:
        raise ValueError(
            f"Unknown clear_phones type {unknown[0][:20]!r}. Use any of: {', '.join(_PHONE_TYPES)}."
        )
    if phone is not None and phone_type in clear_phones:
        raise ValueError(
            f"clear_phones names {phone_type!r}, the type phone is being written to. "
            f"Clear the other types; writing {phone_type!r} already replaces it."
        )


def _clear_phones(contact: Any, clear_phones: list[str]) -> None:
    """Empty each named phone field — the other half of moving a number (#120).

    The lists clear as ``[]``, which kiota serializes as written. ``mobilePhone``
    needs an explicit JSON null, and ``contact.mobile_phone = None`` does not
    produce one: the backing store emits it under its Python name. So it goes
    through ``additional_data``, as ``update_event``'s ``remove_recurrence`` does.
    """
    if "mobile" in clear_phones:
        contact.additional_data = {**(contact.additional_data or {}), "mobilePhone": None}
    if "business" in clear_phones:
        contact.business_phones = []
    if "home" in clear_phones:
        contact.home_phones = []


# Exchange keeps a contact's addresses in three fixed slots (EmailAddress1 to
# EmailAddress3 in MS-OXWSCONT), so a merge that comes to more is refused here
# rather than sent to a store that cannot hold it. Whether Graph would refuse a
# fourth or drop it is not known offline; the live_write tier asks.
_MAX_EMAILS = 3


def _email_list(name: str, values: list[str] | None) -> list[str] | None:
    """Validated, de-duplicated addresses from ``add_emails`` / ``remove_emails``.

    Compared without case, as mail addresses are in practice; the first spelling
    given is the one kept. An empty list is refused rather than read as "nothing
    to do": a patch that carries nothing and answers "updated" is #41's shape.
    """
    if values is None:
        return None
    if not values:
        raise ValueError(f"{name} is empty. Omit it to leave the addresses unchanged.")
    seen: dict[str, str] = {}
    for value in values:
        validate_email(value)
        seen.setdefault(value.lower(), value)
    return list(seen.values())


def _merged_emails(current: Any, adding: list[str], removing: list[str]) -> list[Any]:
    """The contact's addresses with ``removing`` taken out and ``adding`` appended.

    Kept addresses go back as they were read, ``name`` included, in their stored
    order; an added one goes on the end with no name, as ``email`` writes it. The
    entries are new models built field by field, not the ones the read returned,
    so what is sent is exactly what is listed here.
    """
    from msgraph.generated.models.email_address import EmailAddress

    stored = [
        (address, getattr(entry, "name", None))
        for entry in (getattr(current, "email_addresses", None) or [])
        if (address := getattr(entry, "address", None))
    ]
    held = {address.lower() for address, _ in stored}
    absent = [address for address in removing if address.lower() not in held]
    if absent:
        raise ValueError(
            f"The contact has no address {', '.join(absent)}, so nothing was changed. "
            f"Its addresses are: {', '.join(a for a, _ in stored) or 'none'}."
        )

    dropped = {address.lower() for address in removing}
    kept = [(address, name) for address, name in stored if address.lower() not in dropped]
    present = {address.lower() for address, _ in kept}
    merged = kept + [(address, None) for address in adding if address.lower() not in present]
    if len(merged) > _MAX_EMAILS:
        raise ValueError(
            f"A contact holds at most three email addresses, and this would leave "
            f"{len(merged)}. Nothing was changed. Remove one in the same call to make room."
        )

    entries = []
    for address, name in merged:
        entry = EmailAddress()
        entry.address = address
        if name:
            entry.name = name
        entries.append(entry)
    return entries


def _primary_phone(contact: Any) -> str:
    """Pick the most representative phone from the consumer-Graph contact fields.

    Consumer Outlook accounts expose mobilePhone (single), homePhones (list),
    and businessPhones (list) — not the unified ``phones`` collection. Prefer
    mobile, then first home, then first business.

    Spelled out rather than looped over a tuple of names: a ``getattr`` whose
    attribute is a variable is invisible to `test_select_covers_the_formatter`,
    which reads this source to derive what the `$select` must ask for. A loop
    here would take `homePhones` and `businessPhones` out from under the guard
    that exists to keep a formatter and its `$select` in agreement.
    """
    mobile = getattr(contact, "mobile_phone", "") or ""
    if mobile:
        return mobile
    home = getattr(contact, "home_phones", None) or []
    if home:
        return home[0]
    business = getattr(contact, "business_phones", None) or []
    if business:
        return business[0]
    return ""


def _categories(contact: Any) -> list[str]:
    """Graph's categories, sanitized. A contact's own text, so treat it as such."""
    return [sanitize_output(c) for c in (getattr(contact, "categories", None) or [])]


def _format_contact_summary(contact: Any, *, with_categories: bool = False) -> dict:
    """Convert Graph SDK contact to summary dict.

    ``with_categories`` is opt-in because only the listing path can honour it;
    see ``_LIST_SELECT``.
    """
    email = ""
    if contact.email_addresses:
        first_email = contact.email_addresses[0]
        email = getattr(first_email, "address", "") or ""

    summary = {
        "id": contact.id,
        "display_name": sanitize_output(contact.display_name or ""),
        "email": email,
        "phone": _primary_phone(contact),
        "company": sanitize_output(contact.company_name or ""),
    }
    if with_categories:
        summary["categories"] = _categories(contact)
    return summary


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

    return {
        "id": contact.id,
        "first_name": sanitize_output(contact.given_name or ""),
        "middle_name": sanitize_output(getattr(contact, "middle_name", "") or ""),
        "last_name": sanitize_output(contact.surname or ""),
        "display_name": sanitize_output(contact.display_name or ""),
        "email_addresses": email_addresses,
        "mobile_phone": getattr(contact, "mobile_phone", "") or "",
        "home_phones": list(getattr(contact, "home_phones", None) or []),
        "business_phones": list(getattr(contact, "business_phones", None) or []),
        "company": sanitize_output(contact.company_name or ""),
        "title": sanitize_output(contact.title or ""),
        "job_title": sanitize_output(getattr(contact, "job_title", "") or ""),
        "department": sanitize_output(getattr(contact, "department", "") or ""),
        "birthday": str(contact.birthday) if contact.birthday else None,
        "home_address": _format_address(getattr(contact, "home_address", None)),
        "business_address": _format_address(getattr(contact, "business_address", None)),
        "other_address": _format_address(getattr(contact, "other_address", None)),
        "categories": _categories(contact),
        # A note is body-shaped text: Outlook's box is multi-line, and every
        # other body field in the repo keeps its line breaks. Flattened, a note
        # comes back as one run-on line carrying a stray CR that a terminal
        # client will use to overwrite whatever it already printed.
        "personal_notes": sanitize_output(
            getattr(contact, "personal_notes", "") or "", multiline=True
        ),
    }


async def list_contacts(
    graph_client: Any,
    count: int = 25,
    cursor: str | None = None,
) -> dict:
    """List contacts with pagination."""
    query_params: dict[str, Any] = {
        "$orderby": "displayName",
        "$select": _LIST_SELECT,
    }
    query_params = apply_pagination(query_params, count, cursor)

    from msgraph.generated.users.item.contacts.contacts_request_builder import (
        ContactsRequestBuilder,
    )

    req_config = build_request_config(
        ContactsRequestBuilder.ContactsRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.contacts.get(request_configuration=req_config)

    contacts = [_format_contact_summary(c, with_categories=True) for c in (response.value or [])]
    next_cursor = wrap_nextlink(response.odata_next_link)

    return {
        "contacts": contacts,
        "count": len(contacts),
        "has_more": next_cursor is not None,
        "cursor": next_cursor,
    }


# A phone search reads the listing in pages of this size, and at most this many
# pages: 10,000 contacts, past any address book a person keeps. A bigger one is
# searched as far as the cap and says so with `has_more`.
_PHONE_SCAN_PAGE = 1000
_PHONE_SCAN_PAGES = 10

# The shorter of two numbers has to have at least this many digits for one to be
# read as the other with its country or area code. Fewer would let an extension or
# a short code match every number that happens to end in it.
_MIN_PHONE_DIGITS = 7


def _phone_digits(number: str) -> str:
    """The digits of a phone number, as written."""
    return "".join(ch for ch in number if ch in "0123456789")


def _same_number(wanted: str, stored: str) -> bool:
    """Whether ``stored`` is the number whose digits are ``wanted``, written with or
    without its country or area code: the shorter's digits end the longer's.

    Compared as written and again without leading zeros, so an international
    ``00`` or a trunk ``0`` on one side does not hide the match (``+44 20 7946
    0018`` is ``020 7946 0018``), while a suffix that itself starts with 0
    (``010 0199`` of ``+1 555 010 0199``) still matches.
    """
    digits = _phone_digits(stored)
    for one, other in ((wanted, digits), (wanted.lstrip("0"), digits.lstrip("0"))):
        shorter, longer = sorted((one, other), key=len)
        if len(shorter) >= _MIN_PHONE_DIGITS and longer.endswith(shorter):
            return True
    return False


def _holds_number(contact: Any, wanted: str) -> bool:
    """Whether any of the contact's three consumer phone fields holds ``wanted``.

    Spelled out field by field, as in ``_primary_phone``, so every field read is
    one the scan's ``$select`` asks for.
    """
    numbers = [getattr(contact, "mobile_phone", "") or ""]
    numbers += list(getattr(contact, "home_phones", None) or [])
    numbers += list(getattr(contact, "business_phones", None) or [])
    return any(_same_number(wanted, number) for number in numbers if number)


async def _search_by_phone(graph_client: Any, wanted: str, count: int) -> dict:
    """Contacts holding the number whose digits are ``wanted``, from the listing.

    Graph's contact ``$search`` finds a number only as it is stored, and phone
    fields cannot be filtered (``$filter`` on contacts reaches only
    ``emailAddresses/any(a:a/address eq ...)``). So this reads the listing, only
    the summary fields, and compares digits here. Nothing is kept after the call.
    The pages are followed by their raw ``@odata.nextLink``, as
    ``todo.fetch_all_task_lists`` follows its pages.
    """
    from msgraph.generated.users.item.contacts.contacts_request_builder import (
        ContactsRequestBuilder,
    )

    req_config = build_request_config(
        ContactsRequestBuilder.ContactsRequestBuilderGetQueryParameters,
        {"$top": _PHONE_SCAN_PAGE, "$select": _SUMMARY_SELECT},
    )
    response = await graph_client.me.contacts.get(request_configuration=req_config)
    matches: list[Any] = []
    pages = 1
    while True:
        matches.extend(c for c in (response.value or []) if _holds_number(c, wanted))
        next_link = getattr(response, "odata_next_link", None)
        if not isinstance(next_link, str) or not next_link or pages >= _PHONE_SCAN_PAGES:
            break
        pages += 1
        response = await graph_client.me.contacts.with_url(next_link).get()
    unread = isinstance(next_link, str) and bool(next_link)

    found = [_format_contact_summary(c) for c in matches[:count]]
    return {
        "contacts": found,
        "count": len(found),
        # No cursor: a scan cut short at the cap, or more matches than `count`,
        # is all `has_more` can say.
        "has_more": len(matches) > count or unread,
        "cursor": None,
    }


async def search_contacts(
    graph_client: Any,
    query: str | None = None,
    count: int = 25,
    phone: str | None = None,
) -> dict:
    """Search contacts using KQL, or by phone number (see ``_search_by_phone``)."""
    count = _clamp(count, 1, 100)
    if (query is None) == (phone is None):
        raise ValueError(
            "Pass query or phone, not both: query is a KQL search, phone finds a number."
        )
    if phone is not None:
        wanted = _phone_digits(validate_phone(phone))
        if len(wanted) < _MIN_PHONE_DIGITS:
            raise ValueError(
                f"phone needs at least {_MIN_PHONE_DIGITS} digits to tell one number from another."
            )
        return await _search_by_phone(graph_client, wanted, count)

    safe_query = sanitize_kql(query)

    query_params: dict[str, Any] = {
        "$top": count,
        "$search": safe_query,
        "$select": _SUMMARY_SELECT,
    }

    from msgraph.generated.users.item.contacts.contacts_request_builder import (
        ContactsRequestBuilder,
    )

    req_config = build_request_config(
        ContactsRequestBuilder.ContactsRequestBuilderGetQueryParameters, query_params
    )
    response = await graph_client.me.contacts.get(request_configuration=req_config)

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
    job_title: str | None = None,
    phone_type: str = "mobile",
    birthday: str | None = None,
    middle_name: str | None = None,
    personal_notes: str | None = None,
    categories: list[str] | None = None,
    *,
    config: Config,
) -> dict:
    """Create a new contact.

    ``title`` is Graph's name prefix ("Dr.", "Ms."), which Graph folds into the
    display name; the job goes in ``job_title``. The two used to be one argument
    in practice — a model filling in a business card put the job in ``title``
    and got "Engineer Jane Doe" back (#119).

    ``birthday`` is a date, stored at 11:59 UTC (see ``validate_birthday``).
    """
    check_permission(config, CATEGORY_CONTACTS_WRITE, "outlook_create_contact")

    if email:
        validate_email(email)
    if phone:
        validate_phone(phone)
    _check_phone_type(phone, phone_type)
    born = validate_birthday(birthday) if birthday is not None else None

    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.email_address import EmailAddress

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
        _assign_phone(new_contact, phone, phone_type)
    if company:
        new_contact.company_name = company
    if title:
        new_contact.title = title
    if job_title:
        new_contact.job_title = job_title
    if born is not None:
        new_contact.birthday = born
    if middle_name:
        new_contact.middle_name = middle_name
    if personal_notes:
        new_contact.personal_notes = personal_notes
    if categories:
        new_contact.categories = categories

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
    phone_type: str = "mobile",
    company: str | None = None,
    title: str | None = None,
    job_title: str | None = None,
    home_address: dict | None = None,
    business_address: dict | None = None,
    other_address: dict | None = None,
    clear_phones: list[str] | None = None,
    birthday: str | None = None,
    middle_name: str | None = None,
    personal_notes: str | None = None,
    categories: list[str] | None = None,
    add_emails: list[str] | None = None,
    remove_emails: list[str] | None = None,
    *,
    config: Config,
) -> dict:
    """Update an existing contact (partial patch).

    Each address takes the shape :func:`get_contact` returns — any subset of
    ``street``, ``city``, ``state``, ``postal_code``, ``country_or_region`` —
    and Graph **replaces** that whole address object rather than merging into
    it: parts not supplied come back empty, not preserved. So pass back every
    part you intend to keep, which is why the read and write halves name the
    same five fields. Omit an address entirely to leave it untouched.

    ``phone_type`` follows the same rule for the list-valued numbers: a business
    or home write replaces that list with ``[phone]``. ``clear_phones`` empties
    the named types, so a number filed as mobile moves to business in one patch.
    ``title`` is the name prefix, here so a job that landed in it can be
    corrected in place (#119). ``birthday`` sets or replaces the date; there is
    no clearing it. ``categories`` replaces the contact's whole list, as
    ``outlook_categorize_message`` does a message's, so ``[]`` removes them all.

    ``email`` replaces every address with the one given. ``add_emails`` and
    ``remove_emails`` change only the addresses they name, so they are merged
    into the addresses the contact has now: the contact is read, the merge is
    computed from that read, and the PATCH is pinned to it with ``If-Match``.
    That is #87's rule for a write that sends back something it read: an edit
    made in between is refused with ``412`` rather than overwritten, and a read
    with no change key goes out unpinned. Every argument is checked before the
    read; the checks that need the stored addresses (a remove the contact does
    not have, more than three in all) come after it and before any write.
    """
    check_permission(config, CATEGORY_CONTACTS_WRITE, "outlook_update_contact")
    contact_id = validate_graph_id(contact_id)

    # `is not None`, not truthiness: `email=""` is a caller asking for something
    # this tool cannot do, and it used to skip validation and reach Graph as a
    # blank address. One meaning for "provided" per argument.
    if email is not None:
        validate_email(email)
    if phone is not None:
        validate_phone(phone)
    _check_phone_type(phone, phone_type)
    if clear_phones is not None:
        _check_clear_phones(clear_phones, phone, phone_type)
    born = validate_birthday(birthday) if birthday is not None else None
    if email is not None and (add_emails is not None or remove_emails is not None):
        raise ValueError(
            "Pass email to replace every address, or add_emails / remove_emails to "
            "change only the ones named — not both."
        )
    adding = _email_list("add_emails", add_emails)
    removing = _email_list("remove_emails", remove_emails)
    both = sorted({a.lower() for a in adding or []} & {a.lower() for a in removing or []})
    if both:
        raise ValueError(
            f"{', '.join(both)} is in both add_emails and remove_emails. Name each address once."
        )

    from msgraph.generated.models.contact import Contact
    from msgraph.generated.models.email_address import EmailAddress

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
        _assign_phone(patch_body, phone, phone_type)
    if clear_phones:
        _clear_phones(patch_body, clear_phones)
    if company is not None:
        patch_body.company_name = company
    if title is not None:
        patch_body.title = title
    if job_title is not None:
        patch_body.job_title = job_title
    if born is not None:
        patch_body.birthday = born
    if middle_name is not None:
        patch_body.middle_name = middle_name
    if personal_notes is not None:
        patch_body.personal_notes = personal_notes
    # A list like the business/home numbers: Graph replaces it with what is sent,
    # so `[]` removes every category and None leaves them alone.
    if categories is not None:
        patch_body.categories = categories

    if home_address is not None:
        patch_body.home_address = _build_address("home", home_address)
    if business_address is not None:
        patch_body.business_address = _build_address("business", business_address)
    if other_address is not None:
        patch_body.other_address = _build_address("other", other_address)

    request_builder = graph_client.me.contacts.by_contact_id(contact_id)
    if_match: str | None = None
    if adding is not None or removing is not None:
        # Last, so every check above has run before anything is read. No
        # $select, as in get_contact: the whole contact, change key included.
        current = await request_builder.get()
        patch_body.email_addresses = _merged_emails(current, adding or [], removing or [])
        if_match = (getattr(current, "additional_data", None) or {}).get("@odata.etag")

    if if_match:
        from kiota_abstractions.base_request_configuration import RequestConfiguration

        patch_config = RequestConfiguration()
        patch_config.headers.add("If-Match", if_match)
        await request_builder.patch(patch_body, request_configuration=patch_config)
    else:
        await request_builder.patch(patch_body)

    return {"status": "updated"}


async def delete_contact(
    graph_client: Any,
    contact_id: str,
    *,
    config: Config,
) -> dict:
    """Delete a contact by ID."""
    check_permission(config, CATEGORY_CONTACTS_WRITE, "outlook_delete_contact")
    contact_id = validate_graph_id(contact_id)

    await graph_client.me.contacts.by_contact_id(contact_id).delete()

    return {"status": "deleted"}
