"""Tests for contact tools."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest
from mcp.client import Client

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import encode_cursor
from outlook_mcp.server import mcp
from outlook_mcp.tools.contacts import (
    _ADDRESS_FIELDS,
    _LIST_SELECT,
    _SUMMARY_SELECT,
    _build_address,
    _format_address,
    _format_contact_detail,
    _format_contact_summary,
    create_contact,
    delete_contact,
    get_contact,
    list_contacts,
    search_contacts,
    update_contact,
)
from tests.test_write_payloads_reach_the_wire import wire

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _make_mock_contact(**overrides):
    """Factory for mock Graph SDK contact objects.

    Matches the consumer Outlook contact shape: ``mobile_phone`` (single
    string), ``home_phones`` (list[str]), ``business_phones`` (list[str]).
    """
    contact = MagicMock(
        spec=[
            "id",
            "display_name",
            "given_name",
            "middle_name",
            "surname",
            "company_name",
            "title",
            "job_title",
            "department",
            "birthday",
            "email_addresses",
            "mobile_phone",
            "home_phones",
            "business_phones",
            "home_address",
            "business_address",
            "other_address",
            "categories",
            "personal_notes",
        ]
    )
    contact.id = overrides.get("id", "contact123")
    contact.display_name = overrides.get("display_name", "John Doe")
    contact.given_name = overrides.get("given_name", "John")
    contact.middle_name = overrides.get("middle_name", "")
    contact.surname = overrides.get("surname", "Doe")
    contact.company_name = overrides.get("company_name", "Acme")
    contact.title = overrides.get("title", "Dr.")
    contact.job_title = overrides.get("job_title", "Engineer")
    contact.department = overrides.get("department", "")
    contact.birthday = overrides.get("birthday", None)

    email = MagicMock()
    email.address = overrides.get("email_address", "john@test.com")
    email.name = overrides.get("email_name", "John")
    contact.email_addresses = overrides.get("email_addresses", [email])

    contact.mobile_phone = overrides.get("mobile_phone", "+1234567890")
    contact.home_phones = overrides.get("home_phones", [])
    contact.business_phones = overrides.get("business_phones", [])

    # Graph returns an empty address object, not null, for an address the
    # contact does not have — so the default here is what the SDK really hands
    # back for a contact with no addresses at all.
    for field in ("home_address", "business_address", "other_address"):
        setattr(contact, field, overrides.get(field, _make_mock_address()))
    contact.categories = overrides.get("categories", [])
    contact.personal_notes = overrides.get("personal_notes", "")

    return contact


def _make_mock_address(**fields):
    """Factory for a Graph physicalAddress; empty by default, as Graph sends it."""
    address = MagicMock(spec=["street", "city", "state", "postal_code", "country_or_region"])
    for field in ("street", "city", "state", "postal_code", "country_or_region"):
        setattr(address, field, fields.get(field, ""))
    return address


def _make_contacts_mock(contacts, next_link=None):
    """Build a mock Graph client for contacts list/search queries."""
    response = MagicMock(value=contacts, odata_next_link=next_link)
    client = MagicMock()
    client.me.contacts.get = AsyncMock(return_value=response)
    return client


def _make_contact_by_id_mock(contact):
    """Build a mock Graph client for single-contact operations."""
    contact_obj = MagicMock()
    contact_obj.get = AsyncMock(return_value=contact)
    contact_obj.patch = AsyncMock()
    contact_obj.delete = AsyncMock()
    client = MagicMock()
    client.me.contacts.by_contact_id = MagicMock(return_value=contact_obj)
    return client


class TestListContacts:
    async def test_list_returns_contacts(self):
        """list_contacts returns structured contact list."""
        mock_contact = _make_mock_contact()
        mock_client = _make_contacts_mock([mock_contact])

        result = await list_contacts(mock_client)
        assert result["count"] == 1
        assert result["contacts"][0]["id"] == "contact123"
        assert result["contacts"][0]["display_name"] == "John Doe"
        assert result["contacts"][0]["email"] == "john@test.com"
        assert result["contacts"][0]["phone"] == "+1234567890"
        assert result["contacts"][0]["company"] == "Acme"

    async def test_list_select_uses_consumer_phone_fields(self):
        """list_contacts $select must use mobilePhone/homePhones/businessPhones,
        not the unsupported ``phones`` aggregate (Bug #1)."""
        mock_client = _make_contacts_mock([])
        await list_contacts(mock_client)

        call_kwargs = mock_client.me.contacts.get.call_args
        select = call_kwargs.kwargs["request_configuration"].query_parameters.select
        select_str = ",".join(select) if isinstance(select, list) else select
        assert "phones" not in select_str.split(",")
        assert "mobilePhone" in select_str
        assert "homePhones" in select_str
        assert "businessPhones" in select_str

    async def test_list_summary_falls_back_to_home_phone(self):
        """When mobile_phone is empty, summary falls back to first home phone."""
        contact = _make_mock_contact(mobile_phone="", home_phones=["+15551112222"])
        mock_client = _make_contacts_mock([contact])

        result = await list_contacts(mock_client)
        assert result["contacts"][0]["phone"] == "+15551112222"

    async def test_list_summary_falls_back_to_business_phone(self):
        """When mobile and home are empty, falls back to first business phone."""
        contact = _make_mock_contact(
            mobile_phone="",
            home_phones=[],
            business_phones=["+15553334444"],
        )
        mock_client = _make_contacts_mock([contact])

        result = await list_contacts(mock_client)
        assert result["contacts"][0]["phone"] == "+15553334444"

    async def test_list_with_cursor(self):
        """list_contacts passes cursor to pagination."""
        mock_client = _make_contacts_mock([])
        cursor = encode_cursor(25)
        result = await list_contacts(mock_client, cursor=cursor)
        assert result["count"] == 0

        # Verify skip was passed via request_configuration
        call_kwargs = mock_client.me.contacts.get.call_args
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert qp.skip == 25

    async def test_list_has_more_with_next_link(self):
        """has_more is True and cursor returned when odata_next_link present."""
        mock_client = _make_contacts_mock(
            [_make_mock_contact()],
            next_link="https://graph.microsoft.com/v1.0/me/contacts?$skip=25",
        )
        result = await list_contacts(mock_client, count=1)
        assert result["has_more"] is True
        assert result["cursor"] is not None

    async def test_list_no_more(self):
        """has_more is False and cursor is None when no next link."""
        mock_client = _make_contacts_mock([_make_mock_contact()])
        result = await list_contacts(mock_client)
        assert result["has_more"] is False
        assert result["cursor"] is None


class TestSearchContacts:
    async def test_search_sanitizes_query(self):
        """search_contacts sanitizes KQL before sending to Graph."""
        mock_client = _make_contacts_mock([])
        result = await search_contacts(mock_client, query='John" OR (hack)')
        assert result["count"] == 0

        # Verify the search param was sanitized. Only the string-literal
        # boundary chars are stripped; `(` is legitimate KQL grouping.
        call_kwargs = mock_client.me.contacts.get.call_args
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert '"' not in qp.search.strip('"')
        assert "\\" not in qp.search

    async def test_search_returns_contacts(self):
        """search_contacts returns matching contacts."""
        mock_contact = _make_mock_contact()
        mock_client = _make_contacts_mock([mock_contact])
        result = await search_contacts(mock_client, query="John")
        assert result["count"] == 1
        assert result["contacts"][0]["display_name"] == "John Doe"

    async def test_search_select_uses_consumer_phone_fields(self):
        """search_contacts $select must use consumer phone fields (Bug #1)."""
        mock_client = _make_contacts_mock([])
        await search_contacts(mock_client, query="John")

        call_kwargs = mock_client.me.contacts.get.call_args
        select = call_kwargs.kwargs["request_configuration"].query_parameters.select
        select_str = ",".join(select) if isinstance(select, list) else select
        assert "phones" not in select_str.split(",")
        assert "mobilePhone" in select_str
        assert "homePhones" in select_str
        assert "businessPhones" in select_str


class TestGetContact:
    async def test_get_returns_full_detail(self):
        """get_contact returns full contact detail using consumer phone fields."""
        mock_contact = _make_mock_contact(
            home_phones=["+15551112222"],
            business_phones=["+15553334444"],
        )
        mock_client = _make_contact_by_id_mock(mock_contact)

        result = await get_contact(mock_client, "contact123")
        assert result["id"] == "contact123"
        assert result["first_name"] == "John"
        assert result["last_name"] == "Doe"
        assert result["display_name"] == "John Doe"
        assert result["company"] == "Acme"
        assert result["title"] == "Dr."
        assert result["job_title"] == "Engineer"
        assert len(result["email_addresses"]) == 1
        assert result["email_addresses"][0]["address"] == "john@test.com"
        assert result["mobile_phone"] == "+1234567890"
        assert result["home_phones"] == ["+15551112222"]
        assert result["business_phones"] == ["+15553334444"]
        assert "phones" not in result, "old aggregate 'phones' field must not appear"

    async def test_get_handles_empty_phone_fields(self):
        """get_contact returns empty defaults when phone fields are missing."""
        mock_contact = _make_mock_contact(
            mobile_phone="",
            home_phones=[],
            business_phones=[],
        )
        mock_client = _make_contact_by_id_mock(mock_contact)

        result = await get_contact(mock_client, "contact123")
        assert result["mobile_phone"] == ""
        assert result["home_phones"] == []
        assert result["business_phones"] == []

    async def test_get_validates_id(self):
        """get_contact rejects invalid contact IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await get_contact(mock_client, "bad id with spaces!")


class TestCreateContact:
    async def test_create_contact(self):
        """create_contact posts to Graph and returns contact."""
        mock_contact = _make_mock_contact()
        mock_client = MagicMock()
        mock_client.me.contacts.post = AsyncMock(return_value=mock_contact)

        result = await create_contact(
            mock_client,
            first_name="John",
            last_name="Doe",
            email="john@test.com",
            phone="+1234567890",
            company="Acme",
            title="Engineer",
            config=_CFG,
        )
        assert result["status"] == "created"
        assert result["id"] == "contact123"
        mock_client.me.contacts.post.assert_called_once()

    async def test_create_contact_writes_mobile_phone_not_phones(self):
        """create_contact must set mobile_phone (not the unsupported 'phones'
        collection) on consumer Graph (Bug #1)."""
        mock_client = MagicMock()
        mock_client.me.contacts.post = AsyncMock(return_value=_make_mock_contact())

        await create_contact(
            mock_client,
            first_name="John",
            phone="+1234567890",
            config=_CFG,
        )

        payload = mock_client.me.contacts.post.call_args.args[0]
        assert payload.mobile_phone == "+1234567890"
        # The unsupported 'phones' aggregate must not be set
        assert getattr(payload, "phones", None) is None

    async def test_create_validates_email(self):
        """create_contact rejects invalid email."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid email"):
            await create_contact(
                mock_client,
                first_name="John",
                email="not-an-email",
                config=_CFG,
            )

    async def test_create_validates_phone(self):
        """create_contact rejects invalid phone number."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid phone"):
            await create_contact(
                mock_client,
                first_name="John",
                phone="not a phone!!!",
                config=_CFG,
            )

    async def test_create_raises_read_only(self):
        """create_contact raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await create_contact(mock_client, first_name="John", config=_CFG_RO)


class TestUpdateContact:
    async def test_update_patches_partial(self):
        """update_contact patches only provided fields."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())

        result = await update_contact(
            mock_client,
            contact_id="contact123",
            first_name="Jane",
            config=_CFG,
        )
        assert result["status"] == "updated"

        # Verify patch was called
        contact_obj = mock_client.me.contacts.by_contact_id.return_value
        contact_obj.patch.assert_called_once()

    async def test_update_writes_mobile_phone_not_phones(self):
        """update_contact must set mobile_phone (not 'phones') on consumer Graph."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())

        await update_contact(
            mock_client,
            contact_id="contact123",
            phone="+19998887777",
            config=_CFG,
        )
        contact_obj = mock_client.me.contacts.by_contact_id.return_value
        payload = contact_obj.patch.call_args.args[0]
        assert payload.mobile_phone == "+19998887777"
        assert getattr(payload, "phones", None) is None

    async def test_unknown_phone_type_is_refused(self):
        """A misspelt type must not fall back to mobile and report success."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ValueError, match="phone_type.*mobile, business, home"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                phone="+19998887777",
                phone_type="work",
                config=_CFG,
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_phone_type_without_phone_is_refused(self):
        """Without a number there is nothing to file under the type — a patch
        that carried nothing would report "updated" and change nothing (#41's shape)."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ValueError, match="phone_type needs phone"):
            await update_contact(
                mock_client, contact_id="contact123", phone_type="business", config=_CFG
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_clearing_an_unknown_phone_type_is_refused(self):
        """A misspelt type would clear nothing and report "updated"."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ValueError, match="clear_phones.*mobile, business, home"):
            await update_contact(
                mock_client, contact_id="contact123", clear_phones=["work"], config=_CFG
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_clearing_the_type_being_written_is_refused(self):
        """Write and clear the same field: one of them would be silently lost."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ValueError, match="clear_phones.*business"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                phone="+19998887777",
                phone_type="business",
                clear_phones=["business"],
                config=_CFG,
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_create_refuses_unknown_phone_type(self):
        mock_client = MagicMock()
        mock_client.me.contacts.post = AsyncMock()
        with pytest.raises(ValueError, match="phone_type"):
            await create_contact(
                mock_client, first_name="John", phone="+1234567890", phone_type="fax", config=_CFG
            )
        mock_client.me.contacts.post.assert_not_called()

    async def test_update_validates_id(self):
        """update_contact rejects invalid contact IDs."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await update_contact(
                mock_client,
                contact_id="bad id!",
                first_name="Jane",
                config=_CFG,
            )

    async def test_update_raises_read_only(self):
        """update_contact raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await update_contact(
                mock_client,
                contact_id="contact123",
                first_name="Jane",
                config=_CFG_RO,
            )


class TestDeleteContact:
    async def test_delete_contact(self):
        """delete_contact calls delete on Graph."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())

        result = await delete_contact(mock_client, "contact123", config=_CFG)
        assert result["status"] == "deleted"

        contact_obj = mock_client.me.contacts.by_contact_id.return_value
        contact_obj.delete.assert_called_once()

    async def test_delete_raises_read_only(self):
        """delete_contact raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_contact(mock_client, "contact123", config=_CFG_RO)


class TestContactDetailCarriesEverythingStored:
    """The detail formatter used to drop fields Graph had already returned.

    `get_contact` sends no `$select`, so Graph returns the whole contact — the
    addresses, categories and notes were arriving and being discarded on the
    way out. A contact with a home address and two categories read back as
    having neither, which is indistinguishable from not having them.
    """

    def _detail(self, **overrides):
        return _format_contact_detail(_make_mock_contact(**overrides))

    def test_home_address_is_returned_field_by_field(self):
        detail = self._detail(
            home_address=_make_mock_address(
                street="693 7th St S",
                city="Kirkland",
                state="WA",
                postal_code="98033",
                country_or_region="USA",
            )
        )
        assert detail["home_address"] == {
            "street": "693 7th St S",
            "city": "Kirkland",
            "state": "WA",
            "postal_code": "98033",
            "country_or_region": "USA",
        }

    def test_an_empty_address_object_reads_as_no_address(self):
        """Graph sends an empty object, not null — that must not become an empty dict."""
        assert self._detail()["home_address"] is None

    def test_a_partial_address_keeps_the_parts_it_has(self):
        detail = self._detail(business_address=_make_mock_address(city="Seattle"))
        assert detail["business_address"] == {
            "street": "",
            "city": "Seattle",
            "state": "",
            "postal_code": "",
            "country_or_region": "",
        }

    def test_all_three_address_slots_are_carried(self):
        detail = self._detail(
            home_address=_make_mock_address(city="Kirkland"),
            business_address=_make_mock_address(city="Redmond"),
            other_address=_make_mock_address(city="Bellevue"),
        )
        assert [detail[f"{slot}_address"]["city"] for slot in ("home", "business", "other")] == [
            "Kirkland",
            "Redmond",
            "Bellevue",
        ]

    def test_categories_are_returned(self):
        detail = self._detail(categories=["Christmas Card", "Microsoft Party"])
        assert detail["categories"] == ["Christmas Card", "Microsoft Party"]

    def test_no_categories_is_an_empty_list_not_none(self):
        assert self._detail()["categories"] == []

    def test_personal_notes_are_returned(self):
        assert self._detail(personal_notes="Met at the 2025 offsite")["personal_notes"] == (
            "Met at the 2025 offsite"
        )

    def test_a_multi_line_note_keeps_its_line_breaks(self):
        """A note is body-shaped text, and every other body field passes multiline=True.

        Flattened it is one run-on line — and since `_CONTROL_CHARS` leaves
        `\x0d` alone, one with a stray CR in it, which a terminal client will
        use to overwrite whatever it has already printed.
        """
        note = "Met at the 2025 offsite.\r\nFollow up in Q3."
        assert self._detail(personal_notes=note)["personal_notes"] == note

    def test_a_multi_line_street_keeps_its_line_breaks(self):
        """Outlook's Street box is multi-line, and here that is not cosmetic.

        The docstring tells the caller to read the parts back and hand them in
        to keep them, so whatever this returns is what lands in the mailbox.
        """
        street = "Apt 4\r\n693 7th St S"
        detail = self._detail(home_address=_make_mock_address(street=street))
        assert detail["home_address"]["street"] == street

    def test_control_characters_are_still_stripped_from_a_note(self):
        """multiline=True keeps line breaks; it does not stop sanitizing."""
        noisy = "Met\x07 at\x1b[31m the offsite"
        assert self._detail(personal_notes=noisy)["personal_notes"] == "Met at the offsite"

    def test_address_content_is_sanitized_like_every_other_echoed_field(self):
        """A contact is attacker-influenced text; sanitize_output strips ANSI and
        control characters from it, exactly as it does for every other field here."""
        detail = self._detail(home_address=_make_mock_address(street="693 7th[31m St S"))
        assert detail["home_address"]["street"] == "693 7th St S"


class TestSummarySelectMatchesTheSummaryFormatter:
    """A $select narrower than the formatter silently blanks the difference."""

    def test_every_field_the_summary_reads_is_selected(self):
        assert {
            "id",
            "displayName",
            "emailAddresses",
            "mobilePhone",
            "homePhones",
            "businessPhones",
            "companyName",
        } <= set(_SUMMARY_SELECT.split(","))

    def test_nothing_is_selected_that_the_summary_never_reads(self):
        """Over-selection is bytes nobody reads; givenName/surname/title were that."""
        assert {"givenName", "surname", "title"}.isdisjoint(set(_SUMMARY_SELECT.split(",")))

    def test_only_the_listing_select_asks_for_categories(self):
        assert "categories" in _LIST_SELECT.split(",")
        assert "categories" not in _SUMMARY_SELECT.split(",")

    def test_listing_summaries_carry_categories(self):
        summary = _format_contact_summary(
            _make_mock_contact(categories=["Christmas Card"]), with_categories=True
        )
        assert summary["categories"] == ["Christmas Card"]

    def test_search_summaries_omit_the_key_rather_than_report_it_empty(self):
        """Graph's $search never returns categories, so an empty list would be a lie."""
        summary = _format_contact_summary(_make_mock_contact(categories=["Christmas Card"]))
        assert "categories" not in summary


class TestUpdateContactWritesTheAddressesItCanRead:
    """The write-side twin of TestContactDetailCarriesEverythingStored.

    An address that `get_contact` returns but `update_contact` cannot set is
    only half a fix: the field is readable and not correctable. These pin that
    the two sides speak the same five field *names* — not merely five fields in
    the same order, which is all a positional round-trip could prove.
    """

    @staticmethod
    def _patched(mock_client):
        """The Contact body handed to Graph."""
        contact_obj = mock_client.me.contacts.by_contact_id.return_value
        contact_obj.patch.assert_called_once()
        return contact_obj.patch.call_args[0][0]

    async def test_home_address_reaches_the_patch_body(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client,
            contact_id="contact123",
            home_address={
                "street": "2821 252nd Ave SE",
                "city": "Sammamish",
                "state": "WA",
                "postal_code": "98075",
                "country_or_region": "USA",
            },
            config=_CFG,
        )
        address = self._patched(mock_client).home_address
        assert address.street == "2821 252nd Ave SE"
        assert address.city == "Sammamish"
        assert address.state == "WA"
        assert address.postal_code == "98075"
        assert address.country_or_region == "USA"

    @pytest.mark.parametrize("slot", ["home_address", "business_address", "other_address"])
    async def test_every_slot_that_reads_back_is_writable(self, slot):
        """All three come back from get_contact, so all three are settable."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", config=_CFG, **{slot: {"city": "Kirkland"}}
        )
        assert getattr(self._patched(mock_client), slot).city == "Kirkland"

    async def test_the_slots_do_not_bleed_into_each_other(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client,
            contact_id="contact123",
            home_address={"city": "Kirkland"},
            other_address={"city": "Bellevue"},
            config=_CFG,
        )
        patched = self._patched(mock_client)
        assert patched.home_address.city == "Kirkland"
        assert patched.other_address.city == "Bellevue"
        assert patched.business_address is None

    async def test_a_partial_address_is_still_written(self):
        """Someone may know the city and not the street.

        Graph replaces the whole address rather than merging, so the parts not
        supplied come back empty — verified live. That belongs in the docstring
        the model reads, not in a silent difference between what was asked for
        and what was stored.
        """
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_address={"city": "Bothell"}, config=_CFG
        )
        address = self._patched(mock_client).home_address
        assert address.city == "Bothell"
        assert address.street is None

    async def test_omitting_the_addresses_leaves_them_alone(self):
        """Partial patch: a name-only update must not blank a stored address."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(mock_client, contact_id="contact123", first_name="Jane", config=_CFG)
        patched = self._patched(mock_client)
        assert patched.home_address is None
        assert patched.business_address is None
        assert patched.other_address is None

    async def test_values_are_stripped(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_address={"city": "  Kirkland  "}, config=_CFG
        )
        assert self._patched(mock_client).home_address.city == "Kirkland"

    async def test_a_part_that_is_only_whitespace_is_not_written(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client,
            contact_id="contact123",
            home_address={"city": "Bothell", "street": "   "},
            config=_CFG,
        )
        address = self._patched(mock_client).home_address
        assert address.city == "Bothell"
        assert address.street is None

    async def test_read_only_mode_refuses_the_write(self):
        """The new parameters must not open a path around the read-only gate."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ReadOnlyError):
            await update_contact(
                mock_client,
                contact_id="contact123",
                home_address={"city": "Bothell"},
                config=_CFG_RO,
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()


class TestAnAddressThatWouldPatchNothingIsRefused:
    """A status of "updated" has to mean something was updated.

    An address that collapses to nothing — every part empty, every part
    whitespace, or keys this tool does not know — would otherwise skip the
    assignment, send a PATCH without it, and answer `{"status": "updated"}`.
    That is the silent no-op this module exists to stop telling, and the caller
    most likely to hit it is a model asked to clear an address.
    """

    @staticmethod
    def _client():
        return _make_contact_by_id_mock(_make_mock_contact())

    @pytest.mark.parametrize(
        "address",
        [{}, {"street": ""}, {"street": "   ", "city": ""}, {"street": None, "city": None}],
        ids=["empty-dict", "empty-value", "whitespace", "explicit-nulls"],
    )
    async def test_an_address_with_no_content_is_an_error(self, address):
        client = self._client()
        with pytest.raises(ValueError, match="cannot clear an address"):
            await update_contact(client, contact_id="contact123", home_address=address, config=_CFG)
        client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_an_unknown_part_is_an_error_naming_the_valid_ones(self):
        """A misspelt key that patched nothing would be #41's shape again."""
        client = self._client()
        with pytest.raises(ValueError, match="country_or_region") as caught:
            await update_contact(
                client,
                contact_id="contact123",
                home_address={"city": "Bothell", "country": "USA", "zip": "98011"},
                config=_CFG,
            )
        assert "country" in str(caught.value) and "zip" in str(caught.value)
        client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    async def test_the_slot_is_named_in_the_error(self):
        client = self._client()
        with pytest.raises(ValueError, match="business_address"):
            await update_contact(
                client, contact_id="contact123", business_address={"zip": "98011"}, config=_CFG
            )

    @pytest.mark.parametrize("address", ["693 7th St S", ["Kirkland"], 98033])
    async def test_an_address_that_is_not_an_object_is_an_error(self, address):
        client = self._client()
        with pytest.raises(ValueError, match="takes an object"):
            await update_contact(client, contact_id="contact123", home_address=address, config=_CFG)

    async def test_a_part_that_is_not_a_string_is_an_error(self):
        client = self._client()
        with pytest.raises(ValueError, match="must be a string"):
            await update_contact(
                client, contact_id="contact123", home_address={"postal_code": 98033}, config=_CFG
            )

    async def test_an_empty_email_is_rejected_rather_than_sent(self):
        """`email=""` used to skip validation and reach Graph as a blank address."""
        client = self._client()
        with pytest.raises(ValueError, match="Invalid email"):
            await update_contact(client, contact_id="contact123", email="", config=_CFG)
        client.me.contacts.by_contact_id.return_value.patch.assert_not_called()


class TestTheReadAndWriteHalvesNameTheSameParts:
    """Not "five fields in the same order" — the same keys, by name.

    The docstring tells the caller to hand an address straight back to keep the
    parts it is not changing. That is one dict splat, and it resolves only if
    every key the read path emits is a key the write path accepts. A positional
    round-trip cannot see a rename on either side.
    """

    _PARTS = {
        "street": "693 7th St S",
        "city": "Kirkland",
        "state": "WA",
        "postal_code": "98033",
        "country_or_region": "USA",
    }

    def test_what_the_write_path_stores_the_read_path_returns(self):
        assert _format_address(_build_address("home", self._PARTS)) == self._PARTS

    def test_the_detail_shape_is_accepted_verbatim_by_the_write_path(self):
        """The documented round trip, performed literally."""
        detail = _format_contact_detail(
            _make_mock_contact(home_address=_make_mock_address(**self._PARTS))
        )
        assert _format_address(_build_address("home", detail["home_address"])) == self._PARTS

    def test_the_read_path_emits_exactly_the_keys_the_write_path_accepts(self):
        detail = _format_contact_detail(
            _make_mock_contact(home_address=_make_mock_address(city="Kirkland"))
        )
        assert set(detail["home_address"]) == set(_ADDRESS_FIELDS)

    def test_an_address_of_only_whitespace_reads_as_no_address(self):
        """The two halves have to agree on what counts as content.

        `_build_address` refuses a blank part, so a stored address of spaces
        reported here as real would be one the docstring tells the caller to
        hand back and the write path then rejects.
        """
        assert _format_address(_make_mock_address(street="   ", city=" ")) is None


class TestPartialAddressDoesNotLeakNulls:
    """A part we were not given must be left unset, never assigned ``None``.

    Assigning ``None`` to the omitted parts looked equivalent and was not: the
    backing store emits an explicitly-``None`` field, and for a nested model it
    emits it onto the **parent**, under its Python name. Every partial address
    went out as::

        {"country_or_region": null, "postal_code": null, "state": null,
         "street": null, "homeAddress": {"city": "Bothell"}}

    which Graph rejects — ``400 The property 'country_or_region' does not exist
    on type 'microsoft.graph.contact'``. The full five-part write returned 200,
    so the tool worked for the case anyone would test by hand and failed for the
    one it was written for.

    ``assert_on_wire`` now rejects leaked keys and stray nulls for every write
    tool (#64), so this class keeps only what is specific to addresses: the
    partial write, where the leak actually happened, asserted key by key.
    """

    async def test_a_city_only_patch_sends_the_city_and_nothing_else(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_address={"city": "Bothell"}, config=_CFG
        )
        body = json.loads(
            wire(mock_client.me.contacts.by_contact_id.return_value.patch.call_args[0][0])
        )

        # By key, not by substring: a contact on Nullah Road fails `"null" not in
        # body`, and comparing serialized text additionally rides on kiota's
        # separators and key order.
        assert set(body) == {"@odata.type", "homeAddress"}
        assert body["homeAddress"] == {"city": "Bothell"}

    async def test_every_slot_patches_only_itself(self):
        """Three addresses, three chances for a leak onto the parent."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client,
            contact_id="contact123",
            business_address={"city": "Redmond", "postal_code": "98052"},
            config=_CFG,
        )
        body = json.loads(
            wire(mock_client.me.contacts.by_contact_id.return_value.patch.call_args[0][0])
        )

        assert set(body) == {"@odata.type", "businessAddress"}
        assert body["businessAddress"] == {"city": "Redmond", "postalCode": "98052"}


class TestGetContactAsksGraphForEverything:
    """The contract that makes the detail formatter safe.

    `_format_contact_detail` reads seventeen fields, and the reason it can is
    that `get_contact` sends no `$select` — Graph then returns the whole
    contact. The listing paths pair a formatter with a `$select` and are
    guarded by `test_select_covers_the_formatter`; this path has no `$select`
    to guard, so what needs pinning is that it stays that way. Narrow it and
    every field nobody thought to list comes back empty, which is
    indistinguishable from genuinely empty.
    """

    async def test_no_select_is_sent(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await get_contact(mock_client, "contact123")

        builder = mock_client.me.contacts.by_contact_id.return_value
        builder.get.assert_awaited_once()
        assert builder.get.call_args.args == ()
        assert builder.get.call_args.kwargs == {}

    async def test_get_contact_returns_the_middle_name(self):
        """The write tools take `middle_name`, so the read returns it under the same name."""
        detail = await get_contact(
            _make_contact_by_id_mock(_make_mock_contact(middle_name="Quincy")), "contact123"
        )

        assert detail["middle_name"] == "Quincy"
        assert detail["first_name"] == "John"

    async def test_every_field_the_detail_formatter_reads_survives_the_round_trip(self):
        """A cheap structural echo of the guard the listing paths get for free."""
        detail = await get_contact(_make_contact_by_id_mock(_make_mock_contact()), "contact123")

        assert set(detail) == {
            "id",
            "first_name",
            "middle_name",
            "last_name",
            "display_name",
            "email_addresses",
            "mobile_phone",
            "home_phones",
            "business_phones",
            "company",
            "title",
            "job_title",
            "department",
            "birthday",
            "home_address",
            "business_address",
            "other_address",
            "categories",
            "personal_notes",
        }


@pytest.mark.asyncio
@pytest.mark.parametrize("tool", ["outlook_create_contact", "outlook_update_contact"])
async def test_title_is_described_as_the_name_prefix_where_the_model_reads_it(tool):
    """#119: an undescribed `title` gets filled with the job. The tool description is the
    only text a client sees, so both write tools must say what `title` is there."""
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    description = tools[tool].description
    assert "name prefix" in description
    assert "not the job" in description


_BIRTHDAY_STORED = "1990-05-01T11:59:00+00:00"


class TestBirthday:
    """`outlook_get_contact` has always returned `birthday`; nothing could write one."""

    async def test_create_writes_the_birthday(self):
        mock_client = MagicMock()
        mock_client.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))
        await create_contact(mock_client, first_name="Jane", birthday="1990-05-01", config=_CFG)
        posted = mock_client.me.contacts.post.call_args[0][0]
        assert posted.birthday.isoformat() == _BIRTHDAY_STORED

    async def test_update_writes_the_birthday(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", birthday="1990-05-01", config=_CFG
        )
        patched = mock_client.me.contacts.by_contact_id.return_value.patch.call_args[0][0]
        assert patched.birthday.isoformat() == _BIRTHDAY_STORED

    async def test_omitting_it_leaves_it_alone(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(mock_client, contact_id="contact123", first_name="Jane", config=_CFG)
        patched = mock_client.me.contacts.by_contact_id.return_value.patch.call_args[0][0]
        assert patched.birthday is None

    async def test_create_refuses_a_bad_date_before_anything_is_sent(self):
        mock_client = MagicMock()
        mock_client.me.contacts.post = AsyncMock()
        with pytest.raises(ValueError, match="YYYY-MM-DD"):
            await create_contact(mock_client, first_name="Jane", birthday="1 May", config=_CFG)
        mock_client.me.contacts.post.assert_not_called()

    async def test_update_refuses_a_bad_date_before_anything_is_sent(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ValueError, match="YYYY-MM-DD"):
            await update_contact(
                mock_client, contact_id="contact123", birthday="1990-13-01", config=_CFG
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()

    def test_what_the_write_stores_the_read_returns_as_that_date(self):
        """The read path returns Graph's instant; its date is the date written."""
        from outlook_mcp.validation import validate_birthday

        stored = validate_birthday("1990-05-01")
        detail = _format_contact_detail(_make_mock_contact(birthday=stored))
        assert detail["birthday"].startswith("1990-05-01")


class TestBirthdayThroughTheServer:
    """`server.py` passes the contact arguments positionally, so a `birthday` in its
    neighbour's slot (`phone_type` on create, `clear_phones` on update) would be
    refused there, or read as something else."""

    async def test_create_birthday_lands_in_birthday(self):
        from unittest.mock import patch

        from outlook_mcp import server as server_mod

        sdk = MagicMock()
        sdk.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))
        client = MagicMock()
        client.sdk_client = sdk
        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=_CFG),
        ):
            await server_mod.outlook_create_contact(
                MagicMock(), first_name="Jane", phone="+15555550100", birthday="1990-05-01"
            )
        posted = sdk.me.contacts.post.call_args[0][0]
        assert posted.birthday.isoformat() == _BIRTHDAY_STORED
        assert posted.mobile_phone == "+15555550100"

    async def test_update_birthday_lands_in_birthday(self):
        from unittest.mock import patch

        from outlook_mcp import server as server_mod

        sdk = _make_contact_by_id_mock(_make_mock_contact())
        client = MagicMock()
        client.sdk_client = sdk
        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=_CFG),
        ):
            await server_mod.outlook_update_contact(
                MagicMock(), contact_id="contact123", birthday="1990-05-01"
            )
        patched = sdk.me.contacts.by_contact_id.return_value.patch.call_args[0][0]
        assert patched.birthday.isoformat() == _BIRTHDAY_STORED
        assert patched.business_phones is None
        assert patched.home_phones is None


@pytest.mark.asyncio
@pytest.mark.parametrize("tool", ["outlook_create_contact", "outlook_update_contact"])
async def test_birthday_is_described_as_a_date_where_the_model_reads_it(tool):
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    assert "birthday" in tools[tool].input_schema["properties"]
    assert "YYYY-MM-DD" in tools[tool].description


#: A contact's change key in the shape Graph returns it in ``@odata.etag``.
_CONTACT_ETAG = 'W/"EQAAABYAAAA9sMfE2cIhRb3f6CTJkTSZAAD0dvmb"'


def _stored_addresses(*pairs):
    """Graph ``emailAddress`` entries, as (address, name) pairs."""
    entries = []
    for address, name in pairs:
        entry = MagicMock(spec=["address", "name"])
        entry.address = address
        entry.name = name
        entries.append(entry)
    return entries


def _contact_holding(*pairs, etag=_CONTACT_ETAG):
    """A contact read back from Graph with these addresses and this change key.

    ``additional_data`` is set explicitly: ``_make_mock_contact``'s spec leaves it
    out, and an auto-created MagicMock attribute would hand the code a change key
    no Graph response contains.
    """
    contact = _make_mock_contact(email_addresses=_stored_addresses(*pairs))
    contact.additional_data = {"@odata.etag": etag} if etag is not None else {}
    return contact


class TestAddAndRemoveEmails:
    """`email` replaces every address with one; these change only the ones named.

    The merge is computed from a read, so the PATCH sends back addresses this
    call did not write. By #87's rule a write that sends back what it read pins
    the version it read, so the patch carries `If-Match`.
    """

    @staticmethod
    def _builder(mock_client):
        return mock_client.me.contacts.by_contact_id.return_value

    def _sent(self, mock_client):
        """(address, name) for every address in the PATCH, in order."""
        builder = self._builder(mock_client)
        builder.patch.assert_called_once()
        sent = builder.patch.call_args[0][0].email_addresses
        return [(e.address, e.name) for e in sent]

    async def test_an_added_address_keeps_the_ones_already_there(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        await update_contact(
            mock_client, contact_id="contact123", add_emails=["jane@example.com"], config=_CFG
        )
        assert self._sent(mock_client) == [("jane@contoso.com", "Jane"), ("jane@example.com", None)]

    async def test_a_removed_address_takes_only_itself_out(self):
        mock_client = _make_contact_by_id_mock(
            _contact_holding(
                ("jane@contoso.com", "Jane"),
                ("old@example.com", "old@example.com"),
                ("jane@example.net", "Jane (home)"),
            )
        )
        await update_contact(
            mock_client, contact_id="contact123", remove_emails=["old@example.com"], config=_CFG
        )
        assert self._sent(mock_client) == [
            ("jane@contoso.com", "Jane"),
            ("jane@example.net", "Jane (home)"),
        ]

    async def test_add_and_remove_in_one_patch(self):
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("jane@contoso.com", "Jane"), ("old@example.com", None))
        )
        await update_contact(
            mock_client,
            contact_id="contact123",
            add_emails=["jane@example.com"],
            remove_emails=["old@example.com"],
            config=_CFG,
        )
        assert self._sent(mock_client) == [("jane@contoso.com", "Jane"), ("jane@example.com", None)]

    async def test_addresses_compare_without_case(self):
        """Mail addresses are case-insensitive in practice; a re-add in other case is
        the same address, and a remove in other case removes it."""
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("Jane@Contoso.com", "Jane"), ("Old@Example.com", None))
        )
        await update_contact(
            mock_client,
            contact_id="contact123",
            add_emails=["jane@contoso.com", "NEW@example.com", "new@example.com"],
            remove_emails=["OLD@example.COM"],
            config=_CFG,
        )
        assert self._sent(mock_client) == [("Jane@Contoso.com", "Jane"), ("NEW@example.com", None)]

    async def test_removing_the_last_address_sends_an_empty_list(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        await update_contact(
            mock_client, contact_id="contact123", remove_emails=["jane@contoso.com"], config=_CFG
        )
        assert self._sent(mock_client) == []

    async def test_a_remove_of_an_address_the_contact_lacks_is_refused(self):
        """A misspelt remove would otherwise report "updated" and leave the real
        address in place — #41's shape. The message names what is stored."""
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ValueError, match=r"no address jnae@contoso\.com.*jane@contoso\.com"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                remove_emails=["jnae@contoso.com"],
                config=_CFG,
            )
        self._builder(mock_client).patch.assert_not_called()

    async def test_more_than_three_addresses_is_refused(self):
        """Exchange keeps three email slots per contact; a fourth is refused here
        rather than sent to a store that cannot hold it."""
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("a@contoso.com", None), ("b@contoso.com", None))
        )
        with pytest.raises(ValueError, match="three"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                add_emails=["c@contoso.com", "d@contoso.com"],
                config=_CFG,
            )
        self._builder(mock_client).patch.assert_not_called()

    async def test_three_addresses_are_allowed(self):
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("a@contoso.com", None), ("b@contoso.com", None))
        )
        await update_contact(
            mock_client, contact_id="contact123", add_emails=["c@contoso.com"], config=_CFG
        )
        assert [a for a, _ in self._sent(mock_client)] == [
            "a@contoso.com",
            "b@contoso.com",
            "c@contoso.com",
        ]

    async def test_a_swap_at_the_limit_is_allowed(self):
        """The limit is on the result, not on the stored count plus the additions."""
        mock_client = _make_contact_by_id_mock(
            _contact_holding(
                ("a@contoso.com", None), ("b@contoso.com", None), ("c@contoso.com", None)
            )
        )
        await update_contact(
            mock_client,
            contact_id="contact123",
            add_emails=["d@contoso.com"],
            remove_emails=["a@contoso.com"],
            config=_CFG,
        )
        assert [a for a, _ in self._sent(mock_client)] == [
            "b@contoso.com",
            "c@contoso.com",
            "d@contoso.com",
        ]

    async def test_the_patch_is_pinned_to_the_version_it_read(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        await update_contact(
            mock_client, contact_id="contact123", add_emails=["jane@example.com"], config=_CFG
        )
        builder = self._builder(mock_client)
        # The header itself: a request configuration carrying nothing, or another
        # change key, would pass a weaker assertion.
        config_sent = builder.patch.call_args.kwargs["request_configuration"]
        assert config_sent.headers.get("If-Match") == {_CONTACT_ETAG}

    async def test_without_a_change_key_the_patch_goes_unpinned(self):
        """#87's rule for a missing key: send unpinned, as every other write does."""
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("jane@contoso.com", "Jane"), etag=None)
        )
        await update_contact(
            mock_client, contact_id="contact123", add_emails=["jane@example.com"], config=_CFG
        )
        builder = self._builder(mock_client)
        assert "request_configuration" not in builder.patch.call_args.kwargs
        assert self._sent(mock_client)[-1] == ("jane@example.com", None)

    async def test_a_patch_without_them_reads_nothing_and_is_not_pinned(self):
        """The control: every other argument is the caller's value, last-writer-wins."""
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        await update_contact(mock_client, contact_id="contact123", first_name="Janet", config=_CFG)
        builder = self._builder(mock_client)
        builder.get.assert_not_called()
        assert "request_configuration" not in builder.patch.call_args.kwargs
        assert builder.patch.call_args[0][0].email_addresses is None

    async def test_the_other_fields_travel_in_the_same_pinned_patch(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        await update_contact(
            mock_client,
            contact_id="contact123",
            job_title="Engineer",
            add_emails=["jane@example.com"],
            config=_CFG,
        )
        builder = self._builder(mock_client)
        assert builder.patch.call_args[0][0].job_title == "Engineer"
        assert builder.patch.call_args.kwargs["request_configuration"].headers.get("If-Match")

    @pytest.mark.parametrize(
        "lists",
        [
            {"add_emails": ["jane@example.com"]},
            {"remove_emails": ["jane@contoso.com"]},
        ],
    )
    async def test_email_with_either_list_is_refused_before_anything_is_read(self, lists):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ValueError, match="not both"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                email="jane@example.org",
                config=_CFG,
                **lists,
            )
        self._builder(mock_client).get.assert_not_called()
        self._builder(mock_client).patch.assert_not_called()

    async def test_one_address_in_both_lists_is_refused(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ValueError, match="both add_emails and remove_emails"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                add_emails=["Jane@contoso.com"],
                remove_emails=["jane@contoso.com"],
                config=_CFG,
            )
        self._builder(mock_client).get.assert_not_called()

    @pytest.mark.parametrize("name", ["add_emails", "remove_emails"])
    async def test_an_empty_list_is_refused(self, name):
        """A list naming nothing is a patch that carries nothing and says "updated"."""
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ValueError, match=f"{name} is empty"):
            await update_contact(mock_client, contact_id="contact123", config=_CFG, **{name: []})
        self._builder(mock_client).get.assert_not_called()

    @pytest.mark.parametrize("name", ["add_emails", "remove_emails"])
    async def test_every_address_is_validated_before_anything_is_read(self, name):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ValueError, match="Invalid email address"):
            await update_contact(
                mock_client,
                contact_id="contact123",
                config=_CFG,
                **{name: ["jane@example.com", "not an address"]},
            )
        self._builder(mock_client).get.assert_not_called()

    async def test_read_only_mode_refuses_before_reading(self):
        mock_client = _make_contact_by_id_mock(_contact_holding(("jane@contoso.com", "Jane")))
        with pytest.raises(ReadOnlyError):
            await update_contact(
                mock_client,
                contact_id="contact123",
                add_emails=["jane@example.com"],
                config=_CFG_RO,
            )
        self._builder(mock_client).get.assert_not_called()

    async def test_email_still_replaces_every_address(self):
        """Today's behaviour, unchanged: one address in, the rest gone, nothing read."""
        mock_client = _make_contact_by_id_mock(
            _contact_holding(("jane@contoso.com", "Jane"), ("old@example.com", None))
        )
        await update_contact(
            mock_client, contact_id="contact123", email="jane@example.org", config=_CFG
        )
        assert [a for a, _ in self._sent(mock_client)] == ["jane@example.org"]
        self._builder(mock_client).get.assert_not_called()


class TestAddAndRemoveEmailsThroughTheServer:
    """`server.py` hands `update_contact` its arguments positionally, and the two
    lists have the same type, so a swap would add what was meant to be removed.
    These run the real handler behind the real wrapper, one list at a time."""

    @staticmethod
    async def _call(**arguments):
        from unittest.mock import patch

        from outlook_mcp import server as server_mod

        sdk = _make_contact_by_id_mock(
            _contact_holding(("jane@contoso.com", "Jane"), ("old@example.com", None))
        )
        client = MagicMock()
        client.sdk_client = sdk
        ctx = MagicMock()
        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=_CFG),
        ):
            await server_mod.outlook_update_contact(ctx, contact_id="contact123", **arguments)
        builder = sdk.me.contacts.by_contact_id.return_value
        return [e.address for e in builder.patch.call_args[0][0].email_addresses]

    async def test_add_emails_adds(self):
        assert await self._call(add_emails=["jane@example.com"]) == [
            "jane@contoso.com",
            "old@example.com",
            "jane@example.com",
        ]

    async def test_remove_emails_removes(self):
        assert await self._call(remove_emails=["old@example.com"]) == ["jane@contoso.com"]


@pytest.mark.asyncio
async def test_update_says_email_replaces_and_the_lists_merge_where_the_model_reads_it():
    """The replace-everything behaviour of `email` is silent data loss for a caller who
    meant "add"; the tool description is the only text a client sees."""
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    description = tools["outlook_update_contact"].description
    assert "add_emails" in description
    assert "remove_emails" in description
    assert "REPLACES" in description


_NEXT_PAGE = "https://graph.microsoft.com/v1.0/me/contacts?$top=1000&$select=id&$skip=1000"


def _with_phones(display_name, mobile="", business=(), home=()):
    return _make_mock_contact(
        id=f"id-{display_name}",
        display_name=display_name,
        mobile_phone=mobile,
        business_phones=list(business),
        home_phones=list(home),
    )


def _paged_contacts_mock(*pages):
    """A contacts listing that pages: each page but the last carries a nextLink,
    followed through `with_url` as the other nextLink walks here do."""
    responses = [
        MagicMock(value=list(page), odata_next_link=_NEXT_PAGE if i < len(pages) - 1 else None)
        for i, page in enumerate(pages)
    ]
    client = MagicMock()
    client.me.contacts.get = AsyncMock(return_value=responses[0])
    follow = MagicMock()
    follow.get = AsyncMock(side_effect=responses[1:] or [None])
    client.me.contacts.with_url = MagicMock(return_value=follow)
    return client


class TestPhoneSearch:
    """Graph's contact `$search` finds a number only as it is stored.

    Checked against an Exchange Online mailbox: a whole number as stored is found,
    the same number without its country code is not, and neither is a part of one.
    Graph documents no contact property for `$search` and lets `$filter` reach only
    `emailAddresses/any(a:a/address eq ...)`. So `phone` reads the listing and
    compares digits: a number matches when the digits of the shorter of the two end
    the longer's, as written or without leading zeros (a `00` prefix, a trunk 0),
    and the shorter has at least seven.
    """

    @staticmethod
    def _found(result):
        return [c["display_name"] for c in result["contacts"]]

    async def test_a_number_without_its_country_code_finds_the_international_one(self):
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+44 7700 900123")])
        result = await search_contacts(client, phone="07700 900123")
        assert self._found(result) == ["Jane"]

    async def test_an_international_number_finds_a_national_one_with_a_trunk_zero(self):
        client = _paged_contacts_mock([_with_phones("Jane", mobile="020 7946 0018")])
        result = await search_contacts(client, phone="+44 20 7946 0018")
        assert self._found(result) == ["Jane"]

    async def test_a_00_prefix_is_the_same_as_a_plus(self):
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+1 (555) 010-0199")])
        result = await search_contacts(client, phone="001 555 010 0199")
        assert self._found(result) == ["Jane"]

    async def test_business_and_home_numbers_are_compared_too(self):
        client = _paged_contacts_mock(
            [
                _with_phones("Office", business=["+1 555 010 0100", "+1 555 010 0101"]),
                _with_phones("Home", home=["+1 555 010 0102"]),
                _with_phones("Other", mobile="+1 555 010 0103"),
            ]
        )
        assert self._found(await search_contacts(client, phone="5550100101")) == ["Office"]
        client = _paged_contacts_mock([_with_phones("Home", home=["+1 555 010 0102"])])
        assert self._found(await search_contacts(client, phone="555-010-0102")) == ["Home"]

    async def test_a_different_number_is_not_found(self):
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+44 7700 900123")])
        result = await search_contacts(client, phone="07700 900124")
        assert result["contacts"] == []
        assert result["count"] == 0
        assert result["has_more"] is False

    async def test_digits_inside_a_number_do_not_match(self):
        """Only the end of a number identifies it; seven digits from its middle are
        the end of nobody's number."""
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+1 555 010 0199")])
        result = await search_contacts(client, phone="555 0100")
        assert result["contacts"] == []

    async def test_a_short_stored_number_never_matches(self):
        """An extension or a short code ends many long numbers; seven digits is the floor."""
        client = _paged_contacts_mock([_with_phones("Desk", business=["4567"])])
        result = await search_contacts(client, phone="+1 555 010 4567")
        assert result["contacts"] == []

    async def test_seven_digits_match_and_six_are_refused(self):
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+1 555 123 4567")])
        assert self._found(await search_contacts(client, phone="123 4567")) == ["Jane"]
        with pytest.raises(ValueError, match="at least 7 digits"):
            await search_contacts(client, phone="23 4567")

    async def test_a_short_number_is_refused_before_anything_is_read(self):
        client = _paged_contacts_mock([])
        with pytest.raises(ValueError, match="at least 7 digits"):
            await search_contacts(client, phone="+1 23 45")
        client.me.contacts.get.assert_not_called()

    async def test_the_last_digits_find_a_number_even_when_they_start_with_zero(self):
        """Leading zeros are ignored only as a second try: stripped first, `010 0199`
        would be six digits and refused, though it ends the stored number."""
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+1 555 010 0199")])
        result = await search_contacts(client, phone="010 0199")
        assert self._found(result) == ["Jane"]

    async def test_a_number_with_other_characters_is_refused(self):
        client = _paged_contacts_mock([])
        with pytest.raises(ValueError, match="Invalid phone number"):
            await search_contacts(client, phone="call 555 010 0199")
        client.me.contacts.get.assert_not_called()

    async def test_query_and_phone_together_are_refused(self):
        client = _paged_contacts_mock([])
        with pytest.raises(ValueError, match="query or phone"):
            await search_contacts(client, query="Jane", phone="+1 555 010 0199")
        client.me.contacts.get.assert_not_called()

    async def test_neither_is_refused(self):
        client = _paged_contacts_mock([])
        with pytest.raises(ValueError, match="query or phone"):
            await search_contacts(client)
        client.me.contacts.get.assert_not_called()

    async def test_the_listing_is_read_without_search_in_large_pages(self):
        client = _paged_contacts_mock([])
        await search_contacts(client, phone="+1 555 010 0199")
        qp = client.me.contacts.get.call_args.kwargs["request_configuration"].query_parameters
        assert qp.search is None
        assert qp.top == 1000
        assert ",".join(qp.select) == _SUMMARY_SELECT

    async def test_every_page_is_followed_by_its_next_link(self):
        client = _paged_contacts_mock(
            [_with_phones("First", mobile="+1 555 010 0100")],
            [_with_phones("Second", mobile="+1 555 010 0199")],
        )
        result = await search_contacts(client, phone="5550100199")
        assert self._found(result) == ["Second"]
        client.me.contacts.with_url.assert_called_once_with(_NEXT_PAGE)
        assert result["has_more"] is False

    async def test_the_scan_stops_at_its_cap_and_says_the_rest_was_not_read(self):
        from outlook_mcp.tools import contacts as contacts_mod

        endless = MagicMock(value=[_with_phones("Jane", mobile="+1 555 010 0199")])
        endless.odata_next_link = _NEXT_PAGE
        client = MagicMock()
        client.me.contacts.get = AsyncMock(return_value=endless)
        follow = MagicMock()
        follow.get = AsyncMock(return_value=endless)
        client.me.contacts.with_url = MagicMock(return_value=follow)

        result = await search_contacts(client, phone="5550100199", count=100)

        assert follow.get.await_count == contacts_mod._PHONE_SCAN_PAGES - 1
        assert result["count"] == contacts_mod._PHONE_SCAN_PAGES
        # No cursor to hand back, so `has_more` alone says the scan was cut short.
        assert result["has_more"] is True
        assert result["cursor"] is None

    async def test_count_caps_the_matches_returned(self):
        client = _paged_contacts_mock(
            [
                _with_phones("A", mobile="+1 555 010 0199"),
                _with_phones("B", business=["555 010 0199"]),
                _with_phones("C", home=["15550100199"]),
            ]
        )
        result = await search_contacts(client, phone="+1 555 010 0199", count=2)
        assert self._found(result) == ["A", "B"]
        assert result["count"] == 2
        assert result["has_more"] is True
        assert result["cursor"] is None

    async def test_matches_have_the_search_shape(self):
        """The same rows a `query` search returns: no `categories`, since the
        search path never claims to know them."""
        client = _paged_contacts_mock([_with_phones("Jane", mobile="+44 7700 900123")])
        result = await search_contacts(client, phone="7700900123")
        assert set(result["contacts"][0]) == {"id", "display_name", "email", "phone", "company"}

    async def test_a_query_search_is_unchanged(self):
        client = _make_contacts_mock([_make_mock_contact()])
        result = await search_contacts(client, query="John")
        qp = client.me.contacts.get.call_args.kwargs["request_configuration"].query_parameters
        assert qp.search == '"John"'
        assert result["count"] == 1


class TestPhoneSearchThroughTheServer:
    """`server.py` passes `query`, `count` and `phone` positionally."""

    async def test_phone_reaches_the_phone_search(self):
        from unittest.mock import patch

        from outlook_mcp import server as server_mod

        sdk = _paged_contacts_mock([_with_phones("Jane", mobile="+44 7700 900123")])
        client = MagicMock()
        client.sdk_client = sdk
        with patch.object(server_mod, "_get_graph_client", return_value=client):
            result = await server_mod.outlook_search_contacts(MagicMock(), phone="07700 900123")
        assert [c["display_name"] for c in result["contacts"]] == ["Jane"]

    async def test_count_still_reaches_count(self):
        from unittest.mock import patch

        from outlook_mcp import server as server_mod

        sdk = _make_contacts_mock([])
        client = MagicMock()
        client.sdk_client = sdk
        with patch.object(server_mod, "_get_graph_client", return_value=client):
            await server_mod.outlook_search_contacts(MagicMock(), query="John", count=7)
        qp = sdk.me.contacts.get.call_args.kwargs["request_configuration"].query_parameters
        assert qp.top == 7


@pytest.mark.asyncio
async def test_phone_search_is_described_where_the_model_reads_it():
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    tool = tools["outlook_search_contacts"]
    assert "phone" in tool.input_schema["properties"]
    assert "query" not in tool.input_schema.get("required", [])
    assert "phone" in tool.description
