"""Tests for contact tools."""

import json
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import encode_cursor
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
            "surname",
            "company_name",
            "title",
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
    contact.surname = overrides.get("surname", "Doe")
    contact.company_name = overrides.get("company_name", "Acme")
    contact.title = overrides.get("title", "Engineer")
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
        assert result["title"] == "Engineer"
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

    async def test_every_field_the_detail_formatter_reads_survives_the_round_trip(self):
        """A cheap structural echo of the guard the listing paths get for free."""
        detail = await get_contact(_make_contact_by_id_mock(_make_mock_contact()), "contact123")

        assert set(detail) == {
            "id",
            "first_name",
            "last_name",
            "display_name",
            "email_addresses",
            "mobile_phone",
            "home_phones",
            "business_phones",
            "company",
            "title",
            "department",
            "birthday",
            "home_address",
            "business_address",
            "other_address",
            "categories",
            "personal_notes",
        }
