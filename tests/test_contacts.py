"""Tests for contact tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from kiota_abstractions.store import BackingStoreSerializationWriterProxyFactory
from kiota_serialization_json.json_serialization_writer_factory import (
    JsonSerializationWriterFactory,
)

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import encode_cursor
from outlook_mcp.tools.contacts import (
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

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)


def _make_mock_contact(**overrides):
    """Factory for mock Graph SDK contact objects.

    Matches the consumer Outlook contact shape: ``mobile_phone`` (single
    string), ``home_phones`` (list[str]), ``business_phones`` (list[str]).
    """
    contact = MagicMock(spec=[
        "id", "display_name", "given_name", "surname",
        "company_name", "title", "department", "birthday",
        "email_addresses", "mobile_phone", "home_phones", "business_phones",
        "home_address", "business_address", "other_address",
        "categories", "personal_notes",
    ])
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
            mobile_phone="", home_phones=[], business_phones=["+15553334444"],
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
            mobile_phone="", home_phones=[], business_phones=[],
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
            mock_client, first_name="John", phone="+1234567890", config=_CFG,
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
                mock_client, first_name="John", email="not-an-email", config=_CFG,
            )

    async def test_create_validates_phone(self):
        """create_contact rejects invalid phone number."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid phone"):
            await create_contact(
                mock_client, first_name="John", phone="not a phone!!!", config=_CFG,
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
            mock_client, contact_id="contact123", first_name="Jane", config=_CFG,
        )
        assert result["status"] == "updated"

        # Verify patch was called
        contact_obj = mock_client.me.contacts.by_contact_id.return_value
        contact_obj.patch.assert_called_once()

    async def test_update_writes_mobile_phone_not_phones(self):
        """update_contact must set mobile_phone (not 'phones') on consumer Graph."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())

        await update_contact(
            mock_client, contact_id="contact123", phone="+19998887777", config=_CFG,
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
                mock_client, contact_id="bad id!", first_name="Jane", config=_CFG,
            )

    async def test_update_raises_read_only(self):
        """update_contact raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await update_contact(
                mock_client, contact_id="contact123", first_name="Jane", config=_CFG_RO,
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

    def test_address_content_is_sanitized_like_every_other_echoed_field(self):
        """A contact is attacker-influenced text; sanitize_output strips ANSI and
        control characters from it, exactly as it does for every other field here."""
        detail = self._detail(
            home_address=_make_mock_address(street="693 7th[31m St S")
        )
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


class TestUpdateContactWritesTheAddressItCanRead:
    """The write-side twin of TestContactDetailCarriesEverythingStored.

    A home address that `get_contact` returns but `update_contact` cannot set is
    only half a fix: the field is readable and not correctable. These pin that
    the two sides speak the same five fields.
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
            home_street="2821 252nd Ave SE",
            home_city="Sammamish",
            home_state="WA",
            home_postal_code="98075",
            home_country="USA",
            config=_CFG,
        )
        address = self._patched(mock_client).home_address
        assert address.street == "2821 252nd Ave SE"
        assert address.city == "Sammamish"
        assert address.state == "WA"
        assert address.postal_code == "98075"
        assert address.country_or_region == "USA"

    async def test_a_partial_address_is_still_written(self):
        """Someone may know the city and not the street.

        Graph replaces the whole address rather than merging, so the parts not
        supplied come back empty — verified live. That belongs in the docstring
        the model reads, not in a silent difference between what was asked for
        and what was stored.
        """
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_city="Bothell", config=_CFG,
        )
        address = self._patched(mock_client).home_address
        assert address.city == "Bothell"
        assert address.street is None

    async def test_omitting_every_part_leaves_the_address_alone(self):
        """Partial patch: a name-only update must not blank a stored address."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", first_name="Jane", config=_CFG,
        )
        assert self._patched(mock_client).home_address is None

    async def test_whitespace_only_parts_count_as_omitted(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_street="   ", config=_CFG,
        )
        assert self._patched(mock_client).home_address is None

    async def test_values_are_stripped(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_city="  Kirkland  ", config=_CFG,
        )
        assert self._patched(mock_client).home_address.city == "Kirkland"

    async def test_read_and_write_agree_on_the_same_five_fields(self):
        """Round-trip: what _build_address writes, _format_address reads back."""
        parts = {
            "street": "693 7th St S",
            "city": "Kirkland",
            "state": "WA",
            "postal_code": "98033",
            "country_or_region": "USA",
        }
        built = _build_address(
            parts["street"], parts["city"], parts["state"],
            parts["postal_code"], parts["country_or_region"],
        )
        assert _format_address(built) == parts

    async def test_read_only_mode_refuses_the_write(self):
        """The new parameters must not open a path around the read-only gate."""
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        with pytest.raises(ReadOnlyError):
            await update_contact(
                mock_client, contact_id="contact123", home_city="Bothell", config=_CFG_RO,
            )
        mock_client.me.contacts.by_contact_id.return_value.patch.assert_not_called()


def _adapter_wire(model) -> str:
    """The JSON the Graph request adapter would actually send for this model.

    Deliberately not a bare ``JsonSerializationWriter``: the adapter serializes
    through the backing-store proxy, and only that path emits a field that was
    explicitly assigned ``None``. ``tests/test_write_payloads_reach_the_wire.py``
    asserts values are *present*, which the bare writer answers correctly; this
    helper exists for the opposite question — what got in that we never asked for.
    """
    factory = BackingStoreSerializationWriterProxyFactory(JsonSerializationWriterFactory())
    writer = factory.get_serialization_writer("application/json")
    writer.write_object_value(None, model)
    return writer.get_serialized_content().decode()


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
    so the tool worked for the case anyone would test by hand and failed for
    the one it was written for. Only the live tier and this serializer see it.
    """

    async def test_a_city_only_patch_sends_the_city_and_nothing_else(self):
        mock_client = _make_contact_by_id_mock(_make_mock_contact())
        await update_contact(
            mock_client, contact_id="contact123", home_city="Bothell", config=_CFG,
        )
        body = _adapter_wire(
            mock_client.me.contacts.by_contact_id.return_value.patch.call_args[0][0]
        )

        assert '"homeAddress": {"city": "Bothell"}' in body
        assert "null" not in body, f"a part we were never given reached the wire: {body}"
        for python_name in ("country_or_region", "postal_code"):
            assert python_name not in body

    async def test_the_serializer_this_uses_is_the_one_that_can_see_it(self):
        """Pin why the helper is not a plain JsonSerializationWriter.

        If kiota ever stops emitting explicitly-``None`` fields, this fails and
        the distinction above can be dropped.
        """
        from msgraph.generated.models.contact import Contact
        from msgraph.generated.models.physical_address import PhysicalAddress

        address = PhysicalAddress()
        address.city = "Bothell"
        address.street = None
        contact = Contact()
        contact.home_address = address

        assert '"street": null' in _adapter_wire(contact)

