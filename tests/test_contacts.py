"""Tests for contact tools."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import ReadOnlyError
from outlook_mcp.pagination import encode_cursor
from outlook_mcp.tools.contacts import (
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

    return contact


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

        # Verify the search param was sanitized (quotes/parens stripped)
        call_kwargs = mock_client.me.contacts.get.call_args
        qp = call_kwargs.kwargs["request_configuration"].query_parameters
        assert '"' not in qp.search.strip('"')
        assert "(" not in qp.search.strip('"')

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
