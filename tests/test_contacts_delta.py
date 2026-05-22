"""Tests for contacts delta tool."""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.tools.contacts_delta import _format_contact_delta, list_contacts_delta

# ── Helpers ──────────────────────────────────────────────────────────


def _raw_contact(**overrides) -> dict:
    base = {
        "id": "CONT_AAA",
        "displayName": "Alice Example",
        "emailAddresses": [
            {"address": "alice@example.com", "name": "Alice"},
        ],
        "mobilePhone": "555-1212",
        "homePhones": [],
        "businessPhones": [],
        "companyName": "Example Inc",
    }
    base.update(overrides)
    return base


def _mock_graph_client():
    client = MagicMock()
    client.credential = MagicMock()
    return client


def _http_response(body: dict, status: int = 200):
    r = MagicMock()
    r.status_code = status
    r.json = MagicMock(return_value=body)
    r.raise_for_status = MagicMock()
    return r


def _async_client_with(responses):
    responses = list(responses)

    async def fake_get(url, headers=None):
        fake_get.last_call = {"url": url, "headers": headers}
        return responses.pop(0)

    fake_get.last_call = None

    fake_client = MagicMock()
    fake_client.get = AsyncMock(side_effect=fake_get)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    return (
        patch(
            "outlook_mcp.tools._delta.httpx.AsyncClient",
            return_value=fake_client,
        ),
        fake_client,
        fake_get,
    )


# ── Formatter ────────────────────────────────────────────────────────


class TestFormatContactDelta:
    def test_maps_basic_fields(self):
        out = _format_contact_delta(_raw_contact())
        assert out["id"] == "CONT_AAA"
        assert out["display_name"] == "Alice Example"
        assert out["email"] == "alice@example.com"
        assert out["phone"] == "555-1212"
        assert out["company"] == "Example Inc"

    def test_prefers_mobile_then_home_then_business(self):
        out = _format_contact_delta(
            _raw_contact(mobilePhone="", homePhones=["111"], businessPhones=["222"])
        )
        assert out["phone"] == "111"
        out2 = _format_contact_delta(
            _raw_contact(mobilePhone="", homePhones=[], businessPhones=["222"])
        )
        assert out2["phone"] == "222"

    def test_missing_emails(self):
        out = _format_contact_delta(_raw_contact(emailAddresses=[]))
        assert out["email"] == ""


# ── First call: no params allowed ─────────────────────────────────────


@pytest.mark.asyncio
async def test_first_call_uses_prefer_header_and_no_query_params():
    body = {
        "value": [_raw_contact(id="c1"), _raw_contact(id="c2")],
        "@odata.deltaLink": "https://graph/contacts-delta",
    }
    patch_client, _, fake_get = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_contacts_delta(_mock_graph_client(), page_size=25)

    assert len(result["contacts"]) == 2
    assert result["delta_token"] == "https://graph/contacts-delta"
    assert result["has_more"] is False

    # No query string on the contacts/delta URL
    url = fake_get.last_call["url"]
    assert "?" not in url, f"expected bare URL, got {url}"

    # maxpagesize must be in Prefer
    headers = fake_get.last_call["headers"]
    assert headers.get("Prefer") == "odata.maxpagesize=25"


# ── Subsequent call ─────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_subsequent_call_uses_delta_token_url_verbatim():
    body = {
        "value": [_raw_contact(id="changed")],
        "@odata.deltaLink": "https://graph/contacts-delta-new",
    }
    prior = "https://graph/contacts-delta-prior"
    patch_client, _, fake_get = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_contacts_delta(_mock_graph_client(), delta_token=prior)

    assert fake_get.last_call["url"] == prior
    assert result["delta_token"] == "https://graph/contacts-delta-new"


# ── Safety cap ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_cap_reached_returns_nextlink_and_has_more():
    page = lambda i, link: {  # noqa: E731
        "value": [_raw_contact(id=f"c{i*50 + n}") for n in range(50)],
        "@odata.nextLink": link,
    }
    responses = [
        _http_response(page(0, "https://graph/p2")),
        _http_response(page(1, "https://graph/p3")),
        _http_response(page(2, "https://graph/p4")),
        _http_response(page(3, "https://graph/p5")),
    ]
    patch_client, _, _ = _async_client_with(responses)
    with patch_client:
        result = await list_contacts_delta(_mock_graph_client(), page_size=50)

    assert len(result["contacts"]) == 200
    assert result["has_more"] is True
    assert result["delta_token"] == "https://graph/p5"


# ── Final page ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_follows_nextlink_until_deltalink():
    responses = [
        _http_response({
            "value": [_raw_contact(id="c1")],
            "@odata.nextLink": "https://graph/p2",
        }),
        _http_response({
            "value": [_raw_contact(id="c2")],
            "@odata.deltaLink": "https://graph/contacts-delta-final",
        }),
    ]
    patch_client, _, _ = _async_client_with(responses)
    with patch_client:
        result = await list_contacts_delta(_mock_graph_client())

    assert len(result["contacts"]) == 2
    assert result["delta_token"] == "https://graph/contacts-delta-final"
    assert result["has_more"] is False


# ── Tombstones ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_removed_contact_collapses_to_id_only():
    body = {
        "value": [
            _raw_contact(id="c1"),
            {"id": "c2-deleted", "@removed": {"reason": "deleted"}},
        ],
        "@odata.deltaLink": "https://graph/contacts-delta",
    }
    patch_client, _, _ = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_contacts_delta(_mock_graph_client())

    tomb = result["contacts"][1]
    assert tomb == {"id": "c2-deleted", "is_deleted": True}
    assert set(tomb.keys()) == {"id", "is_deleted"}
    assert result["contacts"][0]["is_deleted"] is False


# ── No-changes case ─────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_no_changes_returns_empty_list_and_delta_token():
    body = {"value": [], "@odata.deltaLink": "https://graph/contacts-delta-same"}
    patch_client, _, _ = _async_client_with([_http_response(body)])
    with patch_client:
        result = await list_contacts_delta(
            _mock_graph_client(),
            delta_token="https://graph/contacts-delta-prior",
        )

    assert result["contacts"] == []
    assert result["delta_token"] == "https://graph/contacts-delta-same"
    assert result["has_more"] is False
