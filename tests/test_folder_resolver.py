"""Tests for folder_resolver.resolve_folder_id."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.folder_resolver import resolve_folder_id


def _mock_folder(folder_id: str, display_name: str) -> MagicMock:
    folder = MagicMock()
    folder.id = folder_id
    folder.display_name = display_name
    return folder


def _mock_list_response(folders: list[MagicMock]) -> MagicMock:
    response = MagicMock()
    response.value = folders
    return response


class TestCanonicalWellKnown:
    """Canonical well-known names short-circuit without a Graph call."""

    @pytest.mark.parametrize(
        "name",
        ["inbox", "drafts", "sentitems", "deleteditems", "junkemail", "archive", "outbox"],
    )
    async def test_canonical_passthrough(self, name):
        mock_client = MagicMock()
        result = await resolve_folder_id(mock_client, name)
        assert result == name
        mock_client.me.mail_folders.get.assert_not_called()

    async def test_canonical_case_insensitive(self):
        mock_client = MagicMock()
        result = await resolve_folder_id(mock_client, "INBOX")
        assert result == "inbox"
        mock_client.me.mail_folders.get.assert_not_called()


class TestWellKnownDisplayAliases:
    """Display forms like 'Junk Email' and 'Sent Items' resolve to canonical names."""

    @pytest.mark.parametrize(
        "input_name, expected",
        [
            ("Junk Email", "junkemail"),
            ("junk email", "junkemail"),
            ("JUNK EMAIL", "junkemail"),
            ("Sent Items", "sentitems"),
            ("sent items", "sentitems"),
            ("Deleted Items", "deleteditems"),
            ("Inbox", "inbox"),
            ("Drafts", "drafts"),
            ("Archive", "archive"),
            ("Outbox", "outbox"),
        ],
    )
    async def test_display_alias_to_canonical(self, input_name, expected):
        mock_client = MagicMock()
        result = await resolve_folder_id(mock_client, input_name)
        assert result == expected
        mock_client.me.mail_folders.get.assert_not_called()


class TestGraphIdPassthrough:
    """Valid Graph IDs pass through without a lookup."""

    async def test_graph_id_returned_as_is(self):
        mock_client = MagicMock()
        graph_id = "AAMkAGVmMDEzMTM4LTZmYWUtNGY1ZC1iZjRkLTc3YmMxY2U5YjBhNgAuAAAAAADi"
        result = await resolve_folder_id(mock_client, graph_id)
        assert result == graph_id
        mock_client.me.mail_folders.get.assert_not_called()


class TestDisplayNameLookup:
    """User-created folder display names trigger a Graph lookup."""

    async def test_single_match_resolves_to_id(self):
        tldr = _mock_folder("AAMkAG_tldr_id=", "TLDR")
        inbox = _mock_folder("AAMkAG_inbox=", "Inbox")
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([inbox, tldr]))

        result = await resolve_folder_id(mock_client, "TLDR")

        assert result == "AAMkAG_tldr_id="
        mock_client.me.mail_folders.get.assert_called_once()

    async def test_case_insensitive_match(self):
        tldr = _mock_folder("AAMkAG_tldr=", "TLDR Product")
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([tldr]))

        result = await resolve_folder_id(mock_client, "tldr product")

        assert result == "AAMkAG_tldr="

    async def test_no_match_raises_helpful_error(self):
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([_mock_folder("id1", "Receipts")])
        )

        with pytest.raises(ValueError, match="not found.*outlook_list_folders"):
            await resolve_folder_id(mock_client, "NonexistentFolder")

    async def test_ambiguous_match_raises(self):
        """Two top-level folders with the same display name → error, not silent pick."""
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response(
                [
                    _mock_folder("id_a", "Archive Old"),
                    _mock_folder("id_b", "Archive Old"),
                ]
            )
        )

        with pytest.raises(ValueError, match="ambiguous"):
            await resolve_folder_id(mock_client, "Archive Old")

    async def test_empty_folder_list(self):
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([]))

        with pytest.raises(ValueError, match="not found"):
            await resolve_folder_id(mock_client, "Anything")


class TestEdgeCases:
    async def test_empty_string_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="must not be empty"):
            await resolve_folder_id(mock_client, "")

    async def test_whitespace_only_raises(self):
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="must not be empty"):
            await resolve_folder_id(mock_client, "   ")

    async def test_leading_trailing_whitespace_stripped(self):
        mock_client = MagicMock()
        result = await resolve_folder_id(mock_client, "  Inbox  ")
        assert result == "inbox"
