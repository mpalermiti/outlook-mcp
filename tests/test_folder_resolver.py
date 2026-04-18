"""Tests for folder_resolver.resolve_folder_id."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.folder_resolver import resolve_folder_id


def _mock_folder(
    folder_id: str,
    display_name: str,
    *,
    child_count: int = 0,
) -> MagicMock:
    folder = MagicMock()
    folder.id = folder_id
    folder.display_name = display_name
    folder.child_folder_count = child_count
    return folder


def _setup_child_folders(mock_client: MagicMock, children_by_parent: dict) -> None:
    """Wire mock_client.me.mail_folders.by_mail_folder_id(pid).child_folders.get()."""
    def _by_id(parent_id: str):
        builder = MagicMock()
        kids = children_by_parent.get(parent_id, [])
        builder.child_folders.get = AsyncMock(return_value=_mock_list_response(kids))
        return builder
    mock_client.me.mail_folders.by_mail_folder_id = MagicMock(side_effect=_by_id)


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

        with pytest.raises(ValueError, match="not found"):
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


class TestSubfolderLookup:
    """Subfolder names resolve via BFS walk when no top-level match exists."""

    async def test_single_subfolder_match_resolves(self):
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        domains = _mock_folder("domains_id", "Domains")
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([receipts])
        )
        _setup_child_folders(mock_client, {"receipts_id": [domains]})

        result = await resolve_folder_id(mock_client, "Domains")

        assert result == "domains_id"

    async def test_nested_subfolder_match_resolves(self):
        """Folder nested two levels deep still resolves via BFS."""
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        services = _mock_folder("services_id", "Services", child_count=1)
        domains = _mock_folder("domains_id", "Domains")
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([receipts])
        )
        _setup_child_folders(
            mock_client,
            {"receipts_id": [services], "services_id": [domains]},
        )

        result = await resolve_folder_id(mock_client, "Domains")

        assert result == "domains_id"

    async def test_top_level_match_wins_over_subfolder(self):
        """Top-level match short-circuits before walking subfolders."""
        top_tldr = _mock_folder("top_tldr_id", "TLDR")
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([top_tldr, receipts])
        )
        _setup_child_folders(mock_client, {"receipts_id": [_mock_folder("nested_tldr", "TLDR")]})

        result = await resolve_folder_id(mock_client, "TLDR")

        assert result == "top_tldr_id"
        mock_client.me.mail_folders.by_mail_folder_id.assert_not_called()

    async def test_ambiguous_subfolder_matches_raise(self):
        a = _mock_folder("a_id", "A", child_count=1)
        b = _mock_folder("b_id", "B", child_count=1)
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([a, b])
        )
        _setup_child_folders(
            mock_client,
            {
                "a_id": [_mock_folder("dom_a", "Domains")],
                "b_id": [_mock_folder("dom_b", "Domains")],
            },
        )

        with pytest.raises(ValueError, match="ambiguous.*across tree"):
            await resolve_folder_id(mock_client, "Domains")

    async def test_not_found_after_full_walk(self):
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(
            return_value=_mock_list_response([receipts])
        )
        _setup_child_folders(
            mock_client,
            {"receipts_id": [_mock_folder("bills_id", "Bills")]},
        )

        with pytest.raises(ValueError, match="not found.*recursive=True"):
            await resolve_folder_id(mock_client, "Domains")
