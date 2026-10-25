"""Tests for folder_resolver.resolve_folder_id."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.folder_resolver import folder_path_part, resolve_folder_id, split_folder_path


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
    response.odata_next_link = None
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
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([receipts]))
        _setup_child_folders(mock_client, {"receipts_id": [domains]})

        result = await resolve_folder_id(mock_client, "Domains")

        assert result == "domains_id"

    async def test_nested_subfolder_match_resolves(self):
        """Folder nested two levels deep still resolves via BFS."""
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        services = _mock_folder("services_id", "Services", child_count=1)
        domains = _mock_folder("domains_id", "Domains")
        mock_client = MagicMock()
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([receipts]))
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
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([a, b]))
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
        mock_client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response([receipts]))
        _setup_child_folders(
            mock_client,
            {"receipts_id": [_mock_folder("bills_id", "Bills")]},
        )

        with pytest.raises(ValueError, match="not found.*recursive=True"):
            await resolve_folder_id(mock_client, "Domains")


# ── Paths (#144) ─────────────────────────────────────────────────────────────
# A "/" separates the parts of a path; "//" is a "/" inside a name.


def _tree_client(top_level, children_by_parent):
    """A client whose top level and childFolders listings answer from dicts.

    Returns the client; ``client.me.mail_folders.get`` and
    ``client.me.mail_folders.by_mail_folder_id`` record what was listed.
    """
    client = MagicMock()
    client.me.mail_folders.get = AsyncMock(return_value=_mock_list_response(top_level))
    _setup_child_folders(client, children_by_parent)
    return client


def _listed_parents(client):
    return [c.args[0] for c in client.me.mail_folders.by_mail_folder_id.call_args_list]


class TestSplitFolderPath:
    @pytest.mark.parametrize(
        "ref,parts",
        [
            ("Inbox/Travel/Flights", ["Inbox", "Travel", "Flights"]),
            ("Inbox/Q1//Q2", ["Inbox", "Q1/Q2"]),
            ("Q1//Q2", ["Q1/Q2"]),
            ("Inbox/A // B", ["Inbox", "A / B"]),
            # Greedy from the left: "///" is an escaped "/" then a separator.
            ("a///b", ["a/", "b"]),
            ("//lead/x", ["/lead", "x"]),
            ("  Inbox / Travel  ", ["Inbox", "Travel"]),
        ],
    )
    def test_parts(self, ref, parts):
        assert split_folder_path(ref) == parts

    @pytest.mark.parametrize("ref", ["/Inbox", "Inbox/", "Inbox/ /Travel", "Inbox/Travel/"])
    def test_an_empty_part_is_refused_with_the_rule(self, ref):
        with pytest.raises(ValueError, match="empty part.*'//'"):
            split_folder_path(ref)

    def test_a_part_doubles_its_slashes(self):
        assert folder_path_part("Q1/Q2") == "Q1//Q2"
        assert folder_path_part("Travel") == "Travel"

    @pytest.mark.parametrize("names", [["Inbox", "Q1/Q2", "A / B"], ["Projects", "a/", "x"]])
    def test_parts_round_trip(self, names):
        assert split_folder_path("/".join(folder_path_part(n) for n in names)) == names


class TestPathResolution:
    async def test_a_well_known_first_part_lists_only_the_levels_below_it(self):
        travel = _mock_folder("travel_id", "Travel", child_count=1)
        client = _tree_client(
            [],
            {"inbox": [travel], "travel_id": [_mock_folder("flights_id", "Flights")]},
        )

        assert await resolve_folder_id(client, "Inbox/Travel/Flights") == "flights_id"

        assert _listed_parents(client) == ["inbox", "travel_id"]
        client.me.mail_folders.get.assert_not_awaited()

    async def test_a_display_alias_first_part_is_its_well_known_name(self):
        client = _tree_client([], {"sentitems": [_mock_folder("y2026_id", "2026")]})

        assert await resolve_folder_id(client, "Sent Items/2026") == "y2026_id"
        assert _listed_parents(client) == ["sentitems"]

    async def test_a_user_folder_first_part_is_found_at_the_top_level(self):
        projects = _mock_folder("projects_id", "Projects", child_count=1)
        client = _tree_client([projects], {"projects_id": [_mock_folder("c_id", "Clients")]})

        assert await resolve_folder_id(client, "Projects/Clients") == "c_id"
        client.me.mail_folders.get.assert_awaited_once()
        assert _listed_parents(client) == ["projects_id"]

    async def test_each_part_matches_case_insensitively(self):
        client = _tree_client(
            [],
            {
                "inbox": [_mock_folder("travel_id", "Travel")],
                "travel_id": [_mock_folder("f", "Flights")],
            },
        )

        assert await resolve_folder_id(client, "inbox/TRAVEL/flights") == "f"

    async def test_a_path_tells_two_folders_of_one_name_apart(self):
        client = _tree_client(
            [_mock_folder("archive_folder", "Old Mail", child_count=1)],
            {
                "inbox": [_mock_folder("inbox_travel", "Travel")],
                "archive_folder": [_mock_folder("old_travel", "Travel")],
            },
        )

        assert await resolve_folder_id(client, "Old Mail/Travel") == "old_travel"
        assert await resolve_folder_id(client, "Inbox/Travel") == "inbox_travel"

    async def test_a_name_with_a_slash_is_written_doubled(self):
        client = _tree_client([], {"inbox": [_mock_folder("q_id", "Q1/Q2")]})

        assert await resolve_folder_id(client, "Inbox/Q1//Q2") == "q_id"

    async def test_a_single_escaped_name_is_looked_up_like_any_name(self):
        receipts = _mock_folder("receipts_id", "Receipts", child_count=1)
        client = _tree_client([receipts], {"receipts_id": [_mock_folder("q_id", "Q1/Q2")]})

        assert await resolve_folder_id(client, "Q1//Q2") == "q_id"

    async def test_a_missing_lower_part_names_where_it_looked_and_stops(self):
        client = _tree_client(
            [],
            {"inbox": [_mock_folder("travel_id", "Travel", child_count=1)], "travel_id": []},
        )

        with pytest.raises(
            ValueError, match="No folder 'Flights' in 'Inbox/Travel'.*recursive=True"
        ):
            await resolve_folder_id(client, "Inbox/Travel/Flights/2026")

        assert _listed_parents(client) == ["inbox", "travel_id"]

    async def test_a_missing_first_part_says_a_path_starts_at_the_top(self):
        client = _tree_client([_mock_folder("p", "Projects")], {})

        with pytest.raises(
            ValueError, match="No folder 'Travel' at the top level. A path starts at the top"
        ):
            await resolve_folder_id(client, "Travel/Flights")

        assert _listed_parents(client) == []

    async def test_two_children_of_one_name_are_ambiguous(self):
        client = _tree_client(
            [], {"inbox": [_mock_folder("a", "Travel"), _mock_folder("b", "travel")]}
        )

        with pytest.raises(ValueError, match="ambiguous"):
            await resolve_folder_id(client, "Inbox/Travel")

    async def test_an_empty_part_is_refused_before_any_call(self):
        client = _tree_client([], {})

        with pytest.raises(ValueError, match="empty part"):
            await resolve_folder_id(client, "Inbox/ /Travel")

        client.me.mail_folders.get.assert_not_awaited()
        assert _listed_parents(client) == []


class TestIdsWithASlash:
    """validate_graph_id accepts "/", and a "/" used to mean "an id". An id that
    no path matches still goes through as one."""

    _ID = "AAMkAGI2TG93AAA" + "x" * 30 + "/AAEMAAA="

    async def test_an_id_shaped_reference_no_path_matches_is_passed_through(self):
        client = _tree_client([_mock_folder("p", "Projects")], {})

        assert await resolve_folder_id(client, self._ID) == self._ID

    async def test_a_short_path_that_matches_nothing_is_not_taken_for_an_id(self):
        client = _tree_client([_mock_folder("p", "Projects")], {})

        with pytest.raises(ValueError, match="No folder 'Travl' at the top level"):
            await resolve_folder_id(client, "Travl/Flights")

    async def test_a_long_path_with_a_space_is_not_taken_for_an_id(self):
        client = _tree_client([_mock_folder("p", "Projects")], {})
        ref = "Projects and Clients/" + "x" * 40

        with pytest.raises(ValueError, match="at the top level"):
            await resolve_folder_id(client, ref)

    async def test_a_long_mistyped_path_below_a_folder_that_exists_is_not_an_id(self):
        """40+ id-alphabet characters, but the first part matched: a path gone wrong,
        so the resolver's message, not Graph's error for a bad id."""
        client = _tree_client([], {"inbox": [_mock_folder("travel_id", "Travel")]})
        ref = "Inbox/TravelPlansForTheYearAheadAndTheOneAfter"

        with pytest.raises(ValueError, match="No folder 'TravelPlans.*' in 'Inbox'"):
            await resolve_folder_id(client, ref)

    async def test_below_a_user_folder_too(self):
        client = _tree_client(
            [_mock_folder("projects_id", "Projects", child_count=1)],
            {"projects_id": [_mock_folder("c", "Clients")]},
        )

        with pytest.raises(ValueError, match="No folder 'ClientsAndPartners.*' in 'Projects'"):
            await resolve_folder_id(client, "Projects/ClientsAndPartnersOfTheLastFiveYears")

    async def test_an_ambiguous_first_part_is_not_an_id(self):
        client = _tree_client([_mock_folder("a", "Shared"), _mock_folder("b", "shared")], {})

        with pytest.raises(ValueError, match="ambiguous"):
            await resolve_folder_id(client, "Shared/" + "x" * 40)

    async def test_a_path_split_folder_path_refuses_is_refused_before_any_call(self):
        """An id has no empty part to refuse, so an id-shaped reference with one is a
        path, mistyped."""
        client = _tree_client([], {})

        with pytest.raises(ValueError, match="empty part"):
            await resolve_folder_id(client, self._ID + "/")

        client.me.mail_folders.get.assert_not_awaited()
        assert _listed_parents(client) == []

    async def test_an_id_with_a_doubled_slash_no_folder_has_is_passed_through(self):
        """One part, "//" read as a "/" in a name: no folder anywhere has that name."""
        ref = self._ID.replace("/", "//")
        client = _tree_client([_mock_folder("p", "Projects")], {})

        assert await resolve_folder_id(client, ref) == ref

    async def test_a_short_id_with_a_plus_is_still_an_id(self):
        """`+` still marks an id on its own, now that `/` marks a path instead."""
        client = _tree_client([], {})

        assert await resolve_folder_id(client, "AAMk+AAA") == "AAMk+AAA"
        client.me.mail_folders.get.assert_not_awaited()

    async def test_a_reference_without_a_slash_still_never_lists_folders(self):
        client = _tree_client([], {})

        assert await resolve_folder_id(client, "AAMkAGI2TG93AAA=") == "AAMkAGI2TG93AAA="
        client.me.mail_folders.get.assert_not_awaited()


class TestAmbiguousNamesListTheirPaths:
    async def test_the_error_names_each_candidate_path(self):
        client = _tree_client(
            [
                _mock_folder("inbox_id", "Inbox", child_count=1),
                _mock_folder("old_id", "Old Mail", child_count=1),
            ],
            {
                "inbox_id": [_mock_folder("t1", "Travel")],
                "old_id": [_mock_folder("t2", "Travel")],
            },
        )

        with pytest.raises(ValueError) as err:
            await resolve_folder_id(client, "Travel")

        assert "'Inbox/Travel'" in str(err.value)
        assert "'Old Mail/Travel'" in str(err.value)
        assert "Pass one of those paths" in str(err.value)

    async def test_paths_in_the_error_double_a_slash_in_a_name(self):
        client = _tree_client(
            [_mock_folder("a", "A/B", child_count=1), _mock_folder("c", "C", child_count=1)],
            {"a": [_mock_folder("x1", "X")], "c": [_mock_folder("x2", "X")]},
        )

        with pytest.raises(ValueError, match="'A//B/X'"):
            await resolve_folder_id(client, "X")

    async def test_the_error_lists_five_and_counts_the_rest(self):
        parents = [_mock_folder(f"p{i}", f"P{i}", child_count=1) for i in range(7)]
        client = _tree_client(
            parents, {f"p{i}": [_mock_folder(f"t{i}", "Travel")] for i in range(7)}
        )

        with pytest.raises(ValueError) as err:
            await resolve_folder_id(client, "Travel")

        text = str(err.value)
        assert "7 matches" in text
        assert "'P4/Travel'" in text
        assert "'P5/Travel'" not in text
        assert "and 2 more" in text
