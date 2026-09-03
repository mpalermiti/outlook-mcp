"""Tests for input validation — ported from olkcli patterns."""

import itertools

import pytest

from outlook_mcp.validation import (
    sanitize_kql,
    sanitize_output,
    validate_datetime,
    validate_email,
    validate_folder_name,
    validate_graph_id,
    validate_phone,
)


class TestGraphIdValidation:
    def test_valid_id(self):
        assert validate_graph_id("AAMkAGI2TG93AAA=") == "AAMkAGI2TG93AAA="

    def test_valid_id_with_slashes(self):
        assert validate_graph_id("AAMkAG/test+id=") == "AAMkAG/test+id="

    def test_rejects_empty(self):
        with pytest.raises(ValueError, match="empty"):
            validate_graph_id("")

    def test_rejects_too_long(self):
        with pytest.raises(ValueError, match="too long"):
            validate_graph_id("A" * 1025)

    def test_rejects_special_chars(self):
        with pytest.raises(ValueError, match="invalid"):
            validate_graph_id("id with spaces")

    def test_rejects_injection(self):
        with pytest.raises(ValueError, match="invalid"):
            validate_graph_id("../../etc/passwd")


class TestEmailValidation:
    def test_valid_email(self):
        assert validate_email("user@outlook.com") == "user@outlook.com"

    def test_rejects_no_at(self):
        with pytest.raises(ValueError):
            validate_email("notanemail")

    def test_rejects_injection(self):
        with pytest.raises(ValueError):
            validate_email("user@evil.com' OR 1=1--")


class TestDatetimeValidation:
    def test_valid_iso_utc(self):
        result = validate_datetime("2026-04-12T10:30:00Z")
        assert result == "2026-04-12T10:30:00Z"

    def test_valid_iso_with_offset(self):
        result = validate_datetime("2026-04-12T10:30:00+05:00")
        # Should parse and re-serialize to UTC
        assert "Z" in result or "+" in result  # Valid ISO output

    def test_valid_date_only(self):
        """Date-only input gets interpreted as midnight UTC."""
        result = validate_datetime("2026-04-12")
        assert "2026-04-12" in result

    def test_rejects_garbage(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("not-a-date")

    def test_rejects_injection(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("2026-04-12' OR 1=1--")

    def test_rejects_odata_injection(self):
        with pytest.raises(ValueError, match="Invalid datetime"):
            validate_datetime("2026-04-12T00:00:00Z eq true")


class TestKqlSanitization:
    def test_simple_query(self):
        assert sanitize_kql("budget report") == '"budget report"'

    def test_preserves_alphanumeric(self):
        result = sanitize_kql("meeting notes 2026")
        assert "meeting" in result
        assert "notes" in result
        assert "2026" in result

    def test_strips_kql_operators(self):
        # `&`, `|`, `!` are not Graph operators — the uppercase words are.
        # The symbol forms silently zero out an otherwise-matching query.
        result = sanitize_kql("test & hack | evil")
        assert "&" not in result
        assert "|" not in result

    # ── Property restrictions must survive (the #30 regression) ──

    def test_preserves_property_restriction_colon(self):
        """Stripping `:` turned every documented query into a 0-result phrase."""
        assert sanitize_kql("subject:Unlock") == '"subject:Unlock"'

    def test_preserves_all_documented_query_forms(self):
        for query in (
            "from:sarah@acme.com",
            "hasattachment:true",
            "received>=2026-01-01",
            "subject:a AND subject:b",
            "subject:a OR subject:b",
            "NOT subject:a",
        ):
            assert sanitize_kql(query) == f'"{query}"', query

    def test_preserves_grouping_parens(self):
        assert sanitize_kql("subject:(a OR b)") == '"subject:(a OR b)"'

    # ── Security: the two chars that must never survive ──

    def test_strips_quote_that_would_neutralize_search(self):
        """An embedded quote makes Graph silently discard $search and return
        the whole mailbox — 200, no error. This is the real injection vector."""
        result = sanitize_kql('Haverhill" OR "Wayfair')
        assert result == '"Haverhill OR Wayfair"'
        assert result.count('"') == 2

    def test_strips_backslash_escape_metachar(self):
        """`\\` is a real string-literal escape: `\\s` is a 400 and a trailing
        `\\` escapes our own closing quote (400, unterminated literal)."""
        assert sanitize_kql("back\\slash") == '"backslash"'
        assert "\\" not in sanitize_kql("Acrisure\\")

    def test_strips_bare_wildcard(self):
        """Graph already prefix-matches (`subject:Amaz` == `subject:Amaz*`), so
        `*` buys nothing. Allowing it would turn a bare `*` into a silent
        whole-mailbox read; stripping it leaves an empty query we reject."""
        with pytest.raises(ValueError, match="empty after sanitization"):
            sanitize_kql("*")

    def test_rejects_query_that_sanitizes_to_empty(self):
        """$search="" is an opaque Graph BadRequest — fail clearly instead."""
        for query in ('"""', "&|!", "   "):
            with pytest.raises(ValueError, match="empty after sanitization"):
                sanitize_kql(query)

    def test_wrapper_invariant_holds_for_any_input(self):
        """The invariant that actually prevents the vulnerability: whatever goes
        in, exactly two quotes come out and no backslash survives."""
        for combo in itertools.product('ab":\\()&|!*<>= ', repeat=3):
            try:
                result = sanitize_kql("".join(combo))
            except ValueError:
                continue  # rejected outright — also safe
            assert result.count('"') == 2, result
            assert "\\" not in result, result
            assert result.startswith('"') and result.endswith('"'), result


class TestFolderNameValidation:
    def test_wellknown_folders(self):
        assert validate_folder_name("inbox") == "inbox"
        assert validate_folder_name("drafts") == "drafts"
        assert validate_folder_name("sentitems") == "sentitems"
        assert validate_folder_name("deleteditems") == "deleteditems"
        assert validate_folder_name("junkemail") == "junkemail"
        assert validate_folder_name("archive") == "archive"

    def test_case_insensitive_wellknown(self):
        assert validate_folder_name("Inbox") == "inbox"
        assert validate_folder_name("DRAFTS") == "drafts"

    def test_custom_folder_id(self):
        """Custom folder IDs pass through graph ID validation."""
        assert validate_folder_name("AAMkAGFolderId=") == "AAMkAGFolderId="

    def test_rejects_invalid(self):
        with pytest.raises(ValueError):
            validate_folder_name("../../evil")


class TestPhoneValidation:
    def test_valid_phone(self):
        assert validate_phone("+1 (555) 123-4567") == "+1 (555) 123-4567"

    def test_rejects_letters(self):
        with pytest.raises(ValueError):
            validate_phone("call me maybe")

    def test_rejects_too_long(self):
        with pytest.raises(ValueError):
            validate_phone("1" * 31)


class TestOutputSanitization:
    def test_strips_control_chars(self):
        assert sanitize_output("normal text") == "normal text"
        assert sanitize_output("evil\x1b[31mred\x1b[0m") == "evilred"
        assert sanitize_output("tab\there") == "tab here"

    def test_preserves_newlines_in_multiline(self):
        result = sanitize_output("line1\nline2", multiline=True)
        assert "\n" in result

    def test_strips_null_bytes(self):
        assert sanitize_output("null\x00byte") == "nullbyte"
