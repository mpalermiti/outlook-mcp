"""Tests for input validation — ported from olkcli patterns."""

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

    def test_strips_dangerous_chars(self):
        result = sanitize_kql('from:boss" OR (hack)')
        assert '"' not in result.strip('"')
        assert "(" not in result.strip('"')
        assert ")" not in result.strip('"')

    def test_preserves_alphanumeric(self):
        result = sanitize_kql("meeting notes 2026")
        assert "meeting" in result
        assert "notes" in result
        assert "2026" in result

    def test_strips_kql_operators(self):
        result = sanitize_kql("test & hack | evil")
        assert "&" not in result
        assert "|" not in result


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
