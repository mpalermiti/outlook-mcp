"""Tests for cursor-based pagination."""

import base64
import json

import pytest

from outlook_mcp.pagination import (
    apply_pagination,
    decode_cursor,
    encode_cursor,
    wrap_nextlink,
)


class TestEncodeDecode:
    def test_roundtrip(self):
        cursor = encode_cursor(25)
        assert decode_cursor(cursor) == 25

    def test_roundtrip_zero(self):
        cursor = encode_cursor(0)
        assert decode_cursor(cursor) == 0

    def test_decode_invalid_base64(self):
        with pytest.raises(ValueError, match="Invalid pagination cursor"):
            decode_cursor("not-valid-base64!!!")

    def test_decode_invalid_json(self):
        bad = base64.urlsafe_b64encode(b"not json").decode()
        with pytest.raises(ValueError, match="Invalid pagination cursor"):
            decode_cursor(bad)

    def test_decode_negative_skip(self):
        bad = base64.urlsafe_b64encode(json.dumps({"skip": -1}).encode()).decode()
        with pytest.raises(ValueError, match="Invalid skip value"):
            decode_cursor(bad)


class TestWrapNextlink:
    def test_none_returns_none(self):
        assert wrap_nextlink(None) is None

    def test_empty_returns_none(self):
        assert wrap_nextlink("") is None

    def test_extracts_skip(self):
        url = "https://graph.microsoft.com/v1.0/me/messages?$skip=50&$top=25"
        cursor = wrap_nextlink(url)
        assert cursor is not None
        assert decode_cursor(cursor) == 50


class TestApplyPagination:
    def test_first_page(self):
        params: dict = {}
        result = apply_pagination(params, count=25)
        assert result["$top"] == 25
        assert "$skip" not in result

    def test_with_cursor(self):
        cursor = encode_cursor(50)
        params: dict = {}
        result = apply_pagination(params, count=25, cursor=cursor)
        assert result["$top"] == 25
        assert result["$skip"] == 50

    def test_clamps_count_high(self):
        params: dict = {}
        result = apply_pagination(params, count=200)
        assert result["$top"] == 100

    def test_clamps_count_low(self):
        params: dict = {}
        result = apply_pagination(params, count=0)
        assert result["$top"] == 1

    def test_invalid_cursor_raises(self):
        with pytest.raises(ValueError):
            apply_pagination({}, count=25, cursor="garbage")
