"""The ``catch_up`` cursor (contracts §7.1, §7.2, §7.4): ``intents/cursor.py``.

It comes back from the agent, so it is untrusted: anything malformed is refused,
a link that isn't Graph's own delta endpoint is an error before any token goes
to it, and an unknown version or another account's cursor starts again from the
window. What counts as a stale link is broad on purpose (§7.4).
"""

from __future__ import annotations

import base64
import json

import pytest

from outlook_mcp.errors import ForeignCursorError, UntrustedURLError
from outlook_mcp.intents import cursor
from outlook_mcp.intents.cursor import CalState, CursorState, MailState
from outlook_mcp.tools.digest import is_stale_delta

MAIL = "https://graph.microsoft.com/v1.0/me/mailFolders('inbox')/messages/delta?$deltatoken=abc"
CAL = "https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=xyz"
ME = "me@example.com"


def state(**overrides) -> CursorState:
    fields = {
        "acct": cursor.account_tag(ME),
        "made": "2026-11-26T12:00:00Z",
        "mail": MailState(link=MAIL, created_only=True, bootstrap_from=None),
        "cal": CalState(link=CAL, start="2026-11-25T12:00:00Z", end="2026-12-26T12:00:00Z"),
    }
    fields.update(overrides)
    return CursorState(**fields)


def raw(data: object) -> str:
    text = json.dumps(data).encode()
    return cursor.PREFIX + base64.urlsafe_b64encode(text).decode().rstrip("=")


def test_a_round_trip():
    token = cursor.encode(state())
    assert token.startswith("cu1.")
    assert "=" not in token
    read, status = cursor.read(token, username=ME)
    assert status == "ok"
    assert read.as_dict() == state().as_dict()


def test_no_cursor_is_none():
    assert cursor.read(None, username=ME) == (None, "none")


@pytest.mark.parametrize(
    "token",
    ["junk", "cu1.", "cu1.!!!", "cu2.abc", raw([1, 2]), raw("text"), "cu1.bm90IGpzb24"],
)
def test_junk_is_refused(token):
    with pytest.raises(ValueError, match="Call again without cursor"):
        cursor.read(token, username=ME)


def test_a_v1_object_missing_its_parts_is_refused():
    data = state().as_dict()
    del data["mail"]["created_only"]
    with pytest.raises(ValueError, match="Invalid catch_up cursor"):
        cursor.read(raw(data), username=ME)
    data = state().as_dict()
    data["made"] = "yesterday"
    with pytest.raises(ValueError):
        cursor.read(raw(data), username=ME)


def test_an_unknown_version_starts_from_the_window():
    data = {**state().as_dict(), "v": 2}
    assert cursor.read(raw(data), username=ME) == (None, "reset_unrecognised")


def test_another_accounts_cursor_starts_from_the_window():
    token = cursor.encode(state())
    assert cursor.read(token, username="someone@example.org") == (None, "reset_other_account")
    # The account is compared case-blind, as Graph compares addresses.
    assert cursor.read(token, username="Me@Example.com")[1] == "ok"


@pytest.mark.parametrize(
    ("link", "where", "error"),
    [
        ("https://evil.example/v1.0/me/calendarView/delta", "cal", UntrustedURLError),
        ("https://graph.microsoft.com/v1.0/me/messages", "mail", ForeignCursorError),
        (
            "https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=a",
            "mail",
            ForeignCursorError,
        ),
        (
            "https://graph.microsoft.com/v1.0/me/calendarView/delta?$select=subject",
            "cal",
            ForeignCursorError,
        ),
    ],
)
def test_a_link_that_is_not_the_endpoints_own_raises(link, where, error):
    data = state().as_dict()
    data[where]["link"] = link
    with pytest.raises(error):
        cursor.read(raw(data), username=ME)


def test_an_empty_link_means_not_started():
    data = state().as_dict()
    data["mail"]["link"] = ""
    read, status = cursor.read(raw(data), username=ME)
    assert status == "ok" and read.mail.link == ""


def test_the_created_only_link_may_carry_its_change_type():
    """`_delta.py`'s one new option (§7.1): a mail delta link with changeType."""
    data = state().as_dict()
    data["mail"]["link"] = MAIL.replace("?", "?changeType=created&")
    assert cursor.read(raw(data), username=ME)[1] == "ok"


@pytest.mark.parametrize("status", [400, 404, 409, 410])
def test_a_client_error_on_a_delta_link_is_stale(status):
    assert is_stale_delta(status) is True


@pytest.mark.parametrize("status", [401, 403, 429, 500, 503, 200])
def test_sign_in_permission_throttling_and_server_errors_are_not_stale(status):
    assert is_stale_delta(status) is False


def test_the_size_thresholds(caplog):
    assert cursor.encode(state()) is not None
    big = state(mail=MailState(link=MAIL + "x" * 7000, created_only=True, bootstrap_from=None))
    with caplog.at_level("WARNING"):
        assert cursor.encode(big) is not None
    assert "over 8192" in caplog.text
    huge = state(mail=MailState(link=MAIL + "x" * 13000, created_only=True, bootstrap_from=None))
    assert cursor.encode(huge) is None


def test_the_account_is_a_hash_not_an_address():
    tag = cursor.account_tag(ME)
    assert len(tag) == 16 and "@" not in tag
    decoded = base64.urlsafe_b64decode(cursor.encode(state())[4:] + "==").decode()
    assert ME not in decoded and "example.com" not in decoded
