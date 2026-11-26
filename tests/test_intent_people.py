"""People: names, lookalikes, the sent table and the person score (contracts §2.3, §2.5, §2.6).

Every function here is pure. The rules they hold are the ones that decide whether
a name is acted on without asking:

- Only letters the owner sent make someone known. Meeting requests and responses
  don't, and the owner's own addresses never count as recipients.
- Received mail never adds to a score.
- A display name or address built to look like someone the owner writes to is
  flagged, and ordinary names (a surname in two scripts, two webmail providers)
  are not.
"""

from datetime import datetime, timezone

import pytest

from outlook_mcp.intents import people
from outlook_mcp.intents.people import (
    PROVIDERS,
    build_sent_table,
    damerau_levenshtein,
    domain_skeleton,
    is_meeting_message,
    mixed_script,
    name_match,
    norm_name,
    person_score,
    person_warnings,
    recency_score,
    registrable,
    written_score,
)
from tests.intent_fakes import ME, letter, meeting_response

NOW = datetime(2026, 11, 26, 12, 0, tzinfo=timezone.utc)


def codes(warnings: list[dict]) -> list[str]:
    return [w["code"] for w in warnings]


# ── Normalization ─────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("raw", "normal"),
    [
        ("Sam Jones", "sam jones"),
        ("Sаm Jones", "sam jones"),  # Cyrillic а
        ("José Núñez", "jose nunez"),
        ("Ｓａｍ", "sam"),  # fullwidth
        ("O'Brien,  Pat", "o brien pat"),
        ("οlga", "olga"),  # Greek ο
        ("Tanaka 田中", "tanaka"),
        ("田中", ""),
    ],
)
def test_norm_name(raw, normal):
    assert norm_name(raw) == normal


@pytest.mark.parametrize(
    ("lookalike", "real"),
    [
        ("rnodern.com", "modern.com"),
        ("paypa1.com", "paypal.com"),
        ("g00gle.com", "google.com"),
    ],
)
def test_domain_skeleton_folds_lookalike_letters(lookalike, real):
    assert domain_skeleton(lookalike) == domain_skeleton(real)


@pytest.mark.parametrize(
    ("domain", "base"),
    [
        ("example.com", "example.com"),
        ("mail.example.com", "example.com"),
        ("a.b.example.com", "example.com"),
        ("mail.example.co.uk", "example.co.uk"),
        ("shop.foo.com.au", "foo.com.au"),
        ("example.de", "example.de"),
    ],
)
def test_registrable(domain, base):
    assert registrable(domain) == base


@pytest.mark.parametrize(
    ("a", "b", "distance"),
    [
        ("jones", "jonse", 1),
        ("abc", "abc", 0),
        ("ab", "ba", 1),
        ("kitten", "sitting", 3),
        ("", "ab", 2),
    ],
)
def test_damerau_levenshtein(a, b, distance):
    assert damerau_levenshtein(a, b) == distance


@pytest.mark.parametrize(
    ("name", "mixed"),
    [
        ("Sаm Jones", True),
        ("Ρaul", True),  # Greek Ρ with Latin
        ("Tanaka 田中", False),
        ("田中 たなか カタカナ", False),
        ("김 Kim", False),
        ("Иван Петров", False),
        ("Sam Jones", False),
    ],
)
def test_mixed_script(name, mixed):
    assert mixed_script(name) is mixed


def test_providers_are_the_contracts():
    assert {"gmail.com", "outlook.com", "ymail.com", "mail.com", "proton.me"} <= PROVIDERS


# ── The sent table ────────────────────────────────────────────────────────


def test_meeting_messages_are_recognised_by_type():
    assert is_meeting_message({"@odata.type": "#microsoft.graph.eventMessageResponse"})
    assert is_meeting_message({"@odata.type": "#microsoft.graph.eventMessageRequest"})
    assert is_meeting_message({"@odata.type": "#microsoft.graph.eventMessage"})
    assert not is_meeting_message({"@odata.type": "#microsoft.graph.message"})
    assert not is_meeting_message({})


def test_the_sent_table_counts_letters_per_recipient():
    rows = [
        letter([("sam@example.com", "Sam Jones")], sent="2026-11-20T10:00:00Z"),
        letter(["sam@example.com"], cc=["ann@example.org"], sent="2026-10-01T10:00:00Z"),
        letter([], bcc=["SAM@example.com"], sent="2026-09-01T10:00:00Z"),
    ]
    table = build_sent_table(rows, signed_in=None)
    sam = table.entries["sam@example.com"]
    assert sam.count == 3
    assert sam.last == datetime(2026, 11, 20, 10, tzinfo=timezone.utc)
    assert sam.first == datetime(2026, 9, 1, 10, tzinfo=timezone.utc)
    assert sam.name == "Sam Jones"
    assert table.entries["ann@example.org"].count == 1
    assert table.scanned == 3


def test_sent_table_drops_meeting_messages_and_own_addresses():
    rows = [
        meeting_response("stranger@example.net"),
        letter(["me.alias@example.com", "friend@example.com"], sender="me.alias@example.com"),
        letter([ME]),
    ]
    table = build_sent_table(rows, signed_in="Me.Other@Example.com")
    assert "stranger@example.net" not in table.entries
    assert table.meeting_messages_dropped == 1
    assert {"me@example.com", "me.alias@example.com", "me.other@example.com"} <= table.me
    assert set(table.entries) == {"friend@example.com"}


def test_an_unreadable_row_is_skipped_not_fatal():
    rows = [
        {"toRecipients": [{"emailAddress": None}], "sentDateTime": "not a date"},
        letter(["a@example.com"]),
    ]
    table = build_sent_table(rows, signed_in=None)
    assert set(table.entries) == {"a@example.com"}


# ── Scores ────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("count", "score"),
    [(0, 0), (1, 5), (2, 10), (3, 10), (4, 15), (8, 20), (15, 20), (16, 25), (900, 25)],
)
def test_written_score(count, score):
    assert written_score(count) == score


def test_recency_score():
    assert recency_score(None, NOW) == 0
    assert recency_score(datetime(2026, 9, 1, tzinfo=timezone.utc), NOW) == 10
    assert recency_score(datetime(2026, 2, 1, tzinfo=timezone.utc), NOW) == 5
    assert recency_score(datetime(2025, 1, 1, tzinfo=timezone.utc), NOW) == 0


@pytest.mark.parametrize(
    ("words", "name", "address", "expected"),
    [
        ("sam@example.com", "Sam Jones", "sam@example.com", (100, "address")),
        ("SAM@EXAMPLE.COM", "", "sam@example.com", (100, "address")),
        ("sam jones", "Sam Jones", "sj@example.com", (40, "full")),
        ("Sаm Jones", "Sam Jones", "sj@example.com", (40, "full")),
        ("jones", "Sam Jones", "sj@example.com", (30, "first_or_last")),
        ("sam", "Sam Jones", "sj@example.com", (30, "first_or_last")),
        ("jon", "Sam Jones", "sj@example.com", (20, "prefix")),
        ("sa jo", "Sam Jones", "sj@example.com", (20, "prefix")),
        ("am jo", "Sam Jones", "sj@example.com", (10, "contains")),
        ("sjones", "Someone", "sjones@example.com", (10, "local_part")),
        ("田中", "田中 太郎", "t@example.jp", (30, "first_or_last")),
        ("zed", "Sam Jones", "sj@example.com", (0, None)),
    ],
)
def test_name_match(words, name, address, expected):
    assert name_match(words, name, address) == expected


def test_person_score_adds_the_parts_and_received_mail_is_not_one():
    score = person_score(
        name=40,
        count=4,
        last=datetime(2026, 11, 1, tzinfo=timezone.utc),
        saved_contact=True,
        now=NOW,
    )
    assert score == 40 + 15 + 10 + 10
    assert person_score(name=30, count=0, last=None, saved_contact=False, now=NOW) == 30
    # The signature has no room for received mail: it cannot add to the score.
    assert "received" not in person_score.__code__.co_varnames


# ── Warnings ──────────────────────────────────────────────────────────────


def _table(*rows):
    return build_sent_table(list(rows), signed_in=None)


KNOWN = _table(
    letter([("sam.jones@example.com", "Sam Jones")]),
    letter([("pat@paypal.com", "Pat")]),
    letter([("dev@modern.com", "Dev")]),
    letter([("friend@gmail.com", "Friend")]),
)


def _warnings(address, name, written=False, contacts=()):
    return codes(person_warnings(address, name, KNOWN, written=written, contacts=list(contacts)))


def test_blocking_warnings_prevent_single_inputs():
    assert "mixed_script_name" in _warnings("sam@other.example", "Sаm Jones")
    assert "lookalike_domain" in _warnings("pat@paypa1.com", "Pat")
    assert "lookalike_domain" in _warnings("dev@rnodern.com", "Dev")
    assert "idn_domain" in _warnings("x@xn--pypal-4ve.com", "X")
    assert "name_contains_other_address" in _warnings("x@evil.example", "billing@bank.example.com")
    assert "lookalike_name" in _warnings("sam@elsewhere.example", "Sam Jones")


def test_a_lookalike_carries_what_it_looks_like():
    found = person_warnings("pat@paypa1.com", "Pat", KNOWN, written=False, contacts=[])
    assert {"code": "lookalike_domain", "like": "paypal.com"} in found
    found = person_warnings("sam.jonse@example.com", "S", KNOWN, written=False, contacts=[])
    assert {"code": "lookalike_address", "like": "sam.jones@example.com"} in found


def test_ordinary_names_are_not_lookalikes():
    assert "mixed_script_name" not in _warnings("t@example.jp", "Tanaka 田中")
    assert "mixed_script_name" not in _warnings("t@example.jp", "田中 たなか")
    assert not {"lookalike_domain"} & set(_warnings("someone@ymail.com", "Someone"))
    assert not {"lookalike_domain"} & set(_warnings("someone@mail.com", "Someone"))
    assert "lookalike_address" in _warnings("sam.jonse@example.com", "S")
    # Two Japanese names that both normalize to nothing are not each other's lookalike.
    table = _table(letter([("a@example.jp", "田中")]))
    assert "lookalike_name" not in codes(
        person_warnings("b@example.jp", "山田", table, written=False, contacts=[])
    )


def test_a_known_address_gets_no_lookalike_warning():
    assert _warnings("sam.jones@example.com", "Sam Jones", written=True) == []


def test_an_unknown_address_is_first_time():
    assert "first_time" in _warnings("new@example.org", "New Person")
    assert "first_time" not in _warnings("new@example.org", "New Person", written=True)


def test_gmail_dots_are_the_same_mailbox_not_a_lookalike():
    assert "lookalike_address" not in _warnings("fri.end@gmail.com", "F")
    assert "lookalike_address" in _warnings("frend@gmail.com", "F")


def test_a_saved_contact_name_at_another_address_is_a_lookalike():
    found = _warnings("ann@other.example", "Ann Lee", contacts=[("Ann Lee", "ann@example.com")])
    assert "lookalike_name" in found


def test_an_invalid_address_is_flagged_and_nothing_else_is_computed_on_it():
    assert _warnings("not an address", "Sam Jones") == ["invalid_address"]


def test_warning_order_is_stable():
    first = _warnings("sam@paypa1.com", "Sаm Jones billing@bank.example.com")
    assert first == _warnings("sam@paypa1.com", "Sаm Jones billing@bank.example.com")
    assert people.WARNING_ORDER.index(first[0]) <= people.WARNING_ORDER.index(first[-1])
