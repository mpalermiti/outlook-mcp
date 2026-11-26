"""Structured mail selection, search mode (contracts §4): the KQL ``find`` sends.

``find`` always has words, so it always searches. Search mode sends ``$search``
alone: whether Graph accepts it next to ``$filter`` or ``$orderby`` on messages is
unsettled (probes P-F2a/b), so every field is written as KQL, and anything KQL
says only roughly (dates, to the day) is checked again on what comes back. The
agent's words can only narrow the search: quotes, parentheses and the KQL
operators are taken out, so nothing in them can widen it past the sender or the
dates.
"""

from datetime import datetime, timezone

import pytest

from outlook_mcp.intents.selection import MailSelection, kql_words, matches_locally, search_kql

UTC = timezone.utc


def test_words_alone():
    assert search_kql(MailSelection(words="budget report")) == '"budget report"'


def test_every_field_in_order():
    selection = MailSelection(
        words="budget",
        from_=["sam@example.com"],
        newer_than=datetime(2026, 11, 10, 8, tzinfo=UTC),
        older_than=datetime(2026, 11, 20, 23, tzinfo=UTC),
        has_attachments=True,
    )
    assert search_kql(selection) == (
        '"budget from:sam@example.com received>=2026-11-09 received<=2026-11-21 '
        'hasAttachments:true"'
    )


def test_several_senders_are_one_or_group():
    selection = MailSelection(words="plan", from_=["a@example.com", "b@example.org"])
    assert search_kql(selection) == '"plan (from:a@example.com OR from:b@example.org)"'


def test_dates_are_widened_a_day_each_side():
    # KQL dates are whole days in a zone Graph does not state, so the query asks for
    # a day more on each side and the exact bound is applied to what comes back.
    selection = MailSelection(words="x", newer_than=datetime(2026, 1, 1, 0, 0, tzinfo=UTC))
    assert search_kql(selection) == '"x received>=2025-12-31"'


def test_has_attachments_false_is_checked_locally_only():
    assert search_kql(MailSelection(words="x", has_attachments=False)) == '"x"'


@pytest.mark.parametrize(
    ("words", "kept"),
    [
        ('budget" OR from:evil@example.com', "budget from:evil@example.com"),
        ("x) OR (y", "x y"),
        ("a AND b NOT c NEAR d", "a b c d"),
        ("cats and dogs or birds", "cats and dogs or birds"),
        ("subject:budget", "subject:budget"),
        ("pay*ment & co | x !", "payment co x"),
    ],
)
def test_the_words_cannot_widen_the_search(words, kept):
    assert kql_words(words) == kept


def test_words_that_leave_nothing_are_refused():
    with pytest.raises(ValueError, match="nothing to search"):
        search_kql(MailSelection(words='" OR ( )'))


def test_a_selection_without_words_is_not_a_search():
    with pytest.raises(ValueError):
        search_kql(MailSelection(words=None))


def test_a_sender_must_be_an_address():
    with pytest.raises(ValueError):
        search_kql(MailSelection(words="x", from_=["sam jones"]))


def _row(received="2026-11-15T10:00:00Z", sender="sam@example.com", attachments=True):
    return {
        "id": "m1",
        "receivedDateTime": received,
        "from": {"emailAddress": {"address": sender, "name": "Sam"}},
        "hasAttachments": attachments,
    }


def test_the_local_check_applies_the_exact_bounds():
    selection = MailSelection(
        words="x",
        newer_than=datetime(2026, 11, 15, 10, tzinfo=UTC),
        older_than=datetime(2026, 11, 15, 11, tzinfo=UTC),
    )
    assert matches_locally(_row("2026-11-15T10:00:00Z"), selection)
    assert not matches_locally(_row("2026-11-15T09:59:59Z"), selection)
    assert not matches_locally(_row("2026-11-15T11:00:00Z"), selection)
    assert not matches_locally(_row(received=None), selection)


def test_the_local_check_applies_the_sender_and_attachments():
    selection = MailSelection(words="x", from_=["SAM@example.com"], has_attachments=True)
    assert matches_locally(_row(), selection)
    assert not matches_locally(_row(sender="pam@example.com"), selection)
    assert not matches_locally(_row(attachments=False), selection)
    assert matches_locally(_row(attachments=False), MailSelection(words="x", has_attachments=False))


def test_the_echo_names_the_mode_and_the_query():
    selection = MailSelection(words="budget", from_=["sam@example.com"], folder="inbox")
    echo = selection.echo()
    assert echo == {
        "mode": "search",
        "words": "budget",
        "folder": "inbox",
        "from": ["sam@example.com"],
        "newer_than": None,
        "older_than": None,
        "has_attachments": None,
        "query": '"budget from:sam@example.com"',
    }
