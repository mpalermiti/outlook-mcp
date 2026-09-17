"""Tier 3 of the silent-no-op audit: is the ``$select`` as wide as its formatter?

``$select`` and the function that reads the response are two halves of one
contract, written in two places, in two spellings — camelCase on the wire,
snake_case on the SDK model. Nothing connects them, so they drift silently and
in the quietest possible direction: Graph returns exactly what was asked for,
the SDK leaves the unrequested attribute as ``None``, and the formatter turns
that into ``""``. The caller sees a field that is present and empty, which is
indistinguishable from a field that is genuinely empty.

Two real instances, neither of which failed a single test:

- ``_format_message_summary`` reads ``inference_classification``, but of its
  four call sites only ``list_inbox`` selected it. ``search_mail``,
  ``list_drafts`` and ``get_thread`` reported every message as unclassified.
- ``_format_contact_detail`` read twelve fields of a response fetched with no
  ``$select`` at all, and dropped five of them (#63).

So this file reads the formatter's source and asserts the ``$select`` beside it
covers every field it touches. Add a row to ``_PAIRS`` for each new formatter.
"""

from __future__ import annotations

import ast
import inspect
import textwrap

import pytest

from outlook_mcp.tools import contacts, mail_read

# (label, formatter function, the $select string that feeds it)
#
# `_format_contact_summary` is paired with the listing's wider select: it is the
# one that can reach every field the formatter reads. The search path sends
# `_SUMMARY_SELECT`, which is deliberately narrower, and the formatter omits the
# key it cannot honour rather than reporting it empty (`with_categories`).
_PAIRS = [
    ("mail summary", mail_read._format_message_summary, mail_read.SUMMARY_SELECT),
    ("contact summary", contacts._format_contact_summary, contacts._LIST_SELECT),
]

# The SDK renames exactly one field to dodge the Python keyword.
_SDK_RENAMES = {"from_": "from"}


def _graph_name(attr: str) -> str:
    """The Graph property name for an SDK attribute: snake_case -> camelCase."""
    if attr in _SDK_RENAMES:
        return _SDK_RENAMES[attr]
    head, *rest = attr.split("_")
    return head + "".join(word.title() for word in rest)


def _fields_read_from(func, _seen: frozenset[str] = frozenset()) -> set[str]:
    """Every attribute the function reads off its first parameter.

    Nested reads (``msg.from_.email_address.address``) contribute only the
    outermost name, which is the one ``$select`` names.

    Reads through a helper count as reads: a call to a function in the same
    module, handed the model as its first argument, is followed into.

    ``_format_contact_summary`` gets its phone from ``_primary_phone(contact)``
    and its categories from ``_categories(contact)``. Without following those
    calls the guard does not go quiet — it fails, and points at the wrong fix::

        contact summary: the $select asks for ['businessPhones', 'categories',
        'homePhones', 'mobilePhone'], which _format_contact_summary never
        reads. Drop them, or read them.

    Do what that says and four fields leave the ``$select``: phone stops
    resolving on the listing, and the ``categories`` gap that was the actual
    bug is now "fixed" by no longer asking Graph for the field. Following the
    helpers, the same broken ``$select`` reports the real defect instead —
    ``_format_contact_summary reads ['categories'], which the $select never
    asks Graph for``. A guard that misdirects the repair is worse than one that
    stays silent, which is the argument for the extension.
    """
    tree = ast.parse(textwrap.dedent(inspect.getsource(func)))
    fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef))
    param = fn.args.args[0].arg
    module = inspect.getmodule(func)
    seen = _seen | {func.__qualname__}

    found: set[str] = set()
    for node in ast.walk(fn):
        # msg.subject
        if (
            isinstance(node, ast.Attribute)
            and isinstance(node.value, ast.Name)
            and node.value.id == param
        ):
            found.add(node.attr)
        # getattr(msg, "inference_classification", None)
        elif (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "getattr"
            and len(node.args) >= 2
            and isinstance(node.args[0], ast.Name)
            and node.args[0].id == param
            and isinstance(node.args[1], ast.Constant)
            and isinstance(node.args[1].value, str)
        ):
            found.add(node.args[1].value)
        # _primary_phone(contact) — a helper in this module, handed the model
        elif (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.args
            and isinstance(node.args[0], ast.Name)
            and node.args[0].id == param
        ):
            helper = getattr(module, node.func.id, None)
            if (
                callable(helper)
                and inspect.getmodule(helper) is module
                and helper.__qualname__ not in seen
            ):
                found |= _fields_read_from(helper, seen)
    return found


@pytest.mark.parametrize("label,formatter,select", _PAIRS, ids=[p[0] for p in _PAIRS])
def test_the_select_covers_every_field_its_formatter_reads(label, formatter, select):
    selected = {field.strip() for field in select.split(",")}
    required = {_graph_name(attr) for attr in _fields_read_from(formatter)}

    missing = sorted(required - selected)
    assert not missing, (
        f"{label}: {formatter.__name__} reads {missing}, which the $select never "
        f"asks Graph for. Those come back None and the formatter reports them as "
        f"empty — the caller cannot tell that from genuinely empty.\n"
        f"$select was: {select}"
    )


@pytest.mark.parametrize("label,formatter,select", _PAIRS, ids=[p[0] for p in _PAIRS])
def test_the_select_asks_for_nothing_its_formatter_ignores(label, formatter, select):
    """The other end of the same mistake: paying Graph for unread fields."""
    selected = {field.strip() for field in select.split(",")}
    read = {_graph_name(attr) for attr in _fields_read_from(formatter)}

    unused = sorted(selected - read)
    assert not unused, (
        f"{label}: the $select asks for {unused}, which {formatter.__name__} never "
        f"reads. Drop them, or read them."
    )


def test_a_field_read_through_a_helper_still_counts_as_read():
    """The extension above, pinned.

    Without it a formatter could move every read into a helper and this guard
    would go quietly blind — passing both directions while selecting fields
    nobody reads and reading fields nobody selected.
    """
    reads = _fields_read_from(contacts._format_contact_summary)
    assert "mobile_phone" in reads, "read through _primary_phone"
    assert "home_phones" in reads, "read through _primary_phone"
    assert "categories" in reads, "read through _categories"


def test_the_summary_select_is_spelled_in_exactly_one_place():
    """Four verbatim copies are how the classification gap opened.

    ``list_inbox`` gained ``inferenceClassification``; the three copies in
    ``search_mail``, ``list_drafts`` and ``get_thread`` did not.
    """
    from pathlib import Path

    src = Path(mail_read.__file__).parent
    # The comma-joined form, which only a $select has. ``mail_delta`` names the
    # same fields as raw dict keys — /me/messages/delta takes no $select and
    # returns the whole resource — so it is not a copy of this list.
    needle = "bodyPreview,hasAttachments"
    offenders = [
        path.relative_to(src.parent.parent).as_posix()
        for path in src.rglob("*.py")
        if needle in path.read_text() and path.name != "mail_read.py"
    ]
    assert not offenders, (
        f"The message-summary field list is spelled again in {offenders}. "
        f"Import mail_read.SUMMARY_SELECT instead — copies drift, and they drift "
        f"silently."
    )
