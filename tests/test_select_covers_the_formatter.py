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

from outlook_mcp.tools import calendar_read, contacts, mail_read

# (label, formatter function, the $select string that feeds it)
#
# `_format_contact_summary` is paired with the listing's wider select: it is the
# one that can reach every field the formatter reads. The search path sends
# `_SUMMARY_SELECT`, which is deliberately narrower, and the formatter omits the
# key it cannot honour rather than reporting it empty (`with_categories`).
#
# Both calendar rows are the fix for #69, where this guard's own bug class had
# shipped unseen for seven releases — 1.16.0 through 1.22.0 — because calendar
# had never been enrolled: the listing read `type` without selecting it, and
# selected `categories` without reading it — one row, failing in both directions
# at once. `test_every_select_constant_is_enrolled` below is why the next module
# cannot be forgotten the same way.
#
# The concise row is here even though that pair has always agreed. It is not
# coverage for its own sake: the `$select` it pins is one `list_events` already
# sends, so nothing is invented to make the row possible — which is what was
# ruled out when `outlook_get_contact`, which sends no `$select` at all, was
# considered for a row. It also inherits a contract that used to have its own
# test: #73 asserted that concise mode does not pay for `showAs`, and this
# guard's second direction says that generally, for every field, so the one-off
# could go.
_PAIRS = [
    ("mail summary", mail_read._format_message_summary, mail_read.SUMMARY_SELECT),
    ("contact summary", contacts._format_contact_summary, contacts._LIST_SELECT),
    ("event summary", calendar_read._format_event_summary, calendar_read._SUMMARY_SELECT),
    ("event concise", calendar_read._format_event_concise, calendar_read._CONCISE_SELECT),
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


# A ``*_SELECT`` constant that deliberately has no ``_PAIRS`` row, and why.
# Keep this as small as the reasons justify: every entry is a `$select` nothing
# compares against its formatter.
_UNENROLLED = {
    (
        "outlook_mcp.tools.contacts",
        "_SUMMARY_SELECT",
    ): (
        "the search path's select, deliberately narrower than the listing's. "
        "`_format_contact_summary` omits the `categories` key entirely when the "
        "field was not selected, rather than reporting it empty, so the second "
        "direction of the guard would fail on a contract that is correct. The "
        "listing's `_LIST_SELECT` — the widest one that formatter is used with "
        "— carries the row."
    ),
}


def _select_constants() -> dict[tuple[str, str], str]:
    """Every module-level ``*_SELECT`` string under ``outlook_mcp.tools``."""
    import importlib
    import pkgutil

    import outlook_mcp.tools

    found: dict[tuple[str, str], str] = {}
    for info in pkgutil.iter_modules(outlook_mcp.tools.__path__):
        module = importlib.import_module(f"outlook_mcp.tools.{info.name}")
        for name, value in vars(module).items():
            if name.endswith("_SELECT") and isinstance(value, str):
                found[(module.__name__, name)] = value
    return found


def test_every_select_constant_is_enrolled():
    """The general form of #69, which was one module's worth of the same gap.

    Adding a ``_PAIRS`` row for calendar fixes the instance. It does nothing
    about the next module to hoist a ``$select`` to a constant and not think of
    this file — which is precisely how calendar shipped unguarded for seven
    releases while the guard it needed already existed, built for #65.

    So the enrolment is derived rather than remembered: every module-level
    ``*_SELECT`` under ``outlook_mcp.tools`` must appear in ``_PAIRS`` or be
    named in ``_UNENROLLED`` with a reason. Matching is by value, because that
    is what a row actually checks — a second constant holding a string already
    enrolled is the same contract under another name, and is covered.

    ``_UNENROLLED`` is asserted to be live in both directions: an entry naming a
    constant that no longer exists is removed rather than left as a comment
    about nothing.
    """
    constants = _select_constants()
    enrolled = {select for _, _, select in _PAIRS}

    stale = sorted(key for key in _UNENROLLED if key not in constants)
    assert not stale, (
        f"_UNENROLLED names {stale}, which no longer exist. Drop the entry — an "
        f"exemption for a constant that is gone reads as coverage."
    )

    missing = sorted(
        f"{module.split('.')[-1]}.{name}"
        for (module, name), select in constants.items()
        if select not in enrolled and (module, name) not in _UNENROLLED
    )
    assert not missing, (
        f"{missing} is a $select with no _PAIRS row, so nothing checks it against "
        f"the formatter that reads its response. That is #65 and #69 — both shipped "
        f"exactly this way. Add a row, or add it to _UNENROLLED with the reason an "
        f"honest row is impossible."
    )


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
