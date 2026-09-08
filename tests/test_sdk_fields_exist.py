"""Every attribute a tool assigns on an SDK model must be a real field of it.

Found by the wire-level payload tests: ``send_message`` did
``msg.sensitivity = Sensitivity.Private`` for its entire life, and msgraph-sdk's
``Message`` has **no ``sensitivity`` field** — not in the dataclass, not in
``serialize()``, not in the deserializers. The assignment just hung a stray
Python attribute on the instance; kiota never saw it, so it never reached the
wire, so the parameter never did anything. No error anywhere.

That's a whole class: a typo (``msg.is_read_reciept_requested``), a field the
SDK removed in an upgrade, a property that only exists on a sibling model.
Dataclass instances accept arbitrary attributes, so nothing complains.

This test resolves, statically, every ``<var>.<attr> = ...`` in ``tools/``
where ``<var>`` was bound by ``<var> = <SdkClass>()`` in the same function,
and asserts ``<attr>`` is a declared field of ``<SdkClass>``. Offline, no
mocks, no network.
"""

from __future__ import annotations

import ast
import importlib
import pathlib

SRC_TOOLS = pathlib.Path(__file__).resolve().parent.parent / "src" / "outlook_mcp" / "tools"

# Attributes that are legitimately not dataclass fields.
NOT_FIELDS = {"additional_data", "backing_store", "odata_type"}


def _sdk_class_bindings(fn: ast.AST) -> dict[str, tuple[str, str]]:
    """var name -> (module path, class name) for `var = Class()` where Class was
    imported from msgraph.generated.models inside this function or module."""
    imported: dict[str, tuple[str, str]] = {}
    for node in ast.walk(fn):
        if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith(
            "msgraph.generated.models"
        ):
            for alias in node.names:
                imported[alias.asname or alias.name] = (node.module, alias.name)
    bindings: dict[str, tuple[str, str]] = {}
    for node in ast.walk(fn):
        if (
            isinstance(node, ast.Assign)
            and len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and isinstance(node.value, ast.Call)
            and isinstance(node.value.func, ast.Name)
            and node.value.func.id in imported
        ):
            bindings[node.targets[0].id] = imported[node.value.func.id]
    return bindings


def _attribute_assignments(fn: ast.AST) -> list[tuple[int, str, str]]:
    out = []
    for node in ast.walk(fn):
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            target = node.targets[0]
            if isinstance(target, ast.Attribute) and isinstance(target.value, ast.Name):
                out.append((node.lineno, target.value.id, target.attr))
    return out


def _fields_of(module_path: str, class_name: str) -> set[str]:
    cls = getattr(importlib.import_module(module_path), class_name)
    return set(getattr(cls, "__dataclass_fields__", {}))


def test_every_assigned_sdk_attribute_is_a_declared_field():
    offenders: list[str] = []
    checked = 0

    for path in sorted(SRC_TOOLS.glob("*.py")):
        tree = ast.parse(path.read_text())
        module_bindings = _sdk_class_bindings(tree)
        for fn in ast.walk(tree):
            if not isinstance(fn, ast.FunctionDef | ast.AsyncFunctionDef):
                continue
            bindings = {**module_bindings, **_sdk_class_bindings(fn)}
            for lineno, var, attr in _attribute_assignments(fn):
                if var not in bindings or attr in NOT_FIELDS:
                    continue
                module_path, class_name = bindings[var]
                checked += 1
                if attr not in _fields_of(module_path, class_name):
                    offenders.append(
                        f"  {path.relative_to(SRC_TOOLS.parent.parent.parent)}:{lineno}  "
                        f"{var}.{attr} — {class_name} has no field {attr!r}"
                    )

    assert checked > 50, f"sanity: only {checked} assignments resolved — the walker is broken"
    assert not offenders, (
        "Assignment(s) to attributes the SDK model does not declare. These never "
        "serialize and never reach Graph:\n" + "\n".join(offenders)
    )


def test_walker_catches_a_planted_phantom_field(tmp_path):
    """Mutation check: the guard must fail on exactly the sensitivity shape."""
    phantom = SRC_TOOLS / "_phantom_for_test.py"
    phantom.write_text(
        "def build():\n"
        "    from msgraph.generated.models.message import Message\n"
        "    msg = Message()\n"
        "    msg.subject = 'ok'\n"
        "    msg.sensitivity = 'private'\n"
        "    return msg\n"
    )
    try:
        tree = ast.parse(phantom.read_text())
        fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef))
        bindings = _sdk_class_bindings(fn)
        bad = [
            attr
            for _, var, attr in _attribute_assignments(fn)
            if var in bindings and attr not in _fields_of(*bindings[var])
        ]
        assert bad == ["sensitivity"]
    finally:
        phantom.unlink()
