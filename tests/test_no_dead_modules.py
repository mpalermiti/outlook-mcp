"""Guard against dead modules: every file in the package must be reachable.

`src/outlook_mcp/models/` — 22 Pydantic classes across five files — sat
unreferenced by any tool path for months, while 30 tests imported it
directly and passed. Tests that import a module by name make dead code look
alive. It read as the input-validation layer, CLAUDE.md said it was one, and
it is why #41 went unnoticed: the validator that should have rejected a bad
`recurrence` existed, looked correct, and was never called.

`test_no_dead_parameters.py` catches a dead *parameter*. This catches a dead
*module*: it builds the intra-package import graph from the real entry point
and fails on any module file nothing reaches.

Imports are collected from every node in each file, not just the top level,
because this package imports lazily inside functions (`cli` → `server`,
`mail_read` → `mail_drafts`, `permissions` → `config`). A top-level-only walk
would have declared `server.py` dead.
"""

from __future__ import annotations

import ast
import pathlib

SRC = pathlib.Path(__file__).resolve().parent.parent / "src" / "outlook_mcp"
PACKAGE = "outlook_mcp"

# The console script (`outlook-mcp = "outlook_mcp.cli:main"`). `server` is
# reached from here lazily. If a second entry point is ever added, list it.
ROOTS = {f"{PACKAGE}.cli"}

# Modules that exist for reasons other than being imported. Keep short, with
# a reason. There are none today; the set is here so the next one has a home.
ALLOWED: dict[str, str] = {}


def _module_name(path: pathlib.Path) -> str:
    rel = path.relative_to(SRC.parent).with_suffix("")
    parts = list(rel.parts)
    if parts[-1] == "__init__":
        parts = parts[:-1]
    return ".".join(parts)


def _all_modules() -> dict[str, pathlib.Path]:
    return {_module_name(p): p for p in SRC.rglob("*.py")}


def _imports_of(path: pathlib.Path, module: str, known: set[str]) -> set[str]:
    """Package-internal modules this file imports anywhere in its body."""
    tree = ast.parse(path.read_text())
    package_of = module if path.name == "__init__.py" else module.rpartition(".")[0]
    found: set[str] = set()

    def add(candidate: str) -> None:
        # `from a.b import c` may name a submodule (a.b.c) or an attribute of
        # a.b — add whichever actually exists as a file.
        if candidate in known:
            found.add(candidate)

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name.startswith(PACKAGE):
                    add(alias.name)
        elif isinstance(node, ast.ImportFrom):
            if node.level:  # relative import
                base_parts = package_of.split(".")
                if node.level > 1:
                    base_parts = base_parts[: -(node.level - 1)]
                base = ".".join(base_parts + ([node.module] if node.module else []))
            else:
                base = node.module or ""
            if not base.startswith(PACKAGE):
                continue
            add(base)
            for alias in node.names:
                add(f"{base}.{alias.name}")
    return found


def _reachable() -> set[str]:
    modules = _all_modules()
    known = set(modules)
    seen: set[str] = set()
    frontier = list(ROOTS)
    while frontier:
        mod = frontier.pop()
        if mod in seen or mod not in modules:
            continue
        seen.add(mod)
        # Importing a submodule imports its package __init__ too.
        parent = mod.rpartition(".")[0]
        if parent and parent in modules and parent not in seen:
            frontier.append(parent)
        frontier.extend(_imports_of(modules[mod], mod, known) - seen)
    return seen


def test_every_module_is_reachable_from_the_entry_point():
    modules = _all_modules()
    unreachable = sorted(
        m for m in modules if m not in _reachable() and m not in ALLOWED and m not in ROOTS
    )
    assert not unreachable, (
        "Module(s) nothing on the entry-point import graph reaches — dead code that "
        "looks alive whenever a test imports it directly:\n"
        + "\n".join(f"  {m}  ({modules[m].relative_to(SRC.parent.parent)})" for m in unreachable)
        + "\n\nWire it in, delete it, or add it to ALLOWED in this file with a reason."
    )


def test_roots_exist():
    modules = _all_modules()
    missing = ROOTS - set(modules)
    assert not missing, f"ROOTS names modules that don't exist: {sorted(missing)}"


def test_allowlist_entries_still_exist_and_are_still_unreachable():
    modules = _all_modules()
    reachable = _reachable()
    gone = [m for m in ALLOWED if m not in modules]
    now_reachable = [m for m in ALLOWED if m in reachable]
    assert not gone, f"ALLOWED names modules that no longer exist: {gone}"
    assert not now_reachable, (
        f"ALLOWED entries are reachable now and should be removed: {now_reachable}"
    )


def test_guard_would_have_caught_the_models_package(tmp_path, monkeypatch):
    """Mutation check: plant an unreferenced module and make sure it's flagged."""
    ghost = SRC / "tools" / "_ghost_module_for_test.py"
    ghost.write_text('"""Unreferenced on purpose."""\nVALUE = 1\n')
    try:
        modules = _all_modules()
        assert f"{PACKAGE}.tools._ghost_module_for_test" in modules
        assert f"{PACKAGE}.tools._ghost_module_for_test" not in _reachable()
    finally:
        ghost.unlink()
