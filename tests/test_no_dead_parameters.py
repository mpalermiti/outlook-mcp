"""Guard against the #41 failure shape: a parameter declared and never applied.

`outlook_create_event` accepted a `recurrence` argument, passed it to the
handler, and never assigned it to the Graph payload. The call returned
`status: created`, so it looked like it worked, and it survived fourteen
minor releases. The same shape then turned up in `outlook_reply`, where
`is_html` was accepted and ignored so every "HTML" reply went out as plain
text, and in `AuthManager`, where two methods took a `scopes` argument and
resolved scopes internally instead.

All three are statically detectable: the parameter is never read in the body.
This test walks the package and fails if that is true of any parameter, which
would have caught #41 the day it was written — for free, offline, in CI.

It cannot catch the other two shapes of the same bug (Graph accepts a field
and ignores it; the SDK drops a field from the payload). Those need live
coverage and payload-level assertions respectively — see
`tests/test_live_calendar_write.py` and the serialization assertions in
`tests/test_calendar_write.py`.
"""

from __future__ import annotations

import ast
import pathlib

SRC = pathlib.Path(__file__).resolve().parent.parent / "src" / "outlook_mcp"

# Parameters that exist to satisfy a caller's signature and are legitimately
# unread. Keep this list short and justified — every entry is a place the
# guard is switched off. Format: (module suffix, function, parameter, why).
ALLOWED: set[tuple[str, str, str]] = {
    ("server.py", "lifespan", "server"),
    ("auth.py", "_on_device_code", "expires_on"),
}

_REASONS = {
    ("server.py", "lifespan", "server"): (
        "MCPServer calls the lifespan hook with the server instance; we take it "
        "to match the expected signature and read config from elsewhere."
    ),
    ("auth.py", "_on_device_code", "expires_on"): (
        "azure-identity's prompt_callback is invoked with three positional "
        "arguments; only the URI and user code are shown to the user."
    ),
}


def _unread_parameters(path: pathlib.Path) -> list[tuple[int, str, str]]:
    """Return (lineno, function, parameter) for parameters never read in a body."""
    found: list[tuple[int, str, str]] = []
    tree = ast.parse(path.read_text())

    for fn in ast.walk(tree):
        if not isinstance(fn, ast.FunctionDef | ast.AsyncFunctionDef):
            continue

        args = fn.args
        # *args/**kwargs forwarding makes "was it used?" undecidable here.
        if args.vararg or args.kwarg:
            continue

        # Only Name loads count. Deliberately NOT attribute names: counting
        # `event.recurrence` as a use of a `recurrence` parameter is exactly
        # how #41 would slip past this guard.
        loaded = {
            node.id
            for node in ast.walk(fn)
            if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load)
        }

        for param in (a.arg for a in args.posonlyargs + args.args + args.kwonlyargs):
            if param in ("self", "cls") or param in loaded:
                continue
            found.append((fn.lineno, fn.name, param))

    return found


def test_no_parameter_is_declared_and_never_read():
    offenders: list[str] = []

    for path in sorted(SRC.rglob("*.py")):
        for lineno, func, param in _unread_parameters(path):
            if (path.name, func, param) in ALLOWED:
                continue
            rel = path.relative_to(SRC.parent.parent)
            offenders.append(f"  {rel}:{lineno}  {func}()  parameter {param!r} is never read")

    assert not offenders, (
        "Parameter(s) declared but never applied — the #41 shape:\n"
        + "\n".join(offenders)
        + "\n\nEither use the parameter, remove it, or (if a caller's signature "
        "requires it) add it to ALLOWED in this file with a reason."
    )


def test_allowlist_entries_still_exist():
    """A stale allowlist entry silently re-opens the hole it was excusing."""
    live = {
        (path.name, func, param)
        for path in sorted(SRC.rglob("*.py"))
        for _, func, param in _unread_parameters(path)
    }

    stale = ALLOWED - live
    assert not stale, f"Allowlist entries no longer apply and should be deleted: {sorted(stale)}"


def test_every_allowlist_entry_is_justified():
    assert ALLOWED == set(_REASONS), "Every ALLOWED entry needs a reason in _REASONS"
