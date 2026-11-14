"""azure-identity stays below 1.27 until a sign-in saved by 1.26 is moved across (#244).

On macOS, 1.27.0 keeps a named token cache under a Keychain account named after
the cache, where 1.26 and earlier used one shared ``MSALCache`` account. An
install that resolves 1.27 cannot find the sign-in saved before it, and reports
"Not authenticated" until ``outlook-mcp auth`` runs again. Dependabot never
edits bounds, so the cap only moves by hand: lift it, and delete this file, in
the change that teaches ``auth.py`` to move an existing sign-in across.
"""

from __future__ import annotations

import importlib.metadata
import re
from pathlib import Path

PYPROJECT = Path(__file__).resolve().parent.parent / "pyproject.toml"


def _major_minor(version: str) -> tuple[int, int]:
    major, minor = version.split(".")[:2]
    return int(major), int(minor)


def test_the_requirement_caps_azure_identity_below_1_27():
    match = re.search(r'"azure-identity([^"]*)"', PYPROJECT.read_text(encoding="utf-8"))
    assert match, "azure-identity is no longer a direct dependency in pyproject.toml"
    assert "<1.27" in match.group(1).replace(" ", "")


def test_the_installed_azure_identity_is_below_1_27():
    assert _major_minor(importlib.metadata.version("azure-identity")) < (1, 27)
