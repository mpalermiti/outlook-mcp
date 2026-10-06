"""The README's app-registration step and the consent list must name the same set.

#102 was exactly this drift: the first consent asked for what the code listed,
the registration step told the user to grant a different set, and a permission
one tool needs (`MailboxSettings.Read` for `outlook_list_categories`) sat in
neither. Both sides are now generated from the same intent, and this guard
keeps them together — a scope added to one and forgotten in the other, in
either direction, reddens here.
"""

import re
from pathlib import Path

from outlook_mcp import auth as auth_module

README = Path(__file__).resolve().parent.parent / "README.md"


def _readme_registration_scopes() -> set[str]:
    """Parse the delegated-permission bullets out of README step 5.

    Anchored on the step's own heading (``5. Go to **API permissions**``) and
    bounded by the blank line that ends it, so a permission mentioned
    elsewhere in the README never leaks into the set.
    """
    text = README.read_text(encoding="utf-8")
    m = re.search(r"^5\. Go to \*\*API permissions\*\*.*?\n\n", text, re.MULTILINE | re.DOTALL)
    assert m, "README step 5 (API permissions) not found — was it renumbered or moved?"
    scopes: set[str] = set()
    for backticked in re.findall(r"`([^`]+)`", m.group(0)):
        scopes.update(s.strip() for s in backticked.split(",") if s.strip())
    return scopes


def test_the_readme_registration_step_and_the_consent_list_name_the_same_set():
    """offline_access is reserved for MSAL (it appends it itself); everything
    else the README tells the user to grant must be exactly what the first
    consent asks for — and vice versa."""
    readme = _readme_registration_scopes()
    consented = set(auth_module.SCOPES_READWRITE)

    assert "offline_access" in readme, (
        "README step 5 dropped offline_access — MSAL adds it to token requests "
        "itself, but the app registration still needs it for the refresh flow"
    )
    assert "offline_access" not in consented
    assert readme - {"offline_access"} == consented
