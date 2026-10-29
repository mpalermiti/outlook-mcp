"""The README and SKILL.md must describe the surface the server actually registers.

Agents learn the server from these two files — SKILL.md is what a ClawHub install
reads, and the README is where everyone else starts. A count, a group or a tool list
that has drifted from the code sends an agent looking for something that is not
there. The 2026-10 docs audit found exactly that: the "Concise mode" section named
five tools when six take `concise`, the toolset example promised "~30 tools" for a
selection that loads 28, and the errors were described as `{code, message, action}`
responses when the client receives a tool error whose text carries the message and
the hint. Each check here pins one of those classes, read from the code, so the next
tool added or removed reddens a test instead of quietly making a doc wrong.

Prose descriptions are deliberately not pinned. Only facts the code settles are.
"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest
from mcp.client import Client

from outlook_mcp import server, toolsets
from outlook_mcp.permissions import VALID_CATEGORIES
from outlook_mcp.server import mcp

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "SKILL.md").read_text(encoding="utf-8")

# The tool surface the *project* ships, not the one this process registered:
# the `rules` and `documents` groups are opt-in, so a stock import registers 71
# of the 76 tools these files document. Docs describe what exists, so the
# project surface is read from the classification table rather than from a
# process whose environment decides — and `TOOL_GROUPS` is kept honest against
# the real tools by
# `test_server.py::test_tool_groups_names_every_real_tool_and_nothing_else`.
LOADED = set(mcp._tool_manager._tools)
REGISTERED = set(toolsets.TOOL_GROUPS)


def test_default_surface_omits_only_optional_groups():
    """Pin a fresh default process even when this suite enables every group."""
    probe = (
        "import json\n"
        "from outlook_mcp.server import mcp\n"
        "print(json.dumps(sorted(mcp._tool_manager._tools)))\n"
    )
    env = {**os.environ}
    env.pop("OUTLOOK_MCP_TOOLSETS", None)
    result = subprocess.run(
        [sys.executable, "-c", probe],
        env=env,
        capture_output=True,
        text=True,
        timeout=60,
        check=True,
    )
    default = set(json.loads(result.stdout.strip().splitlines()[-1]))
    assert len(default) == 71
    hidden = toolsets.OFF_BY_DEFAULT
    assert default == toolsets.select_kept_tools(list(REGISTERED), None) - {
        name for name in REGISTERED if toolsets.TOOL_GROUPS.get(name) in hidden
    }
    assert "outlook_read_attachment" not in default
    assert not (
        default & {name for name in REGISTERED if toolsets.TOOL_GROUPS.get(name) == "rules"}
    )


def test_current_surface_follows_the_requested_toolsets():
    # Mirror how server.py configures at import: the raw env selection goes
    # through resolve_groups first, so an unset env yields the default group set
    # (opt-in groups excluded) rather than None — which would keep `rules`.
    enabled = toolsets.resolve_groups(
        toolsets.parse_toolsets(os.environ.get("OUTLOOK_MCP_TOOLSETS")),
        server._mailbox_settings_write_is_on(),
    )
    assert LOADED == toolsets.select_kept_tools(list(REGISTERED), enabled)


# The README's "Features" list names its categories in words; their counts are
# only checkable against its own "Tool Reference" tables, which carry the names.
_FEATURE = re.compile(r"^- \*\*(.+?) \((\d+)\)\*\*", re.MULTILINE)
_README_ROW = re.compile(r"^\| `(outlook_\w+)` \|", re.MULTILINE)
_SKILL_BULLET = re.compile(r"^- `(outlook_\w+)` —", re.MULTILINE)


def _section(text: str, heading: str) -> str:
    """The body under a markdown heading, up to the next heading of any level."""
    m = re.search(rf"^#+ {re.escape(heading)}\n(.*?)(?=^#|\Z)", text, re.MULTILINE | re.DOTALL)
    assert m, f"heading {heading!r} not found — was it renamed?"
    return m.group(1)


def _tool_reference() -> str:
    m = re.search(r"^## Tool Reference\n(.*?)(?=^## )", README, re.MULTILINE | re.DOTALL)
    assert m, "README '## Tool Reference' section not found"
    return m.group(1)


def test_every_registered_tool_has_exactly_one_readme_row():
    rows = _README_ROW.findall(_tool_reference())
    assert len(rows) == len(set(rows)), "a tool is listed twice in the README"
    assert set(rows) == REGISTERED


def test_every_registered_tool_has_exactly_one_skill_bullet():
    bullets = _SKILL_BULLET.findall(SKILL)
    assert len(bullets) == len(set(bullets)), "a tool is listed twice in SKILL.md"
    assert set(bullets) == REGISTERED


def test_each_readme_feature_count_matches_its_reference_table():
    reference = _tool_reference()
    features = _FEATURE.findall(README)
    assert features, "README feature list not found"
    for name, stated in features:
        rows = _README_ROW.findall(_section(reference, name))
        assert int(stated) == len(rows), f"README says {name} ({stated}), its table has {len(rows)}"
    assert sum(int(n) for _, n in features) == len(REGISTERED)
    m = re.search(r"\*\*(\d+) tools\*\* across (\d+) categories", README)
    assert m, "README '**N tools** across M categories' line not found"
    assert (int(m.group(1)), int(m.group(2))) == (len(REGISTERED), len(features))


@pytest.mark.parametrize("doc", ["README.md", "SKILL.md"])
def test_every_total_tool_count_is_the_registered_count(doc):
    """`72 tools`, `Tools (72)` and `all 72` describe the full declared surface.

    A historical measurement names its own surface with a hyphen (`the 62-tool
    surface`) and is not matched. The selected/default comparison is checked
    separately against those two registrations below.
    """
    text = (ROOT / doc).read_text(encoding="utf-8")
    totals = re.findall(
        r"\b(\d+) (?:typed Graph |discrete )?tools\b(?! instead)"
        r"|Tools \((\d+)\)|\b[Aa]ll (\d+)\b",
        text,
    )
    numbers = {int(n) for groups in totals for n in groups if n}
    assert numbers, f"{doc} states no tool total"
    assert numbers == {len(REGISTERED)}, f"{doc} states {sorted(numbers)}"


def test_the_readme_toolset_example_loads_what_it_says():
    """Compare the example's selected and default registrations, including `account`."""
    m = re.search(r'OUTLOOK_MCP_TOOLSETS="([a-z,]+)"', README)
    assert m, "README OUTLOOK_MCP_TOOLSETS example not found"
    selected = toolsets.select_kept_tools(list(REGISTERED), set(m.group(1).split(",")))
    default = toolsets.select_kept_tools(list(REGISTERED), None)
    stated = re.search(r"(\d+) tools instead of (\d+)\b", README)
    assert stated, "README selected/default tool comparison not found"
    assert (int(stated.group(1)), int(stated.group(2))) == (len(selected), len(default))


def test_the_readme_toolset_groups_are_the_selectable_ones():
    m = re.search(r"^Groups: (.+?)\. ", README, re.MULTILINE)
    assert m, "README 'Groups:' line not found"
    named = set(re.findall(r"`(\w+)`", m.group(1)))
    assert named == set(toolsets.TOOL_GROUPS.values()) - toolsets.ALWAYS_ON_GROUPS


def test_the_concise_tools_named_are_the_ones_that_take_concise():
    takes_concise = {
        name
        for name, tool in mcp._tool_manager._tools.items()
        if "concise" in tool.parameters.get("properties", {})
    }
    readme = re.search(r"^- \*\*Concise mode\*\*.*?\((.*?)\)", README, re.MULTILINE)
    skill = re.search(r"^Pass `concise=True` to read tools \((.*?)\)", SKILL, re.MULTILINE)
    assert readme and skill, "the concise-mode tool lists were not found"
    for found in (readme, skill):
        assert set(re.findall(r"`(outlook_\w+)`", found.group(1))) == takes_concise


def test_the_counts_no_table_carries_match_the_surface():
    # Two places state the tool count in prose, and nothing above reads either: the
    # MCP registry's description in server.json, and the README's "isn't for you"
    # line. A tool added without them leaves the registry showing the old number.
    server = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    registry = re.search(r"\((\d+) tools\)", server["description"])
    assert registry, "server.json description no longer says '(N tools)'"
    assert int(registry.group(1)) == len(REGISTERED), "server.json's tool count is stale"
    prose = re.search(r"this has (\d+) tools", README)
    assert prose, "README 'this has N tools' line not found"
    assert int(prose.group(1)) == len(REGISTERED), "the README's 'this has N tools' is stale"


def test_the_registry_description_fits_the_registry():
    # The MCP registry refuses a description over 100 characters with a 422, and
    # it only says so at publish time, after PyPI already has the release: 1.30.0
    # reached PyPI and not the registry for one character.
    server = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert len(server["description"]) <= 100, (
        f"server.json description is {len(server['description'])} characters; "
        "the MCP registry accepts at most 100"
    )


def test_the_readme_category_table_is_the_accepted_set():
    table = _section(README, "Granular Write Permissions (optional)")
    rows = set(re.findall(r"^\| `(\w+)` \|", table, re.MULTILINE))
    assert rows == VALID_CATEGORIES
    words = {7: "seven", 8: "eight"}
    assert f"only the {words[len(VALID_CATEGORIES)]} names above" in table
    assert f"the {len(VALID_CATEGORIES)} categories" in SKILL


@pytest.mark.asyncio
async def test_the_documented_prompts_are_the_registered_ones():
    async with Client(mcp) as client:
        prompts = {
            p.name: {a.name for a in p.arguments or []}
            for p in (await client.list_prompts()).prompts
        }

    table = _section(README, "Prompts")
    readme = {
        name: set(re.findall(r"`(\w+)` \(default", args))
        for name, args in re.findall(r"^\| `(\w+)` \| (.*?) \|", table, re.MULTILINE)
    }
    assert readme == prompts

    skill_section = _section(SKILL, f"Prompts ({len(prompts)})")
    skill = {
        name: set(re.findall(r"(\w+)=", args))
        for name, args in re.findall(r"^- `(\w+)\((.*?)\)`", skill_section, re.MULTILINE)
    }
    assert skill == prompts
