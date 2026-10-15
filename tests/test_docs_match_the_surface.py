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

import re
from pathlib import Path

import pytest
from mcp.client import Client

from outlook_mcp import toolsets
from outlook_mcp.permissions import VALID_CATEGORIES
from outlook_mcp.server import mcp

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "SKILL.md").read_text(encoding="utf-8")

REGISTERED = set(mcp._tool_manager._tools)

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
    """`68 tools`, `Tools (68)`, `all 68` and `instead of 68` all mean the whole surface.

    A historical measurement names its own surface with a hyphen (`the 62-tool
    surface`) and is not matched.
    """
    text = (ROOT / doc).read_text(encoding="utf-8")
    totals = re.findall(
        r"\b(\d+) (?:typed Graph |discrete )?tools\b(?! instead)"
        r"|Tools \((\d+)\)|\b[Aa]ll (\d+)\b|instead of (\d+)\b",
        text,
    )
    numbers = {int(n) for groups in totals for n in groups if n}
    assert numbers, f"{doc} states no tool total"
    assert numbers == {len(REGISTERED)}, f"{doc} states {sorted(numbers)}"


def test_the_readme_toolset_example_loads_what_it_says():
    """The example's tool count is what that selection, plus `account`, registers."""
    m = re.search(r'OUTLOOK_MCP_TOOLSETS="([a-z,]+)"', README)
    assert m, "README OUTLOOK_MCP_TOOLSETS example not found"
    groups = set(m.group(1).split(",")) | toolsets.ALWAYS_ON_GROUPS
    loaded = sum(1 for g in toolsets.TOOL_GROUPS.values() if g in groups)
    stated = re.search(r"(\d+) tools instead of", README)
    assert stated and int(stated.group(1)) == loaded


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


def test_the_readme_category_table_is_the_accepted_set():
    table = _section(README, "Granular Write Permissions (optional)")
    rows = set(re.findall(r"^\| `(\w+)` \|", table, re.MULTILINE))
    assert rows == VALID_CATEGORIES
    words = {7: "seven"}
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
