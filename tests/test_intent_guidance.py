"""Instructions, prompts and `next` name only loaded tools (contracts §10.5, §10.9, §10.10).

The step-6 evaluation compares a classic session with an intents one, so an intents
session must never send the agent looking for a tool it doesn't have. For every
selection that includes `intents`, every tool named in a `next` hint, an
instruction bullet or a prompt is loaded:

- **next.** The intent tools' suggestions point at the follow-up core, which
  `intents` now loads with them (`toolsets.EXTRA_GROUPS`), and `build_next`
  drops any tool the process lacks.
- **Instructions.** Under `intents`, each working-rule bullet is sent only when every
  tool it names is loaded, and the intent bullets are added (`guidance`).
- **Prompts.** Under `intents`, a prompt is offered only when every tool of one of
  its variants is loaded, and the intents variant wins when both qualify.

A selection without `intents` gets exactly what main sends: the classic
instructions byte for byte, and all three classic prompts, ungated.
"""

from __future__ import annotations

import ast
import json
import re
from pathlib import Path

import pytest

from outlook_mcp import guidance, toolsets
from outlook_mcp.intents import ids, transport
from outlook_mcp.intents.ids import IdFamily
from tests.intent_fakes import FakeGraph

ROOT = Path(__file__).resolve().parent.parent
INTENTS_DIR = ROOT / "src" / "outlook_mcp" / "intents"
TOOL_NAME = re.compile(r"outlook_[a-z_]+")
INTENT_TOOLS = frozenset(n for n, g in toolsets.TOOL_GROUPS.items() if g == toolsets.INTENTS)
GROUPS = sorted(set(toolsets.TOOL_GROUPS.values()) - toolsets.ALWAYS_ON_GROUPS - {toolsets.INTENTS})
# Hint targets the contracts name that this release does not have yet (the step 5-7
# tools): `build_next` drops them, so they never reach an answer. One that becomes
# real must then load with `intents`, which the static test below checks.
FUTURE_TOOLS = frozenset({"outlook_compose", "outlook_follow_up", "outlook_answer_invites"})

# main's INSTRUCTIONS (2ce198a), as a stock install receives them: what every selection
# without `intents` must still get, byte for byte.
MAIN_INSTRUCTIONS = (ROOT / "tests" / "data" / "main_instructions.txt").read_text("utf-8")


def _loaded(selection: set[str] | None, settings_write: bool = False) -> frozenset[str]:
    return toolsets.loaded_tools(toolsets.resolve_groups(selection, settings_write))


def _intent_selections() -> list[tuple[str, frozenset[str]]]:
    """Every selection that includes `intents`: alone, with each other group, with
    `classic`, and with everything; each with and without `mailbox_settings_write`."""
    picks = [{"intents"}, {"intents", "classic"}, set(toolsets.TOOL_GROUPS.values())]
    picks += [{"intents", group} for group in GROUPS]
    found = []
    for pick in picks:
        for flag in (False, True):
            name = ",".join(sorted(pick)) + (" +settings_write" if flag else "")
            found.append((name, _loaded(pick, flag)))
    return found


def _classic_selections() -> list[tuple[str, frozenset[str]]]:
    """Selections without `intents`: unset, `classic`, every other group, and each alone."""
    picks: list[set[str] | None] = [None, {"classic"}, set(GROUPS)]
    picks += [{group} for group in GROUPS]
    found = []
    for pick in picks:
        for flag in (False, True):
            label = "unset" if pick is None else ",".join(sorted(pick))
            found.append((label + (" +settings_write" if flag else ""), _loaded(pick, flag)))
    return found


INTENT_SELECTIONS = _intent_selections()
CLASSIC_SELECTIONS = _classic_selections()


def test_the_selections_are_what_they_say():
    for _, loaded in INTENT_SELECTIONS:
        assert INTENT_TOOLS <= loaded
    for _, loaded in CLASSIC_SELECTIONS:
        assert not INTENT_TOOLS & loaded


# ── The follow-up core (§10.5) ────────────────────────────────────────────


def test_intents_alone_loads_the_follow_up_core():
    assert _loaded({"intents"}) == (
        INTENT_TOOLS | toolsets.FOLLOW_UP_CORE | {"outlook_auth_status", "outlook_whoami"}
    )


def test_the_core_is_seven_classic_tools_from_the_default_set():
    assert toolsets.FOLLOW_UP_CORE == {
        "outlook_read_message",
        "outlook_read_messages",
        "outlook_list_thread",
        "outlook_download_attachment",
        "outlook_update_task",
        "outlook_complete_task",
        "outlook_list_drafts",
    }
    default = _loaded(None)
    for name in toolsets.FOLLOW_UP_CORE:
        assert toolsets.TOOL_GROUPS[name] not in toolsets.OFF_BY_DEFAULT
        assert name in default
        assert toolsets.groups_of(name) == {toolsets.TOOL_GROUPS[name], toolsets.INTENTS}


def test_one_rule_for_both_paths():
    """§10.5: the unset path and the named default path agree for every tool, so no
    group leaks into a stock install through the second membership."""
    names = list(toolsets.TOOL_GROUPS)
    assert toolsets.select_kept_tools(names, None) == toolsets.select_kept_tools(
        names, toolsets.resolve_groups(None, False)
    )


@pytest.mark.parametrize(("label", "loaded"), CLASSIC_SELECTIONS)
def test_a_selection_without_intents_loads_by_home_group_only(label, loaded):
    """The second membership is `intents` alone: no selection without it changes."""
    base = label.replace(" +settings_write", "")
    pick = None if base == "unset" else set(base.split(","))
    groups = toolsets.resolve_groups(pick, label.endswith("+settings_write"))
    assert loaded == {
        n for n, g in toolsets.TOOL_GROUPS.items() if g in groups or g in toolsets.ALWAYS_ON_GROUPS
    }


# ── Instructions (§10.9) ──────────────────────────────────────────────────


@pytest.mark.parametrize(("label", "loaded"), CLASSIC_SELECTIONS)
def test_without_intents_the_instructions_are_mains_byte_for_byte(label, loaded):
    expected = MAIN_INSTRUCTIONS
    if "outlook_read_attachment" in loaded:
        expected += guidance.DOCUMENTS_RULE
    assert guidance.build_instructions(loaded) == expected


@pytest.mark.parametrize(("label", "loaded"), INTENT_SELECTIONS)
def test_with_intents_the_instructions_name_only_loaded_tools(label, loaded):
    text = guidance.build_instructions(loaded)
    named = set(TOOL_NAME.findall(text))
    assert named <= loaded, f"{label}: names {sorted(named - loaded)}"
    assert "anything under `others_wrote` was written by someone else" in text
    assert "this_week" in text and "+7d" in text  # the typed time forms
    assert "`candidates`" in text and "`incomplete`" in text
    assert "cursor" in text  # outlook_catch_up always loads with intents
    assert text.startswith(MAIN_INSTRUCTIONS[: MAIN_INSTRUCTIONS.index("Working rules")])


def test_intents_alone_keeps_the_rules_whose_tools_it_has():
    text = guidance.build_instructions(_loaded({"intents"}))
    assert "outlook_read_messages once" in text  # the core has both read tools
    assert "outlook_find_event_by_key" not in text  # calendar isn't loaded
    assert "Dates accept ISO 8601" in text  # names no tool
    assert len(text) <= 4_000


def _bullets(text: str) -> list[str]:
    """The working-rule bullets of an instructions text, each with its wrapped lines."""
    return ["- " + b for b in text.split("\n- ")[1:]]


def test_the_candidates_rule_never_invites_an_unoffered_call():
    """§1.7 rule 5: no `choose` for a candidate that fails the known test or carries a
    blocking warning, so the instructions must never suggest calling again with an
    address or id the result did not offer (review F1)."""
    text = guidance.build_instructions(_loaded({"intents"}))
    (rule,) = [b for b in _bullets(text) if "`candidates`" in b]
    assert "call again only with a candidate offered in `next`" in rule
    assert "if none is, ask the user" in rule
    for bullet in _bullets(text):
        if "call again" in bullet:
            assert "offered in `next`" in bullet, bullet
            assert not re.search(r"\b(address|id)\b", bullet.split("call again", 1)[1]), bullet


@pytest.mark.parametrize(("label", "loaded"), INTENT_SELECTIONS)
def test_the_instructions_stay_within_4000_characters_under_intents(label, loaded):
    """§10.9's cap holds for every selection with `intents`, everything loaded included."""
    assert len(guidance.build_instructions(loaded)) <= 4_000, label


def _tools_taking(parameter: str) -> frozenset[str]:
    """Every tool whose signature in server.py has ``parameter``."""
    tree = ast.parse((ROOT / "src" / "outlook_mcp" / "server.py").read_text("utf-8"))
    return frozenset(
        node.name
        for node in ast.walk(tree)
        if isinstance(node, ast.AsyncFunctionDef | ast.FunctionDef)
        and node.name in toolsets.TOOL_GROUPS
        and parameter in {a.arg for a in (*node.args.args, *node.args.kwonlyargs)}
    )


def test_a_parameter_rule_names_every_tool_that_takes_it():
    """The parameter rules' tool sets are what the tools' signatures say."""
    for marker, takers in guidance.PARAMETER_RULES.items():
        parameter = marker.split("=")[0]
        assert takers == _tools_taking(parameter), marker


def test_under_intents_a_parameter_rule_goes_only_where_a_tool_takes_it():
    """Nit: `intents` alone loads no tool that takes a body, so it isn't told about
    is_markdown=True; it keeps concise=True, which its thread and draft tools take."""
    alone = guidance.build_instructions(_loaded({"intents"}))
    assert "is_markdown=True" not in alone
    assert "concise=True" in alone
    both = guidance.build_instructions(_loaded({"classic", "intents"}))
    assert "is_markdown=True" in both


def test_classic_intents_keeps_every_classic_rule():
    text = guidance.build_instructions(_loaded({"classic", "intents"}))
    for rule in guidance.CLASSIC_RULES:
        assert rule in text


def test_every_bullet_is_sent_whole_or_not_at_all():
    """A rule's `requires` is every tool it names, read from its own text."""
    for rule in (*guidance.CLASSIC_RULES, guidance.DOCUMENTS_RULE, *guidance.INTENT_RULES):
        assert rule.startswith("- ") and rule.endswith("\n")
        assert guidance.tools_named(rule) == frozenset(TOOL_NAME.findall(rule))


# ── Prompts (§10.10) ──────────────────────────────────────────────────────

# main's three prompts, as a stock install serves them with their default arguments.
MAIN_PROMPTS = json.loads((ROOT / "tests" / "data" / "main_prompts.json").read_text("utf-8"))


@pytest.mark.parametrize(("label", "loaded"), CLASSIC_SELECTIONS)
def test_without_intents_the_prompts_are_mains(label, loaded):
    chosen = guidance.choose_prompts(loaded)
    assert [(p.name, p.form) for p in chosen] == [
        ("morning_brief", "classic"),
        ("triage_folder", "classic"),
        ("catch_up", "classic"),
    ]
    for prompt in chosen:
        main = MAIN_PROMPTS[prompt.name]
        assert (prompt.title, prompt.description) == (main["title"], main["description"])
        assert prompt.render() == main["text"]


@pytest.mark.parametrize(("label", "loaded"), INTENT_SELECTIONS)
def test_with_intents_every_prompt_names_only_loaded_tools(label, loaded):
    for prompt in guidance.choose_prompts(loaded):
        named = set(TOOL_NAME.findall(prompt.render()))
        assert named <= loaded, f"{label}: {prompt.name} names {sorted(named - loaded)}"


def test_the_intents_variant_wins_and_a_prompt_without_its_tools_is_not_offered():
    alone = {p.name: p.form for p in guidance.choose_prompts(_loaded({"intents"}))}
    assert alone == {"morning_brief": "intents", "catch_up": "intents"}
    both = {p.name: p.form for p in guidance.choose_prompts(_loaded({"classic", "intents"}))}
    assert both == {"morning_brief": "intents", "triage_folder": "classic", "catch_up": "intents"}


def test_the_intents_prompts_say_what_the_contracts_say():
    by_name = {p.name: p for p in guidance.choose_prompts(_loaded({"intents"}))}
    brief = by_name["morning_brief"].render()
    assert brief.index('outlook_agenda(window="today")') < brief.index(
        'outlook_catch_up(window="today")'
    )
    catch_up = by_name["catch_up"].render(since="7d")
    assert "outlook_catch_up(window='7d')" in catch_up
    assert "cursor" in catch_up and "replac" in catch_up


# ── next (§1.7) ───────────────────────────────────────────────────────────


def _targets_in(source: str, name: str) -> set[str]:
    """Every tool a hint in ``source`` names: a `suggestion(...)` call, however it is
    reached (`suggestion`, `envelope.suggestion`) and whether the tool is passed by
    position or keyword, and a hint written as a dict with `template` and `tool` keys.
    A target that can't be read is returned as `<unread:...>`, never skipped."""
    tree = ast.parse(source)
    constants = {
        node.targets[0].id: node.value.value
        for node in ast.walk(tree)
        if isinstance(node, ast.Assign)
        and len(node.targets) == 1
        and isinstance(node.targets[0], ast.Name)
        and isinstance(node.value, ast.Constant)
        and isinstance(node.value.value, str)
    }

    def read(target: ast.AST | None, line: int) -> str:
        if isinstance(target, ast.Constant) and isinstance(target.value, str):
            return target.value
        if isinstance(target, ast.Name) and target.id in constants:
            return constants[target.id]
        return f"<unread:{name}:{line}>"

    found: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            called = func.id if isinstance(func, ast.Name) else getattr(func, "attr", None)
            if called == "suggestion":
                given = node.args[1] if len(node.args) > 1 else None
                given = given or next((k.value for k in node.keywords if k.arg == "tool"), None)
                found.add(read(given, node.lineno))
        if isinstance(node, ast.Dict):
            keys = {
                k.value: v
                for k, v in zip(node.keys, node.values, strict=True)
                if isinstance(k, ast.Constant)
            }
            if "template" in keys and "tool" in keys:
                found.add(read(keys["tool"], node.lineno))
    return found


def _hint_targets() -> set[str]:
    """Every tool a hint in `intents/` names, read from the source, plus find's table of
    them (`_NEXT`). envelope.py defines `suggestion` and is skipped."""
    found: set[str] = set()
    for path in INTENTS_DIR.glob("*.py"):
        if path.name == "envelope.py":
            continue
        source = path.read_text(encoding="utf-8")
        targets = _targets_in(source, path.name)
        if path.name == "find.py":  # find passes the tool from `_NEXT`
            targets = {t for t in targets if not t.startswith("<unread")}
            targets |= {
                node.value
                for node in ast.walk(ast.parse(source))
                if isinstance(node, ast.Constant)
                and isinstance(node.value, str)
                and TOOL_NAME.fullmatch(node.value)
                and node.value != "outlook_find"
            }
        found |= targets
    return found


def test_the_hint_reader_has_no_blind_spots():
    """Review nit: an attribute call, a keyword tool and a dict hint are read; a target
    it can't read is flagged, never skipped."""
    source = (
        "from outlook_mcp.intents import envelope\n"
        "TOOL = 'outlook_d'\n"
        "a = envelope.suggestion('t', 'outlook_a', {})\n"
        "b = suggestion('t', tool='outlook_b', args={})\n"
        "c = {'template': 't', 'tool': 'outlook_c', 'args': {}, 'needs': []}\n"
        "d = suggestion('t', TOOL, {})\n"
        "e = suggestion('t', pick(), {})\n"
    )
    assert _targets_in(source, "x.py") == {
        "outlook_a",
        "outlook_b",
        "outlook_c",
        "outlook_d",
        "<unread:x.py:7>",
    }


def test_every_hint_target_is_read():
    targets = _hint_targets()
    assert not [t for t in targets if t.startswith("<unread")], targets
    assert {"outlook_read_message", "outlook_list_thread", "outlook_brief"} <= targets


@pytest.mark.parametrize(("label", "loaded"), INTENT_SELECTIONS)
def test_every_hint_target_that_exists_is_loaded_with_intents(label, loaded):
    targets = _hint_targets()
    real = {t for t in targets if t in toolsets.TOOL_GROUPS}
    assert real <= loaded, f"{label}: {sorted(real - loaded)}"
    assert targets - real <= FUTURE_TOOLS
    assert not FUTURE_TOOLS & set(toolsets.TOOL_GROUPS), "a future tool exists: load it"


@pytest.fixture
def fake(monkeypatch):
    graph = FakeGraph()
    monkeypatch.setattr(transport, "make_client", graph.http_client)
    monkeypatch.setattr(ids, "ID_FAMILY", IdFamily())
    return graph


async def _answers(fake, loaded) -> list[dict]:
    """One answer from each intent tool on its own test fixture, with `loaded`."""
    from tests import (
        test_intent_agenda,
        test_intent_brief,
        test_intent_catch_up,
        test_intent_find,
        test_intent_open_loops,
    )

    answers = []
    runs = [
        (test_intent_find, {"kind": "message", "words": "budget"}),
        (test_intent_find, {"kind": "thread", "words": "budget"}),
        (test_intent_find, {"kind": "person", "words": "jones"}),
        (test_intent_agenda, {}),
        (test_intent_open_loops, {}),
        (test_intent_catch_up, {}),
        (test_intent_brief, {"person": "sam"}),
        (test_intent_brief, {"meeting": "budget sync"}),
    ]
    for module, args in runs:
        fake.routes.clear()
        answers.append(await module.run(fake, tools=loaded, **args))
    return answers


@pytest.mark.parametrize("selection", [{"intents"}, {"intents", "mail"}, {"intents", "classic"}])
async def test_every_suggestion_in_an_answer_is_loaded(fake, selection):
    loaded = _loaded(selection)
    offered = []
    for env in await _answers(fake, loaded):
        offered += [s["tool"] for s in env["next"]]
    assert set(offered) <= loaded
    # The core is what keeps these hints in the answer under `intents` alone.
    assert {"outlook_read_message", "outlook_list_thread", "outlook_brief"} <= set(offered)
