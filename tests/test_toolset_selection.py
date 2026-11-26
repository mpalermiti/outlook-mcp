"""Where the tool selection comes from, and proof that adding a source moved nothing.

A selection can now come from two places: the ``OUTLOOK_MCP_TOOLSETS``
environment variable, as before, or a ``toolsets`` key in config.json. The
variable wins when it is set and not blank; neither means the default set,
which ``classic`` also names. This file pins that table, the ``classic`` alias
and the warning for a name that is not a group.

Most of it is the inertness proof. No config.json written before this change
carries a ``toolsets`` key, so every existing install must load exactly what it
loaded before, for every value of the variable it might have set. ``_ON_MAIN``
below is the selection logic as it stood before (main at de410c5), copied and
frozen rather than imported, so it cannot drift along with the code under test.
The matrix compares it with what a fresh server process really registers, for
each config shape an existing install can have.

The one deliberate difference is ``classic`` itself: it was not a group name
before, so ``OUTLOOK_MCP_TOOLSETS=classic`` loaded only the two account tools.
It now loads the default set, which is what the name says.

``intents`` became a group when ``outlook_find`` landed. It is off by default, so
the frozen copy below lists it with the other opt-in groups, and a selection that
doesn't name it loads what it did. One that names it now gets the intent tools:
it was a name that matched nothing before, warned about at every start. Since the
follow-up core (contracts §10.5), it also brings the seven classic tools the intent
tools' suggestions point to (``toolsets.FOLLOW_UP_CORE``): ``_on_main_or_core`` adds
them for a value that names ``intents``, and for nothing else.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from outlook_mcp import config as config_mod
from outlook_mcp import server, toolsets

ACCOUNT_TOOLS = {"outlook_auth_status", "outlook_whoami"}


# ── The selection as it stood before (main at de410c5), frozen ─────────────


def _on_main(env_value: str | None, mailbox_settings_write: bool) -> set[str]:
    """What main loaded for this environment value and config flag.

    A frozen copy of ``parse_toolsets`` + ``resolve_groups`` + ``select_kept_tools``
    as they were, with the off-by-default groups spelt out. A group that becomes
    off by default later (as the intent tools' will) has to be added to the set
    here too: this pins the surface those installs had, not the constants.
    """
    enabled = None
    if env_value is not None:
        enabled = {part.strip().lower() for part in env_value.split(",") if part.strip()} or None
    if enabled is None:
        groups = set(toolsets.TOOL_GROUPS.values()) - {
            "documents",
            "rules",
            "mime",
            "contact_photos",
            "intents",
        }
        if mailbox_settings_write:
            groups |= {"rules"}
    else:
        groups = enabled
    return {
        name
        for name, group in toolsets.TOOL_GROUPS.items()
        if group == "account" or group in groups
    }


_DEFAULT_GROUPS = sorted(set(toolsets.TOOL_GROUPS.values()) - toolsets.OFF_BY_DEFAULT)
_ALL_GROUPS = sorted(set(toolsets.TOOL_GROUPS.values()))

# Every value an existing install could have set: unset, blank, each group on
# its own, the documented examples, every group, names that match nothing, and
# the spelling variations the parser forgives. `classic` is not here: it is the
# one value whose meaning changes on purpose (see the module docstring).
EXISTING_ENV_VALUES: list[str | None] = [
    None,
    "",
    "   ",
    ",",
    *_ALL_GROUPS,
    "mail,calendar,digest,delta",
    "mail,attachments,documents",
    "mail,calendar,contacts",
    "todo",
    ",".join(_DEFAULT_GROUPS),
    ",".join(_DEFAULT_GROUPS + ["rules"]),
    ",".join(_ALL_GROUPS),
    " Mail , CALENDAR ",
    "mail,,calendar,",
    "foo",
    "mail,foo",
    "intents",
    "intents,mail",
    "default",
    "all",
]


def _names_intents(env_value: str | None) -> bool:
    return env_value is not None and "intents" in {
        part.strip().lower() for part in env_value.split(",")
    }


def _on_main_or_core(env_value: str | None, mailbox_settings_write: bool) -> set[str]:
    """Main's selection, plus the follow-up core for a value that names ``intents``:
    the one deliberate change since, and it reaches no other value."""
    found = _on_main(env_value, mailbox_settings_write)
    return found | toolsets.FOLLOW_UP_CORE if _names_intents(env_value) else found


def _now(env_value: str | None, file_value: str | None, mailbox_settings_write: bool) -> set[str]:
    """What this branch loads, through the same calls server.py makes at import."""
    enabled, _ = toolsets.resolve_selection(env_value, file_value)
    return set(toolsets.loaded_tools(toolsets.resolve_groups(enabled, mailbox_settings_write)))


@pytest.mark.parametrize("settings_write", [False, True])
@pytest.mark.parametrize("env_value", EXISTING_ENV_VALUES)
def test_every_existing_environment_value_loads_what_it_did(env_value, settings_write):
    """No existing config has a ``toolsets`` key, so the file value is None."""
    assert _now(env_value, None, settings_write) == _on_main_or_core(env_value, settings_write)


@pytest.mark.parametrize("settings_write", [False, True])
@pytest.mark.parametrize("file_value", [None, "", "  ", ","])
def test_a_blank_file_value_is_the_same_as_none(file_value, settings_write):
    for env_value in EXISTING_ENV_VALUES:
        assert _now(env_value, file_value, settings_write) == _on_main_or_core(
            env_value, settings_write
        )


# ── The table: environment over file over default ────────────────────────


@pytest.mark.parametrize(
    ("env_value", "file_value", "expected"),
    [
        ("mail", "calendar", ({"mail"}, "env")),
        ("mail", None, ({"mail"}, "env")),
        (None, "calendar", ({"calendar"}, "config")),
        ("", "calendar", ({"calendar"}, "config")),
        ("  ", "calendar", ({"calendar"}, "config")),
        (None, None, (None, "default")),
        ("", "", (None, "default")),
        (None, " , ", (None, "default")),
        (None, "Classic", ({"classic"}, "config")),
    ],
)
def test_resolve_selection_follows_the_table(env_value, file_value, expected):
    assert toolsets.resolve_selection(env_value, file_value) == expected


def test_toolset_names_keep_order_and_drop_repeats():
    """One grammar for the variable, the file key and the CLI: lowercased, stripped."""
    assert toolsets.toolset_names(" Mail, calendar ,MAIL,, todo") == ["mail", "calendar", "todo"]
    assert toolsets.toolset_names(None) == []
    assert toolsets.parse_toolsets(" Mail, calendar ,MAIL") == {"mail", "calendar"}


# ── classic ─────────────────────────────────────────────────────────────


@pytest.mark.parametrize("settings_write", [False, True])
def test_classic_means_exactly_what_nothing_set_means(settings_write):
    assert toolsets.resolve_groups({"classic"}, settings_write) == toolsets.resolve_groups(
        None, settings_write
    )
    assert _now(None, "classic", settings_write) == _on_main(None, settings_write)
    assert _now("classic", None, settings_write) == _on_main(None, settings_write)


@pytest.mark.parametrize("settings_write", [False, True])
def test_classic_with_a_named_group_adds_it_to_the_default(settings_write):
    groups = toolsets.resolve_groups({"classic", "documents"}, settings_write)
    assert groups == toolsets.resolve_groups(None, settings_write) | {"documents"}
    assert "classic" not in groups
    assert ("rules" in groups) is settings_write


def test_classic_is_known_and_is_not_a_group():
    assert "classic" in toolsets.known_toolset_names()
    assert "classic" not in set(toolsets.TOOL_GROUPS.values())
    assert toolsets.unknown_toolset_names({"classic", "mail"}) == set()


# ── Unknown names, and the `intents` group ──────────────────────────────


def test_unknown_names_are_the_ones_that_are_neither_groups_nor_aliases():
    assert toolsets.unknown_toolset_names(None) == set()
    assert toolsets.unknown_toolset_names({"mail", "foo", "Bar"}) == {"foo", "Bar"}


@pytest.mark.parametrize("settings_write", [False, True])
def test_intents_is_an_opt_in_group_that_adds_the_intent_tools(settings_write):
    """Named, it adds the intent tools; unnamed, nothing changes, the default included."""
    intent_tools = {n for n, g in toolsets.TOOL_GROUPS.items() if g == toolsets.INTENTS}
    assert {"outlook_find", "outlook_agenda"} <= intent_tools
    assert "intents" in toolsets.known_toolset_names()
    assert toolsets.unknown_toolset_names({"intents"}) == set()
    assert not intent_tools & _now(None, None, settings_write)
    core = toolsets.FOLLOW_UP_CORE
    assert _now("intents", None, settings_write) == ACCOUNT_TOOLS | intent_tools | core
    # The core is in the default set already, so `classic,intents` gains only the intents.
    assert (
        _now(None, "classic,intents", settings_write)
        == _on_main(None, settings_write) | intent_tools
    )
    assert _now("intents,mail", None, settings_write) == _on_main("mail", settings_write) | (
        intent_tools | core
    )
    # Not named: the core's tools load by their home groups only.
    assert _now("todo", None, settings_write) == _on_main("todo", settings_write)


def test_a_group_with_tools_called_intents_must_be_off_by_default():
    """The intent tools are opt-in while they are a preview."""
    assert "intents" in set(toolsets.TOOL_GROUPS.values())
    assert "intents" in toolsets.OFF_BY_DEFAULT


def test_loaded_tools_is_what_configure_keeps():
    for groups in (
        toolsets.resolve_groups(None, False),
        toolsets.resolve_groups(None, True),
        {"mail"},
        {"foo"},
        set(toolsets.TOOL_GROUPS.values()),
    ):
        assert toolsets.loaded_tools(groups) == frozenset(
            toolsets.select_kept_tools(list(toolsets.TOOL_GROUPS), groups)
        )


def test_this_process_registered_what_it_precomputed():
    assert set(server.mcp._tool_manager._tools) == set(server.LOADED_TOOLS)


# ── What the server says at start ─────────────────────────────────────────


@pytest.mark.parametrize(
    ("source", "where"),
    [("env", "OUTLOOK_MCP_TOOLSETS"), ("config", "config.json toolsets")],
)
def test_an_unknown_name_is_warned_about_with_its_source(source, where, caplog):
    loaded = toolsets.loaded_tools({"mail", "foo"})
    with caplog.at_level(logging.WARNING):
        server._report_the_selection({"mail", "foo"}, source, loaded)

    unknown = [r.getMessage() for r in caplog.records if "does not have" in r.getMessage()]
    assert len(unknown) == 1, caplog.records
    assert where in unknown[0] and "'foo'" in unknown[0]
    assert "classic" in unknown[0] and "mail" in unknown[0]


def test_a_selection_that_loads_only_the_account_tools_says_so(caplog):
    with caplog.at_level(logging.WARNING):
        server._report_the_selection({"nothing_here"}, "env", frozenset(ACCOUNT_TOOLS))

    messages = [r.getMessage() for r in caplog.records]
    assert len(messages) == 2, messages
    assert "'nothing_here'" in messages[0]
    assert "account" in messages[1] and "OUTLOOK_MCP_TOOLSETS" in messages[1]


@pytest.mark.parametrize(
    ("source", "enabled"),
    [("default", None), ("env", {"mail"}), ("env", {"classic"}), ("env", {"classic", "documents"})],
)
def test_known_names_from_the_environment_or_the_default_are_quiet(source, enabled, caplog):
    with caplog.at_level(logging.WARNING):
        server._report_the_selection(
            enabled, source, toolsets.loaded_tools(toolsets.resolve_groups(enabled, False))
        )
    assert caplog.records == []


def test_a_selection_from_config_json_is_never_silent(caplog):
    """A key written before this release, ignored until now, takes effect: say where from."""
    loaded = toolsets.loaded_tools({"mail", "calendar"})
    with caplog.at_level(logging.WARNING):
        server._report_the_selection({"mail", "calendar"}, "config", loaded)

    messages = [r.getMessage() for r in caplog.records]
    assert len(messages) == 1, messages
    assert messages[0].startswith("Tool groups from config.json toolsets: calendar, mail")
    assert f"{len(loaded)} of {len(toolsets.TOOL_GROUPS)} tools" in messages[0]
    assert "outlook-mcp toolsets classic" in messages[0]


# ── Reading the file at import ───────────────────────────────────────────


def _write_config(directory: str, settings: dict) -> None:
    path = Path(directory) / "config.json"
    path.write_text(json.dumps(settings), encoding="utf-8")
    path.chmod(0o600)


def test_the_file_value_is_read_quietly_at_import(caplog):
    _write_config(config_mod.DEFAULT_CONFIG_DIR, {"toolsets": "mail", "unknown_xyz": 1})
    with caplog.at_level(logging.WARNING):
        assert server._toolsets_from_config() == "mail"
    assert caplog.records == []


@pytest.mark.parametrize("value", [["mail", "calendar"], 5, 2.5, True, {"mail": True}])
def test_a_toolsets_key_of_another_type_is_ignored_with_one_warning(value, caplog, monkeypatch):
    """Main ignored the key, whatever it held; a list must not now refuse the whole file."""
    monkeypatch.setattr(config_mod, "_toolsets_type_reported", False)
    _write_config(
        config_mod.DEFAULT_CONFIG_DIR,
        {"client_id": "test-client", "read_only": True, "toolsets": value},
    )

    with caplog.at_level(logging.WARNING):
        first = config_mod.load_config()
        second = config_mod.load_config()

    for config in (first, second):
        assert (config.client_id, config.read_only, config.toolsets) == ("test-client", True, None)
    warned = [r.getMessage() for r in caplog.records if r.getMessage().startswith(TYPE_WARNING)]
    assert len(warned) == 1, [r.getMessage() for r in caplog.records]
    assert "string" in warned[0] and "default" in warned[0]
    assert server._toolsets_from_config() is None


def test_the_import_time_peek_does_not_spend_the_one_warning(caplog, monkeypatch):
    monkeypatch.setattr(config_mod, "_toolsets_type_reported", False)
    _write_config(config_mod.DEFAULT_CONFIG_DIR, {"toolsets": ["mail"]})
    with caplog.at_level(logging.WARNING):
        assert server._toolsets_from_config() is None
        assert caplog.records == []
        config_mod.load_config()
    assert [r for r in caplog.records if r.getMessage().startswith(TYPE_WARNING)]


def test_a_config_that_will_not_load_gives_classic_not_account_only():
    (Path(config_mod.DEFAULT_CONFIG_DIR) / "config.json").write_text("{nope", encoding="utf-8")
    assert server._toolsets_from_config() is None
    assert toolsets.resolve_selection(None, server._toolsets_from_config()) == (None, "default")


# ── Fresh processes: what really registers ────────────────────────────────

# Warnings are collected from the package's loggers before the import that
# prints them, rather than read back from stderr, where the SDK's log handler
# wraps lines to the width of a terminal the child does not have. After the
# import, the probe also runs the load `server.main` runs before the transport
# starts, and `outlook-mcp auth`, `status` and `toolsets` run first: a config
# that main loaded must still load, with its other settings intact.
_PROBE = """
import asyncio, json, logging
warnings = []
class Keep(logging.Handler):
    def emit(self, record):
        warnings.append(record.getMessage())
logging.getLogger("outlook_mcp").addHandler(Keep(logging.WARNING))
from outlook_mcp import server
tools = asyncio.run(server.mcp.list_tools())
at_import = list(warnings)
del warnings[:]
try:
    config = server.load_config()
    loaded = {"ok": True, "client_id": config.client_id, "read_only": config.read_only,
              "toolsets": config.toolsets}
except Exception as exc:
    loaded = {"ok": False, "error": type(exc).__name__}
print(json.dumps({
    "tools": sorted(t.name for t in tools),
    "loaded": sorted(server.LOADED_TOOLS),
    "source": server.SELECTION_SOURCE,
    "documents_bullet": "outlook_read_attachment" in server.INSTRUCTIONS,
    "warnings": at_import,
    "config": loaded,
    "load_warnings": list(warnings),
}))
"""


def _probe(settings: str, env_value: str | None) -> dict:
    env = {**os.environ, "OUTLOOK_MCP_CONFIG_DIR": settings}
    env.pop("OUTLOOK_MCP_TOOLSETS", None)
    env.pop("PYTHONPATH", None)
    if env_value is not None:
        env["OUTLOOK_MCP_TOOLSETS"] = env_value
    result = subprocess.run(
        [sys.executable, "-c", _PROBE],
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
        check=True,
    )
    return json.loads(result.stdout.strip().splitlines()[-1])


# The config shapes an existing install can have: (contents, the
# `mailbox_settings_write` value main's import-time peek read, whether main
# loaded it, the selection this branch reads from it). None means no
# config.json at all: a fresh install, or one that runs on the defaults.
#
# The `toolsets` shapes are configs written before the key existed, where main
# warned about it and ignored it. A string is the one this branch honours: that
# is the feature, and the only change to what loads. Any other type is warned
# about and ignored, as main did, and must never stop the config loading.
CONFIG_SHAPES: dict[str, tuple[object, bool, bool, str | None]] = {
    "no file": (None, False, True, None),
    "client id only": ({"client_id": "test-client"}, False, True, None),
    "read only": (
        {"client_id": "test-client", "read_only": True, "allow_categories": ["mail_drafts"]},
        False,
        True,
        None,
    ),
    "rule writes": ({"client_id": "test-client", "mailbox_settings_write": True}, True, True, None),
    "unknown and legacy keys": (
        {"client_id": "test-client", "accounts": [], "toolset": "mail"},
        False,
        True,
        None,
    ),
    "not JSON": ("{not json", False, False, None),
    "invalid value": (
        {"client_id": "test-client", "mailbox_settings_write": True, "read_only_consent": True},
        False,
        False,
        None,
    ),
    "toolsets string": (
        {"client_id": "test-client", "read_only": True, "toolsets": "mail"},
        False,
        True,
        "mail",
    ),
    "toolsets empty string": ({"client_id": "test-client", "toolsets": ""}, False, True, None),
    "toolsets null": ({"client_id": "test-client", "toolsets": None}, False, True, None),
    "toolsets list": (
        {"client_id": "test-client", "read_only": True, "toolsets": ["mail", "calendar"]},
        False,
        True,
        None,
    ),
    "toolsets list with rule writes": (
        {"client_id": "test-client", "mailbox_settings_write": True, "toolsets": ["rules"]},
        True,
        True,
        None,
    ),
    "toolsets number": ({"client_id": "test-client", "toolsets": 5}, False, True, None),
    "toolsets object": (
        {"client_id": "test-client", "toolsets": {"mail": True}},
        False,
        True,
        None,
    ),
}

PROCESS_ENV_VALUES: list[str | None] = [
    None,
    "",
    "mail,calendar,digest,delta",
    "rules",
    "mail,attachments,documents",
    "foo",
    ",".join(_ALL_GROUPS),
]

TYPE_WARNING = "config.json key 'toolsets' ignored"


def _settings_for(base: Path, name: str, shape: object) -> str:
    directory = base / name.replace(" ", "-")
    directory.mkdir()
    if isinstance(shape, dict):
        _write_config(str(directory), shape)
    elif isinstance(shape, str):
        (directory / "config.json").write_text(shape, encoding="utf-8")
    return str(directory)


def test_every_existing_config_shape_loads_and_registers_what_main_did(tmp_path):
    """The matrix, in fresh interpreters, because the selection is made at import."""
    cases = []
    for name, (shape, settings_write, loads, selection) in CONFIG_SHAPES.items():
        settings = _settings_for(tmp_path, name, shape)
        for env_value in PROCESS_ENV_VALUES:
            cases.append((name, shape, settings, settings_write, loads, selection, env_value))

    with ThreadPoolExecutor(max_workers=8) as pool:
        results = list(pool.map(lambda case: _probe(case[2], case[6]), cases))

    for (name, shape, _, settings_write, loads, selection, env_value), measured in zip(
        cases, results
    ):
        label = f"{name} / OUTLOOK_MCP_TOOLSETS={env_value!r}"
        env_wins = toolsets.parse_toolsets(env_value) is not None
        expected = _on_main(env_value if env_wins else selection, settings_write)
        from_file = not env_wins and selection is not None
        assert set(measured["tools"]) == expected, label
        assert set(measured["loaded"]) == expected, label
        assert measured["source"] == ("config" if from_file else "env" if env_wins else "default")
        assert measured["documents_bullet"] is ("outlook_read_attachment" in expected), label

        # It loads wherever main's did, and the other settings come through.
        config = measured["config"]
        assert config["ok"] is loads, (label, config)
        if loads and isinstance(shape, dict):
            assert config["client_id"] == "test-client", label
            assert config["read_only"] is shape.get("read_only", False), label
            given = shape.get("toolsets")
            assert config["toolsets"] == (given if isinstance(given, str) else None), label

        # A key of the wrong type is named once, by the real load, never at import.
        bad_type = isinstance(shape, dict) and not isinstance(
            shape.get("toolsets"), (str, type(None))
        )
        type_warnings = [w for w in measured["load_warnings"] if w.startswith(TYPE_WARNING)]
        assert len(type_warnings) == (1 if bad_type else 0), (label, measured["load_warnings"])
        assert not [w for w in measured["warnings"] if w.startswith(TYPE_WARNING)], label

        # At start: unknown names are warned about, and a selection read from
        # config.json is never silent.
        unknown = [w for w in measured["warnings"] if "does not have" in w]
        notice = [w for w in measured["warnings"] if w.startswith("Tool groups from config.json")]
        assert bool(unknown) is (env_value == "foo"), (label, measured["warnings"])
        assert len(notice) == (1 if from_file else 0), (label, measured["warnings"])


def test_a_symlinked_config_still_registers_the_default(tmp_path):
    real = tmp_path / "real.json"
    real.write_text(json.dumps({"toolsets": "mail", "mailbox_settings_write": True}))
    settings = tmp_path / "settings"
    settings.mkdir()
    (settings / "config.json").symlink_to(real)

    measured = _probe(str(settings), None)

    assert set(measured["tools"]) == _on_main(None, False)
    assert measured["source"] == "default"


def test_the_file_key_selects_and_the_environment_overrides_it(tmp_path):
    settings = _settings_for(tmp_path, "file", {"toolsets": "todo", "mailbox_settings_write": True})

    from_file = _probe(settings, None)
    blank_env = _probe(settings, "  ")
    from_env = _probe(settings, "calendar")

    todo = {n for n, g in toolsets.TOOL_GROUPS.items() if g in {"todo", "account"}}
    calendar = {n for n, g in toolsets.TOOL_GROUPS.items() if g in {"calendar", "account"}}
    assert (set(from_file["tools"]), from_file["source"]) == (todo, "config")
    assert (set(blank_env["tools"]), blank_env["source"]) == (todo, "config")
    assert (set(from_env["tools"]), from_env["source"]) == (calendar, "env")


@pytest.mark.parametrize(("settings_write", "rules"), [(False, False), (True, True)])
def test_classic_in_the_file_is_the_default_with_and_without_rule_writes(
    tmp_path, settings_write, rules
):
    settings = _settings_for(
        tmp_path, "classic", {"toolsets": "classic", "mailbox_settings_write": settings_write}
    )
    measured = _probe(settings, None)

    assert set(measured["tools"]) == _on_main(None, settings_write)
    assert bool(set(measured["tools"]) & {"outlook_create_rule"}) is rules
    assert measured["source"] == "config"


def test_an_unknown_name_in_the_file_warns_once_and_names_the_file(tmp_path):
    settings = _settings_for(tmp_path, "unknown", {"toolsets": "nothing_here"})
    measured = _probe(settings, None)

    assert set(measured["tools"]) == ACCOUNT_TOOLS
    warnings = measured["warnings"]
    assert len(warnings) == 3, warnings
    assert warnings[0].startswith("Tool groups from config.json toolsets: nothing_here")
    assert "'nothing_here'" in warnings[1] and "config.json toolsets" in warnings[1]
    assert "account" in warnings[2] and "config.json toolsets" in warnings[2]
    assert not [w for w in warnings if "OUTLOOK_MCP_TOOLSETS" in w], warnings
