"""CLI entry point: `outlook-mcp serve`, `auth`, `status`, `logout` and `toolsets`."""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import NoReturn

from azure.core.exceptions import ClientAuthenticationError
from pydantic import ValidationError

from outlook_mcp import toolsets
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import (
    DEFAULT_CONFIG_DIR,
    Config,
    _ensure_dir,
    atomic_write,
    config_repair_lines,
    load_config,
)
from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.toolsets import TOOLSETS_ENV

# The AADSTS65xxx consent refusals (65001 user or admin has not consented,
# 65004 user declined, 65005 the app asks for permissions the resource never
# offered) — the one failure class the app registration can actually cause.
_CONSENT_REFUSAL = re.compile(r"AADSTS65\d{3}")


def _is_consent_refusal(exc: BaseException) -> bool:
    """True for the AADSTS65xxx refusals an app registration can cause.

    Every other refusal — a timed-out device code, a declined sign-in, a
    generic invalid grant — says nothing about the registration, and the
    registration remedy would point the operator at the wrong thing.
    """
    return _CONSENT_REFUSAL.search(str(exc)) is not None


def _load_config_or_exit() -> Config:
    """Load the config, or exit with the repair instead of a traceback.

    The same failure set the server exits on before its transport starts:
    an invalid value, a refused symlink, an unreadable file or directory,
    a file that is neither UTF-8 nor a valid config in the machine's code
    page, a settings path that is a file.
    """
    try:
        return load_config()
    except (ValidationError, OSError, ValueError) as exc:
        for line in config_repair_lines(exc):
            print(line, file=sys.stderr)
        sys.exit(1)


def _print_usage() -> None:
    print("Usage: outlook-mcp <command>")
    print()
    print("Commands:")
    print("  serve    Start the MCP server (default, used by OpenClaw)")
    print("  auth     Authenticate with Microsoft (device code flow)")
    print("  status   Check authentication status")
    print("  logout   Clear cached credentials")
    print("  toolsets Show the tool groups this server loads; with names, set them")


def _selection_lines(config: Config) -> list[str]:
    """The tool selection in force, where it came from, and what it loads.

    The same resolution the server makes at import, from the same two inputs,
    so this can answer without importing (and registering) the server.
    """
    enabled, source = toolsets.resolve_selection(os.environ.get(TOOLSETS_ENV), config.toolsets)
    loaded = toolsets.loaded_tools(toolsets.resolve_groups(enabled, config.mailbox_settings_write))
    if source == "default":
        selection = f'classic, the default ({TOOLSETS_ENV} and config.json "toolsets" are not set)'
    else:
        where = TOOLSETS_ENV if source == "env" else 'config.json "toolsets"'
        selection = f"{', '.join(sorted(enabled or ()))} (from {where})"
    lines = [
        f"Toolsets:  {selection}",
        f"Groups:    {', '.join(sorted({toolsets.TOOL_GROUPS[name] for name in loaded}))}",
        f"Tools:     {len(loaded)} of {len(toolsets.TOOL_GROUPS)}",
    ]
    unknown = toolsets.unknown_toolset_names(enabled)
    if unknown:
        names = ", ".join(repr(name) for name in sorted(unknown))
        what = "not a tool group" if len(unknown) == 1 else "not tool groups"
        load = "it loads" if len(unknown) == 1 else "they load"
        lines.append(f"Unknown:   {names} ({what} in this release; {load} nothing)")
    return lines


def _known_names_line() -> str:
    groups = ", ".join(sorted(set(toolsets.TOOL_GROUPS.values())))
    return f"Tool groups: {groups}; or {' or '.join(sorted(toolsets.ALIASES))} for the default set."


def _refuse(code: int, *lines: str) -> NoReturn:
    for line in lines:
        print(line, file=sys.stderr)
    sys.exit(code)


def _write_toolsets(value: str) -> Path:
    """Set ``toolsets`` in config.json to ``value``, and nothing else.

    An existing file must be a regular file that loads with the normal loader;
    its other keys and values stay as written, in their order, with no defaults
    filled in. A missing file is created holding only this key. Written through
    the same owner-only directory and atomic file write as every settings file.
    """
    config_dir = DEFAULT_CONFIG_DIR
    path = Path(config_dir) / "config.json"
    # Before loading: the loader refuses a symlink too, but reading a FIFO
    # would block, and a directory here is a settings mistake, not a config.
    if path.is_symlink() or (path.exists() and not path.is_file()):
        _refuse(
            1,
            f"Refusing to change {path}: it is a symlink or not a regular file.",
            "Replace it with a real file, then run this again. Nothing was written.",
        )
    try:
        load_config(config_dir)
    except (ValidationError, OSError, ValueError) as exc:
        _refuse(1, *config_repair_lines(exc), "Nothing was written.")
    settings: dict = {}
    if path.exists():
        try:
            # UTF-8 only. The loader still reads a file saved in this machine's
            # code page (as releases before #99 did), but writing one back here
            # would change the bytes of every other value in it, so the person
            # is asked to re-save it as UTF-8 first (UnicodeDecodeError's repair).
            settings = json.loads(path.read_bytes().decode("utf-8-sig"))
        except (OSError, ValueError) as exc:
            _refuse(1, *config_repair_lines(exc), "Nothing was written.")
    settings["toolsets"] = value
    atomic_write(
        _ensure_dir(config_dir) / "config.json",
        json.dumps(settings, indent=2, ensure_ascii=False) + "\n",
    )
    return path


def cmd_toolsets(args: list[str]) -> None:
    """Show the selection in force, or write config.json's ``toolsets``.

    With no names: the selection, its source and the tool count. With names
    (comma-separated, or as separate arguments): each must be a tool group or
    ``classic``, or nothing is written and the exit code is 2 — at server start
    an unknown name is only warned about, but here the person is choosing.
    """
    if not args:
        for line in _selection_lines(_load_config_or_exit()):
            print(line)
        return
    names = toolsets.toolset_names(",".join(args))
    if not names:
        _refuse(
            2, "Name at least one tool group, or classic. Nothing was written.", _known_names_line()
        )
    unknown = [name for name in names if name not in toolsets.known_toolset_names()]
    if unknown:
        _refuse(
            2,
            f"Not a tool group in this release: {', '.join(repr(n) for n in unknown)}. "
            "Nothing was written.",
            _known_names_line(),
        )
    value = ",".join(names)
    path = _write_toolsets(value)
    print(f'Wrote "toolsets": "{value}" to {path}.')
    if toolsets.parse_toolsets(os.environ.get(TOOLSETS_ENV)) is not None:
        print(f"{TOOLSETS_ENV} is set here and overrides this file.")
    print("Restart your MCP client to load the new tools.")


def cmd_auth() -> None:
    """Interactive device code auth — run this in a terminal."""
    config = _load_config_or_exit()
    if not config.client_id:
        print("Error: client_id not configured.")
        print(f"Set client_id in {DEFAULT_CONFIG_DIR}/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    # The read-write set whatever read_only says: that flag gates the tools,
    # not the token, and a consent narrowed by it could never be widened once
    # the config flips. Only the explicit read_only_consent key asks for less.
    mode = "read-only" if config.read_only_consent else "read-write"
    print(f"Authenticating with the {mode} scopes...")
    if config.read_only and not config.read_only_consent:
        # The one setup this can bite: client_id pointing at a second, read-only
        # app registration, which 1.23.0 signed in with `.default` and so never
        # asked for write access. Said before the browser opens, because
        # accepting that consent screen is not undone by changing the config.
        print(
            "Note: read_only is set, but this sign-in asks for write access too. "
            "If client_id is a read-only app registration, stop here (Ctrl-C), set "
            "read_only_consent: true, and run `outlook-mcp auth` again."
        )
    print()

    try:
        auth.login_interactive()
    except OutlookMCPError as exc:
        # Structured failures carry their own remedy (e.g. the unencrypted
        # cache refusal names the config flag and the system packages) —
        # print it, not a traceback.
        print(str(exc), file=sys.stderr)
        sys.exit(1)
    except ClientAuthenticationError as exc:
        # The sign-in itself was refused. The concrete scopes make a consent
        # refusal reachable for a mis-registered app: a registration missing
        # one of the delegated permissions fails here and now, where a
        # .default consent used to "succeed" and strand the session instead.
        # Only the AADSTS65xxx refusals say anything about the registration,
        # so only those get that remedy; Azure's text is printed either way.
        print(f"Sign-in was refused: {exc}", file=sys.stderr)
        if _is_consent_refusal(exc):
            print(
                "Check that the app registration carries every delegated "
                "permission the README's registration step lists, then run "
                "`outlook-mcp auth` again.",
                file=sys.stderr,
            )
        sys.exit(1)
    print()
    print("Done. The MCP server will use this cached token automatically.")


def cmd_status() -> None:
    """Check if a cached token exists and is usable."""
    config = _load_config_or_exit()
    if not config.client_id:
        print(f"Not configured — set client_id in {DEFAULT_CONFIG_DIR}/config.json")
        sys.exit(1)

    auth = AuthManager(config)

    print(f"Client ID: {config.client_id[:8]}...")
    print(f"Tenant:    {config.tenant_id}")
    print(f"Mode:      {'read-only' if config.read_only else 'read-write'}")
    for line in _selection_lines(config):
        print(line)
    print()

    try:
        authenticated = auth.try_cached_token()
    except OutlookMCPError as exc:
        # A refusal rather than a stale token (the plaintext-cache one is
        # re-raised): it carries its own remedy, which a traceback would bury.
        print("Status: not authenticated")
        print(str(exc))
        return
    if authenticated:
        print("Status: authenticated (cached token valid)")
    else:
        print("Status: not authenticated")
        if auth.startup_error is not None:
            # The refresh already named its own remedy (e.g. the AADSTS70000
            # dead end, whose only exit is a fresh login) — print it instead
            # of the generic line, which reads as "any re-auth will do".
            print(str(auth.startup_error))
        else:
            print("Run: outlook-mcp auth")


def cmd_logout() -> None:
    """Remove this instance's auth record; report what stays behind."""
    auth = AuthManager(_load_config_or_exit())
    if auth.logout()["record_removed"]:
        print(f"Removed this instance's auth record ({DEFAULT_CONFIG_DIR}/auth_record.json).")
    else:
        print("No auth record was present for this instance.")
    print("This server will ask for `outlook-mcp auth` again on next start.")
    print()
    print(
        "The encrypted token cache the OS keeps for azure-identity "
        "(Keychain item Microsoft.Developer.IdentityService on macOS, its"
    )
    print("equivalent on other systems) is shared across apps and left in place; its tokens")
    print("age out on their own.")


def cmd_serve() -> None:
    """Start the MCP server (stdio by default; optional streamable HTTP)."""
    from outlook_mcp.server import main as serve_main

    serve_main()


def main() -> None:
    """CLI dispatcher."""
    args = sys.argv[1:]

    if not args or args[0] == "serve":
        cmd_serve()
    elif args[0] == "auth":
        cmd_auth()
    elif args[0] == "status":
        cmd_status()
    elif args[0] == "logout":
        cmd_logout()
    elif args[0] == "toolsets":
        cmd_toolsets(args[1:])
    elif args[0] in ("-h", "--help", "help"):
        _print_usage()
    else:
        # Unknown arg — assume it's the MCP server (backwards compat)
        cmd_serve()
