"""CLI entry point: `outlook-mcp auth` and `outlook-mcp serve`."""

from __future__ import annotations

import sys

from pydantic import ValidationError

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import DEFAULT_CONFIG_DIR, Config, config_repair_lines, load_config
from outlook_mcp.errors import OutlookMCPError


def _load_config_or_exit() -> Config:
    """Load the config, or exit with the repair instead of a traceback.

    The same failure set the server exits on before its transport starts:
    an invalid value, a refused symlink, an unreadable file or directory,
    non-UTF-8 bytes, a settings path that is a file.
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


def cmd_auth() -> None:
    """Interactive device code auth — run this in a terminal."""
    config = _load_config_or_exit()
    if not config.client_id:
        print("Error: client_id not configured.")
        print(f"Set client_id in {DEFAULT_CONFIG_DIR}/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    mode = "read-only" if config.read_only else "read-write"
    print(f"Authenticating with {mode} scopes...")
    print()

    try:
        auth.login_interactive()
    except OutlookMCPError as exc:
        # Structured failures carry their own remedy (e.g. the unencrypted
        # cache refusal names the config flag and the system packages) —
        # print it, not a traceback.
        print(str(exc), file=sys.stderr)
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
    print()

    if auth.try_cached_token():
        print("Status: authenticated (cached token valid)")
    else:
        print("Status: not authenticated")
        print("Run: outlook-mcp auth")


def cmd_logout() -> None:
    """Remove this instance's auth record; report what stays behind."""
    auth = AuthManager(_load_config_or_exit())
    if auth.logout()["record_removed"]:
        print("Removed this instance's auth record "
              f"({DEFAULT_CONFIG_DIR}/auth_record.json).")
    else:
        print("No auth record was present for this instance.")
    print("This server will ask for `outlook-mcp auth` again on next start.")
    print()
    print("The encrypted token cache the OS keeps for azure-identity "
          "(Keychain item Microsoft.Developer.IdentityService on macOS, its")
    print("equivalent on other systems) is shared across apps and left in "
          "place; its tokens")
    print("age out on their own.")


def cmd_serve() -> None:
    """Start the MCP stdio server."""
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
    elif args[0] in ("-h", "--help", "help"):
        _print_usage()
    else:
        # Unknown arg — assume it's the MCP server (backwards compat)
        cmd_serve()
