"""CLI entry point: `outlook-mcp auth` and `outlook-mcp serve`."""

from __future__ import annotations

import sys

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import load_config


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
    config = load_config()
    if not config.client_id:
        print("Error: client_id not configured.")
        print("Set client_id in ~/.outlook-mcp/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    scopes = auth.get_scopes()
    mode = "read-only" if config.read_only else "read-write"
    print(f"Authenticating with {mode} scopes...")
    print()

    auth.login_interactive(scopes)
    print()
    print("Done. The MCP server will use this cached token automatically.")


def cmd_status() -> None:
    """Check if a cached token exists and is usable."""
    config = load_config()
    if not config.client_id:
        print("Not configured — set client_id in ~/.outlook-mcp/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    scopes = auth.get_scopes()

    print(f"Client ID: {config.client_id[:8]}...")
    print(f"Tenant:    {config.tenant_id}")
    print(f"Mode:      {'read-only' if config.read_only else 'read-write'}")
    print()

    if auth.try_cached_token(scopes):
        print("Status: authenticated (cached token valid)")
    else:
        print("Status: not authenticated")
        print("Run: outlook-mcp auth")


def cmd_logout() -> None:
    """Clear cached credentials."""
    # Token cache is in the system keychain under "outlook-mcp".
    # DeviceCodeCredential doesn't expose a cache-clear API, so we
    # just inform the user.
    print("To fully clear cached tokens, remove 'outlook-mcp' from")
    print("Keychain Access (macOS) or the credential store on your OS.")
    print()
    print("The MCP server will require re-authentication on next start.")


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
