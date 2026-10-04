"""CLI entry point: `outlook-mcp auth` and `outlook-mcp serve`."""

from __future__ import annotations

import re
import sys

from azure.core.exceptions import ClientAuthenticationError
from pydantic import ValidationError

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import DEFAULT_CONFIG_DIR, Config, config_repair_lines, load_config
from outlook_mcp.errors import OutlookMCPError

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
