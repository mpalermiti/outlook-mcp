"""CLI entry point: `outlook-mcp auth` and `outlook-mcp serve`."""

from __future__ import annotations

import json
import sys

from pydantic import ValidationError

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config, load_config


def _print_usage() -> None:
    print("Usage: outlook-mcp <command> [account]")
    print()
    print("Commands:")
    print("  serve    Start the MCP server (default, used by OpenClaw)")
    print("  auth     Authenticate with Microsoft (device code flow)")
    print("           Optional account name when 'accounts' is configured:")
    print("           outlook-mcp auth net")
    print("  status   Check authentication status (all accounts)")
    print("  logout   Clear cached credentials")
    print("           Optional account name: outlook-mcp logout net")


def _load_config_or_exit() -> Config:
    """Load config, or exit with the fix spelled out.

    Config shapes older versions accepted are normalized inside the
    validators (accept-and-warn); reaching an error here means the file is
    genuinely unparseable — and every command, `logout` included, must say
    what to fix rather than raw-traceback (review of #61).
    """
    try:
        return load_config()
    except ValidationError as exc:
        print("Error: ~/.outlook-mcp/config.json is invalid:")
        for e in exc.errors():
            field = ".".join(str(p) for p in e["loc"]) or "config"
            print(f"  {field}: {e['msg']}")
        print("Fix the file and run the command again.")
        sys.exit(2)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"Error: cannot read ~/.outlook-mcp/config.json: {exc}")
        sys.exit(2)


def _resolve_account_arg(config: Config, arg: str | None) -> str | None:
    """Validate an account argument against the config, or exit with help."""
    if arg is None:
        return None
    names = [acc.name for acc in config.accounts]
    if not names:
        print("Error: no 'accounts' configured in ~/.outlook-mcp/config.json —")
        print("single-account installs authenticate without an account name.")
        sys.exit(1)
    if arg not in names:
        print(f"Error: unknown account '{arg}'. Configured accounts: {', '.join(names)}")
        sys.exit(1)
    return arg


def cmd_auth(account: str | None = None) -> None:
    """Interactive device code auth — run this in a terminal."""
    config = _load_config_or_exit()
    # Validate an explicit name before printing anything: the device-code
    # banner must not appear for an account that cannot exist.
    account = _resolve_account_arg(config, account)
    if config.accounts and account is None:
        account = config.default_account
    client_id = config.client_id if account is None else None
    if account is not None:
        client_id = next((a.client_id for a in config.accounts if a.name == account), None)
    if not client_id:
        print("Error: client_id not configured.")
        print("Set client_id in ~/.outlook-mcp/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    mode = "read-only" if config.read_only else "read-write"
    target = f"account '{account}'" if account else "default account"
    print(f"Authenticating {target} with {mode} scopes...")
    if account:
        print("Sign in to THIS account's identity in the browser when prompted.")
    print()

    auth.login_interactive(account)
    print()
    print("Done. The MCP server will use this cached token automatically.")


def cmd_status() -> None:
    """Check if cached tokens exist and are usable."""
    config = _load_config_or_exit()
    if not config.client_id and not config.accounts:
        print("Not configured — set client_id in ~/.outlook-mcp/config.json")
        sys.exit(1)

    auth = AuthManager(config)
    mode = "read-only" if config.read_only else "read-write"

    print(f"Mode:      {mode}")
    if config.capability_accounts:
        print("Routing:")
        for capability, name in sorted(config.capability_accounts.items()):
            print(f"  {capability:<10} -> {name}")
        print(f"  {'(other)':<10} -> {config.default_account}")
    print()

    if not config.accounts:
        print(f"Client ID: {config.client_id[:8]}...")
        print(f"Tenant:    {config.tenant_id}")
        print()
        if auth.try_cached_token():
            print("Status: authenticated (cached token valid)")
        else:
            print("Status: not authenticated")
            print("Run: outlook-mcp auth")
        return

    ok = auth.try_cached_token()
    authenticated = set(auth.authenticated_accounts)
    for acc in config.accounts:
        status = "authenticated" if acc.name in authenticated else "not authenticated"
        marker = " (default)" if acc.name == config.default_account else ""
        print(f"  {acc.name}: {status}{marker}")
    if not ok:
        print()
        print("Run: outlook-mcp auth <account> for each account you need.")


def cmd_logout(account: str | None = None) -> None:
    """Clear cached credentials for one account (default: the default account)."""
    config = _load_config_or_exit()
    if account:
        _resolve_account_arg(config, account)
    auth = AuthManager(config)
    result = auth.logout(account)
    target = account or (
        config.default_account if config.accounts else "the (single) configured account"
    )
    print(result["message"])
    print()
    if config.accounts:
        print(f"Account '{target}' is logged out: its auth record is deleted and")
        print("the server will raise auth errors for it on next start — it does")
        print("NOT fall back to another account. Other accounts keep working.")
        print()
        print("To clear every account, run logout once per account name.")
    else:
        print("The MCP server will require re-authentication on next start.")
    print()
    # All accounts share ONE token cache entry (the MSAL cache is
    # multi-account; per-account cache names do not exist on macOS — see
    # auth.CACHE_NAME). DeviceCodeCredential exposes no cache-clear API.
    print("Auth records are removed by this command. Cached tokens remain in")
    print('the system store under "outlook-mcp" — remove that entry from')
    print("Keychain Access (macOS) or the credential store on your OS to wipe them.")


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
        cmd_auth(args[1] if len(args) > 1 else None)
    elif args[0] == "status":
        cmd_status()
    elif args[0] == "logout":
        cmd_logout(args[1] if len(args) > 1 else None)
    elif args[0] in ("-h", "--help", "help"):
        _print_usage()
    else:
        # Unknown arg — assume it's the MCP server (backwards compat)
        cmd_serve()
