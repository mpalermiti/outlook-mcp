# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in outlook-mcp, please report it responsibly.

**Do NOT open a public GitHub issue for security vulnerabilities.**

### How to Report

Use [GitHub Security Advisories](https://github.com/mpalermiti/outlook-mcp/security/advisories/new) to privately report vulnerabilities.

### Response Timeline

- **Acknowledgment:** Within 48 hours
- **Initial assessment:** Within 7 days
- **Fix timeline:** Depends on severity, typically within 30 days

### Scope

The following are in scope for security reports:

- Token leakage or credential exposure
- Authentication bypass
- Input injection (OData, KQL, path traversal)
- Unauthorized file system access
- Supply chain vulnerabilities in dependencies
- Symlink attacks on config files

### Out of Scope

- Social engineering attacks
- Denial of service against Microsoft Graph API endpoints
- Issues in Microsoft Graph API itself
- Issues requiring physical access to the machine

### Security Design

outlook-mcp is designed with security in mind:

- **Tokens:** Stored in the OS keyring via azure-identity — macOS Keychain,
  Windows Credential Store, and libsecret/gnome-keyring on Linux. Where no
  encrypted store is available (Linux without libsecret), the server refuses to
  persist the cache rather than fall back to cleartext; set
  `allow_unencrypted_token_cache: true` in `~/.outlook-mcp/config.json` to
  accept plaintext storage instead.
- **Delta cursors:** `delta_token` is caller-held state and is treated as
  untrusted input. Every URL that receives a Graph bearer token — the cursor and
  each `@odata.nextLink` — is parsed and required to be https on
  `graph.microsoft.com`.
- **Read-only mode is a tool gate, not a token scope.** `read_only: true` blocks this
  server's write tools. It does not narrow the OAuth token, which is acquired with
  `.default` and carries whatever the Azure app was consented for. A `read_only` server
  still holds a write-capable Graph token, and the setting is a config-file value rather
  than anything Microsoft enforces. For a credential that genuinely cannot write, consent a
  separate Azure app to the read scopes only.
- **Input validation:** All Graph IDs, emails, dates, KQL queries, and folder names are validated before use
- **Atomic writes:** Config files are written atomically to prevent corruption
- **Symlink rejection:** Config loader refuses symlinked files
- **No telemetry:** Zero data sent to any third party
- **No caching:** Email and calendar data is never written to disk

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest  | Yes       |
| < Latest | Best effort |
