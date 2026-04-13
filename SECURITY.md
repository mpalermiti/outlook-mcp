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

- **Tokens:** Stored in OS keyring via azure-identity (never in plain files)
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
