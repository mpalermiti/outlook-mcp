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
  `graph.microsoft.com`, and its path must be the delta endpoint of the tool it
  was handed to: a mail cursor is only followed by the mail delta tool, and no
  delta tool will fetch any other Graph path.
- **Read-only mode is a tool gate, not a token scope.** `read_only: true` blocks this
  server's write tools. It does not narrow the OAuth token: sign-in consents the read-write
  scopes, and every refresh asks for `.default`, which carries whatever the Azure app was
  consented for. A `read_only` server still holds a write-capable Graph token, and the
  setting is a config-file value rather than anything Microsoft enforces. For a credential
  that genuinely cannot write, register a separate Azure app with the read permissions only
  and set `read_only_consent: true` beside `read_only: true`, which makes sign-in ask for
  the read scopes and nothing else. A saved sign-in is only used with the app registration
  it was made for.
- **Sender authentication:** `include_sender_auth=True` on the read tools reports
  Microsoft's verdict on a message's From domain. It is Microsoft's verdict, not a
  check this server makes, and it vouches for the From domain only, never for the
  display name; `false` means unverified, not spoofed. Only the topmost
  `Authentication-Results` header counts: a sender can write that header into a
  message, and Microsoft adds its own above everything the sender wrote. It is
  parsed as RFC 8601, a pass must be aligned with the From domain, and a header
  that names another service, fails DMARC or composite authentication, states
  either twice, or cannot be parsed authenticates nothing.
- **HTTP transport is opt-in, and never runs without a token.** With
  `OUTLOOK_MCP_TRANSPORT=streamable-http`, one request to the port reaches the mailbox, so the
  server refuses to start without a token on any bind, loopback included, and answers `401` to
  any request without it. The token comes from `OUTLOOK_MCP_HTTP_TOKEN`, or from
  `OUTLOOK_MCP_HTTP_TOKEN_FILE`, an owner-only file that is not a symlink. Host and Origin are
  checked against an allow-list on every bind (the loopback names on a loopback bind, otherwise
  the bind address or `OUTLOOK_MCP_HTTP_ALLOWED_HOSTS`), and no token or request line is
  logged. It speaks plain HTTP, so TLS belongs to a reverse proxy in front; the README's
  [Remote / HTTP transport](README.md#remote--http-transport-optional--outlook_mcp_transport)
  section has the rest.
- **Attachment paths are confined.** Every tool that reads or writes a file on the host —
  attaching to a message, draft or event, uploading to a task, saving a download — resolves
  the path inside `attachments_dir` (`resolve_attachment_path`) and refuses anything that lands
  outside it. Paths come from the model, which reads email, so confinement is by resolving
  (absolute paths, `..` and symlinks collapse first), never by comparing strings. On Windows a
  network or device path is refused before it is resolved, since resolving it would connect to
  the host. A download lands through an owner-only temp file moved into place, so a symlink
  planted at the target cannot redirect the bytes.
- **Attachments as content touch no disk.** `as_content=True` on the downloads returns the
  file in the tool result, 10 MiB at most, and inline uploads (`attachments`) are sent straight
  from the arguments. The name and type are the sender's, so the bytes decide how a file comes
  back: as an image only when it is a complete PNG, JPEG, GIF or WEBP of at most 5 MiB, as text
  only when it is a text type that decodes in its declared charset with no control characters,
  and otherwise as an embedded resource holding the original bytes.
- **Upload URLs stay out of errors.** A file over 3 MB goes up through a Graph upload session,
  whose URL carries its own authorization. Every chunk's reply is checked, nothing is sent after
  a failed upload, and a refusal reaches the model as Graph's status and error body only, never
  the URL.
- **Document reading is opt-in and contained, not sandboxed.** `outlook_read_attachment` is in
  the `documents` group, which loads only when `OUTLOOK_MCP_TOOLSETS` names it, and its PDF and
  spreadsheet parsers come from the optional `documents` extra. Each PDF or XLSX is parsed in a
  short-lived worker process, one at a time, started with a clean environment so it inherits no
  credential variables; the server stops it past 256 MiB of memory, 30 seconds or 8 MiB of
  output. The worker still has the user's filesystem, network and credential-store access.
  An attachment over 10 MiB is refused before it is downloaded, nothing is written to disk,
  there is no OCR, and formulas are never recalculated (Excel's saved result is reported).
  Results are marked `untrusted_external_content`: the text is the sender's.
- **Input validation:** All Graph IDs, emails, dates, KQL queries, folder names and event keys are validated before use
- **Atomic writes:** Config files are written atomically to prevent corruption
- **Symlink rejection:** Config loader refuses symlinked files
- **No telemetry:** Zero data sent to any third party
- **No caching:** Email and calendar data is never written to disk, except an attachment a
  download is asked to save in `attachments_dir`

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest  | Yes       |
| < Latest | Best effort |
