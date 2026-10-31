<!-- mcp-name: io.github.mpalermiti/outlook-mcp -->

# outlook-mcp

MCP server for Microsoft Outlook personal accounts via Microsoft Graph API.

[![PyPI](https://img.shields.io/pypi/v/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-graph-mcp.svg)](https://pypi.org/project/outlook-graph-mcp/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![MCP Registry](https://img.shields.io/badge/MCP_Registry-listed-green)](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti)

> **Personal Microsoft accounts only** — `@outlook.com`, `@hotmail.com`, `@live.com`. Work/school accounts (Entra ID) are not supported: the server is built and tested on personal accounts. Fixes that make a tool behave on a work account, without changing what a personal account sees, are welcome.

> **Disclaimer:** Independent open-source project. Not affiliated with, endorsed by, or supported by Microsoft Corporation. "Outlook" and "Microsoft Graph" are trademarks of Microsoft.

---

## Who this is for

You'll like this if you're:

- An **agent builder** wiring Outlook into your own infra (OpenClaw, Claude Code, Cursor, custom MCP host) and want a typed tool surface — not stdout you have to parse
- Building on **personal Microsoft accounts** (Outlook.com / Hotmail / Live) and want full control: BYO Azure app, no enterprise consent flow, no shared client ID
- Looking for **real coverage** — mail, calendar, contacts, to-do, drafts, folders, batch ops, threading — instead of a mail-only or calendar-only wrapper
- Security-conscious: tokens in the OS keyring (Keychain on macOS, libsecret on Linux -- never cleartext unless you opt in), granular `allow_categories`, optional `read_only` mode, zero telemetry

This **isn't for you** if you need work/school M365 accounts (use Microsoft's official tooling — Entra ID auth and admin-consent flows are out of scope here), or if a basic mail-only client would suffice (this has 77 tools — way more than you need for "read my inbox").

### How it differs from other Outlook tools you'll find

This is the only **first-class MCP server** in the personal-Outlook space — most alternatives are bash scripts or skill-shaped CLI wrappers the agent shells out to. That distinction matters: the agent gets typed tool schemas with structured args/returns, not stdout it has to parse. Other things you won't find elsewhere: `/$batch`-optimized triage (10-20× faster on bulk ops), recursive folder ops with name resolution, granular per-category permissions, multiple mailboxes (one server per account via `OUTLOOK_MCP_CONFIG_DIR`), and full attachment write paths including >3MB upload sessions for drafts and calendar events.

---

## What This Enables

Give your AI agent full Outlook access. Example prompts that just work:

- *"Summarize my unread email from the past 24 hours and flag anything time-sensitive."*
- *"What's in my Focused Inbox right now? Anything in Other that looks like it belongs up top?"*
- *"Any shipping updates in my inbox? Track what I'm waiting on and when it's supposed to arrive."*
- *"Scan my email for upcoming subscription renewals — what's about to auto-charge in the next two weeks?"*
- *"I've got a trip to Seattle next week — check my calendar for the itinerary and create a To Do task with a packing checklist."*
- *"Draft a reply to the last message from my sister saying I'll call her this weekend."*
- *"Move all newsletter and promotional email from this week to a 'Read Later' folder — batch 20 at a time."*

The server exposes 77 discrete tools (71 load by default; the four rule tools, document reader and MIME export are [opt-in](#toolset-selection-optional--outlook_mcp_toolsets)) so the agent can compose its own workflow — read, triage, write, schedule, track tasks — without hardcoded macros.

## Works With

- **[OpenClaw](https://openclaw.ai)** — native MCP support, available via [ClawHub](https://clawhub.ai/skills?q=outlook-mcp)
- **[Claude Code](https://claude.com/claude-code)** — `claude mcp add outlook -- outlook-mcp`
- **[Cursor](https://cursor.com)** — MCP-compatible
- **Any MCP client** — it's a standard stdio MCP server, with an opt-in [streamable-HTTP transport](#remote--http-transport-optional--outlook_mcp_transport) for clients that cannot spawn a local process

Listed on the [official MCP Registry](https://registry.modelcontextprotocol.io/v0/servers?search=mpalermiti) as `io.github.mpalermiti/outlook-mcp`.

---

## What's new

The last few releases, in brief. Every change is in the [CHANGELOG](CHANGELOG.md).

- **Export an email** (1.31): get a message's original file exactly as it arrived, attachments included, up to 10 MB. Off unless you turn it on.
- **Inbox rules** (1.30): list, create, change and delete server-side inbox rules, and see what each rule leaves out (1.30.2). Off unless you turn it on, and a rule that forwards mail also needs send permission.
- **Attachments as text** (1.28): PDFs, spreadsheets and text files come back as readable text, parsed in a separate process with memory and time limits. Off unless you turn it on.
- **Calendar** (1.27–1.29): turn a meeting invitation into its calendar event, attach files to events, tag an event so you can find it again, create events in any of your calendars, and see who owns each calendar.
- **Mail** (1.27–1.28): write in Markdown, and replies keep their formatting and the quoted original. Start a draft as a reply or forward, name folders by path (`Inbox/Travel/Flights`), and see whether Microsoft verified a sender.
- **Remote clients** (1.27–1.28): an optional HTTP mode serves a client on another machine and refuses to start without a token. Downloads can come back in the reply instead of as a file on the server.
- **Safer sends** (1.27–1.28): a send that times out is never sent twice on its own, and a large attachment that fails to upload no longer lets the mail go without it.
- **Accounts and settings** (1.29): pin a server to one sign-in, and set options from environment variables, which can only tighten security.
- **Cheaper every turn** (1.29): the tool list every client reads on each turn is a quarter smaller, and features that would add to it are off by default.

---

## Features

**77 tools** across 14 categories. 71 load by default; the four rule tools,
document reader and MIME export are opt-in groups (see [Toolset selection](#toolset-selection-optional--outlook_mcp_toolsets)).

- **Auth (1)** -- auth status check (login is via CLI)
- **Mail Read (8)** -- list inbox (with Focused Inbox and uncategorized filters), read message, opt-in original MIME export, bulk read by ID via `$batch`, search (KQL), list folders, delta-sync inbox changes, composed "since last call" digest across mail/events/contacts
- **Mail Write (3)** -- send, reply/reply-all, forward
- **Mail Triage (9)** -- move, delete (soft by default), flag, categorize, mark read/unread, reclassify (Focused Inbox), list/set/delete per-sender Focused Inbox overrides
- **Mail Rules (4)** -- list rules in execution order, create from sender/subject/body conditions, update (reads the rule first, so other actions survive a re-point), delete — opt-in toolset group
- **Calendar Read (4)** -- list events (with recurring expansion), get event details, find events by key, delta-sync event changes
- **Calendar Write (6)** -- create, update, delete, RSVP (accept/decline/tentative), attach and remove event files
- **Contacts (7)** -- list, search, get, create, update, delete, delta-sync changes
- **To Do (14)** -- task lists, tasks (list/get/create/update/complete/delete), checklist items (add/update/delete), task attachments (list/download/upload/delete)
- **Drafts (5)** -- list, create (new, or a reply, reply-all or forward in the message's thread), update, send, delete
- **Attachments (6)** -- list, download, read document text (opt-in), send-with-attachments, attach-to-draft, remove-draft-attachment
- **Folder Management (3)** -- create, rename, delete mail folders
- **Threading and Batch (3)** -- list thread, copy message, batch triage
- **User and Admin (4)** -- whoami, list calendars, list categories, mail tips

**Design principles:**

- **BYOID** -- Bring Your Own ID. You register your own Azure AD app. No shared client ID.
- **Zero telemetry** -- no analytics, no local caching, no third-party calls.
- **Token storage** -- OS keyring via `azure-identity` (macOS Keychain, Windows Credential Store, Linux Secret Service).
- **Input validation** -- all inputs validated (email, Graph IDs, OData, KQL, datetimes) before any API call.
- **Read-only mode** -- set `read_only: true` in config to block all write operations. Note this limits the *tools*, not the *token* -- see [What `read_only` does and does not do](#what-read_only-does-and-does-not-do).
- **Soft delete** -- delete moves to Deleted Items by default. Hard delete requires explicit `permanent: true`.
- **Timezone-aware** -- calendar operations respect your configured IANA timezone.
- **Relative dates** -- every datetime parameter takes ISO 8601 or an offset: `7d` is seven days ago, `+7d` is seven days from now, `now` is this moment. Units: `m`, `h`, `d`, `w`.
- **Bounded attachments** -- attachment reads and writes are confined to `attachments_dir`, so a message that asks an agent to mail a file elsewhere on disk cannot be obeyed.
- **Bounded delta cursors** -- a `delta_token` is caller-held state, so it is untrusted input. Every URL that would carry a Graph bearer token is parsed and required to be https on `graph.microsoft.com`, which is what stops a poisoned cursor from redirecting your mailbox token to someone else. Its path must also be the delta endpoint of the tool it was handed to, so a cursor cannot point a delta tool at some other part of the mailbox.
- **Workflow prompts** -- `morning_brief`, `triage_folder` and `catch_up` ship as MCP prompts, so the common sequences do not have to be reconstructed call by call.

### Agent-friendly shape (1.8.0)

Two pure-code upgrades that make the tools cheaper and more recoverable for AI agents:

- **Concise mode** — pass `concise=True` to the six high-volume read tools (`outlook_list_inbox`, `outlook_read_message`, `outlook_read_messages`, `outlook_search_mail`, `outlook_list_events`, `outlook_list_thread`) to drop bulky fields: full message bodies, quoted prior-message text in threads, body previews and categories on inbox listings — typical payload reduction ~10×. On `outlook_list_events` the trade is different and smaller: `organizer`, `response_status`, `type`, `show_as` and `key` come off, `is_organizer` and `attendees_count` are added, and `is_online` is spelled `is_online_meeting`, so a concise scan cannot tell a recurring occurrence from a one-off. Default `concise=False` preserves the existing response shape — strict backward compat.

- **Structured Graph errors** — every tool turns msgraph SDK exceptions into a tool error whose text carries Graph's message and a recovery hint: re-auth on 401, a link to the repo's [ROADMAP dead-ends list](https://github.com/mpalermiti/outlook-mcp/blob/main/ROADMAP.md#investigated-and-not-viable) on 403/`ErrorAccessDenied`, re-list on 404/`ErrorItemNotFound`, back-off on 429, retry on 503 — except on the tools that can email someone (mail, invitations, event updates, RSVPs), where a 5xx says the write may already have gone through and where to check, because those writes are never re-sent after one. `OutlookMCPError` subclasses and validation errors pass through unchanged.

---

## Azure AD App Registration

You need to register a free Azure AD app to get a client ID.

### Prerequisites (Personal Microsoft Accounts)

Microsoft has deprecated app registration for personal accounts without an Azure AD tenant. You need to create a free Azure account first:

1. Go to [azure.microsoft.com/free](https://azure.microsoft.com/free) and sign up with your personal `@outlook.com` account. Requires a credit card for identity verification but **won't charge you**. This creates a proper Azure AD tenant.

### Register the App

1. Go to [App Registrations](https://go.microsoft.com/fwlink/?linkid=2083908) and sign in with your `@outlook.com` account.

2. Click **"+ New registration"** and fill in:
   - **Name:** anything except Microsoft-branded terms (e.g. `mp-outlook-mcp` — names like "Outlook MCP" will be rejected)
   - **Supported account types:** select **"Personal Microsoft accounts only"**
   - **Redirect URI:** leave blank

3. Click **Register**. Copy the **Application (client) ID** from the overview page.

4. Go to **Authentication (Preview)** → **Settings** tab → toggle **"Allow public client flows"** to **Yes** → **Save**.

5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → add:
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`, `Tasks.ReadWrite`
   - `MailboxSettings.Read`
   - `User.Read`, `offline_access`

No client secret is needed. The device code flow uses public client auth.

---

## Quick Start

### Install

**Option A — from PyPI (recommended):**

```bash
uv tool install outlook-graph-mcp
# or: pipx install outlook-graph-mcp
# or: pip install outlook-graph-mcp
```

To read PDF and XLSX attachments as text, install the optional `documents` extra instead
(`uv tool install 'outlook-graph-mcp[documents]'`) and enable the opt-in `documents` group;
see [Attachments](#attachments).

**Option B — from source:**

```bash
git clone https://github.com/mpalermiti/outlook-mcp.git
cd outlook-mcp
uv sync
```

### Configure

Create `~/.outlook-mcp/config.json`:

```json
{
  "client_id": "YOUR_APPLICATION_CLIENT_ID",
  "tenant_id": "consumers",
  "timezone": "America/Los_Angeles",
  "read_only": true,
  "attachments_dir": "~/.outlook-mcp/attachments"
}
```

The only required field is `client_id`. Everything else has sensible defaults. Start with `read_only: true` — flip to `false` when you're comfortable.

### Register with your MCP client

**If installed from PyPI:**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

**If installed from source:**

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["--directory", "/path/to/outlook-mcp", "run", "outlook-mcp"]
    }
  }
}
```

**For OpenClaw**, use the `openclaw mcp` CLI — it writes to `mcp.servers` in `~/.openclaw/openclaw.json` for you:

```bash
# If installed from PyPI:
openclaw mcp set outlook '{"command":"outlook-mcp"}'

# If installed from source:
openclaw mcp set outlook '{"command":"uv","args":["--directory","/path/to/outlook-mcp","run","outlook-mcp"]}'

# Verify:
openclaw mcp list
openclaw mcp show outlook --json
```

Restart the OpenClaw gateway after registering. See the [OpenClaw MCP docs](https://docs.openclaw.ai/cli/mcp) for SSE/HTTP transport variants.

### Authenticate

Run this once on the machine where the MCP server will run:

```bash
outlook-mcp auth          # installed from PyPI
uv run outlook-mcp auth   # from a source checkout
```

You'll get a URL and a code. Open the URL in any browser, enter the code, and sign in with your Microsoft account. Tokens are cached in the OS keyring — the MCP server picks them up automatically.

The consent screen lists the full read-write set (`Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`, `Tasks.ReadWrite`, `MailboxSettings.Read`, `User.Read`) — with `mailbox_settings_write: true` the settings scope becomes `MailboxSettings.ReadWrite` — even when the config starts with `read_only: true`, because `read_only` gates the tools, not the token (see [What `read_only` does and does not do](#what-read_only-does-and-does-not-do)), and the scopes a first consent leaves out can never be granted later without logging in again. (The one exception is `read_only_consent: true`, which asks for the read permissions only; it exists for a separate read-only app, described under the same heading.) Every request afterwards — silent refresh and each Graph call — uses the `.default` scope, which on an already-consented account means exactly "the set you granted". That ordering is deliberate: on personal accounts a first consent asking only for `.default` can land a session with no delegated permissions, which then can't be redeemed (`AADSTS70000`) without logging in again ([#82](https://github.com/mpalermiti/outlook-mcp/issues/82)).

Other CLI commands (prefix `uv run` from a source checkout):

```bash
outlook-mcp status   # Check auth status
outlook-mcp logout   # Remove this instance's auth record (the OS token cache is left to age out)
outlook-mcp serve    # Start MCP server (default; stdio, or HTTP with OUTLOOK_MCP_TRANSPORT)
```

---

## Troubleshooting

### `me-token-to-replace is invalid` on every call

You're on 1.22.0 installed after 2026-09-18. A fresh install of that version resolves `microsoft-kiota-*` 1.13 or later, which `msgraph-core` 1.5.1 doesn't handle, so every `/me` request reaches Graph as `/users/me-token-to-replace`. Fixed in 1.22.1, which caps kiota below 1.13 ([#80](https://github.com/mpalermiti/outlook-mcp/issues/80)). In 1.26.0 that cap is replaced by a `msgraph-core>=1.5.2` floor, which carries the fix and works with kiota 1.13+:

```bash
uv tool upgrade outlook-graph-mcp
# or: pipx upgrade outlook-graph-mcp
# or: pip install --upgrade outlook-graph-mcp
```

Your config and sign-in are untouched; no re-auth needed.

### `SSL: CERTIFICATE_VERIFY_FAILED` on Linux

If auth fails with `[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate`, your Python environment can't find the system CA bundle. This is common on minimal/container Linux images and with the isolated venv from `uv tool install`.

Point Python at your system CA bundle. Set **both** variables — auth (via `azure-identity` → `requests`) reads `REQUESTS_CA_BUNDLE`, while the delta/`$batch` paths (via `httpx`) read `SSL_CERT_FILE`:

```bash
export SSL_CERT_FILE=/etc/ssl/certs/ca-certificates.crt      # httpx + Python ssl
export REQUESTS_CA_BUNDLE=/etc/ssl/certs/ca-certificates.crt  # azure-identity auth
```

The path varies by distro: Debian/Ubuntu use `/etc/ssl/certs/ca-certificates.crt`; RHEL/Fedora use `/etc/pki/tls/certs/ca-bundle.crt`. If the file is missing, install your distro's CA package (`ca-certificates`). Set these in the same environment your MCP client launches the server from so they apply at runtime, not just to the one-time `auth` command.

### Token cache stored unencrypted (Linux)

On Linux without an encrypted store (`libsecret`/PyGObject not importable, or no running Secret Service), `outlook-mcp auth` and the server refuse to write the token cache in cleartext, and every tool call says why. Install the system packages for an encrypted store — see [Privacy and Security](#privacy-and-security) — or set `allow_unencrypted_token_cache: true` to accept a plaintext cache file, after which the server logs a one-time warning at startup.

---

## Tool Reference

**Dates.** Every datetime parameter (`after`, `before`, `start`, `end`, `due`,
`deferred_send_datetime`) accepts ISO 8601 — `2026-10-22` or `2026-10-22T14:30:00Z` — or a
relative offset: `7d` is seven days **ago**, `+7d` is seven days **from now**, and `now` is
this moment. Units are `m`, `h`, `d`, `w`. Bare means *ago*, matching the usual CLI
convention, so a due date in the future needs the `+`. Zone-less input is interpreted in your
configured `timezone`; responses are always UTC.

**Folders.** Every `folder` argument, `outlook_create_folder`'s `parent_folder` and
`outlook_batch_triage`'s move target take a well-known name (`inbox`, `junkemail`), a display
name (`Junk Email`), a Graph ID, or a path: `Inbox/Travel/Flights` is Flights under Travel under
Inbox. A `/` inside a folder's own name is written `//` (`Inbox/Q1//Q2`). A display name finds a
top-level folder of that name first; a name that several nested folders share is refused, and
the error lists their paths. `outlook_rename_folder` and `outlook_delete_folder` take an ID.

### Auth

| Tool | Description |
|------|-------------|
| `outlook_auth_status` | Check if authenticated and whether read-only mode is active. |

> **Note:** Authentication is handled via the CLI (`outlook-mcp auth`), not through MCP tools. See [Authenticate](#authenticate) above.

### Mail Read

| Tool | Description |
|------|-------------|
| `outlook_list_inbox` | List messages in a folder. `folder` accepts display names, well-known names, paths (`Inbox/Travel/Flights`), or Graph IDs. Filter by read status, sender, date range, Focused Inbox classification, or no categories (`uncategorized_only`). Pagination via `cursor` (or `skip`). |
| `outlook_read_message` | Get full message by ID, or by its Message-ID header value (`<abc@example.com>`, matched exactly; the response's `id` is the Graph ID the other tools take). Format: `text` (Graph's plain-text rendering), `html`, or `full` — the stored body, usually HTML, in both `body` and `body_html`; read `text` for plain text. Pass `include_deferred_send=True` to also surface the draft's scheduled delivery time. Pass `include_sender_auth=True` to add `sender_authentication`, Microsoft's verdict on the From domain: `authenticated`, `from_domain`, the `dkim`, `dmarc` and `compauth` results, and a short `reason`. Only Microsoft's own topmost `Authentication-Results` header counts, and `authenticated` is true only for a pass aligned with the From domain. It vouches for `from_domain` only, never for the display name, and `false` means unverified, not spoofed. |
| `outlook_read_messages` | Bulk read up to 20 messages by ID via Graph `$batch` in one round-trip. Per-message shape matches `outlook_read_message` byte-for-byte for the same `(format, concise, include_deferred_send, include_sender_auth)`. Partial-failure tolerant: 404s on some IDs surface in `failures[]` without failing the whole call. Use NOT N `outlook_read_message` calls. |
| `outlook_search_mail` | Search mail using KQL query. Optionally scope to a folder by name, path or ID. A `cursor` continues the search in the folder it started in. Graph has been seen to end a search a little past 250 results without saying so (255–275 on Exchange Online; 275 also observed on Outlook.com), so a last page past 250 carries `incomplete`, a note that older matches may be missing and how to reach them. |
| `outlook_list_folders` | List mail folders with counts, `parent_id`, and `child_count`. Pass `recursive=true` to walk the full folder tree (subfolders included); each folder then carries its `path`, the form every `folder` argument accepts. |
| `outlook_list_inbox_delta` | List only inbox changes since the last call. `folder` (default `inbox`) takes a name, path or ID. First call returns a full snapshot plus a `delta_token`; subsequent calls (token passed back) return only added/updated/deleted items. Deletes come back as `{id, is_deleted: True}`. Cursor is stateless — agent persists and replays. |
| `outlook_changes_since` | One structured "since last call" digest composing mail/events/contacts deltas. Returns counts + `urgent_flagged` mail + top-5 `by_sender` + new/cancelled events. Each resource has an independent `delta_token`; stale-token recovery (HTTP 410) auto-resyncs that resource and surfaces `_meta.resync`. The first call covers the last `fallback_window_hours` (default 24): Mail requests Graph's received-time filter and enforces it client-side when ignored (`_meta.mail_window_fallback`); contacts are kept by last-modified time, and events by the calendar view (the window plus the next 7 days). Where Graph honours the mail filter, later calls keep it: mail received before the window reports only deletes, moves out of the inbox and read/unread changes, not new flags or categories. `has_more` names any resource the per-call cap stopped mid-sync; replace the stored `delta_tokens` dictionary with the returned one and pass it back; do not merge the dictionaries. It keeps `mail_bootstrap_from` while a mail bootstrap is unfinished, so continuation keeps the original bound and reports it in `window.from`. A reply with no new cursor hands back the token passed in. Designed for recurring agent loops. |
| `outlook_export_message` | **Opt-in** (`mime` group). Export original MIME by Graph message ID as a `message/rfc822` resource blob after JSON metadata (`size_bytes`, `sha256`); read-only, bounded to 10 MiB. |

Original MIME export uses the existing sign-in and needs no additional scope or package.
Enable it with `OUTLOOK_MCP_TOOLSETS=mail,mime` (plus any other groups you need) and restart.
It preserves the original bytes, including attachments; it does not parse or rewrite them.
The result contains a small JSON metadata block followed by an embedded resource with a
base64 blob, never decoded text. Messages larger than 10 MiB are refused while streaming.
A throttled request (HTTP 429) returns an error; bulk exporters must pace their calls and
handle throttling rather than relying on automatic retries.

### Mail Write

| Tool | Description |
|------|-------------|
| `outlook_send_message` | Send email. Supports TO/CC/BCC, HTML or Markdown body, importance level, `reply_to` and a read receipt (`request_read_receipt`). |
| `outlook_reply` | Reply or reply-all to a message, the text above the quoted original. Plain text keeps its paragraphs, line breaks and `- ` lists; a plain-text thread gets a plain-text reply. |
| `outlook_forward` | Forward a message to one or more recipients with optional comment, which keeps its line breaks like a reply's text. |

Bodies are plain text unless `is_html=True` (HTML, sent as written) or `is_markdown=True`. Markdown is converted to HTML for a safe subset: paragraphs, line breaks, `-`/`*`/`+` and numbered lists, `**bold**`, `*italic*`, `` `code` ``, and `[text](url)` links to `https:` and `mailto:` addresses only. An image, `![alt](url)`, becomes a link to it, never an embedded picture. Everything else, raw HTML included, is shown as written. `outlook_send_message`, `outlook_reply`, `outlook_forward` (its `comment`), `outlook_create_draft` and `outlook_update_draft` take it; on a plain-text thread a Markdown reply is sent as written, as plain text.

### Mail Triage

| Tool | Description |
|------|-------------|
| `outlook_move_message` | Move a message to a folder by name, path or ID. |
| `outlook_delete_message` | Delete a message. Soft delete (Deleted Items) by default. `permanent: true` for hard delete. |
| `outlook_flag_message` | Set follow-up flag: `flagged`, `complete`, or `notFlagged`. |
| `outlook_categorize_message` | Set categories on a message. |
| `outlook_mark_read` | Mark a message as read or unread. |
| `outlook_reclassify_message` | Move a message between Focused Inbox and Other (`focused` / `other`). |
| `outlook_list_inbox_overrides` | List Focused Inbox per-sender override rules. |
| `outlook_set_inbox_override` | Upsert a per-sender Focused Inbox override (`focused` / `other`). Case-insensitive sender matching; PATCH-if-exists, else POST. |
| `outlook_delete_inbox_override` | Delete a Focused Inbox override by ID. |

### Mail Rules

Server-side rules: they run whether or not any Outlook client is open. A rule is **durable, not a one-off** — create one and it keeps applying to all future mail long after the agent that made it has disconnected. That is what makes rules useful, and it is also why `outlook_create_rule` and `outlook_update_rule` are annotated destructive — re-enabling one resumes whatever it does — and why `forward_to`, and re-enabling a rule that forwards or redirects, additionally require the `mail_send` category: a forwarding rule keeps mailing a third party after the session ends, which is the standard way a compromised mailbox gets drained. Hard delete is not offered here — `delete` moves matching mail to Deleted Items, which is recoverable. The same `mail_send` gate covers the indirect routes to a forwarding rule starting work: reordering one, or disabling, reordering or deleting a `stop_processing_rules` or `permanentDelete` rule that was hiding it — `permanentDelete` removes the message before later rules see it. Repairing an errored forwarding rule needs `mail_send` too: the repair puts it back into service. And a listing that stopped early (`complete: false`) counts as "cannot tell", so it gates the same way rather than assuming nothing is hidden. A rule that never runs until the rule above it changes is still a rule that forwards mail.

| Tool | Description |
|------|-------------|
| `outlook_list_rules` | List inbox rules in execution order, with conditions, exceptions and actions, and whether Exchange reports an error applying the rule. |
| `outlook_create_rule` | Create a rule from sender/subject/body/attachment conditions and move, copy, mark-read, delete, forward, categorize or stop-processing actions. |
| `outlook_update_rule` | Change a rule's name, enabled state, sequence, target folder or mark-read action. Other actions are preserved; re-enabling a rule that forwards or redirects needs `mail_send`. |
| `outlook_delete_rule` | Delete a rule by ID. Irreversible. |

The three write tools are gated by the `mail_rules` permission category — add it to `allow_categories` (or leave `allow_categories` unset for a full read-write install).

The four rule tools are opt-in, off in a stock install: with `OUTLOOK_MCP_TOOLSETS` unset none of them is loaded — not even `outlook_list_rules` — and the surface every user pays for is unchanged. Name `rules` in `OUTLOOK_MCP_TOOLSETS` to load them, or set `mailbox_settings_write: true`, which adds them only when `OUTLOOK_MCP_TOOLSETS` is unset — an explicit selection is what you asked for, so it wins.

A rule hides as well as forwards, and the hiding needs no permission beyond `mail_rules`: a rule at sequence 1 that deletes, marks read and stops processing on "security", "sign-in" or "password" removes that mail before any later rule sees it. Read the rules of a mailbox you did not configure before trusting what it shows you.

**Not every personal mailbox allows rules.** One answered all four operations; another refused even a plain listing with 403, though its token reads categories under the same `MailboxSettings.Read` scope. A refused listing says it is either the consent or the mailbox. If a fresh sign-in does not change it, rules cannot be managed there.

**Order is Graph's to set.** On a live personal mailbox, `sequence` was ignored on both create and update: a rule created at sequence 500 and then patched to 500 stayed at the sequence Graph had assigned (128 of 127). The tools accept the field because Graph's schema does, and they return the sequence Graph actually stored — treat that returned value as authoritative, not the one you asked for. When the two differ, the reply also carries `sequence_requested`, so an ignored request is visible. The ordering and tie logic in the `mail_send` gate is defensive against a listing that reports equal or missing sequences, not something this API was observed to produce.

Creating, updating or deleting a rule also needs `MailboxSettings.ReadWrite` on the token. The default read-write consent set requests `MailboxSettings.Read` only, so the write tools answer `ErrorAccessDenied` until the token carries the scope. To enable the write tools:

1. **Ask for the write scope at sign-in.** Set `mailbox_settings_write: true` in your config. The default read-write consent asks for `MailboxSettings.Read` only, so without this every rule write answers `ErrorAccessDenied` no matter what the app registration allows.
2. **Re-run `outlook-mcp auth`.** Consent is granted per *requested* scope and an existing token keeps the scopes it was issued with, so the new scope only lands on a fresh consent.
3. **Restart the server.** Tools are chosen when the process starts, so the rule tools appear only after a restart — a config change mid-session does nothing to the surface, and the write scope alone does not add them.
4. **If your tenant restricts it**, add `MailboxSettings.ReadWrite` (delegated, Microsoft Graph) to the app registration in the Azure portal and grant admin consent before step 2. A personal account does not need this.

`outlook_mcp` cannot detect a half-configured setup: a token without the scope authenticates fine and reads keep working, and every rule write fails at Microsoft's end with `ErrorAccessDenied`. If rule writes fail while `outlook_list_rules` works, the scope is the cause.

### Calendar Read

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List events in a date range. Expands recurring events. Each event carries `type` — `occurrence` or `exception` for an instance of a series, `singleInstance` for a one-off — so a listing tells recurring work apart without a second call. A `seriesMaster` never appears here: `calendarView` returns expanded instances, so use `outlook_get_event` to see a master. Each event also carries `show_as`, the free/busy status Outlook labels "Show as". Configurable via `days`, `after`, `before`. `calendar` selects which calendar to read: omit (or `"primary"`) for the default, otherwise a display name (case-insensitive; not-found and ambiguous errors name what exists) or an ID from `outlook_list_calendars`. A `cursor` continues the listing it came from, so later pages need neither `calendar` nor a second lookup. An event that has a `key` (see `outlook_create_event`) carries it; one without has no `key` field. `concise=True` omits `type`, `show_as` and `key`. |
| `outlook_get_event` | Get full event details: attendees, body, online meeting URL, recurrence, `type` (`singleInstance` / `seriesMaster` / `occurrence` / `exception`), `show_as`, `key` (`null` when the event has none), and `attachments` — `id`, `name`, `size`, `content_type` and `is_inline` of each file, never its bytes (`[]` when there are none; read with a second request only when Graph reports the event has attachments). Pass either `event_id` or an invitation's `message_id` to resolve the exact event in your calendar (never by subject). Returns current details and invitation staleness, plus nullable `is_organizer`, `allow_new_time_proposals`, `is_cancelled` and attendee `proposed_new_time` slots with their zones. Missing metadata stays null; organizer status is relative to the calendar owner. |
| `outlook_find_event_by_key` | Find the events whose `key` is the one given — the caller's own reference set by `outlook_create_event` or `outlook_update_event`. No date window: a booking months away is found as readily as tomorrow's. Searches the default calendar, or `calendar` (resolved as for `outlook_list_events`). Returns events in the default listing's shape, `key` included: single events and series masters, never occurrences, so the ID found is the one to update. Nothing matching returns an empty list; `has_more` is `true` when more than 25 events share the key. |
| `outlook_list_events_delta` | List only event changes inside a window since the last call. `start` and `end` (ISO 8601) required on the first call (Graph constraint — no whole-calendar sync). Each changed event carries every field the **default** `outlook_list_events` listing returns except `key`, `type` and `show_as` included, plus `is_deleted` (`concise=True` has its own narrower shape, which the delta tool does not mirror). Unlike the listing, this endpoint *does* return `seriesMaster` items, so a caller seeding from `outlook_list_events` should expect ids here that the seed never held. Deletes come back as `{id, is_deleted: True}` and nothing else. Cursor is stateless. |

### Calendar Write

| Tool | Description |
|------|-------------|
| `outlook_create_event` | Create event with location and attendees. (`is_online` has no effect on personal accounts — Graph ignores `isOnlineMeeting` for consumer mailboxes.) Pass `recurrence` to create a **series**: a shorthand (`daily`, `weekdays`, `weekly`, `monthly`, `yearly`, anchored on `start`) or a full [Graph recurrence object](https://learn.microsoft.com/graph/api/resources/patternedrecurrence) for anything else. `range.startDate` defaults to the event's start date. `timezone` (an IANA name such as `America/Los_Angeles`, never an abbreviation) is the zone the event is anchored in and a series is expanded in; omit it for the config `timezone`. `show_as` sets the free/busy status Outlook labels "Show as" — `free`, `tentative`, `busy`, `oof` (out of office), `workingElsewhere`, or `unknown`; omit it and Graph applies its own default of `busy`. `calendar` creates it in another calendar, resolved as `outlook_list_events` resolves it: omit (or `"primary"`) for the default, otherwise a display name or an ID from `outlook_list_calendars`. A calendar Graph lists as read-only (`canEdit: false`) is refused before anything is sent. The tools that take an event ID need no `calendar`: an ID reaches an event on any of your own calendars. `key` stores your own reference for the event (a booking number, say: 1–128 letters, digits and `. _ : @ -`) in a hidden extended property, so `outlook_find_event_by_key` can find the event again when a change to the same booking arrives. Nothing stops two events sharing a key; look it up before creating. |
| `outlook_update_event` | Update event fields (subject, time, location, body, attendees, all-day, `show_as`). Only patches changed fields. Pass `recurrence` to turn a single event into a series, or `remove_recurrence=True` to turn a series back into a single event. `attendees` **replaces** the whole guest list and emails invitations/cancellations; `is_all_day` needs `start`+`end` in the same call, while `show_as` patches on its own. Patching a time keeps the zone the event is anchored in; `timezone` (with `start`+`end`) re-anchors it elsewhere, which is how a series created before zones existed gets repaired. A `start`, `end` or `recurrence` patch to a series with edited or deleted occurrences is refused, naming them — Graph would silently discard every one. `key` sets or replaces the event's key and patches on its own; there is no clearing one. |
| `outlook_delete_event` | Delete a calendar event. |
| `outlook_rsvp` | RSVP to an event: `accept`, `decline`, or `tentative`. Optionally include a message. |
| `outlook_attach_to_event` | Attach files to an event — an e-ticket on a flight, say. `attachment_paths` resolve inside `attachments_dir`, as for `outlook_attach_to_draft`; a missing file, or one path outside it, attaches nothing. Files over 3 MB use an upload session; a chunk Graph refuses, or an upload it does not report complete, stops the call with an error rather than reporting the file attached. Returns the new attachment IDs (none yet for an upload-session file). Attendees are not sent the files when they are attached; they get them with the event's next update, so on an event with attendees this needs `mail_send` under a restricted `allow_categories`. |
| `outlook_remove_event_attachment` | Remove one attachment from an event, by the ID `outlook_attach_to_event` returned or `outlook_get_event` lists. |

### Contacts

| Tool | Description |
|------|-------------|
| `outlook_list_contacts` | List contacts with cursor pagination. Summaries carry `categories`; `outlook_search_contacts` omits the key because Graph's `$search` does not return it. |
| `outlook_search_contacts` | Search contacts by name or email. `phone`, instead of `query`, finds a number however it is written: with or without its country code, a leading `00` or trunk `0`, spaces or dashes, from at least its last seven digits. Graph's contact `$search` finds a number only as it is stored, and phone fields can't be filtered, so a phone search reads the contact listing (pages of 1,000, up to 10,000 contacts) and compares digits. `has_more` is true when there were more matches than `count`, or contacts past the 10,000th. |
| `outlook_get_contact` | Get full contact details by ID, including `middle_name`, `job_title`, `birthday`, home/business/other addresses, categories and personal notes. |
| `outlook_create_contact` | Create a new contact. `title` is the name prefix ("Dr.", "Ms."), which Outlook shows in the display name; the job goes in `job_title`. `phone_type` (`mobile`, the default, `business` or `home`) says which kind of number `phone` is. `birthday` is a date (`YYYY-MM-DD`), stored at 11:59 UTC so that it is the same date in every time zone from UTC−11 to UTC+12; `outlook_get_contact` returns that timestamp. `middle_name`, `personal_notes` and `categories` are set as given. |
| `outlook_update_contact` | Update contact fields, including `middle_name`, `company`, `title`, `job_title` and `personal_notes`. `categories` **replaces** the contact's whole list, as `outlook_categorize_message` does a message's, so `[]` removes every category; read the contact first to add one. `phone_type` works as on create; a `business` or `home` number **replaces** that list of numbers with the one given. `clear_phones` empties the types it names, so `phone_type="business", clear_phones=["mobile"]` moves a number saved as mobile to business in one update. `home_address`, `business_address` and `other_address` take the shape `outlook_get_contact` returns — any subset of `street`, `city`, `state`, `postal_code`, `country_or_region` — and **replace** that whole address, so pass back every part you want to keep. Omit one to leave it untouched. `birthday` sets or replaces the date, as on create. `email` **replaces** every address with the one given; `add_emails` and `remove_emails` change only the addresses they name, keeping the rest (a contact holds three at most). They are merged with what the contact holds now, and the update is pinned to that read with `If-Match`, so an edit made in between is refused with `412` rather than overwritten. Removing an address the contact doesn't have is refused. |
| `outlook_delete_contact` | Delete a contact. |
| `outlook_list_contacts_delta` | List only contact changes since the last call. Deletes come back as `{id, is_deleted: True}`. Cursor is stateless. |

### To Do

| Tool | Description |
|------|-------------|
| `outlook_list_task_lists` | List To Do lists. |
| `outlook_list_tasks` | List tasks with status filter and pagination. |
| `outlook_get_task` | Get full task details: notes (`body`), checklist items (ordered unchecked-first), due, recurrence flag. |
| `outlook_create_task` | Create task with due date, importance, recurrence. |
| `outlook_update_task` | Update task fields. |
| `outlook_complete_task` | Mark task as completed. |
| `outlook_delete_task` | Delete a task. |
| `outlook_add_checklist_item` | Add a checklist item (sub-step) to a task. |
| `outlook_update_checklist_item` | Update a checklist item — mark done (`is_checked`) or rename (partial patch). |
| `outlook_delete_checklist_item` | Delete a checklist item from a task. |
| `outlook_list_task_attachments` | List attachments on a To Do task (id, name, size, content_type) with pagination. |
| `outlook_download_task_attachment` | Download a task attachment's content to a local file (confined to `attachments_dir`), or return it with `as_content=True`. |
| `outlook_upload_task_attachment` | Attach a local file (from `attachments_dir`) to a task via inline base64 POST (1 byte – 20 MiB). |
| `outlook_delete_task_attachment` | Remove an attachment from a To Do task. |

### Drafts

| Tool | Description |
|------|-------------|
| `outlook_list_drafts` | List draft messages with pagination. |
| `outlook_create_draft` | Create a draft, with an HTML (`is_html=True`) or Markdown (`is_markdown=True`) body if wanted. With `reply_to_message_id` (`reply_all=True` for reply-all) or `forward_message_id`, the draft is created in that message's thread with the original quoted below, and nothing is sent; Graph fills a reply's recipients and subject, and a forward needs `to`. Supports scheduled delivery via `deferred_send_datetime` (server-side, Outlook-desktop-compatible "Delay Delivery"). |
| `outlook_update_draft` | Update draft fields. Accepts `is_html=True` for HTML bodies, `is_markdown=True` for Markdown ones, and `deferred_send_datetime` to set or clear the scheduled delivery time. |
| `outlook_send_draft` | Send an existing draft. |
| `outlook_delete_draft` | Delete a draft. |

### Attachments

> **Since 1.20.0, these tools only reach `attachments_dir`** (default `~/.outlook-mcp/attachments`).
> A bare filename resolves inside it; a path outside it is refused, including via a symlink.
> To email a file, move it there first — or widen `attachments_dir`, understanding that
> anything reachable from it can be sent. Before 1.20.0 these tools could read any file the
> server process could read, which meant an email asking an agent to attach one could be obeyed.
>
> The same fence covers the To Do attachment tools (`outlook_upload_task_attachment`,
> `outlook_download_task_attachment`): uploads read from `attachments_dir` and downloads
> write into it, so neither can sweep arbitrary files off disk. (Downloads are not
> `todo_write`-gated — they are reads, like `outlook_download_attachment`; the fence, not
> the category, is what confines them.)
>
> **A client that can't see the server's disk**, such as a remote or sandboxed one, can carry
> the bytes in the call instead, and then nothing is read from or written to disk.
> `as_content=True` on either download returns the attachment in the result, up to 10 MB:
> PNG, JPEG, GIF and WebP up to 5 MB as images, text up to 100,000 characters as text, and
> anything else, PDFs included, as an embedded resource holding the base64 bytes and their
> type. An image cut short (a PNG with no end chunk, say) comes back as that resource too.
> Text that carries control characters, or a lone surrogate its charset decoded to,
> also comes back as that resource, so it arrives unaltered. `attachments` on
> `outlook_send_with_attachments` and `outlook_attach_to_draft` takes inline files,
> `[{name, content_base64, content_type?}]`, up to 3 MB each (Graph's simple-attachment
> size). The fence still governs every path argument. An inline file is bytes the agent
> already holds and could already put in a message body, so it widens nothing.
> `outlook_attach_to_event` and `outlook_upload_task_attachment` take paths only.

| Tool | Description |
|------|-------------|
| `outlook_list_attachments` | List attachments on a message. |
| `outlook_download_attachment` | Download an attachment and save decoded bytes into `attachments_dir`, or return it with `as_content=True`. |
| `outlook_read_attachment` | **Opt-in** (`documents` group, not loaded by default; see below). Read PDF pages, XLSX cells/formulas/saved results, or text in its declared charset (UTF-8 by default); `offset` / `max_chars` page the text, no file saved. |
| `outlook_send_with_attachments` | Send a message with attachments read from `attachments_dir` (auto upload session for >3MB) or passed inline in `attachments`. If a large upload fails, nothing is sent, and the error names the draft left in Drafts so it can be deleted or finished with `outlook_attach_to_draft` instead of sent twice. |
| `outlook_attach_to_draft` | Add attachments from `attachments_dir` (auto upload session for >3MB) or inline in `attachments` to an existing draft. A failed large upload is an error, never a file reported as attached. Returns the new attachment IDs (none yet for an upload-session file). |
| `outlook_remove_draft_attachment` | Remove a single attachment from a draft. |

For **readable document contents** rather than original bytes, use `outlook_read_attachment`.
The saved path from a download belongs to the server, so a remote client may not be able to open it.
`as_content=True` avoids that path and preserves the original file; the reading tool instead returns
extracted text for clients that cannot interpret the embedded binary resource. Existing downloads
and their defaults are unchanged.

PDF and XLSX parsing is optional: install `outlook-graph-mcp[documents]` (for example,
`uv tool install 'outlook-graph-mcp[documents]'`, `pipx install 'outlook-graph-mcp[documents]'`,
`pip install 'outlook-graph-mcp[documents]'`, or `uv sync --extra documents` in a source
checkout). Enable the opt-in `documents` group with
`OUTLOOK_MCP_TOOLSETS=mail,attachments,documents` (plus any other groups you need) and restart.
Naming groups loads only those, so list every group you want; see
[Toolset selection](#toolset-selection-optional--outlook_mcp_toolsets).
When enabled, a base install can read text in its declared charset (UTF-8 by default),
and returns an installation hint for PDF/XLSX. Parser dependencies have bounded versions
and are imported only in a short-lived worker; they add no requirement to the base install.

`outlook_read_attachment(message_id, attachment_id, offset=0, max_chars=20000)` returns `text`,
`next_offset` (null at the end), `total_chars`, file `bytes` / `sha256`, and `warnings`. Continue with
the same IDs and the returned offset until null. Each continuation downloads and parses the
whole file again, then returns the requested slice; there is no document cache. Use a larger
`max_chars` (up to 50000) to reduce repeated work. PDF output marks page numbers; a page without a
text layer is reported as blank or needing OCR. Extracted **PDF** text lines are prefixed with
`| `, so only the PDF reader can emit an unprefixed `[Page N]` structural marker in PDF output.
This guarantee is PDF-only: XLSX and text chunks can start mid-cell or mid-line, and apparent
cell addresses or page markers in their contents are untrusted text. No OCR is performed. XLSX output includes hidden
sheets, cell addresses, number formats, formulas and `saved_result` (Excel's stored value, null when
absent). Formulas are **not recalculated**, and charts/images are not extracted. All attachment
contents remain untrusted data.

Worksheet metadata (`sheets`) is returned only at offset 0, and only when its serialized list
plus at least one text character fits `max_chars`; otherwise the list is omitted and worksheet
names remain in the paged text. A workbook above 100 sheets is refused; displayed sheet names
are cut to Excel's 31-character limit, with `… [N]` marking a cut name and its one-based
worksheet position so shared prefixes remain distinguishable. Names within the limit are unchanged.
`max_chars` budgets text plus serialized worksheet metadata;
the fixed response envelope and bounded warnings are separate. A PDF slice starting mid-line
receives a `| ` presentation prefix when it fits. `continuation_prefix` identifies that prefix:
strip it before joining chunks to reconstruct the text. Offsets and `total_chars` count the
canonical text, excluding this presentation prefix.

Reading is limited to 10 MiB of file bytes, 200 PDF pages, 100,000 visited spreadsheet cells, one
million extracted characters and 50 MiB / 2,000 entries of expanded XLSX content. `max_chars`
has a ceiling of 50,000. Text is decoded in-process; all extracted text is cleaned of control
characters before paging. PDF/XLSX use a **subprocess for crash and resource containment,
not a sandbox**: server credential environment variables are absent,
but it runs as the same user with filesystem, network and Keychain access. A parent watchdog
polls RSS every 10 ms and kills a parser above 256 MiB on Linux, macOS and Windows; brief
overshoot between polls is possible. Only one parser runs at a time, with a 30-second wall timeout.
OS resource limits supplement these checks where supported.
A limit is an explicit error, never a silently partial document.

### Folder Management

| Tool | Description |
|------|-------------|
| `outlook_create_folder` | Create mail folder (top-level, or nested under `parent_folder`: a name, path or ID). |
| `outlook_rename_folder` | Rename a mail folder, by ID. |
| `outlook_delete_folder` | Delete a mail folder by ID (refuses well-known folders). |

### Threading and Batch

| Tool | Description |
|------|-------------|
| `outlook_list_thread` | Get all messages in a conversation thread. |
| `outlook_copy_message` | Copy a message to another folder (name, path or ID). |
| `outlook_batch_triage` | Batch move/flag/categorize/mark_read (max 20 per call). Single Graph `/$batch` round-trip — 10-20× faster than per-message calls for large triage; a message Graph throttles is re-sent after its `Retry-After` rather than reported as failed. `mark_read` takes `true`/`false` (or `read`/`unread`). |

### User and Admin

| Tool | Description |
|------|-------------|
| `outlook_whoami` | Get current user profile. |
| `outlook_list_calendars` | List available calendars, including `owner` (`name`, `address`) and `can_view_private_items`. Missing owner or private-item permission metadata is `null`; `false` means you cannot see private items on that calendar. Use the calendar ID when names repeat. |
| `outlook_list_categories` | List category definitions with colors. |
| `outlook_get_mail_tips` | Pre-send check (OOF, delivery restrictions). |

---

## Prompts

Three workflows ship as MCP prompts, so the common sequences do not have to be
reconstructed call by call. Any MCP client that supports prompts will list them; in most
clients they appear as slash commands or a prompt picker.

| Prompt | Arguments | What it does |
|--------|-----------|--------------|
| `morning_brief` | `folder` (default `inbox`) | Today's events, unread mail and tasks due, in the cheapest order — one scan each, `concise=True`, batched reads. |
| `triage_folder` | `folder` (default `inbox`), `count` (default 50) | One cheap scan of a folder, sorted into reply / archive / junk, applied with a single `outlook_batch_triage` call rather than one call per message. |
| `catch_up` | `since` (default `24h`) | What changed in mail, calendar and contacts, via the delta path — roughly ten times cheaper than re-scanning on a schedule. |

They cost nothing until invoked: `prompts/list` carries only a name and one line each, and
the body is fetched on use.

---

## Configuration

Config lives at `~/.outlook-mcp/config.json` (created with `0600` permissions on macOS and Linux; see **Config permissions** below for Windows). It is read as UTF-8 on every platform, and a byte-order mark is accepted. Set the `OUTLOOK_MCP_CONFIG_DIR` environment variable to move that settings directory (config.json, auth record, and the attachments default move with it) — see [Two accounts, two instances](#two-accounts-two-instances-optional--outlook_mcp_config_dir) below.

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `client_id` | `string` | `null` | Azure AD application (client) ID. Required for auth. |
| `tenant_id` | `string` | `"consumers"` | Azure AD tenant. Use `"consumers"` for personal Microsoft accounts. |
| `expected_username` | `string` | `null` | Since 1.29: optional sign-in username for this instance. A saved sign-in for another username is refused before refreshing it; a new sign-in is checked before saving its record or using its credential. Case-insensitive; compares Microsoft's sign-in username literally, without resolving aliases. Unset keeps the existing behaviour. Older versions ignore this key with a warning and do not enforce the guard. |
| `timezone` | `string` | `"UTC"` | IANA timezone (e.g. `"America/New_York"`). Interprets zone-less dates, **and anchors every event you create** — a recurring event is expanded in this zone, so on the default `"UTC"` a 09:00 weekly meeting shifts an hour when the clocks change. Set it to where you are. |
| `read_only` | `bool` | `false` | When `true`, all write tools (send, reply, move, delete, create, update, RSVP) return an error. Gates the tools, not the Microsoft token -- see below. |
| `mailbox_settings_write` | `bool` | `false` | When `true`, `outlook-mcp auth` asks for `MailboxSettings.ReadWrite` instead of `MailboxSettings.Read`, which is what the inbox-rule write tools need. Consent is per requested scope, so the Azure-portal permission alone does not put it in the token. It also loads the opt-in `rules` toolset group: asking for the scope is asking to manage rules, so the four rule tools come with it. Off by default: it widens the token to every mailbox setting, not only rules. Refused in combination with `read_only_consent`, which asks for the read scopes only — that pairing would silently never request the write scope. |
| `read_only_consent` | `bool` | `false` | When `true`, `outlook-mcp auth` asks Microsoft for the read permissions only, instead of the read-write set. For a second, read-only app registration -- see [What `read_only` does and does not do](#what-read_only-does-and-does-not-do). Requires `read_only: true`; the config is refused without it. |
| `attachments_dir` | `string` | `"~/.outlook-mcp/attachments"` | The only directory the attachment tools may read from or write to. Every path an agent supplies is resolved and must land inside it — a symlink out or a `..` is refused. Widen it only if you understand that anything reachable can be emailed. |
| `allow_categories` | `list[string]` | `[]` | Optional. Restrict write tools to specific categories (see below). Empty list = all writes allowed when `read_only: false`. |
| `allow_unencrypted_token_cache` | `bool` | `false` | Permit the OAuth token cache to be written in cleartext when the platform has no encrypted store (Linux without libsecret). Off by default: authentication stops with an explanation rather than silently persisting a reusable Graph token in plaintext. macOS and Windows always encrypt and are unaffected. |

### Environment configuration (optional)

`auth`, `serve` and `status` share the same loader. It first loads and validates `config.json`
(or the existing defaults when no file exists), then reads these optional variables. A bad,
unreadable or symlinked config still refuses startup; environment values cannot bypass it.
No override is automatically written back to the file.

| Variable | Input | Relation to the file/default value |
|----------|-------|-----------------------------------|
| `OUTLOOK_MCP_CLIENT_ID` | Nonempty string | Supplies `client_id` only when the file has none; a different app registration is refused |
| `OUTLOOK_MCP_TENANT_ID` | Nonempty string | Overrides `tenant_id` |
| `OUTLOOK_MCP_TIMEZONE` | IANA timezone | Overrides `timezone` |
| `OUTLOOK_MCP_READ_ONLY` | `true` or `false` | May enable a restriction; cannot disable one |
| `OUTLOOK_MCP_READ_ONLY_CONSENT` | `true` or `false` | May enable read-only consent; requires effective `read_only: true` |
| `OUTLOOK_MCP_ALLOW_CATEGORIES` | JSON array of category names | May narrow the permitted categories; cannot widen them |
| `OUTLOOK_MCP_ATTACHMENTS_DIR` | Directory path | Must resolve to the configured directory or a descendant |
| `OUTLOOK_MCP_ALLOW_UNENCRYPTED_TOKEN_CACHE` | `true` or `false` | May prohibit plaintext storage; cannot enable it unless the file already permits it |

Unset variables do nothing. A supported variable set to an empty value refuses startup,
including an empty Compose `${VAR}` pass-through. Booleans accept only `true`/`false`
(case-insensitive, surrounding whitespace allowed); list input must be JSON, for example
`["mail_drafts","contacts_write"]`.
An empty category list means **all categories**, so it cannot replace a restricted file list.
Each security setting is checked independently: enabling read-only mode does not excuse widening
another setting. Security permissions that require explicit consent in the file still require it.

Attachment roots are expanded and resolved without creating a directory. An override is kept as
the checked absolute path; `..`, sibling prefixes and symlinks cannot escape the configured root.
Every attachment path check rechecks that root, including when a directory becomes a symlink
after startup.
On Windows, a different drive/share or an unapproved UNC/device location is refused before
resolution can contact it.
Invalid inputs, misspelt security variables and attempts to loosen policy stop startup and name
the variable to fix in the launch environment. Existing `OUTLOOK_MCP_CONFIG_DIR`, toolset and
transport variables retain their meanings. With none of the new variables, file loading is unchanged.
These checks guard the supported setting variables, not the launch environment as a whole:
anyone who controls it can also retarget the settings file with `OUTLOOK_MCP_CONFIG_DIR` or `HOME`.

### Toolset selection (optional) — `OUTLOOK_MCP_TOOLSETS`

The 71 default tool schemas load into the client's context every turn (the chars/4 proxy `test_tool_surface_budget.py` measures with; a different yardstick than the ~8.6k o200k figure in ROADMAP for the 62-tool surface). A client that only needs part of the surface can set the `OUTLOOK_MCP_TOOLSETS` environment variable to a comma-separated list of tool groups, and only those load. The `account` group (auth / identity) is always available.
Groups: `mail`, `drafts`, `attachments`, `calendar`, `contacts`, `todo`, `folders`, `digest`, `delta`, `admin`, `documents`, `rules`, and `mime`. Unset loads every group except the opt-in ones — `documents`, `rules`, and `mime` — so a stock install gets 71 of the 77 tools. Name `rules` here, or set `mailbox_settings_write: true` with `OUTLOOK_MCP_TOOLSETS` unset, to get the four rule tools; that setting is itself the request for rule management. A selection here always wins: the flag does not add `rules` to a narrower list. For example, `OUTLOOK_MCP_TOOLSETS="mail,calendar,digest,delta"` loads 31 tools instead of 71, cutting the per-turn schema cost to about 40% of the default.

Naming any group loads only the groups named, plus `account`, and a name that is not a group matches nothing. To add all opt-in groups to the default set, name the default groups and the opt-ins: `OUTLOOK_MCP_TOOLSETS=mail,drafts,attachments,calendar,contacts,todo,folders,digest,delta,admin,documents,rules,mime`.

The groups are not the headings of the Tool Reference. `mail` holds the mail read, write and triage tools, `outlook_list_thread`, `outlook_copy_message` and `outlook_batch_triage`, but `outlook_list_folders` is in `folders`; the Focused Inbox tools (`outlook_reclassify_message` and the three override tools) are in `admin`, with categories and mail tips; the three `*_delta` tools are in `delta` and `outlook_changes_since` in `digest`; `outlook_list_calendars` is in `calendar`; the two draft-attachment tools are in `drafts`, `outlook_read_attachment` in `documents`, and the four rule tools in `rules`, and MIME export in `mime`; and `account` is `outlook_auth_status` and `outlook_whoami`.

### Two accounts, two instances (optional) — `OUTLOOK_MCP_CONFIG_DIR`

One server process serves one mailbox. To work against two accounts, register **two client entries** and give each its own settings directory with `OUTLOOK_MCP_CONFIG_DIR` — pair it with `OUTLOOK_MCP_TOOLSETS` so each instance also only loads the tool groups it needs:

```json
{
  "mcpServers": {
    "outlook-net": {
      "command": "outlook-mcp",
      "env": {
        "OUTLOOK_MCP_TOOLSETS": "mail,calendar,contacts",
        "OUTLOOK_MCP_CONFIG_DIR": "~/.outlook-mcp-net"
      }
    },
    "outlook-neko": {
      "command": "outlook-mcp",
      "env": {
        "OUTLOOK_MCP_TOOLSETS": "todo",
        "OUTLOOK_MCP_CONFIG_DIR": "~/.outlook-mcp-neko"
      }
    }
  }
}
```

Then run `outlook-mcp auth` once per instance, with the same env set, to write each auth record in its own directory. Each instance reads its own `config.json` (own `client_id`, `timezone`, permissions) from its own directory.

**Only move the config directory — never `HOME`.** The token cache is not in it: it stays in the OS keyring, and on macOS every azure-identity cache on the host shares one Keychain item, coordinated through a signal file — on this server, `~/.IdentityService/outlook-mcp.nocae` (azure-identity appends `.nocae` to every non-CAE cache name; a CAE cache would be a different file over the same item). Both processes must keep consulting that same signal file so their cache writes lock and merge into the one shared entry — which is exactly what moving the config directory preserves and redirecting `HOME` (or the cache location) would break: two signal files that each believe they own the Keychain item overwrite each other's token. `OUTLOOK_MCP_CONFIG_DIR` deliberately moves only where config.json, the auth record, and attachments live; unset or empty keeps the default `~/.outlook-mcp`.

Since 1.29, to require a particular sign-in for an instance, add
`"expected_username": "owner@example.com"` to that instance's `config.json`. Older versions ignore
this key with a warning and do not enforce the guard. This optional guard compares Microsoft's
saved sign-in username case-insensitively; it does not resolve email aliases. A cached mismatch is
refused before a credential is built or refreshed, and `outlook-mcp status` shows the refusal and
remedy without either username: sign in as the expected user, or correct `expected_username` if it
is wrong. A new sign-in is checked after Microsoft's interactive sign-in completes but before its
record is saved or credential used. A refused sign-in leaves the previous auth record unchanged,
but azure-identity may already have added that account's refresh token to the local shared OS token
cache. On macOS this is the same Keychain item, with the previous account's tokens still present.
Nothing in outlook-mcp, including `logout`, removes these cached tokens.
Unset/null preserves today's behavior. Blank values and misspellings such as `expected_usename`
stop startup instead of silently disabling the guard.

### Remote / HTTP transport (optional) — `OUTLOOK_MCP_TRANSPORT`

stdio is the default, and every client entry above uses it. For a server that runs on another machine or in a container, or a client that cannot spawn a local process, `OUTLOOK_MCP_TRANSPORT=streamable-http` serves MCP's streamable-HTTP transport at `/mcp` instead:

```bash
export OUTLOOK_MCP_TRANSPORT=streamable-http
export OUTLOOK_MCP_HTTP_HOST=127.0.0.1   # default 127.0.0.1
export OUTLOOK_MCP_HTTP_PORT=8000      # default 8000
export OUTLOOK_MCP_HTTP_TOKEN=...      # python -c "import secrets; print(secrets.token_urlsafe(32))"
outlook-mcp serve
```

The client then connects to `https://<your proxy>/mcp` and sends `Authorization: Bearer <token>`.

`OUTLOOK_MCP_TRANSPORT` takes `stdio` (the default, also when unset or empty) or `streamable-http`. Any other value, a port outside 1–65535, a malformed `OUTLOOK_MCP_HTTP_ALLOWED_HOSTS`, or a missing or malformed token stops the server at startup with the reason, before anything binds.

For a container, explicitly set `OUTLOOK_MCP_HTTP_HOST=0.0.0.0` and `OUTLOOK_MCP_HTTP_ALLOWED_HOSTS=mcp.example.com` for the public hostname your proxy forwards. Every public bind prints a plain-HTTP warning before it starts. A wildcard bind (`0.0.0.0` or `::`) without allowed hosts also warns that normal client Host headers will receive 421: the literal wildcard address is not the hostname or address clients use. The token protects against other OS users and browser pages; code running as your OS user can read the client configuration or launch the server itself. Forwarded headers are not trusted.

A secret file can replace the environment token: set `OUTLOOK_MCP_HTTP_TOKEN_FILE` to a regular file of at most 4096 bytes containing only the token (an ending newline is fine), and unset `OUTLOOK_MCP_HTTP_TOKEN`. Both set is refused. On POSIX the file must belong to the server user with no group or other permissions (`chmod 600`); symlinks are refused. On Windows, restrict the file ACL to the server account yourself; the server does not validate Windows ACLs. File contents and paths are never included in token-file errors.

**Container secret mounts need the same file protections.**

- **Kubernetes:** projected Secret keys are symlinks, so their usual volume paths are refused. A file mounted with `subPath` avoids that symlink, but must still belong to the server UID and have mode `0400` or `0600`; `subPath` alone does not fix ownership. For a non-root server, an init container can copy the projected key into a writable `emptyDir` as the server UID with mode `0400`, then the application mounts that prepared directory read-only. Recreate the Pod after rotation to refresh the copy or `subPath` mount; [subPath mounts do not receive Secret updates](https://kubernetes.io/docs/concepts/configuration/secret/#using-secrets-as-files-from-a-pod).
- **Docker:** [Swarm secrets](https://docs.docker.com/engine/swarm/secrets/) default to root-owned, world-readable files, which are refused. Set the secret's `uid` to the server UID, `gid` to its group and `mode` to `0400`. For [Compose secrets backed by a local file](https://docs.docker.com/reference/compose-file/services/#secrets), `uid`, `gid` and `mode` are ignored: prepare a regular host file with the matching container UID and mode `0400`, and mount it read-only. Set `OUTLOOK_MCP_HTTP_TOKEN_FILE` to the resulting in-container file path. Restart the server after replacing a token; it reads the file only at startup.

**This changes who can reach the mailbox.** Over stdio, only the process that started the server can. Over HTTP, one request to the port is enough for anything that holds a token, so the token is the whole of the access control:

- **A token is required on every bind, loopback included.** The server refuses to start without a token from `OUTLOOK_MCP_HTTP_TOKEN` or `OUTLOOK_MCP_HTTP_TOKEN_FILE`. On loopback, a port with no token would let other OS users or browser pages read the mailbox with one HTTP request. Every request must carry `Authorization: Bearer <token>` or gets `401`. The token is one value of at least 32 characters from `A–Z a–z 0–9 -._~+/`, at least 8 of them different (trailing `=` padding is allowed). This rejects obvious low-variety values, not every predictable token; use the generator above. To change it, restart the server with the new one.
- **Plain HTTP sends the token in the clear.** The server does not terminate TLS. On any bind other machines can reach, the token and everything the tools return cross the network readable by anything on the path, unless a reverse proxy in front of the server terminates TLS. Put it behind one, and let the proxy authenticate callers too if more than one client will reach it.
- **Host and Origin checks apply to every bind.** Loopback binds accept loopback hosts and origins. `OUTLOOK_MCP_HTTP_ALLOWED_HOSTS` optionally adds comma-separated hostnames or IPs, with an optional port (for example `mcp.example.com` or `mcp.example.com:8443`). On a public bind it selects the permitted hosts; unset permits only the literal bind address. Bare names and `:*` entries permit any port, including a bare Host on the default HTTP/HTTPS port. Invalid Host is refused with 421; a supplied Origin must match an allowed host over HTTP or HTTPS, or gets 403. Clients without Origin work. A same-host proxy can send `Host: 127.0.0.1:<port>`.
- **File paths are on the server.** Attach and send-with-attachments paths resolve inside the server’s `attachments_dir`; a remote client cannot hand over its own local path. A downloaded file and `saved_to` stay on the server, and a path refusal may name the server’s absolute attachments directory. Inline `attachments` and `as_content=True` carry the bytes in the call instead (see [Attachments](#attachments)).
- **It is stateless.** The server issues no `Mcp-Session-Id` and answers each request with JSON rather than an event stream. No tool sends anything to the client mid-call, so nothing is lost.
- **One process still serves one account.** Pair it with `OUTLOOK_MCP_CONFIG_DIR`, as for stdio (one process, one port and one settings directory per mailbox), and run `outlook-mcp auth` where the server runs, with the same environment. A container usually has no keyring; see [Token cache stored unencrypted (Linux)](#token-cache-stored-unencrypted-linux).
- No token, request URL or request line is logged.

### What `read_only` does and does not do

`read_only: true` stops outlook-mcp's write tools from running. Ask it to send mail and it
refuses.

**It does not make your Microsoft credential read-only.** The first sign-in consents the full read-write set, even with `read_only: true` in the config — the scopes a first consent leaves out can never be added without logging in again (every refresh afterwards uses `.default` -- "everything this account has already consented" -- so a session granted only read scopes would fail every write with 403 no matter what the config says). The stored token can send mail whether `read_only` is on or off, and a session consented before you turned `read_only` on keeps its write scopes.

Two consequences worth understanding:

- `read_only` is a line in a text file. Anything able to edit `~/.outlook-mcp/config.json`
  turns it off and has write access immediately -- no re-authentication, no new consent
  prompt.
- The enforcement lives in this server's Python code. Any other process holding the cached
  token is unaffected by it.

So treat `read_only` as a guardrail against an agent doing something rash, **not as a
security boundary**. If you want a credential that genuinely cannot write, have Microsoft
enforce it rather than us:

1. Register a **second** Azure app and give it the read permissions only: `Mail.Read`,
   `Calendars.Read`, `Contacts.Read`, `Tasks.Read`, `MailboxSettings.Read`, `User.Read`. It has
   to be an app this account has never granted write access to. Microsoft remembers consent
   per app, and a refresh returns everything that app was ever granted.
2. Point `client_id` at it and set both keys:

   ```json
   { "client_id": "<the read-only app>", "read_only": true, "read_only_consent": true }
   ```

3. Run `outlook-mcp auth`. With `read_only_consent` the consent screen lists the read
   permissions only. Without it, sign-in would ask this app for the read-write set too.

`read_only_consent` is refused without `read_only: true`: a sign-in that asked only for read
access cannot write, so a server expecting writes would fail on every one. And a sign-in
belongs to the app it was made with. After `client_id` changes, the server will not use the
old one — `outlook-mcp status` and `outlook_auth_status` say so until you run
`outlook-mcp auth` again.

### Granular Write Permissions (optional)

By default, `read_only: false` unlocks **all** write tools. For finer control, set `allow_categories` to restrict write access to specific categories. Read tools (list, search, get) are always allowed — `allow_categories` only narrows the write surface.

**Available categories:**

| Category | Tools | Risk |
|---|---|---|
| `mail_drafts` | create/update/delete draft, attach to/remove from a draft | Safe — drafts only, no send |
| `mail_triage` | move, delete (soft), flag, categorize, mark read, copy, batch, reclassify and set/delete Focused Inbox overrides | Moderate — reversible except hard delete |
| `mail_folders` | create/rename/delete folder | Moderate |
| `mail_rules` | create/update/delete inbox rule | **Dangerous** — rules run server-side, so one keeps moving, forwarding or deleting mail with every client closed |
| `mail_send` | send, reply, forward, send_draft, send_with_attachments | **Dangerous** — sends email on your behalf |
| `calendar_write` | create/update/delete event, RSVP, attach/remove event files | Moderate — your own calendar, but `outlook_attach_to_event` reads local files from `attachments_dir` and pushes their bytes to Graph. The parts that email other people text the agent wrote need `mail_send` as well: inviting attendees, rewording (subject, body, location) any event that already has them, attaching a file to such an event (they get it with its next update), and adding a message to an RSVP. A bare RSVP, a time change and a cancellation still notify the people involved, but carry nothing the agent wrote |
| `contacts_write` | create/update/delete contact | Moderate |
| `todo_write` | create/update/complete/delete task, checklist items; upload/delete task attachments | Moderate — your own task list, but `outlook_upload_task_attachment` reads local files from `attachments_dir` and pushes their bytes to Graph, and task/checklist/attachment deletes are irreversible. Listing and downloading attachments are plain reads, gated like every other read (not at all) and fenced to `attachments_dir` |

**Example policies:**

**Draft-only assistant** (agent can compose drafts, you review and send):

```json
{ "read_only": false, "allow_categories": ["mail_drafts", "mail_triage", "todo_write"] }
```

(Note that `todo_write` includes the *write-side* task-attachment tools — file reads
from `attachments_dir`, uploads to Graph, and irreversible deletes. Listing and
downloading task attachments are reads and are not write-gated, like the mail
attachment reads — see the table above.)

**Calendar-only** (agent can manage your schedule, nothing else):

```json
{ "read_only": false, "allow_categories": ["calendar_write"] }
```

(It cannot invite anyone. An invitation is an email, so attendees need `mail_send` too — add it
if the agent should set up meetings with other people, not just block out your own time.)

**Full write access** (agent can do everything):

```json
{ "read_only": false }
```

**Read-only** (safest default, no writes):

```json
{ "read_only": true }
```

When `allow_categories` is set, any tool in a non-allowed category returns a permission-denied error (`PermissionDeniedError`) naming the blocked category. When `allow_categories` is empty (or unset) and `read_only` is false, all write tools are permitted. `read_only: true` always takes precedence — if set, all writes are blocked regardless of `allow_categories`. Unknown category names are rejected at config load time with a validation error; only the eight names above are accepted.

---

## Privacy and Security

- **Zero telemetry.** No analytics, no tracking, no usage data collected.
- **Zero local caching.** Every call goes directly to Microsoft Graph. No local email/calendar storage. (One carve-out: the To Do default-list id is resolved once per process and kept for the session — an id, not content; see `outlook_list_tasks`.)
- **Zero third-party calls.** The server only talks to `graph.microsoft.com` and `login.microsoftonline.com`.
- **Token storage.** OAuth tokens are persisted via `azure-identity`'s `TokenCachePersistenceOptions`. On macOS the OS Keychain is used; on Windows, DPAPI; on Linux with PyGObject/libsecret available, gnome-keyring. On Linux *without* libsecret (e.g. the isolated venv created by `uv tool install`), the server refuses to persist tokens unless `allow_unencrypted_token_cache` is `true`; with it, they fall back to a `0600` plaintext file at `~/.IdentityService/` and the MCP logs a one-time warning at startup. For encrypted storage on Linux, install `python3-gi gnome-keyring libsecret-1-0` and re-create the venv with `--system-site-packages`.
- **No logging of sensitive data.** Message bodies, recipient addresses, and tokens are never logged.
- **Config permissions.** On macOS and Linux the config directory is created `0700` and the config file `0600`, and a loose mode on the file is repaired on load. On Windows those POSIX modes cannot be enforced — `os.chmod` there sets only the read-only attribute — so access is governed by the path's Windows ACL, including whatever it inherits from the directory it was created under, which this server neither applies nor verifies. Symlinked configs are rejected on every platform.
- **Input validation.** All user inputs (email addresses, Graph IDs, OData filters, KQL queries, datetimes) are validated and sanitized before reaching the Graph API.
- **HTTP transport (opt-in).** It changes who can reach the mailbox: a bearer token is required on every bind and checked on every request, and the server speaks plain HTTP, so TLS belongs to a proxy in front. The whole security story is in [Remote / HTTP transport](#remote--http-transport-optional--outlook_mcp_transport).

---

## Development

```bash
# Install dev dependencies and the documents extra, as CI does (without it the PDF/XLSX tests skip)
uv sync --all-extras

# Run tests
uv run pytest

# Lint
uv run ruff check src/ tests/ scripts/

# Format (CI fails if `ruff format --check` would change a file)
uv run ruff format src/ tests/ scripts/

# Run server locally (stdio)
uv run outlook-mcp
```

**Requirements:** Python 3.10+

---

## Roadmap

- **Advanced mail** -- internet message headers
- **Calendar** -- cancel event (with attendee notification)
- **Work/school-only tools** -- free/busy, meeting times, automatic replies, shared mailboxes: not planned for now (#133). If any are taken later, they come as an opt-in group that personal accounts never load

Committed direction, and what was investigated and ruled out (mailbox settings), is in [ROADMAP.md](ROADMAP.md).

---

## License

MIT. See [LICENSE](LICENSE).
