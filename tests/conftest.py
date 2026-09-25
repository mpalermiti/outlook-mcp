"""Shared test fixtures for outlook-mcp.

Two tiers here need real credentials and are skipped without them:

``-m integration``
    End-to-end plumbing smoke tests — does each tool return its documented
    response shape against real Graph.

``-m live``
    Query-shape regression guards. These exist because the default suite mocks
    the Graph client and therefore *cannot* observe a query that Graph rejects
    (400) or silently mis-evaluates (200 with wrong results). Every bug fixed
    in #30, #31 and the ``list_thread`` repair was invisible to 558 green mock
    tests. See ``tests/test_live_query_shape.py``.

``-m live_write``
    The only tier that mutates a real mailbox, added for #41. Some
    payloads can only be validated by Graph accepting them — a recurrence that
    is well-formed to the SDK still 400s if the range disagrees with the start
    — and a mock cannot see that. Kept as narrow as possible.

Tiers 1 and 2 are READ-ONLY: never add a fixture or test there that writes to
the mailbox — no sends, drafts, category or folder mutations, no deletes.

Rules for ``live_write``, which exists precisely because it breaks that:

* Double-gated — the marker is deselected by default *and* the tier skips
  unless ``OUTLOOK_MCP_LIVE_WRITE=1``. A cached token alone must never be
  enough to write to someone's calendar.
* Calendar, contacts and To Do only — the three surfaces whose writes have no
  outward side effect. No mail: nothing here may send, and a stray send is not
  recoverable.
* No attendees, ever. A recurring invite emails real people on every
  occurrence.
* Bounded ranges only (``numbered``/``endDate``) — never ``noEnd``.
* Every created item is removed in a ``finally``, and subjects/names carry
  ``LIVE_WRITE_SUBJECT`` / ``LIVE_WRITE_NAME`` so anything a crash leaks is
  greppable in the UI.
"""

import os

import pytest

# Prefix for every item this tier creates, so an orphan left by a hard crash is
# obvious in the calendar UI and easy to search for.
LIVE_WRITE_SUBJECT = "[outlook-mcp live-write guard] safe to delete"
# Contacts and tasks have no subject; their names carry the same marker.
LIVE_WRITE_NAME = "LiveWriteGuard-SafeToDelete"


@pytest.fixture
def mock_graph_client():
    """Mock Microsoft Graph client for unit tests."""
    pass


# ── Live-credential fixtures (shared by the integration and live tiers) ──


@pytest.fixture(scope="session")
def real_config():
    """Load real config from ~/.outlook-mcp/config.json."""
    from outlook_mcp.config import load_config

    config = load_config()
    if not config.client_id:
        pytest.skip("No client_id configured — run the Azure AD app setup first")
    return config


@pytest.fixture(scope="session")
def real_auth(real_config):
    """AuthManager backed by the cached token, or skip.

    Uses ``try_cached_token`` — the same silent path the MCP server uses on
    startup. It must never trigger an interactive device-code prompt, since
    these run unattended.
    """
    from outlook_mcp.auth import AuthManager

    auth = AuthManager(real_config)
    if not auth.try_cached_token():
        pytest.skip("Not authenticated — run `outlook-mcp auth` on this host first")
    return auth


@pytest.fixture
def real_graph_client(real_auth):
    """Real Graph client built from the cached credential.

    Function-scoped on purpose: the underlying kiota/httpx transport binds to
    the running event loop, and pytest-asyncio gives each test a fresh one. A
    session-scoped client raises ``RuntimeError: Event loop is closed`` on the
    second test that uses it. ``real_config``/``real_auth`` stay session-scoped
    — they hold no loop-bound state, so the token is fetched once.
    """
    from outlook_mcp.graph import GraphClient

    return GraphClient(real_auth.get_credential())


# Every personal Microsoft account lives in this one well-known tenant; a work
# or school account carries its organisation's tenant instead. It is a public
# constant published by Microsoft, not a secret.
MSA_CONSUMER_TENANT = "9188040d-6c67-4c5b-b112-36a304b66dad"


def consumer_skip_reason(record) -> str | None:
    """Why this mailbox cannot answer a consumer-only assertion — None if it can.

    Split out from the fixture so it can be tested without a real token.
    """
    if record is None:
        return (
            "No auth record on disk, so the account type is unknown — this assertion "
            "is only meaningful on a personal Microsoft account"
        )
    if record.tenant_id != MSA_CONSUMER_TENANT:
        return (
            f"Work/school account (tenant {record.tenant_id}). This assertion is a claim "
            f"about personal Microsoft accounts (tenant {MSA_CONSUMER_TENANT}); Graph may "
            f"legitimately behave differently here, so a failure would say nothing about "
            f"the behaviour under test"
        )
    return None


@pytest.fixture(scope="session")
def consumer_mailbox_only(real_auth):
    """Skip unless the cached token belongs to a personal Microsoft account.

    Some live assertions are claims about *consumer* mailbox behaviour — Graph
    accepting a field and silently dropping it, say. On a work or school account
    Graph may honour the very thing the test says is unsupported, so the test
    fails correctly and tells us nothing about the claim.

    That is not hypothetical. ``test_online_meeting_is_not_supported_on_personal_accounts``
    reached two separate contributor PRs as a red test each author had to
    explain was not theirs, while passing on the maintainer's mailbox
    throughout. The tier already refuses to assume whose mailbox it is running
    against; this extends that to what *kind* of mailbox it is.

    The skip names the tenant it found, because a skip that does not say why it
    skipped is indistinguishable from coverage.
    """
    from outlook_mcp.auth import _load_auth_record

    reason = consumer_skip_reason(_load_auth_record())
    if reason:
        pytest.skip(reason)


@pytest.fixture
def live_write_config(real_config, tmp_path):
    """Real config for the write tier, or skip.

    Gates on an explicit environment opt-in beyond the marker: these tests
    create and delete real calendar events on whatever account the cached token
    belongs to.

    The config handed back is a sandboxed copy: `attachments_dir` is pointed
    at a per-test tmp_path so the attachment round-trip never stages files in
    the user's real settings directory (which a shell-exported
    `OUTLOOK_MCP_CONFIG_DIR` may have moved somewhere with other uses).
    """
    if os.environ.get("OUTLOOK_MCP_LIVE_WRITE") != "1":
        pytest.skip(
            "Write tier is opt-in — set OUTLOOK_MCP_LIVE_WRITE=1 to let it "
            "create and delete events, contacts and tasks on the authenticated account"
        )
    if real_config.read_only:
        pytest.skip("Config is read_only — the write tier cannot run")

    # The real config may whitelist only some write categories (`allow_categories`
    # is a policy for the *server*, e.g. "my agent may not touch contacts"). The
    # tier's job is to exercise Graph, and the permission gate is unit-tested on
    # its own, so open the three surfaces this tier is allowed to write on an
    # in-process copy. ~/.outlook-mcp/config.json is never modified.
    update = {"attachments_dir": str(tmp_path / "attachments")}
    if real_config.allow_categories:
        from outlook_mcp.permissions import (
            CATEGORY_CALENDAR_WRITE,
            CATEGORY_CONTACTS_WRITE,
            CATEGORY_TODO_WRITE,
        )

        needed = {CATEGORY_CALENDAR_WRITE, CATEGORY_CONTACTS_WRITE, CATEGORY_TODO_WRITE}
        update["allow_categories"] = sorted(set(real_config.allow_categories) | needed)
    return real_config.model_copy(update=update)
