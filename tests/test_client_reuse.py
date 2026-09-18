"""Tests for Graph client reuse across tool calls (connection pooling).

`_get_graph_client` should cache one `GraphClient` per account in the lifespan
context and reuse it while that account's credential is unchanged, rebuilding
only when auth swaps the credential (switch_account / re-auth). Building a
GraphServiceClient — auth provider, request adapter, TLS pool — per call is
wasteful on recurring loops. The account for a call comes from the calling
tool's name via the routing table, so one session routes mail tools and todo
tools to different (cached) clients in a split setup.

The last test drives the REAL SDK client (`Client(mcp)`) so the
request-context plumbing `_calling_tool_name` depends on is exercised as
shipped, not as a hand-rolled MagicMock approximates it (review of #61).
"""

from unittest.mock import MagicMock

import pytest
from mcp.client import Client

from outlook_mcp import server as server_mod
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import AccountConfig, Config
from outlook_mcp.errors import OutlookMCPError


def _split_config() -> Config:
    return Config(
        accounts=[
            AccountConfig(name="net", client_id="id1-abcd"),
            AccountConfig(name="neko", client_id="id2-efgh"),
        ],
        capability_accounts={"mail": "net", "todo": "neko"},
    )


def _ctx_with_auth(auth, tool_name="outlook_list_inbox", config=None):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {
        "auth": auth,
        "config": config if config is not None else Config(),
    }
    if tool_name is None:
        # What a real context looks like when the SDK did not expose params.
        ctx.request_context.params = None
    else:
        ctx.request_context.params = {"name": tool_name, "arguments": {}}
    return ctx


def test_get_graph_client_reuses_instance_for_same_credential():
    """Repeated calls with the same credential return the same GraphClient."""
    cred = MagicMock()
    auth = MagicMock()
    auth.resolve_capability_account.return_value = None
    auth.get_account_credential.return_value = cred
    ctx = _ctx_with_auth(auth)

    first = server_mod._get_graph_client(ctx)
    second = server_mod._get_graph_client(ctx)

    assert first is second


def test_get_graph_client_rebuilds_when_credential_changes():
    """A new credential (switch_account / re-auth) rebuilds the client."""
    cred1 = MagicMock()
    cred2 = MagicMock()
    auth = MagicMock()
    auth.resolve_capability_account.return_value = None
    # same cred twice, then a switched credential
    auth.get_account_credential.side_effect = [cred1, cred1, cred2]
    ctx = _ctx_with_auth(auth)

    first = server_mod._get_graph_client(ctx)
    second = server_mod._get_graph_client(ctx)
    third = server_mod._get_graph_client(ctx)

    assert first is second
    assert third is not first
    assert third.credential is cred2


def test_accounts_get_separate_cached_clients():
    """A split routing (mail->net, todo->neko) caches one client per account."""
    creds = {"net": MagicMock(), "neko": MagicMock()}
    auth = MagicMock()
    auth.resolve_capability_account.side_effect = lambda cap: "net" if cap == "mail" else "neko"
    auth.get_account_credential.side_effect = lambda name: creds[name]
    ctx = _ctx_with_auth(auth, config=_split_config())  # one session, one lifespan cache

    ctx.request_context.params = {"name": "outlook_list_inbox", "arguments": {}}
    mail_client = server_mod._get_graph_client(ctx)
    ctx.request_context.params = {"name": "outlook_list_tasks", "arguments": {}}
    todo_client = server_mod._get_graph_client(ctx)
    ctx.request_context.params = {"name": "outlook_search_mail", "arguments": {}}
    mail_again = server_mod._get_graph_client(ctx)

    assert mail_client is not todo_client
    assert mail_client.credential is creds["net"]
    assert todo_client.credential is creds["neko"]
    assert mail_again is mail_client


def test_client_routing_reads_the_calling_tool_name():
    """The capability — and so the account — comes from the tool actually running."""
    auth = MagicMock()
    auth.resolve_capability_account.side_effect = lambda cap: f"routed:{cap}"
    auth.get_account_credential.return_value = MagicMock()
    mail_ctx = _ctx_with_auth(auth, "outlook_list_events")
    identity_ctx = _ctx_with_auth(auth, "outlook_whoami")

    server_mod._get_graph_client(mail_ctx)
    server_mod._get_graph_client(identity_ctx)

    # calendar tool routed via "calendar"; identity tool follows the active account
    auth.resolve_capability_account.assert_any_call("calendar")
    auth.resolve_capability_account.assert_any_call(None)


# ── Fail closed when the routing inputs are missing ──────────────────


def test_unreadable_tool_name_fails_closed_on_multi_account():
    """Accounts configured + no tool name from the request context: refuse.

    capability_for(None) is None — the same answer identity tools get — so
    without this check the call would route to the active account, writes
    included, with no error and no log (review of #61)."""
    auth = MagicMock()
    ctx = _ctx_with_auth(auth, tool_name=None, config=_split_config())

    with pytest.raises(OutlookMCPError, match="refusing to route"):
        server_mod._get_graph_client(ctx)
    auth.get_account_credential.assert_not_called()


def test_unroutable_tool_fails_closed_on_multi_account():
    """A tool name with no routing entry (drift) must not fall through to the
    active account either."""
    auth = MagicMock()
    ctx = _ctx_with_auth(auth, tool_name="outlook_not_in_the_table", config=_split_config())

    with pytest.raises(OutlookMCPError, match="no capability routing"):
        server_mod._get_graph_client(ctx)


def test_identity_tools_pass_with_accounts_configured():
    """The account-group tools legitimately have capability None — they follow
    the active account and must not trip the fail-closed guard."""
    cred = MagicMock()
    auth = MagicMock()
    auth.resolve_capability_account.return_value = "net"
    auth.get_account_credential.return_value = cred
    ctx = _ctx_with_auth(auth, tool_name="outlook_whoami", config=_split_config())

    assert server_mod._get_graph_client(ctx).credential is cred


def test_single_account_never_needs_the_tool_name():
    """Legacy installs (no accounts) route everything to the one credential,
    fake contexts included — the fail-closed guard is multi-account only."""
    cred = MagicMock()
    auth = MagicMock()
    auth.resolve_capability_account.return_value = None
    auth.get_account_credential.return_value = cred
    ctx = _ctx_with_auth(auth, tool_name=None, config=Config())

    assert server_mod._get_graph_client(ctx).credential is cred


# ── Through the real SDK ──────────────────────────────────────────────


@pytest.mark.asyncio
async def test_real_sdk_client_gives_mail_and_todo_different_credentials(monkeypatch):
    """Drive the server with the real `Client(mcp)`: the inbound call's params
    must reach _calling_tool_name through the SDK's actual request context,
    and a split routing must hand mail and todo different credentials.

    Guards the plumbing (ServerRequestContext.params as a Mapping carrying
    {name, arguments}) against SDK changes, not just our own bookkeeping
    (review of #61)."""
    from outlook_mcp.tools import mail_read
    from outlook_mcp.tools import todo as todo_tools

    config = _split_config()
    monkeypatch.setattr(server_mod, "load_config", lambda: config)

    creds: dict[str, MagicMock] = {}

    def fake_try_cached_token(self):
        for acc in config.accounts:
            creds[acc.name] = MagicMock()
            self._credentials[acc.name] = creds[acc.name]
        self.credential = creds[config.default_account]
        return True

    monkeypatch.setattr(AuthManager, "try_cached_token", fake_try_cached_token)

    real_graph_client = server_mod.GraphClient
    built: list = []

    class CapturingGraphClient(real_graph_client):
        def __init__(self, credential):
            super().__init__(credential)
            built.append(self)

    def credential_of(sdk_client):
        owner = next(c for c in built if c.sdk_client is sdk_client)
        return owner.credential

    seen: dict[str, object] = {}

    async def fake_list_inbox(sdk_client, *args, **kwargs):
        seen["mail"] = credential_of(sdk_client)
        return {}

    async def fake_list_tasks(sdk_client, *args, **kwargs):
        seen["todo"] = credential_of(sdk_client)
        return {}

    monkeypatch.setattr(server_mod, "GraphClient", CapturingGraphClient)
    monkeypatch.setattr(mail_read, "list_inbox", fake_list_inbox)
    monkeypatch.setattr(todo_tools, "list_tasks", fake_list_tasks)

    async with Client(server_mod.mcp) as client:
        await client.call_tool("outlook_list_inbox", {})
        await client.call_tool("outlook_list_tasks", {})

    # Different credentials per surface, each exactly its configured account's.
    assert seen["mail"] is not seen["todo"]
    names_by_cred = {id(cred): name for name, cred in creds.items()}
    assert names_by_cred[id(seen["mail"])] == "net"
    assert names_by_cred[id(seen["todo"])] == "neko"
    # One cached client per account, not one per call.
    assert len(built) == 2
