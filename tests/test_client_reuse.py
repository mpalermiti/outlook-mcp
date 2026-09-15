"""Tests for Graph client reuse across tool calls (connection pooling).

`_get_graph_client` should cache one `GraphClient` per account in the lifespan
context and reuse it while that account's credential is unchanged, rebuilding
only when auth swaps the credential (switch_account / re-auth). Building a
GraphServiceClient — auth provider, request adapter, TLS pool — per call is
wasteful on recurring loops. The account for a call comes from the calling
tool's name via the routing table, so the same ctx routes mail tools and todo
tools to different (cached) clients in a split setup.
"""

from unittest.mock import MagicMock

from outlook_mcp import server as server_mod


def _ctx_with_auth(auth, tool_name="outlook_list_inbox"):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {"auth": auth, "config": MagicMock()}
    ctx.request_context.params = {"name": tool_name, "arguments": {}}
    return ctx


def test_get_graph_client_reuses_instance_for_same_credential():
    """Repeated calls with the same credential return the same GraphClient."""
    cred = MagicMock()
    auth = MagicMock()
    auth.resolve_capability_account.return_value = "net"
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
    auth.resolve_capability_account.return_value = "net"
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
    ctx = _ctx_with_auth(auth)  # one session, one lifespan cache

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
    mail_ctx = _ctx_with_auth(auth, "outlook_list_events")
    identity_ctx = _ctx_with_auth(auth, "outlook_whoami")

    server_mod._get_graph_client(mail_ctx)
    server_mod._get_graph_client(identity_ctx)

    # calendar tool routed via "calendar"; identity tool follows the active account
    auth.resolve_capability_account.assert_any_call("calendar")
    auth.resolve_capability_account.assert_any_call(None)
