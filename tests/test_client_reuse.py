"""Tests for Graph client reuse across tool calls (connection pooling).

`_get_graph_client` should cache one `GraphClient` in the lifespan context and
reuse it while the credential is unchanged, rebuilding only when auth swaps the
credential (switch_account / re-auth). Building a GraphServiceClient — auth
provider, request adapter, TLS pool — per call is wasteful on recurring loops.
"""

from unittest.mock import MagicMock

from outlook_mcp import server as server_mod


def _ctx_with_auth(auth):
    ctx = MagicMock()
    ctx.request_context.lifespan_context = {"auth": auth, "config": MagicMock()}
    return ctx


def test_get_graph_client_reuses_instance_for_same_credential():
    """Repeated calls with the same credential return the same GraphClient."""
    cred = MagicMock()
    auth = MagicMock()
    auth.get_credential.return_value = cred
    ctx = _ctx_with_auth(auth)

    first = server_mod._get_graph_client(ctx)
    second = server_mod._get_graph_client(ctx)

    assert first is second


def test_get_graph_client_rebuilds_when_credential_changes():
    """A new credential (switch_account / re-auth) rebuilds the client."""
    cred1 = MagicMock()
    cred2 = MagicMock()
    auth = MagicMock()
    # same cred twice, then a switched credential
    auth.get_credential.side_effect = [cred1, cred1, cred2]
    ctx = _ctx_with_auth(auth)

    first = server_mod._get_graph_client(ctx)
    second = server_mod._get_graph_client(ctx)
    third = server_mod._get_graph_client(ctx)

    assert first is second
    assert third is not first
    assert third.credential is cred2
