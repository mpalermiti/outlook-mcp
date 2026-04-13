"""Tests for Graph client factory."""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.errors import AuthRequiredError
from outlook_mcp.graph import GraphClient


def test_graph_client_requires_credential():
    """GraphClient raises without credential."""
    with pytest.raises(AuthRequiredError):
        GraphClient(credential=None)


def test_graph_client_init():
    """GraphClient initializes with a credential and creates sdk_client."""
    mock_credential = MagicMock()
    client = GraphClient(credential=mock_credential)
    assert client.sdk_client is not None
