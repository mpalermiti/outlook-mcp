"""Tests for Graph client factory."""

from unittest.mock import MagicMock, patch

import pytest

from outlook_mcp.errors import AuthRequiredError
from outlook_mcp.graph import GraphClient


def test_graph_client_requires_credential():
    """GraphClient raises without credential."""
    with pytest.raises(AuthRequiredError):
        GraphClient(credential=None)


@patch("outlook_mcp.graph.GraphServiceClient")
def test_graph_client_init(mock_gsc_cls):
    """GraphClient initializes with a credential."""
    mock_credential = MagicMock()
    client = GraphClient(credential=mock_credential)
    assert client.sdk_client is not None
    mock_gsc_cls.assert_called_once_with(credentials=mock_credential)
