"""Tests for Focused Inbox per-sender override CRUD tools."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest
from msgraph.generated.models.inference_classification_type import (
    InferenceClassificationType,
)

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.inference_overrides import (
    delete_inbox_override,
    list_inbox_overrides,
    set_inbox_override,
)

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)
_CFG_ALLOW_TRIAGE = Config(client_id="test", allow_categories=["mail_triage"])
_CFG_ALLOW_OTHER = Config(client_id="test", allow_categories=["mail_send"])


def _make_override(
    override_id: str,
    address: str,
    name: str,
    classify_as: InferenceClassificationType,
):
    """Build a mock InferenceClassificationOverride with nested email address."""
    email = MagicMock()
    email.address = address
    email.name = name
    obj = MagicMock()
    obj.id = override_id
    obj.classify_as = classify_as
    obj.sender_email_address = email
    return obj


def _wire_list(mock_client, overrides_value):
    """Wire `graph_client.me.inference_classification.overrides.get()`."""
    overrides_builder = mock_client.me.inference_classification.overrides
    response = MagicMock()
    response.value = overrides_value
    overrides_builder.get = AsyncMock(return_value=response)
    return overrides_builder


class TestListInboxOverrides:
    async def test_empty(self):
        """Returns empty list when SDK .value is empty."""
        mock_client = MagicMock()
        _wire_list(mock_client, [])
        result = await list_inbox_overrides(mock_client)
        assert result == {"overrides": [], "count": 0}

    async def test_maps_two_overrides(self):
        """Returns mapped overrides with correct classify_as strings."""
        mock_client = MagicMock()
        ov1 = _make_override(
            "ABC=", "alice@example.com", "Alice", InferenceClassificationType.Focused
        )
        ov2 = _make_override("DEF=", "bob@example.com", "Bob", InferenceClassificationType.Other)
        _wire_list(mock_client, [ov1, ov2])

        result = await list_inbox_overrides(mock_client)
        assert result["count"] == 2
        assert result["overrides"][0] == {
            "id": "ABC=",
            "sender_email": "alice@example.com",
            "sender_name": "Alice",
            "classify_as": "focused",
        }
        assert result["overrides"][1] == {
            "id": "DEF=",
            "sender_email": "bob@example.com",
            "sender_name": "Bob",
            "classify_as": "other",
        }

    async def test_sanitizes_output(self):
        """Control characters in returned name/email are scrubbed."""
        mock_client = MagicMock()
        ov = _make_override(
            "ABC=",
            "evil\x00@example.com",
            "Bad\x1bName",
            InferenceClassificationType.Focused,
        )
        _wire_list(mock_client, [ov])

        result = await list_inbox_overrides(mock_client)
        assert "\x00" not in result["overrides"][0]["sender_email"]
        assert "\x1b" not in result["overrides"][0]["sender_name"]


class TestSetInboxOverride:
    async def test_post_when_no_existing(self):
        """Creates a new override via POST when none exists for the sender."""
        mock_client = MagicMock()
        overrides_builder = _wire_list(mock_client, [])

        created = _make_override(
            "NEW=", "user@example.com", "", InferenceClassificationType.Focused
        )
        overrides_builder.post = AsyncMock(return_value=created)

        result = await set_inbox_override(
            mock_client,
            sender_email="user@example.com",
            classify_as="focused",
            config=_CFG,
        )

        overrides_builder.post.assert_called_once()
        body = overrides_builder.post.call_args[0][0]
        assert body.classify_as == InferenceClassificationType.Focused
        assert body.sender_email_address.address == "user@example.com"

        assert result == {
            "status": "created",
            "id": "NEW=",
            "sender_email": "user@example.com",
            "classify_as": "focused",
        }

    async def test_patch_when_existing_case_insensitive(self):
        """PATCHes when an override exists for the sender (case-insensitive match)."""
        mock_client = MagicMock()
        existing = _make_override(
            "EXIST=",
            "USER@example.com",
            "User",
            InferenceClassificationType.Other,
        )
        overrides_builder = _wire_list(mock_client, [existing])

        item_builder = MagicMock()
        item_builder.patch = AsyncMock(return_value=existing)
        overrides_builder.by_inference_classification_override_id = MagicMock(
            return_value=item_builder
        )

        result = await set_inbox_override(
            mock_client,
            sender_email="user@example.com",
            classify_as="focused",
            config=_CFG,
        )

        overrides_builder.by_inference_classification_override_id.assert_called_once_with("EXIST=")
        item_builder.patch.assert_called_once()
        body = item_builder.patch.call_args[0][0]
        assert body.classify_as == InferenceClassificationType.Focused
        overrides_builder.post.assert_not_called()

        assert result == {
            "status": "updated",
            "id": "EXIST=",
            "sender_email": "user@example.com",
            "classify_as": "focused",
        }

    async def test_invalid_email_raises(self):
        """Invalid sender_email raises ValueError."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="Invalid email"):
            await set_inbox_override(
                mock_client,
                sender_email="not-an-email",
                classify_as="focused",
                config=_CFG,
            )

    async def test_invalid_classify_as_raises(self):
        """Invalid classify_as raises ValueError."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="classify_as"):
            await set_inbox_override(
                mock_client,
                sender_email="user@example.com",
                classify_as="bogus",
                config=_CFG,
            )

    async def test_raises_read_only(self):
        """Raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await set_inbox_override(
                mock_client,
                sender_email="user@example.com",
                classify_as="focused",
                config=_CFG_RO,
            )

    async def test_raises_permission_denied(self):
        """Raises PermissionDeniedError when mail_triage not in allow_categories."""
        mock_client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await set_inbox_override(
                mock_client,
                sender_email="user@example.com",
                classify_as="focused",
                config=_CFG_ALLOW_OTHER,
            )

    async def test_succeeds_with_mail_triage_allowed(self):
        """Succeeds when 'mail_triage' is in allow_categories."""
        mock_client = MagicMock()
        overrides_builder = _wire_list(mock_client, [])
        created = _make_override(
            "NEW=", "user@example.com", "", InferenceClassificationType.Focused
        )
        overrides_builder.post = AsyncMock(return_value=created)

        result = await set_inbox_override(
            mock_client,
            sender_email="user@example.com",
            classify_as="focused",
            config=_CFG_ALLOW_TRIAGE,
        )
        assert result["status"] == "created"

    async def test_succeeds_with_empty_allow_categories(self):
        """Succeeds when allow_categories is empty (fully open)."""
        mock_client = MagicMock()
        overrides_builder = _wire_list(mock_client, [])
        created = _make_override(
            "NEW=", "user@example.com", "", InferenceClassificationType.Focused
        )
        overrides_builder.post = AsyncMock(return_value=created)

        result = await set_inbox_override(
            mock_client,
            sender_email="user@example.com",
            classify_as="focused",
            config=_CFG,
        )
        assert result["status"] == "created"

    async def test_classify_as_other_maps_to_enum(self):
        """classify_as='other' maps to InferenceClassificationType.Other on POST."""
        mock_client = MagicMock()
        overrides_builder = _wire_list(mock_client, [])
        created = _make_override("NEW=", "user@example.com", "", InferenceClassificationType.Other)
        overrides_builder.post = AsyncMock(return_value=created)

        result = await set_inbox_override(
            mock_client,
            sender_email="user@example.com",
            classify_as="other",
            config=_CFG,
        )

        body = overrides_builder.post.call_args[0][0]
        assert body.classify_as == InferenceClassificationType.Other
        assert result["classify_as"] == "other"


class TestDeleteInboxOverride:
    async def test_delete(self):
        """Calls .delete() on the override item builder."""
        mock_client = MagicMock()
        item_builder = MagicMock()
        item_builder.delete = AsyncMock()
        ic = mock_client.me.inference_classification
        ic.overrides.by_inference_classification_override_id.return_value = item_builder

        result = await delete_inbox_override(mock_client, override_id="ABC=", config=_CFG)
        assert result == {"status": "deleted", "id": "ABC="}
        item_builder.delete.assert_called_once()

    async def test_invalid_id_raises(self):
        """Invalid override_id raises ValueError from validate_graph_id."""
        mock_client = MagicMock()
        with pytest.raises(ValueError, match="invalid characters"):
            await delete_inbox_override(mock_client, override_id="<script>", config=_CFG)

    async def test_raises_read_only(self):
        """Raises ReadOnlyError in read-only mode."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_inbox_override(mock_client, override_id="ABC=", config=_CFG_RO)

    async def test_raises_permission_denied(self):
        """Raises PermissionDeniedError when mail_triage not in allow_categories."""
        mock_client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await delete_inbox_override(mock_client, override_id="ABC=", config=_CFG_ALLOW_OTHER)

    async def test_succeeds_with_mail_triage_allowed(self):
        """Succeeds when 'mail_triage' is in allow_categories."""
        mock_client = MagicMock()
        item_builder = MagicMock()
        item_builder.delete = AsyncMock()
        ic = mock_client.me.inference_classification
        ic.overrides.by_inference_classification_override_id.return_value = item_builder

        result = await delete_inbox_override(
            mock_client, override_id="ABC=", config=_CFG_ALLOW_TRIAGE
        )
        assert result["status"] == "deleted"
