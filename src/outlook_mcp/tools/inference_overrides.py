"""Focused Inbox per-sender override CRUD tools."""

from __future__ import annotations

from typing import Any

from outlook_mcp.config import Config
from outlook_mcp.permissions import CATEGORY_MAIL_TRIAGE, check_permission
from outlook_mcp.validation import sanitize_output, validate_email, validate_graph_id

_VALID_CLASSIFY = ("focused", "other")


def _classify_pairs():
    """Single source of truth for (string, enum) classify_as pairs."""
    from msgraph.generated.models.inference_classification_type import (
        InferenceClassificationType,
    )

    return (
        ("focused", InferenceClassificationType.Focused),
        ("other", InferenceClassificationType.Other),
    )


def _enum_to_str(enum_value: Any) -> str:
    """Map an InferenceClassificationType enum back to 'focused' / 'other'."""
    name = getattr(enum_value, "name", None)
    by_name = {enum.name: s for s, enum in _classify_pairs()}
    return by_name.get(name, "other")


async def list_inbox_overrides(graph_client: Any) -> dict:
    """List Focused Inbox per-sender override rules.

    GET /me/inferenceClassification/overrides. Returns the full list — no
    pagination cursor (override counts are small per user). Names and email
    addresses are passed through `sanitize_output` to scrub control chars.
    """
    response = await graph_client.me.inference_classification.overrides.get()
    values = getattr(response, "value", None) or []

    overrides = []
    for ov in values:
        email = getattr(ov, "sender_email_address", None)
        address = getattr(email, "address", None) or "" if email else ""
        name = getattr(email, "name", None) or "" if email else ""
        overrides.append(
            {
                "id": getattr(ov, "id", "") or "",
                "sender_email": sanitize_output(address),
                "sender_name": sanitize_output(name),
                "classify_as": _enum_to_str(getattr(ov, "classify_as", None)),
            }
        )

    return {"overrides": overrides, "count": len(overrides)}


async def set_inbox_override(
    graph_client: Any,
    sender_email: str,
    classify_as: str,
    *,
    config: Config,
) -> dict:
    """Upsert a Focused Inbox override for a sender.

    Graph rejects a POST when an override for that sender already exists, so
    this lists existing overrides first and PATCHes when a case-insensitive
    sender match is found.

    Race: a concurrent delete between list and PATCH surfaces as a 404 from
    Graph; caller can retry.
    """
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_set_inbox_override")

    sender_email = validate_email(sender_email)
    if classify_as not in _VALID_CLASSIFY:
        raise ValueError(f"classify_as must be one of {_VALID_CLASSIFY!r}, got {classify_as!r}")

    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.inference_classification_override import (
        InferenceClassificationOverride,
    )

    classify_map = dict(_classify_pairs())
    target_enum = classify_map[classify_as]

    overrides_builder = graph_client.me.inference_classification.overrides
    response = await overrides_builder.get()
    existing_values = getattr(response, "value", None) or []
    needle = sender_email.lower()

    existing = None
    # Graph contract: at most one override per sender; first match wins.
    for ov in existing_values:
        email = getattr(ov, "sender_email_address", None)
        address = getattr(email, "address", None) if email else None
        if address and address.lower() == needle:
            existing = ov
            break

    if existing is not None:
        patch_body = InferenceClassificationOverride()
        patch_body.classify_as = target_enum
        item_builder = overrides_builder.by_inference_classification_override_id(existing.id)
        await item_builder.patch(patch_body)
        return {
            "status": "updated",
            "id": existing.id,
            "sender_email": sender_email,
            "classify_as": classify_as,
        }

    body = InferenceClassificationOverride()
    body.classify_as = target_enum
    body.sender_email_address = EmailAddress(address=sender_email)

    created = await overrides_builder.post(body)
    new_id = getattr(created, "id", "") or ""
    return {
        "status": "created",
        "id": new_id,
        "sender_email": sender_email,
        "classify_as": classify_as,
    }


async def delete_inbox_override(
    graph_client: Any,
    override_id: str,
    *,
    config: Config,
) -> dict:
    """Delete a Focused Inbox override by ID."""
    check_permission(config, CATEGORY_MAIL_TRIAGE, "outlook_delete_inbox_override")
    override_id = validate_graph_id(override_id)

    overrides_builder = graph_client.me.inference_classification.overrides
    item_builder = overrides_builder.by_inference_classification_override_id(override_id)
    await item_builder.delete()
    return {"status": "deleted", "id": override_id}
