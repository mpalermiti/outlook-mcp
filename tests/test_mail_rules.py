"""Tests for inbox rule CRUD tools (``/me/mailFolders/inbox/messageRules``)."""

from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.mail_rules import (
    _recipient,
    create_rule,
    delete_rule,
    list_rules,
    update_rule,
)
from tests.test_write_payloads_reach_the_wire import assert_on_wire

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)
_CFG_ALLOW_RULES = Config(client_id="test", allow_categories=["mail_rules"])
_CFG_ALLOW_OTHER = Config(client_id="test", allow_categories=["mail_send"])

_OLD_FOLDER_ID = "AQMkADAwATM0MDAAMS1iMTYzLTExAGNlLTAwAi0wMAoALgAAAyK-old-folder-id-0000"
_NEW_FOLDER_ID = "AQMkADAwATM0MDAAMS1iMTYzLTExAGNlLTAwAi0wMAoALgAAAyK-new-folder-id-0000"


def _make_actions(**kwargs):
    from msgraph.generated.models.message_rule_actions import MessageRuleActions

    actions = MessageRuleActions()
    for key, value in kwargs.items():
        setattr(actions, key, value)
    return actions


def _make_rule(
    rule_id: str = "rule-1",
    name: str = "Move newsletters to Archive",
    sequence: int = 3,
    is_enabled: bool = True,
    actions=None,
    conditions=None,
):
    rule = MagicMock()
    rule.id = rule_id
    rule.display_name = name
    rule.sequence = sequence
    rule.is_enabled = is_enabled
    rule.is_read_only = False
    rule.has_error = False
    rule.actions = actions
    rule.conditions = conditions
    return rule


def _wire_rules(mock_client):
    """Return the ``message_rules`` collection builder for a mock client."""
    return mock_client.me.mail_folders.by_mail_folder_id.return_value.message_rules


def _wire_listing(collection, sequences):
    """Wire the collection's GET with rules carrying these sequence numbers."""
    listing = MagicMock()
    listing.value = [_make_rule(rule_id=f"r{i}", sequence=s) for i, s in enumerate(sequences)]
    collection.get = AsyncMock(return_value=listing)
    return collection


def _wire_single_rule(mock_client, *, current=None, patched=None):
    """Return the single-rule builder wired for get/patch/delete."""
    collection = _wire_rules(mock_client)
    builder = collection.by_message_rule_id.return_value
    builder.get = AsyncMock(return_value=current)
    builder.patch = AsyncMock(return_value=patched)
    builder.delete = AsyncMock(return_value=None)
    return builder


class TestListRules:
    async def test_empty(self):
        client = MagicMock()
        response = MagicMock()
        response.value = []
        _wire_rules(client).get = AsyncMock(return_value=response)

        assert await list_rules(client) == {"count": 0, "rules": []}

    async def test_sorted_by_sequence_and_mapped(self):
        client = MagicMock()
        response = MagicMock()
        response.value = [
            _make_rule(rule_id="b", name="Second", sequence=9),
            _make_rule(rule_id="a", name="First", sequence=2),
        ]
        _wire_rules(client).get = AsyncMock(return_value=response)

        result = await list_rules(client)

        assert result["count"] == 2
        assert [r["id"] for r in result["rules"]] == ["a", "b"]
        assert result["rules"][0] == {
            "id": "a",
            "display_name": "First",
            "is_enabled": True,
            "sequence": 2,
            "is_read_only": False,
            "has_error": False,
            "conditions": {},
            "actions": {},
        }

    async def test_names_are_sanitized(self):
        """A rule name is caller-visible text; control chars are scrubbed."""
        client = MagicMock()
        response = MagicMock()
        response.value = [_make_rule(name="spam\x00rule\x07")]
        _wire_rules(client).get = AsyncMock(return_value=response)

        result = await list_rules(client)

        assert "\x00" not in result["rules"][0]["display_name"]
        assert "\x07" not in result["rules"][0]["display_name"]


class TestCreateRule:
    async def test_requires_a_condition(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="condition"):
            await create_rule(client, "x", config=_CFG, move_to_folder="archive")

    async def test_requires_an_action(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="action"):
            await create_rule(client, "x", config=_CFG, from_addresses=["a@b.com"])

    async def test_empty_name_refused(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="name"):
            await create_rule(
                client, "   ", config=_CFG, from_addresses=["a@b.com"], mark_as_read=True
            )

    async def test_invalid_importance_refused(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="importance"):
            await create_rule(
                client,
                "x",
                config=_CFG,
                from_addresses=["a@b.com"],
                importance="urgent",
                mark_as_read=True,
            )

    async def test_read_only_refused(self):
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await create_rule(
                client, "x", config=_CFG_RO, from_addresses=["a@b.com"], mark_as_read=True
            )

    async def test_disallowed_category_refused(self):
        client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await create_rule(
                client,
                "x",
                config=_CFG_ALLOW_OTHER,
                from_addresses=["a@b.com"],
                mark_as_read=True,
            )

    async def test_allowed_category_passes(self):
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        created = _make_rule(rule_id="new", sequence=100)
        collection.post = AsyncMock(return_value=created)

        result = await create_rule(
            client,
            "Rule",
            config=_CFG_ALLOW_RULES,
            from_addresses=["a@b.com"],
            move_to_folder="archive",
        )

        assert result["status"] == "created"
        assert result["id"] == "new"

    async def test_sequence_appends_after_highest(self):
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [7, 120])
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new", sequence=121))

        await create_rule(
            client, "Rule", config=_CFG, from_addresses=["a@b.com"], mark_as_read=True
        )

        assert collection.post.await_args.args[0].sequence == 121

    async def test_sequence_never_below_100(self):
        """Graph treats low sequences as top-priority; a new rule never jumps the queue."""
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [7, 42])
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new", sequence=100))

        await create_rule(
            client, "Rule", config=_CFG, from_addresses=["a@b.com"], mark_as_read=True
        )

        assert collection.post.await_args.args[0].sequence == 100

    async def test_explicit_sequence_kept(self):
        client = MagicMock()
        collection = _wire_rules(client)
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new", sequence=5))

        result = await create_rule(
            client,
            "Rule",
            config=_CFG,
            from_addresses=["a@b.com"],
            mark_as_read=True,
            sequence=5,
        )

        assert collection.post.await_args.args[0].sequence == 5
        assert result["sequence"] == 5

    async def test_payload_reaches_the_wire(self):
        """The conditions and actions this call was given must all serialize."""
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new"))

        await create_rule(
            client,
            "wire-rule-9f3",
            config=_CFG,
            from_addresses=["wire-sender-9f3@example.com"],
            subject_contains=["wire-subject-9f3"],
            body_contains=["wire-body-9f3"],
            sender_contains=["wire-sender-9f3"],
            has_attachments=True,
            importance="high",
            move_to_folder="archive",
            copy_to_folder="deleteditems",
            mark_as_read=True,
            forward_to=["wire-forward-9f3@example.com"],
            assign_categories=["wire-category-9f3"],
            stop_processing_rules=True,
            sequence=42,
        )

        # Every write argument, not only the convenient ones: an argument that
        # exists on the model and never reaches the payload is invisible until
        # someone trusts the rule it failed to create.
        assert_on_wire(
            collection.post.await_args.args[0],
            "wire-rule-9f3",
            "wire-sender-9f3@example.com",
            "wire-subject-9f3",
            "wire-body-9f3",
            "wire-sender-9f3",
            "wire-forward-9f3@example.com",
            "wire-category-9f3",
            "hasAttachments",
            "markAsRead",
            "stopProcessingRules",
            # move and copy must be distinguishable: same value on both would
            # let one of them silently fail to serialize.
            "archive",
            "deleteditems",
            '"importance": "high"',
            '"sequence": 42',
        )


class TestUpdateRule:
    async def test_repoint_preserves_other_actions(self):
        """Graph replaces `actions` wholesale — the read-modify-write must not drop any."""
        client = MagicMock()
        current = _make_rule(
            actions=_make_actions(
                move_to_folder=_OLD_FOLDER_ID,
                mark_as_read=True,
                stop_processing_rules=True,
            )
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        await update_rule(client, "rule-1", config=_CFG, move_to_folder=_NEW_FOLDER_ID)

        patched_payload = builder.patch.await_args.args[0]
        assert patched_payload.actions.move_to_folder == _NEW_FOLDER_ID
        assert patched_payload.actions.mark_as_read is True
        assert patched_payload.actions.stop_processing_rules is True

    async def test_repoint_payload_reaches_the_wire(self):
        client = MagicMock()
        current = _make_rule(actions=_make_actions(mark_as_read=True, delete=False))
        builder = _wire_single_rule(client, current=current, patched=current)

        await update_rule(client, "rule-1", config=_CFG, move_to_folder=_NEW_FOLDER_ID)

        assert_on_wire(
            builder.patch.await_args.args[0],
            _NEW_FOLDER_ID,
            "markAsRead",
        )

    async def test_mark_as_read_only_preserves_move(self):
        client = MagicMock()
        current = _make_rule(actions=_make_actions(move_to_folder=_OLD_FOLDER_ID))
        builder = _wire_single_rule(client, current=current, patched=current)

        await update_rule(client, "rule-1", config=_CFG, mark_as_read=True)

        actions = builder.patch.await_args.args[0].actions
        assert actions.move_to_folder == _OLD_FOLDER_ID
        assert actions.mark_as_read is True

    async def test_rename_does_not_read_or_patch_actions(self):
        client = MagicMock()
        builder = _wire_single_rule(client, current=_make_rule(), patched=_make_rule())

        await update_rule(client, "rule-1", config=_CFG, name="Renamed")

        payload = builder.patch.await_args.args[0]
        assert payload.display_name == "Renamed"
        assert payload.actions is None
        builder.get.assert_not_awaited()

    async def test_disable(self):
        client = MagicMock()
        builder = _wire_single_rule(
            client, current=_make_rule(), patched=_make_rule(is_enabled=False)
        )

        result = await update_rule(client, "rule-1", config=_CFG, is_enabled=False)

        assert builder.patch.await_args.args[0].is_enabled is False
        assert result["status"] == "updated"

    async def test_bad_sequence_refused(self):
        client = MagicMock()
        _wire_single_rule(client, current=_make_rule(), patched=_make_rule())
        with pytest.raises(ValueError, match="sequence"):
            await update_rule(client, "rule-1", config=_CFG, sequence=0)

    async def test_read_only_refused(self):
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await update_rule(client, "rule-1", config=_CFG_RO, name="x")

    async def test_disallowed_category_refused(self):
        client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_OTHER, name="x")


class TestDeleteRule:
    async def test_deletes_by_id(self):
        client = MagicMock()
        builder = _wire_single_rule(client)

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_awaited_once()
        assert result == {"status": "deleted", "id": "rule-1"}

    async def test_read_only_refused(self):
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await delete_rule(client, "rule-1", config=_CFG_RO)

    async def test_disallowed_category_refused(self):
        client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_OTHER)

    async def test_malformed_id_refused(self):
        client = MagicMock()
        with pytest.raises(ValueError):
            await delete_rule(client, "not an id!!", config=_CFG)


class TestForwardingRuleNeedsMailSend:
    """A forwarding rule emails a third party indefinitely, so mail_send gates it.

    Every other write this server makes happens while an agent is connected and a
    person is nearby. A forwarding rule keeps delivering mail after everyone has
    disconnected — the usual way a compromised mailbox is drained — so it gets the
    same gate the calendar tools use for the part of a call that emails someone.
    """

    async def test_forward_refused_without_mail_send(self):
        client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await create_rule(
                client,
                "Forward to outside",
                config=_CFG_ALLOW_RULES,
                from_addresses=["a@b.com"],
                forward_to=["elsewhere@example.com"],
            )

    async def test_move_only_rule_still_allowed_without_mail_send(self):
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new"))

        result = await create_rule(
            client,
            "Move newsletters",
            config=_CFG_ALLOW_RULES,
            from_addresses=["a@b.com"],
            move_to_folder="archive",
        )

        assert result["status"] == "created"

    async def test_forward_allowed_when_mail_send_granted(self):
        config = Config(client_id="test", allow_categories=["mail_rules", "mail_send"])
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        collection.post = AsyncMock(return_value=_make_rule(rule_id="new"))

        result = await create_rule(
            client,
            "Forward to self",
            config=config,
            from_addresses=["a@b.com"],
            forward_to=["elsewhere@example.com"],
        )

        assert result["status"] == "created"


class TestCreateRuleSurface:
    async def test_permanent_delete_is_not_offered(self):
        """Hard delete of incoming mail is a takeover primitive; delete is recoverable."""
        import inspect

        assert "permanent_delete" not in inspect.signature(create_rule).parameters

    async def test_bad_sequence_is_a_tool_input_error(self):
        """Plain ValueError reaches the model as a bare 'Error executing tool'."""
        from outlook_mcp.errors import ToolInputError

        client = MagicMock()
        with pytest.raises(ToolInputError):
            await create_rule(
                client,
                "Rule",
                config=_CFG,
                from_addresses=["a@b.com"],
                mark_as_read=True,
                sequence=0,
            )

    async def test_bad_importance_is_a_tool_input_error(self):
        from outlook_mcp.errors import ToolInputError

        client = MagicMock()
        with pytest.raises(ToolInputError):
            await create_rule(
                client,
                "Rule",
                config=_CFG,
                from_addresses=["a@b.com"],
                importance="urgent",
                mark_as_read=True,
            )


class TestReEnablingAForwardingRule:
    """Enabling a disabled rule resumes what it does — including forwarding.

    `update_rule` cannot introduce `forward_to`, but `is_enabled=True` turns a
    rule back on without any action being supplied, so a forwarding rule made in
    the web UI (or disabled earlier) would resume mailing a third party with no
    gate. The check therefore looks at the rule, not at the arguments.
    """

    async def test_enabling_a_disabled_forwarding_rule_needs_mail_send(self):
        client = MagicMock()
        current = _make_rule(
            is_enabled=False,
            actions=_make_actions(forward_to=[MagicMock()]),
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=True)

        builder.patch.assert_not_awaited()

    async def test_enabling_a_move_only_rule_is_allowed(self):
        client = MagicMock()
        current = _make_rule(
            is_enabled=False,
            actions=_make_actions(move_to_folder=_OLD_FOLDER_ID),
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=True)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_disabling_a_forwarding_rule_is_not_gated(self):
        """Turning something off never needs the send gate."""
        client = MagicMock()
        current = _make_rule(
            is_enabled=True,
            actions=_make_actions(forward_to=[MagicMock()]),
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_redirect_and_forward_as_attachment_count_too(self):
        for attr in ("redirect_to", "forward_as_attachment_to"):
            client = MagicMock()
            current = _make_rule(is_enabled=False, actions=_make_actions(**{attr: [MagicMock()]}))
            _wire_single_rule(client, current=current, patched=current)

            with pytest.raises(PermissionDeniedError):
                await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=True)


class TestListRuleVisibility:
    """A rule whose behaviour is invisible cannot be audited."""

    async def test_list_returns_conditions_and_actions(self):
        from msgraph.generated.models.message_rule_actions import MessageRuleActions
        from msgraph.generated.models.message_rule_predicates import MessageRulePredicates
        from msgraph.generated.models.size_range import SizeRange

        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])

        predicates = MessageRulePredicates()
        predicates.sender_contains = ["wire-sender-7a1@example.com"]
        predicates.subject_contains = ["wire-subject-7a1"]
        predicates.has_attachments = True
        predicates.within_size_range = SizeRange()
        predicates.within_size_range.minimum_size = 1024
        predicates.within_size_range.maximum_size = 4096

        actions = MessageRuleActions()
        actions.move_to_folder = _OLD_FOLDER_ID
        actions.forward_to = [_recipient("wire-forward-7a1@example.com")]
        actions.assign_categories = ["wire-category-7a1"]
        actions.stop_processing_rules = True

        collection.get = AsyncMock(
            return_value=MagicMock(value=[_make_rule(conditions=predicates, actions=actions)])
        )

        result = await list_rules(client)
        rule = result["rules"][0]

        assert rule["actions"]["forward_to"] == ["wire-forward-7a1@example.com"]
        assert rule["actions"]["move_to_folder"] == _OLD_FOLDER_ID
        assert rule["actions"]["assign_categories"] == ["wire-category-7a1"]
        assert rule["actions"]["stop_processing_rules"] is True
        assert rule["conditions"]["sender_contains"] == ["wire-sender-7a1@example.com"]
        assert rule["conditions"]["has_attachments"] is True
        assert rule["conditions"]["within_size_range"] == {"min": 1024, "max": 4096}

    async def test_unset_conditions_and_actions_are_omitted(self):
        """No empty blocks: a rule reports only what it actually sets."""
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        collection.get = AsyncMock(
            return_value=MagicMock(
                value=[_make_rule(actions=_make_actions(move_to_folder=_OLD_FOLDER_ID))]
            )
        )

        rule = (await list_rules(client))["rules"][0]

        assert rule["actions"] == {"move_to_folder": _OLD_FOLDER_ID}
        assert rule["conditions"] == {}


class TestOverlappingUpdates:
    """Two updaters must not undo each other inside one server."""

    async def test_concurrent_field_updates_both_survive(self):
        client = MagicMock()
        state = {"actions": _make_actions(move_to_folder=_OLD_FOLDER_ID, mark_as_read=False)}

        def _current():
            rule = _make_rule(actions=state["actions"])
            return rule

        async def _get():
            return _current()

        async def _patch(payload):
            actions = payload.actions
            if actions is not None:
                state["actions"] = actions
            return _current()

        builder = _wire_single_rule(client, current=_current(), patched=_current())
        builder.get = AsyncMock(side_effect=_get)
        builder.patch = AsyncMock(side_effect=_patch)

        await asyncio.gather(
            update_rule(client, "rule-lock-1", config=_CFG, move_to_folder=_NEW_FOLDER_ID),
            update_rule(client, "rule-lock-1", config=_CFG, mark_as_read=True),
        )

        assert state["actions"].move_to_folder == _NEW_FOLDER_ID
        assert state["actions"].mark_as_read is True


class TestWriteArgumentsReachTheWire:
    """CLAUDE.md: each write argument must appear in the serialized payload."""

    async def test_forwarding_recipients_survive_a_repoint(self):
        """The copied actions must serialize, not just sit on the model."""
        client = MagicMock()
        current = _make_rule(
            actions=_make_actions(
                move_to_folder=_OLD_FOLDER_ID,
                forward_to=[_recipient("wire-keep-9d4@example.com")],
                assign_categories=["wire-keep-category-9d4"],
            )
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        await update_rule(client, "rule-1", config=_CFG, move_to_folder=_NEW_FOLDER_ID)

        payload = builder.patch.await_args.args[0]
        assert_on_wire(
            payload,
            "wire-keep-9d4@example.com",
            "wire-keep-category-9d4",
            _NEW_FOLDER_ID,
        )
        assert payload.actions.forward_to[0].email_address.address == "wire-keep-9d4@example.com"


class TestRunOrderCannotUnmaskForwarding:
    """Ordering starts forwarding without touching the rule that forwards.

    A forwarding rule sitting behind a `stop_processing_rules` rule never runs.
    Move the forwarding rule up, or move / disable / delete the stop rule, and
    it starts delivering mail — with no `forward_to` argument anywhere in the
    request that did it.
    """

    @staticmethod
    def _forwarder(rule_id="rule-2", sequence=4, enabled=True):
        return _make_rule(
            rule_id=rule_id,
            name="Forward everything",
            sequence=sequence,
            is_enabled=enabled,
            actions=_make_actions(forward_to=[_recipient("wire-out-4c7@example.com")]),
        )

    async def test_moving_a_forwarding_rule_needs_mail_send(self):
        client = MagicMock()
        current = _make_rule(
            sequence=9, actions=_make_actions(forward_to=[_recipient("a@example.com")])
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, sequence=1)

        builder.patch.assert_not_awaited()

    async def test_moving_a_forwarding_rule_is_allowed_with_mail_send(self):
        client = MagicMock()
        current = _make_rule(
            sequence=9, actions=_make_actions(forward_to=[_recipient("a@example.com")])
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(client, "rule-1", config=_CFG, sequence=1)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_disabling_the_stop_rule_that_hides_a_forwarder_needs_mail_send(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        builder.patch.assert_not_awaited()

    async def test_disabling_a_stop_rule_with_no_forwarder_behind_it_is_allowed(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        mover = _make_rule(
            rule_id="rule-3", sequence=4, actions=_make_actions(move_to_folder=_OLD_FOLDER_ID)
        )
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, mover]))

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_deleting_a_masking_stop_rule_needs_mail_send(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_deleting_a_plain_rule_is_not_gated(self):
        client = MagicMock()
        plain = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(move_to_folder=_OLD_FOLDER_ID)
        )
        builder = _wire_single_rule(client, current=plain, patched=plain)
        _wire_rules(client).get = AsyncMock(
            return_value=MagicMock(value=[plain, self._forwarder()])
        )

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        assert result == {"status": "deleted", "id": "rule-1"}
        builder.delete.assert_awaited_once()

    async def test_a_disabled_forwarder_behind_the_stop_rule_does_not_gate(self):
        """Nothing is masked if the forwarding rule is off anyway."""
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        _wire_single_rule(client, current=stop, patched=stop)
        off_forwarder = self._forwarder(enabled=False)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, off_forwarder]))

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        assert result["status"] == "updated"


class TestDisabledStopRuleDoesNotGate:
    """A rule that is switched off hides nothing, so touching it is cleanup."""

    @staticmethod
    def _forwarder(sequence=4):
        return _make_rule(
            rule_id="rule-2",
            name="Forward everything",
            sequence=sequence,
            actions=_make_actions(forward_to=[_recipient("wire-out-4c7@example.com")]),
        )

    async def test_deleting_an_already_disabled_stop_rule_is_allowed(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1",
            sequence=3,
            is_enabled=False,
            actions=_make_actions(stop_processing_rules=True),
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        assert result == {"status": "deleted", "id": "rule-1"}
        builder.delete.assert_awaited_once()

    async def test_disabling_an_already_disabled_stop_rule_is_allowed(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1",
            sequence=3,
            is_enabled=False,
            actions=_make_actions(stop_processing_rules=True),
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_an_enabled_stop_rule_still_gates(self):
        """The control: the same shape, switched on, is still a bypass."""
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)


class TestAuditSweepRegressions:
    """The same classes of defect Greptile found, checked across the feature."""

    async def test_size_range_reports_only_the_bound_it_has(self):
        """A half-open range must not surface as {"min": 10, "max": None}."""
        from msgraph.generated.models.message_rule_predicates import MessageRulePredicates
        from msgraph.generated.models.size_range import SizeRange

        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        predicates = MessageRulePredicates()
        predicates.within_size_range = SizeRange()
        predicates.within_size_range.minimum_size = 1024
        collection.get = AsyncMock(
            return_value=MagicMock(value=[_make_rule(conditions=predicates)])
        )

        rule = (await list_rules(client))["rules"][0]

        assert rule["conditions"]["within_size_range"] == {"min": 1024}

    async def test_unknown_action_fields_survive_an_update(self):
        """`_copy_actions` copies what the rule has, not a hand-kept list."""
        client = MagicMock()
        actions = _make_actions(move_to_folder=_OLD_FOLDER_ID)
        actions.additional_data = {"wireUnknownAction9d1": "keep-me"}
        current = _make_rule(actions=actions)
        builder = _wire_single_rule(client, current=current, patched=current)

        await update_rule(client, "rule-1", config=_CFG, move_to_folder=_NEW_FOLDER_ID)

        patched = builder.patch.await_args.args[0]
        assert patched.actions.additional_data == {"wireUnknownAction9d1": "keep-me"}
        assert patched.actions.move_to_folder == _NEW_FOLDER_ID

    async def test_reordering_a_disabled_forwarder_is_not_gated(self):
        """The mirror of the over-reach: a disabled rule is not delivering mail."""
        client = MagicMock()
        current = _make_rule(
            sequence=9,
            is_enabled=False,
            actions=_make_actions(forward_to=[_recipient("wire-off-2f8@example.com")]),
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, sequence=1)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_list_reports_a_rule_that_exchange_marked_as_broken(self):
        client = MagicMock()
        collection = _wire_listing(_wire_rules(client), [1])
        broken = _make_rule()
        broken.has_error = True
        collection.get = AsyncMock(return_value=MagicMock(value=[broken]))

        rule = (await list_rules(client))["rules"][0]

        assert rule["has_error"] is True
