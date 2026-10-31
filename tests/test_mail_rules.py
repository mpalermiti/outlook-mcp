"""Tests for inbox rule CRUD tools (``/me/mailFolders/inbox/messageRules``)."""

from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import GraphAPIError, PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools import mail_rules
from outlook_mcp.tools.mail_rules import (
    _describe_actions,
    _describe_predicates,
    _forwards_offsite,
    _recipient,
    _stops_processing,
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
    exceptions=None,
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
    rule.exceptions = exceptions
    return rule


def _wire_rules(mock_client):
    """Return the ``message_rules`` collection builder for a mock client.

    ``post`` is wired here because ``create_rule`` sends through
    ``throttle.call_without_replay``, which always awaits the call.
    """
    collection = mock_client.me.mail_folders.by_mail_folder_id.return_value.message_rules
    collection.post = AsyncMock(return_value=_make_rule(rule_id="created-1"))
    collection.get = AsyncMock(return_value=MagicMock(value=[]))
    return collection


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

        assert await list_rules(client) == {"count": 0, "complete": True, "rules": []}

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

    async def test_rename_reads_the_rule_but_does_not_patch_actions(self):
        """It reads now, because a rename can be a repair — see TestRepairingAnErroredForwarder."""
        client = MagicMock()
        builder = _wire_single_rule(client, current=_make_rule(), patched=_make_rule())

        await update_rule(client, "rule-1", config=_CFG, name="Renamed")

        payload = builder.patch.await_args.args[0]
        assert payload.display_name == "Renamed"
        assert payload.actions is None
        assert builder.get.await_count == 1

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
            # Yield, so that without the lock both reads really do happen before
            # either write. A coroutine that returns without suspending cannot
            # interleave, and the test then passes with the lock deleted.
            await asyncio.sleep(0)
            return _current()

        async def _patch(payload):
            await asyncio.sleep(0)
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

    async def test_moving_a_stop_rule_behind_a_forwarder_needs_mail_send(self):
        """Reordering the stop rule does it too, without disabling anything.

        The forwarder keeps its sequence and its enabled state; the rule that was
        hiding it moves down the list, and starts delivering mail.
        """
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, sequence=10)

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


class TestCreateIsNotRetried:
    """A create is not idempotent: a second one is a second rule.

    kiota re-sends a request answered 503/504, POST included. With a forwarding
    rule that means up to four creates, and deleting "the" rule leaves the copies
    forwarding mail nobody is watching.
    """

    async def test_a_503_is_reported_as_outcome_unknown_and_not_resent(self):
        from kiota_abstractions.api_error import APIError

        client = MagicMock()
        collection = _wire_rules(client)
        failed = APIError()
        failed.response_status_code = 503
        collection.post = AsyncMock(side_effect=failed)

        with pytest.raises(GraphAPIError) as excinfo:
            await create_rule(
                client,
                "Forward everything",
                config=_CFG,
                subject_contains=["invoice"],
                forward_to=["outside@example.com"],
            )

        assert collection.post.await_count == 1
        assert "may already exist" in (excinfo.value.action or "")


class TestOrderingGateIsConservative:
    """The stop-rule masking gate must assume the forwarder might run after it."""

    @staticmethod
    def _stop_rule(sequence=5):
        return _make_rule(
            rule_id="stop-1",
            sequence=sequence,
            actions=_make_actions(stop_processing_rules=True),
        )

    @staticmethod
    def _forwarder(rule_id="fwd-1", sequence=5):
        return _make_rule(
            rule_id=rule_id,
            sequence=sequence,
            actions=_make_actions(forward_to=[_recipient("out@example.com")]),
        )

    async def test_a_tied_sequence_counts_as_hidden(self):
        """Exchange orders equal sequences ambiguously: assume the worse case."""
        client = MagicMock()
        current = self._stop_rule(sequence=5)
        builder = _wire_single_rule(client, current=current, patched=current)
        collection = _wire_listing(_wire_rules(client), [])
        collection.get = AsyncMock(return_value=MagicMock(value=[self._forwarder(sequence=5)]))

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "stop-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        builder.patch.assert_not_awaited()

    async def test_an_unknown_sequence_counts_as_hidden(self):
        """A sequence the endpoint omitted is not zero: the order is unknown."""
        client = MagicMock()
        current = self._stop_rule(sequence=5)
        builder = _wire_single_rule(client, current=current, patched=current)
        collection = _wire_rules(client)
        collection.get = AsyncMock(return_value=MagicMock(value=[self._forwarder(sequence=None)]))

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "stop-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        builder.patch.assert_not_awaited()


class TestCreateRuleSequence:
    async def test_a_null_sequence_in_the_listing_does_not_crash(self):
        """`max()` over None raised TypeError and the model saw a bare tool error."""
        client = MagicMock()
        collection = _wire_rules(client)
        _wire_listing(collection, [None, 3])

        result = await create_rule(
            client,
            config=_CFG_ALLOW_RULES,
            name="Archive",
            from_addresses=["a@example.com"],
            move_to_folder=_NEW_FOLDER_ID,
        )

        assert result["status"] == "created"
        # The sequence that goes on the wire, not the one the mock echoes back.
        posted = collection.post.await_args.args[0]
        assert posted.sequence == 100


class TestEmptyUpdateIsRefused:
    async def test_an_update_with_no_fields_does_not_patch(self):
        """An empty PATCH sent `{}` without reading the rule, so no gate ran."""
        client = MagicMock()
        current = _make_rule(actions=_make_actions(forward_to=[_recipient("out@example.com")]))
        current.has_error = True
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(mail_rules.ToolInputError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.get.assert_not_awaited()
        builder.patch.assert_not_awaited()


class TestRepairingAnErroredForwarder:
    """Exchange skips a rule marked `has_error`, so a repair to it is a start."""

    @staticmethod
    def _errored_forwarder():
        rule = _make_rule(actions=_make_actions(forward_to=[_recipient("out-3f@example.com")]))
        rule.has_error = True
        return rule

    async def test_repairing_an_errored_forwarder_needs_mail_send(self):
        client = MagicMock()
        current = self._errored_forwarder()
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(
                client, "rule-1", config=_CFG_ALLOW_RULES, move_to_folder=_NEW_FOLDER_ID
            )

        builder.patch.assert_not_awaited()

    async def test_repairing_an_errored_rule_that_does_not_forward_is_allowed(self):
        client = MagicMock()
        current = _make_rule(actions=_make_actions(mark_as_read=False))
        current.has_error = True
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, mark_as_read=True)

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_a_mark_as_read_repair_needs_mail_send(self):
        """The repair gate is about the write, not about which field it touches."""
        client = MagicMock()
        current = self._errored_forwarder()
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, mark_as_read=True)

        builder.patch.assert_not_awaited()

    async def test_renaming_an_errored_forwarder_needs_mail_send(self):
        """A rename is a write to the rule, and it puts the rule back into service."""
        client = MagicMock()
        current = self._errored_forwarder()
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, name="Mailbox cleanup")

        builder.patch.assert_not_awaited()

    async def test_disabling_an_errored_forwarder_while_repairing_is_allowed(self):
        """The rule is being switched off, so the repaired one cannot send."""
        client = MagicMock()
        current = self._errored_forwarder()
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(
            client,
            "rule-1",
            config=_CFG_ALLOW_RULES,
            move_to_folder=_NEW_FOLDER_ID,
            is_enabled=False,
        )

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_repairing_and_enabling_a_disabled_forwarder_needs_mail_send(self):
        """A repair that also switches the rule on is a start, whatever it was."""
        client = MagicMock()
        current = _make_rule(
            is_enabled=False, actions=_make_actions(forward_to=[_recipient("a@example.com")])
        )
        current.has_error = True
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(
                client,
                "rule-1",
                config=_CFG_ALLOW_RULES,
                move_to_folder=_NEW_FOLDER_ID,
                is_enabled=True,
            )

        builder.patch.assert_not_awaited()

    async def test_renaming_a_disabled_forwarder_is_not_a_repair(self):
        """A repair does not switch a disabled rule on, so it is not a start."""
        client = MagicMock()
        current = _make_rule(
            is_enabled=False, actions=_make_actions(forward_to=[_recipient("a@example.com")])
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(
            client, "rule-1", config=_CFG_ALLOW_RULES, move_to_folder=_NEW_FOLDER_ID
        )

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_renaming_a_healthy_forwarder_is_not_a_repair(self):
        """A healthy rule has nothing to repair, so a rename is not a start.

        Dropping `_is_broken(current)` from the repair gate makes every enabled
        forwarder's rename demand mail_send; this is the test that notices.
        """
        client = MagicMock()
        current = _make_rule(actions=_make_actions(forward_to=[_recipient("a@example.com")]))
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(
            client, "rule-1", config=_CFG_ALLOW_RULES, name="Mailbox cleanup"
        )

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_an_update_with_a_cut_short_list_needs_mail_send(self):
        """The rule after this one cannot be ruled out, so it is not ruled out.

        Dropping `not complete or` from the update gate makes a cut-short walk
        read as "nothing hidden"; this is the test that notices.
        """
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        collection = _wire_rules(client)
        page = MagicMock(value=[stop])
        page.odata_next_link = (
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules?$skiptoken=x"
        )
        collection.get = AsyncMock(return_value=page)
        collection.with_url.return_value.get = AsyncMock(return_value=page)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        builder.patch.assert_not_awaited()


class TestPagingIsAGateConcern:
    """The gates read the rule list, so a second page is a gate, not a display."""

    @staticmethod
    def _two_pages(client, first, second):
        collection = _wire_rules(client)
        page_one = MagicMock(value=[first])
        page_one.odata_next_link = (
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules?$skiptoken=x"
        )
        page_two = MagicMock(value=[second])
        page_two.odata_next_link = None
        collection.get = AsyncMock(return_value=page_one)
        collection.with_url.return_value.get = AsyncMock(return_value=page_two)
        return collection

    async def test_a_forwarder_on_the_second_page_still_gates_a_delete(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=1, actions=_make_actions(stop_processing_rules=True)
        )
        forwarder = _make_rule(
            rule_id="rule-2",
            sequence=5,
            actions=_make_actions(forward_to=[_recipient("out-page2@example.com")]),
        )
        collection = self._two_pages(client, stop, forwarder)
        builder = collection.by_message_rule_id.return_value
        builder.get = AsyncMock(return_value=stop)
        builder.delete = AsyncMock(return_value=None)

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_list_rules_returns_every_page(self):
        client = MagicMock()
        first = _make_rule(rule_id="a", sequence=1)
        second = _make_rule(rule_id="b", sequence=2)
        self._two_pages(client, first, second)

        result = await list_rules(client)

        assert result["count"] == 2
        assert [r["id"] for r in result["rules"]] == ["a", "b"]


class TestUnknownActionKeys:
    """An action this build cannot name is the one an audit must not miss."""

    @staticmethod
    def _rule_with_unknown_action(**rule_kwargs):
        actions = _make_actions()
        actions.additional_data = {"forwardToAThirdParty": True}
        return _make_rule(actions=actions, **rule_kwargs)

    async def test_an_unknown_action_is_shown_in_the_listing(self):
        client = MagicMock()
        rule = self._rule_with_unknown_action()
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[rule]))

        listed = (await list_rules(client))["rules"][0]

        assert listed["actions"]["unknown_actions"] == {"forwardToAThirdParty": True}

    async def test_an_unknown_action_counts_as_possibly_sending(self):
        client = MagicMock()
        current = self._rule_with_unknown_action(is_enabled=False)
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=True)

        builder.patch.assert_not_awaited()


class TestMaskingHearsErroredRules:
    """An errored rule may still run, so it may still hide or forward.

    Whether Exchange skips a rule carrying `hasError` is not established, so the
    masking scan assumes the worse case in its own direction: an errored stop
    rule may still stop, and an errored forwarder may still forward. Both were
    excused once, which let a delete with only `mail_rules` unmask a forwarder.
    """

    @staticmethod
    def _stop(rule_id="rule-1", sequence=3, enabled=True, has_error=False):
        rule = _make_rule(
            rule_id=rule_id,
            sequence=sequence,
            is_enabled=enabled,
            actions=_make_actions(stop_processing_rules=True),
        )
        rule.has_error = has_error
        return rule

    @staticmethod
    def _forwarder(rule_id="rule-2", sequence=4, enabled=True, has_error=False):
        rule = _make_rule(
            rule_id=rule_id,
            sequence=sequence,
            is_enabled=enabled,
            actions=_make_actions(forward_to=[_recipient("wire-out-4c7@example.com")]),
        )
        rule.has_error = has_error
        return rule

    async def test_deleting_an_errored_stop_rule_ahead_of_a_forwarder_needs_mail_send(self):
        client = MagicMock()
        stop = self._stop(has_error=True)
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_an_errored_forwarder_behind_the_stop_rule_still_gates(self):
        client = MagicMock()
        stop = self._stop()
        builder = _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(
            return_value=MagicMock(value=[stop, self._forwarder(has_error=True)])
        )

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_a_disabled_and_errored_stop_rule_still_hides_nothing(self):
        """The exemption is about being switched off, not about being broken."""
        client = MagicMock()
        stop = self._stop(enabled=False, has_error=True)
        _wire_single_rule(client, current=stop, patched=stop)
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, self._forwarder()]))

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        assert result["status"] == "deleted"


class TestPermanentDeleteStopsProcessing:
    """A permanently deleted message is not handed to the rules after it.

    Same gate as `stop_processing_rules`: deleting the rule that deletes the
    message lets a later forwarding rule see mail it never saw.
    """

    @staticmethod
    def _deleter(rule_id="rule-1", sequence=3):
        return _make_rule(
            rule_id=rule_id, sequence=sequence, actions=_make_actions(permanent_delete=True)
        )

    @staticmethod
    def _forwarder(rule_id="rule-2", sequence=4):
        return _make_rule(
            rule_id=rule_id,
            sequence=sequence,
            actions=_make_actions(forward_to=[_recipient("wire-out-4c7@example.com")]),
        )

    async def test_deleting_a_permanent_delete_rule_ahead_of_a_forwarder_needs_mail_send(self):
        client = MagicMock()
        deleter = self._deleter()
        builder = _wire_single_rule(client, current=deleter, patched=deleter)
        _wire_rules(client).get = AsyncMock(
            return_value=MagicMock(value=[deleter, self._forwarder()])
        )

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_disabling_a_permanent_delete_rule_needs_mail_send_too(self):
        client = MagicMock()
        deleter = self._deleter()
        builder = _wire_single_rule(client, current=deleter, patched=deleter)
        _wire_rules(client).get = AsyncMock(
            return_value=MagicMock(value=[deleter, self._forwarder()])
        )

        with pytest.raises(PermissionDeniedError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, is_enabled=False)

        builder.patch.assert_not_awaited()

    async def test_a_permanent_delete_rule_with_nothing_behind_it_is_allowed(self):
        client = MagicMock()
        deleter = self._deleter()
        _wire_single_rule(client, current=deleter, patched=deleter)
        mover = _make_rule(
            rule_id="rule-3", sequence=4, actions=_make_actions(move_to_folder=_OLD_FOLDER_ID)
        )
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[deleter, mover]))

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        assert result["status"] == "deleted"


class TestPagingIsBounded:
    """A list that was cut short cannot be used to say "nothing is hidden"."""

    @staticmethod
    def _self_referential(client):
        collection = _wire_rules(client)
        page = MagicMock()
        page.value = []
        page.odata_next_link = "https://graph.microsoft.com/v1.0/me/messages?$skip=100"
        collection.get = AsyncMock(return_value=page)
        collection.with_url.return_value.get = AsyncMock(return_value=page)
        return collection

    async def test_a_link_that_points_back_at_itself_is_not_followed_forever(self):
        client = MagicMock()
        collection = self._self_referential(client)

        rules = await list_rules(client)

        assert rules["complete"] is False
        assert collection.with_url.call_count == 1

    async def test_the_page_cap_ends_the_walk(self):
        client = MagicMock()
        collection = _wire_rules(client)
        pages = []
        for i in range(30):
            page = MagicMock()
            page.value = []
            page.odata_next_link = f"https://graph.microsoft.com/v1.0/me/messages?$skip={i + 100}"
            pages.append(page)
        collection.get = AsyncMock(return_value=pages[0])
        collection.with_url.return_value.get = AsyncMock(side_effect=pages[1:])

        rules = await list_rules(client)

        assert rules["complete"] is False
        # The cap counts pages fetched, and the first page is already in hand:
        # 10 pages total means 9 followed links. Counting links let the walk
        # fetch one page more than the cap said.
        assert collection.with_url.call_count == mail_rules._MAX_RULE_PAGES - 1

    async def test_a_delete_with_a_cut_short_list_needs_mail_send(self):
        """The rule after this one cannot be ruled out, so it is not ruled out."""
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        builder = _wire_single_rule(client, current=stop, patched=stop)
        self._self_referential(client)

        with pytest.raises(PermissionDeniedError):
            await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        builder.delete.assert_not_awaited()

    async def test_a_complete_list_still_lets_a_harmless_delete_through(self):
        client = MagicMock()
        stop = _make_rule(
            rule_id="rule-1", sequence=3, actions=_make_actions(stop_processing_rules=True)
        )
        _wire_single_rule(client, current=stop, patched=stop)
        mover = _make_rule(
            rule_id="rule-3", sequence=4, actions=_make_actions(move_to_folder=_OLD_FOLDER_ID)
        )
        _wire_rules(client).get = AsyncMock(return_value=MagicMock(value=[stop, mover]))

        result = await delete_rule(client, "rule-1", config=_CFG_ALLOW_RULES)

        assert result["status"] == "deleted"


class TestUnknownActionValues:
    """An unknown action is treated by what it holds, not only by its name."""

    @staticmethod
    def _actions(**extra):
        actions = _make_actions()
        actions.additional_data = extra
        return actions

    def test_a_null_action_does_not_count_as_sending(self):
        assert _forwards_offsite(self._actions(markAsImportant=None)) is False

    def test_a_false_action_does_not_count_as_sending(self):
        assert _forwards_offsite(self._actions(flagForFollowUp=False)) is False

    def test_an_unknown_action_with_content_still_counts(self):
        assert _forwards_offsite(self._actions(somethingNew={"cc": ["a@example.com"]})) is True

    def test_a_null_unknown_action_is_listed_as_null_not_as_text(self):
        described = _describe_actions(self._actions(markAsImportant=None))

        assert "unknown_actions" not in described
        assert "None" not in str(described)


class TestDisplayNameIsSanitized:
    """Every caller-visible string is scrubbed, including the one echoed back."""

    async def test_create_scrubs_the_name_it_returns(self):
        client = MagicMock()
        created = _make_rule(rule_id="created-1", name="Bad\x00Name")
        _wire_rules(client).post = AsyncMock(return_value=created)

        result = await create_rule(
            client,
            config=_CFG,
            name="Bad\x00Name",
            from_addresses=["a@example.com"],
            mark_as_read=True,
            sequence=100,
        )

        assert "\x00" not in result["display_name"]

    async def test_update_scrubs_the_name_it_returns(self):
        client = MagicMock()
        current = _make_rule()
        patched = _make_rule(name="Bad\x00Name")
        _wire_single_rule(client, current=current, patched=patched)

        result = await update_rule(client, "rule-1", config=_CFG, name="Bad\x00Name")

        assert "\x00" not in result["display_name"]


class TestUnknownFieldsInBothDirections:
    """A field this build cannot name is not silently harmless."""

    @staticmethod
    def _actions(**extra):
        actions = _make_actions()
        actions.additional_data = extra
        return actions

    def test_an_unknown_action_counts_as_stopping(self):
        """The mirror of `_forwards_offsite`: unknown might halt processing."""
        assert _stops_processing(self._actions(somethingNew={"stop": True})) is True

    def test_an_empty_unknown_value_is_not_an_action(self):
        """`[]`, `{}` and `""` are not actions, in either direction."""
        for empty in ([], {}, ""):
            assert _forwards_offsite(self._actions(somethingNew=empty)) is False
            assert _stops_processing(self._actions(somethingNew=empty)) is False

    def test_an_unknown_action_key_is_scrubbed_like_a_value(self):
        described = _describe_actions(self._actions(**{"bad\x07key": {"cc": ["a@b.com"]}}))

        (key,) = described["unknown_actions"]
        # Compare the key itself: `str(dict)` reprs a control character as the
        # text "\x07", so a substring check never sees it and passes either way.
        assert key == "badkey"

    def test_unknown_condition_keys_are_listed(self):
        """The docs said nothing Graph adds later is hidden from the read."""
        from msgraph.generated.models.message_rule_predicates import MessageRulePredicates

        conditions = MessageRulePredicates()
        conditions.additional_data = {"somethingNew": ["wire-condition"]}

        described = _describe_predicates(conditions)

        assert described["unknown_conditions"] == {"somethingNew": ["wire-condition"]}


class TestRuleExceptionsAreListed:
    """`exceptions` hangs off the rule, beside `conditions`, not inside it.

    `MessageRulePredicates` has no `exceptions` attribute, so a nested read
    always found nothing and a rule that forwards everything except some mail
    listed as if it caught everything.
    """

    async def test_the_exceptions_block_is_listed(self):
        from msgraph.generated.models.message_rule_predicates import MessageRulePredicates

        exceptions = MessageRulePredicates()
        exceptions.subject_contains = ["wire-exception"]

        client = MagicMock()
        response = MagicMock()
        response.value = [_make_rule(exceptions=exceptions)]
        _wire_rules(client).get = AsyncMock(return_value=response)

        result = await list_rules(client)

        assert result["rules"][0]["exceptions"] == {"subject_contains": ["wire-exception"]}

    async def test_a_rule_without_exceptions_omits_the_key(self):
        """Unset is omitted, not an empty block — like every other field here."""
        client = MagicMock()
        response = MagicMock()
        response.value = [_make_rule()]
        _wire_rules(client).get = AsyncMock(return_value=response)

        result = await list_rules(client)

        assert "exceptions" not in result["rules"][0]

    async def test_the_exceptions_block_is_scrubbed_like_every_other_string(self):
        from msgraph.generated.models.message_rule_predicates import MessageRulePredicates

        exceptions = MessageRulePredicates()
        exceptions.subject_contains = ["wire\x00-exception"]

        client = MagicMock()
        response = MagicMock()
        response.value = [_make_rule(exceptions=exceptions)]
        _wire_rules(client).get = AsyncMock(return_value=response)

        result = await list_rules(client)

        assert result["rules"][0]["exceptions"]["subject_contains"] == ["wire-exception"]


class TestUpdateGuards:
    async def test_a_blank_rename_is_refused(self):
        """A blank name stripped to "" and Graph stored `displayName: ""`."""
        client = MagicMock()
        current = _make_rule()
        builder = _wire_single_rule(client, current=current, patched=current)

        with pytest.raises(mail_rules.ToolInputError):
            await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, name="   ")

        builder.patch.assert_not_awaited()

    async def test_disabling_and_reordering_in_one_call_is_not_a_start(self):
        """The rule ends up off, so moving it cannot deliver anything."""
        client = MagicMock()
        current = _make_rule(
            sequence=5, actions=_make_actions(forward_to=[_recipient("out@example.com")])
        )
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(
            client,
            "rule-1",
            config=_CFG_ALLOW_RULES,
            sequence=9,
            is_enabled=False,
        )

        assert result["status"] == "updated"
        builder.patch.assert_awaited()

    async def test_renaming_a_disabled_errored_forwarder_is_not_a_start(self):
        """`enabled_after`, not `_is_broken` alone: a disabled rule cannot send.

        The existing disabled-rename test uses a healthy rule, so it never
        reaches the `enabled_after` term.
        """
        client = MagicMock()
        current = _make_rule(
            is_enabled=False,
            actions=_make_actions(forward_to=[_recipient("out@example.com")]),
        )
        current.has_error = True
        builder = _wire_single_rule(client, current=current, patched=current)

        result = await update_rule(
            client, "rule-1", config=_CFG_ALLOW_RULES, name="Mailbox cleanup"
        )

        assert result["status"] == "updated"
        builder.patch.assert_awaited()


class TestDeleteRetry:
    """A retried delete meets the SDK's own 404, not a hand-made one.

    `builder.get()` and `builder.delete()` are plain SDK calls, so a 404 arrives
    as the SDK's `ODataError` — only `call_without_replay` converts errors into
    this package's `GraphAPIError`. These tests drive a real `GraphServiceClient`
    over a mock transport, which is the type a live mailbox produces.
    """

    class _Credential:
        def get_token(self, *args, **kwargs):
            from azure.core.credentials import AccessToken

            return AccessToken("synthetic-token", 9_999_999_999)

    RULE_ID = _OLD_FOLDER_ID

    RULE = {
        "id": RULE_ID,
        "displayName": "Move newsletters to Archive",
        "sequence": 3,
        "isEnabled": True,
        "conditions": {"subjectContains": ["wire"]},
        "actions": {"moveToFolder": _OLD_FOLDER_ID, "markAsRead": True},
    }

    @staticmethod
    async def _delete_with(handler):
        import httpx
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as _gra
        from msgraph_core import GraphClientFactory

        async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as raw:
            http_client = GraphClientFactory.create_with_default_middleware(
                options=_gra.options, client=raw
            )
            auth = AzureIdentityAuthenticationProvider(
                TestDeleteRetry._Credential(),
                scopes=["https://graph.microsoft.com/.default"],
            )
            client = GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client))
            return await delete_rule(client, TestDeleteRetry.RULE_ID, config=_CFG_ALLOW_RULES)

    async def test_a_404_on_the_delete_is_not_a_failure(self):
        """Graph has no delete receipt; a retry looks like an id that never was."""
        import httpx

        def handler(request):
            if request.method == "DELETE":
                return httpx.Response(
                    404,
                    json={"error": {"code": "ErrorItemNotFound", "message": "The rule is gone."}},
                )
            return httpx.Response(200, json=TestDeleteRetry.RULE)

        result = await self._delete_with(handler)

        assert result == {"status": "not_found", "id": TestDeleteRetry.RULE_ID}

    async def test_a_404_on_the_read_is_not_a_failure_either(self):
        """The read runs first, so a retry after a landed delete fails there."""
        import httpx

        def handler(request):
            return httpx.Response(
                404,
                json={"error": {"code": "ErrorItemNotFound", "message": "The rule is gone."}},
            )

        result = await self._delete_with(handler)

        assert result == {"status": "not_found", "id": TestDeleteRetry.RULE_ID}

    async def test_a_500_on_the_delete_still_raises(self):
        """Only 404 is the benign case; everything else is a real failure."""
        import httpx
        from kiota_abstractions.api_error import APIError

        def handler(request):
            if request.method == "DELETE":
                return httpx.Response(
                    500, json={"error": {"code": "ErrorInternalServerError", "message": "boom"}}
                )
            return httpx.Response(200, json=TestDeleteRetry.RULE)

        with pytest.raises(APIError) as excinfo:
            await self._delete_with(handler)

        assert excinfo.value.response_status_code == 500


class TestAnIgnoredSequenceIsReported:
    """Graph assigns rule order; a requested `sequence` may not be honoured.

    On a live personal mailbox both POST and PATCH stored the service's own
    value, so a sequence-only update could answer "updated" while the rule never
    moved. The response now carries the request beside the result.
    """

    async def test_an_update_that_does_not_move_reports_the_request(self):
        client = MagicMock()
        current = _make_rule(sequence=3, actions=_make_actions(mark_as_read=True))
        stored = _make_rule(sequence=3, actions=_make_actions(mark_as_read=True))
        _wire_single_rule(client, current=current, patched=stored)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, sequence=500)

        assert result["sequence"] == 3
        assert result["sequence_requested"] == 500

    async def test_an_update_that_moves_reports_nothing_extra(self):
        client = MagicMock()
        current = _make_rule(sequence=3, actions=_make_actions(mark_as_read=True))
        stored = _make_rule(sequence=500, actions=_make_actions(mark_as_read=True))
        _wire_single_rule(client, current=current, patched=stored)

        result = await update_rule(client, "rule-1", config=_CFG_ALLOW_RULES, sequence=500)

        assert result["sequence"] == 500
        assert "sequence_requested" not in result

    async def test_a_create_that_does_not_land_where_asked_reports_the_request(self):
        client = MagicMock()
        collection = _wire_rules(client)
        _wire_listing(collection, [1])
        created = MagicMock()
        created.id = "rule-new"
        created.display_name = "Archive"
        created.sequence = 128
        collection.post = AsyncMock(return_value=created)

        result = await create_rule(
            client,
            config=_CFG_ALLOW_RULES,
            name="Archive",
            from_addresses=["a@example.com"],
            move_to_folder=_NEW_FOLDER_ID,
            sequence=500,
        )

        assert result["sequence"] == 128
        assert result["sequence_requested"] == 500
