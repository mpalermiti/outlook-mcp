"""Inbox rule tools: list, create, update, delete message rules.

Wraps the Graph ``/me/mailFolders/inbox/messageRules`` endpoint. Rules are
server-side: they run whether or not any Outlook client is open, and they
apply to incoming mail in the Inbox. Each rule pairs a set of *conditions*
(what an incoming message must match) with a set of *actions* (what to do
when it matches).

The Graph SDK ships typed builders for this endpoint, so the wrapper uses
them directly rather than hand-rolling OData.
"""

from __future__ import annotations

import asyncio
from typing import Any

from kiota_abstractions.api_error import APIError

from outlook_mcp.config import Config
from outlook_mcp.errors import RULE_MAY_HAVE_BEEN_CREATED, ToolInputError
from outlook_mcp.folder_resolver import resolve_folder_id
from outlook_mcp.permissions import CATEGORY_MAIL_RULES, check_permission, check_sends_mail
from outlook_mcp.throttle import call_without_replay
from outlook_mcp.validation import sanitize_output, validate_email, validate_graph_id


def _recipient(address: str) -> Any:
    """Build a Graph ``Recipient`` from an email address."""
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.recipient import Recipient

    ea = EmailAddress()
    ea.address = address
    r = Recipient()
    r.email_address = ea
    return r


def _recipients(addresses: list[str]) -> list[Any]:
    return [_recipient(a) for a in addresses]


def _copy_actions(existing: Any) -> Any:
    """Copy every action off an existing rule's ``actions`` block.

    A PATCH that carries an ``actions`` object replaces the rule's actions
    wholesale — Graph does not merge this collection — so any action absent
    from the payload is cleared. Re-pointing a rule while preserving its
    mark-as-read / forward / stop-processing behaviour therefore means reading
    the current actions and sending them all back. Dropping one is exactly the
    silent-no-op shape this project's offline guards exist to catch.
    """
    from msgraph.generated.models.message_rule_actions import MessageRuleActions

    actions = MessageRuleActions()
    if existing is None:
        return actions
    # Copy what the rule actually has, not a hand-kept list of names: a field
    # Graph adds later would otherwise be dropped from every update silently, and
    # the rule would lose it the first time anyone re-pointed it.
    #
    # Only what is set: assigning an attribute None marks it changed in the
    # backing store, and the request adapter then emits that null under its
    # *Python* name — the shape that made Graph answer 400 on a partial contact
    # address (see tests/test_write_payloads_reach_the_wire.py).
    for attr, value in vars(existing).items():
        if attr.startswith("_") or attr in {"backing_store", "odata_type"}:
            continue
        if attr == "additional_data":
            if value:
                actions.additional_data = value
            continue
        if value is not None:
            setattr(actions, attr, value)
    return actions


_FORWARDING_ACTIONS = ("forward_to", "redirect_to", "forward_as_attachment_to")

# One lock per rule id. `update_rule` is a read-modify-write: Graph replaces the
# `actions` object wholesale, so the payload is built from a read taken moments
# earlier. Two overlapping calls — a folder change and a mark-read change, the
# obvious pair — each read the old actions and each send a complete replacement,
# so the later request quietly restores the earlier caller's old value in the
# field it never mentioned. The lock is per process because that is the only
# sharing this server controls; see the docstring for the rest.
_rule_locks: dict[str, asyncio.Lock] = {}


def _rule_lock(rule_id: str) -> asyncio.Lock:
    lock = _rule_locks.get(rule_id)
    if lock is None:
        lock = asyncio.Lock()
        _rule_locks[rule_id] = lock
    return lock


def _counts_as_a_value(value: Any) -> bool:
    """Whether an unknown key's value is something, not an empty container.

    `additional_data` keeps whatever Graph sent for a key this build cannot
    name. `[]`, `{}` and `""` are how an unset field comes back, and counting
    them as an action would report a harmless key as possibly sending (and show
    it as one). Shared by the action and condition walks so the two cannot
    drift apart.
    """
    return value is not None and value is not False and value != [] and value != {} and value != ""


def _unknown_keys(extra: dict[str, Any]) -> list[str]:
    """Keys this build cannot name — whatever Graph kept in ``additional_data``."""
    return sorted(
        k
        for k, value in extra.items()
        if not str(k).startswith("@odata.") and _counts_as_a_value(value)
    )


def _unknown_action_keys(actions: Any) -> list[str]:
    """Action keys this build cannot name — whatever Graph kept in ``additional_data``.

    ``_copy_actions`` resends ``additional_data`` verbatim, so an action Graph
    adds after this release reaches the mailbox without appearing in
    ``list_rules`` or in the forwarding gate: a blind spot in both the audit and
    the check. Unknown is treated as possibly sending.
    """
    return _unknown_keys(getattr(actions, "additional_data", None) or {})


def _forwards_offsite(actions: Any) -> bool:
    """True if this rule's actions may send mail outside the mailbox.

    Used on re-enable: a disabled rule can be switched back on without any
    action being supplied, which resumes whatever it does — including forwarding
    to a third party — so the send gate has to look at what the rule *is*, not
    only at what the caller passed.

    An action this build does not recognise counts here. Refusing a known-harmless
    new action costs a permission check; missing a sending one costs the mailbox.
    """
    if actions is None:
        return False
    if any(getattr(actions, attr, None) for attr in _FORWARDING_ACTIONS):
        return True
    return bool(_unknown_action_keys(actions))


_BLOCK_INTERNALS = {"backing_store", "additional_data", "odata_type"}


def _plain(value: Any) -> Any:
    """Turn one Graph value into JSON-able output, scrubbing caller-visible text."""
    if isinstance(value, str):
        return sanitize_output(value)
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, list):
        return [_plain(v) for v in value]
    if hasattr(value, "email_address"):
        return sanitize_output(getattr(value.email_address, "address", None) or "")
    if isinstance(value, dict):
        return {k: _plain(v) for k, v in value.items()}
    return sanitize_output(str(value))


def _describe_block(block: Any) -> dict:
    """Serialize a rule's conditions or actions, omitting what it does not set.

    Driven by the model's own attributes rather than a hand-written list of field
    names: a rule you cannot read the behaviour of cannot be audited — "is
    anything forwarding mail out of this mailbox?" is the question that matters
    and the rule's name does not answer it. A fixed list has the same failure in
    the read direction that a fixed copy list has in the write direction: an
    attribute Graph adds later is silently invisible. Unset fields are omitted
    rather than emitted as null, and a size range reports only the bound it has.
    """
    if block is None:
        return {}
    described: dict[str, Any] = {}
    for attr, value in vars(block).items():
        if attr.startswith("_") or attr in _BLOCK_INTERNALS:
            continue
        if value is None or value == []:
            continue
        if attr == "within_size_range":
            bounds = {
                "min": getattr(value, "minimum_size", None),
                "max": getattr(value, "maximum_size", None),
            }
            described[attr] = {k: v for k, v in bounds.items() if v is not None}
            continue
        described[attr] = _plain(value)
    return described


def _unknown_condition_keys(conditions: Any) -> list[str]:
    """Condition keys this build cannot name — whatever Graph kept in ``additional_data``."""
    return _unknown_keys(getattr(conditions, "additional_data", None) or {})


def _describe_predicates(predicates: Any) -> dict:
    """Serialize a predicates block, including any key this build cannot name.

    Used for both a rule's ``conditions`` and its ``exceptions``: they are the
    same ``MessageRulePredicates`` type, one hanging off the rule directly and
    the other beside it. The read surface claimed a field Graph adds later is
    "neither hidden from the read nor dropped from an update". That held for
    actions only: an unknown condition key was dropped from the listing.
    """
    described = _describe_block(predicates)
    unknown = _unknown_condition_keys(predicates)
    if unknown:
        extra = predicates.additional_data
        described["unknown_conditions"] = {sanitize_output(k): _plain(extra[k]) for k in unknown}
    return described


def _describe_actions(actions: Any) -> dict:
    """Serialize a rule's actions, including any this build cannot name.

    Named rather than dropped: an action the audit cannot see is exactly the one
    that matters, and `_copy_actions` sends it back on every update. Key names are
    caller-visible too, so they get the same scrubbing as values.
    """
    described = _describe_block(actions)
    unknown = _unknown_action_keys(actions)
    if unknown:
        extra = actions.additional_data
        described["unknown_actions"] = {sanitize_output(k): _plain(extra[k]) for k in unknown}
    return described


def _address(recipient: Any) -> str:
    """The address out of a Graph ``Recipient``, scrubbed for display."""
    email = getattr(recipient, "email_address", None)
    return sanitize_output(getattr(email, "address", None) or "")


def _is_on(rule: Any) -> bool:
    """True if the rule is switched on, whatever Exchange makes of it.

    Deliberately not a statement about whether it runs. What MS-OXORULE says
    about a rule carrying ``hasError`` is not established, so each use assumes
    the worse case for its own direction: the masking scan counts an errored
    rule as possibly running, and the repair gate counts an errored forwarder
    as possibly not — see the two callers.
    """
    return rule is not None and bool(getattr(rule, "is_enabled", False))


def _is_broken(rule: Any) -> bool:
    """True if Exchange reports an error applying this rule."""
    return rule is not None and bool(getattr(rule, "has_error", False))


def _stops_processing(actions: Any) -> bool:
    """True if this rule can keep a later rule from seeing the message.

    Two actions do it, not one. ``stop_processing_rules`` halts evaluation
    outright; ``permanentDelete`` removes the message, and MS-OXORULE does not
    hand a permanently deleted message to the rules after it. Only the first was
    checked, so deleting a ``permanentDelete`` rule that sat ahead of a
    forwarder was allowed with only ``mail_rules``.
    """
    if actions is None:
        return False
    if getattr(actions, "stop_processing_rules", None):
        return True
    if getattr(actions, "permanent_delete", None):
        return True
    # An action this build cannot name may halt processing too. In this direction
    # the uncertainty is cheap: assuming it stops costs a permission check when
    # the rule is changed, while assuming it does not can unmask a forwarder
    # ordered behind it. Unknown therefore counts as stopping, the mirror of
    # `_forwards_offsite` counting it as sending.
    return bool(_unknown_action_keys(actions))


def _masks_a_forwarding_rule(current: Any, rules: Any) -> bool:
    """True if `current` halts processing ahead of an enabled forwarding rule.

    A forwarding rule placed after a `stop_processing_rules` rule never runs.
    Changing *that* stop rule — disabling it, reordering it later, deleting it —
    switches the forwarding on without touching the forwarding rule at all, so
    the send gate has to look at the rules the stop rule is hiding, not only at
    the one named in the request.
    """
    # A rule that is switched off stops nothing: nothing it was ordered ahead of
    # is being hidden, and cleaning one up should not demand mail_send. Errored
    # rules are NOT exempt: whether Exchange still evaluates a rule carrying
    # hasError is unconfirmed, and in this direction that uncertainty has to
    # assume it might, or a delete could unmask a forwarder unseen.
    if not _is_on(current):
        return False
    if not _stops_processing(getattr(current, "actions", None)):
        return False
    current_sequence = getattr(current, "sequence", None)
    for rule in rules or []:
        if getattr(rule, "id", None) == getattr(current, "id", None):
            continue
        if not _is_on(rule):
            continue
        if not _forwards_offsite(getattr(rule, "actions", None)):
            continue
        # A sequence the endpoint did not return is not sequence zero. The gate
        # used to read it as 0 while `outlook_list_rules` sorted it last, so the
        # model was told a forwarder runs at the end while the gate treated it
        # as running first. Unknown order means "cannot tell", which in this
        # direction has to assume the forwarder might be behind the stop rule.
        rule_sequence = getattr(rule, "sequence", None)
        if current_sequence is None or rule_sequence is None:
            return True
        # `>=`, not `>`: Exchange orders equal sequences ambiguously, so a
        # forwarder tied with the stop rule may well run after it. Treating a
        # tie as safe would let a stop rule be disabled while the tied
        # forwarder stayed enabled, unmasking it with no mail_send asked for.
        if rule_sequence >= current_sequence:
            return True
    return False


# A mailbox with more rules than this is not a mailbox the gates can reason
# about, and a link that keeps pointing at itself is not a page. Both end the
# walk with `complete=False`, which the gates treat as "cannot tell" rather than
# "nothing hidden": a mailbox nobody has seen all of cannot be pronounced safe.
_MAX_RULE_PAGES = 10


async def _fetch_rules(graph_client: Any) -> tuple[list[Any], bool]:
    """The inbox's rules across pages, and whether that list is the whole list.

    Paging is a gate concern, not a display one: the checks read this list, and
    with the answer split over two pages a stop rule on page one could be
    deleted while an enabled forwarder sat on page two, unmasking it with no
    ``mail_send`` asked for.

    ``complete`` is False when the walk stopped early — the page cap was reached,
    or a ``nextLink`` pointed back at a page already followed. A list that is not
    the whole list cannot answer "is anything hidden after this rule?".
    """
    builder = graph_client.me.mail_folders.by_mail_folder_id("inbox").message_rules
    response = await builder.get()
    collected = list(response.value) if response and response.value else []
    # `isinstance(..., str)`: an absent link has to be absent, not merely falsy —
    # a mock attribute is truthy, and following it turns one page into a hang.
    next_link = getattr(response, "odata_next_link", None)
    followed: set[str] = set()
    # The cap counts pages fetched, not distinct links: the first page is already
    # in hand, so `len(followed)` allowed one more page than the cap said. A
    # mailbox with more pages than this cannot be reasoned about, and the count
    # has to mean what it says or the gate's "cannot tell" is off by one page.
    pages_fetched = 1
    while isinstance(next_link, str) and next_link:
        if next_link in followed or pages_fetched >= _MAX_RULE_PAGES:
            return collected, False
        followed.add(next_link)
        response = await builder.with_url(next_link).get()
        pages_fetched += 1
        collected.extend(list(response.value) if response and response.value else [])
        next_link = getattr(response, "odata_next_link", None)
    return collected, True


async def _rules_in_inbox(graph_client: Any) -> tuple[list[Any], bool]:
    return await _fetch_rules(graph_client)


def _describe_rule(rule: Any) -> dict:
    """One rule as plain JSON, including the exceptions that narrow it.

    ``exceptions`` is a predicates block on the rule itself, beside
    ``conditions`` — ``MessageRulePredicates`` has no ``exceptions`` attribute,
    so a nested read finds nothing. Without it a rule that forwards everything
    except a given set of mail lists as if it caught everything, which is the
    one question a rule listing exists to answer. Omitted rather than emitted as
    an empty block, like every other unset field.
    """
    described = {
        "id": rule.id,
        "display_name": sanitize_output(rule.display_name or ""),
        "is_enabled": rule.is_enabled,
        "sequence": rule.sequence,
        "is_read_only": rule.is_read_only,
        "has_error": getattr(rule, "has_error", None),
        "conditions": _describe_predicates(getattr(rule, "conditions", None)),
        "actions": _describe_actions(getattr(rule, "actions", None)),
    }
    exceptions = _describe_predicates(getattr(rule, "exceptions", None))
    if exceptions:
        described["exceptions"] = exceptions
    return described


async def list_rules(graph_client: Any) -> dict:
    """List all inbox rules, sorted by execution order (sequence).

    Each rule carries its conditions, its exceptions and its actions as plain
    JSON, not only its name and status: the reason to list rules is to find out
    what they *do* — an unexpected forwarding rule is the thing you are looking
    for. Rule names and every caller-visible string inside those blocks are
    scrubbed of control characters the way every other caller-visible string is.

    ``complete`` is false when the listing stopped early (a page cap, or a link
    pointing back at a page already read): every rule Graph showed is here, but
    the list is not necessarily the whole inbox.
    """
    rules, complete = await _fetch_rules(graph_client)
    rules.sort(key=lambda r: r.sequence if r.sequence is not None else 9999)
    return {
        "count": len(rules),
        # False means the walk stopped early (page cap, or a link pointing back):
        # the list is what Graph showed, not necessarily every rule in the inbox.
        "complete": complete,
        "rules": [_describe_rule(r) for r in rules],
    }


async def create_rule(
    graph_client: Any,
    name: str,
    *,
    config: Config,
    from_addresses: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    sender_contains: list[str] | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    mark_as_read: bool | None = None,
    delete: bool | None = None,
    forward_to: list[str] | None = None,
    assign_categories: list[str] | None = None,
    stop_processing_rules: bool | None = None,
    is_enabled: bool = True,
    sequence: int | None = None,
) -> dict:
    """Create an inbox rule from conditions and actions.

    At least one condition and one action are required. ``move_to_folder`` /
    ``copy_to_folder`` accept display names, well-known names, or Graph IDs.
    ``importance`` is "low", "normal", or "high". ``sequence`` controls
    execution order (lower runs first); omitted, the rule is appended after
    the highest existing sequence.

    A rule is durable, not a one-off: it keeps running after this session ends
    and applies to every message that arrives later. ``forward_to`` is also an
    email to somebody else, so it additionally requires the ``mail_send``
    category. Hard delete is deliberately not offered here — ``delete`` moves
    matching mail to Deleted Items, which is recoverable.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_create_rule")
    if forward_to:
        # A forwarding rule emails a third party long after this session ended,
        # which is precisely what `mail_send` gates. Same ordering as the
        # calendar tools: the tool's own category first, then the send gate.
        check_sends_mail(config, "outlook_create_rule", "forward matching mail to another address")
    if not name or not name.strip():
        raise ToolInputError("Rule name is required.")

    from msgraph.generated.models.importance import Importance
    from msgraph.generated.models.message_rule import MessageRule
    from msgraph.generated.models.message_rule_actions import MessageRuleActions
    from msgraph.generated.models.message_rule_predicates import MessageRulePredicates

    # ── Conditions ─────────────────────────────────────────────
    conditions = MessageRulePredicates()
    has_condition = False
    if from_addresses:
        conditions.from_addresses = _recipients([validate_email(a) for a in from_addresses])
        has_condition = True
    if subject_contains:
        conditions.subject_contains = subject_contains
        has_condition = True
    if body_contains:
        conditions.body_contains = body_contains
        has_condition = True
    if sender_contains:
        conditions.sender_contains = sender_contains
        has_condition = True
    if has_attachments is not None:
        conditions.has_attachments = has_attachments
        has_condition = True
    if importance:
        valid = {"low", "normal", "high"}
        if importance not in valid:
            raise ToolInputError(f"importance must be one of {sorted(valid)}; got {importance}")
        conditions.importance = Importance(importance)
        has_condition = True

    # ── Actions ────────────────────────────────────────────────
    actions = MessageRuleActions()
    has_action = False
    if move_to_folder:
        actions.move_to_folder = await resolve_folder_id(graph_client, move_to_folder)
        has_action = True
    if copy_to_folder:
        actions.copy_to_folder = await resolve_folder_id(graph_client, copy_to_folder)
        has_action = True
    if mark_as_read is not None:
        actions.mark_as_read = mark_as_read
        has_action = True
    if delete is not None:
        actions.delete = delete
        has_action = True
    if forward_to:
        actions.forward_to = _recipients([validate_email(a) for a in forward_to])
        has_action = True
    if assign_categories:
        actions.assign_categories = assign_categories
        has_action = True
    if stop_processing_rules is not None:
        actions.stop_processing_rules = stop_processing_rules
        has_action = True

    if not has_condition:
        raise ToolInputError(
            "At least one condition is required: from_addresses, subject_contains, "
            "body_contains, sender_contains, has_attachments, or importance."
        )
    if not has_action:
        raise ToolInputError(
            "At least one action is required: move_to_folder, copy_to_folder, "
            "mark_as_read, delete, forward_to, assign_categories, or "
            "stop_processing_rules."
        )

    # ── Assemble ───────────────────────────────────────────────
    rule = MessageRule()
    rule.display_name = name.strip()
    rule.is_enabled = is_enabled
    rule.conditions = conditions
    rule.actions = actions

    if sequence is not None:
        if sequence < 1:
            raise ToolInputError("sequence must be a positive integer")
        rule.sequence = sequence
    else:
        # Append after the highest existing sequence (default 100 if none).
        existing = await list_rules(graph_client)
        # Filter the `None`s: a rule whose sequence the endpoint omitted made
        # `max()` compare None with int and raise TypeError, which reached the
        # model as "Error executing tool outlook_create_rule" and nothing else.
        sequences = [r["sequence"] for r in existing["rules"] if r.get("sequence") is not None]
        highest = max(sequences, default=0)
        rule.sequence = max(highest + 1, 100)

    # Not a plain POST: kiota re-sends a request answered 503/504, and a rule is
    # not idempotent — a second create is a second rule, and deleting "the" rule
    # afterwards leaves the copy forwarding. So: no replay, and an
    # outcome-unknown message that says to look before trying again.
    created = await call_without_replay(
        lambda once: graph_client.me.mail_folders.by_mail_folder_id("inbox").message_rules.post(
            rule, request_configuration=once
        ),
        outcome_unknown=RULE_MAY_HAVE_BEEN_CREATED,
    )
    response = {
        "status": "created",
        "id": created.id if created else None,
        "display_name": sanitize_output((created.display_name if created else None) or name),
        "sequence": created.sequence if created else rule.sequence,
    }
    # Graph assigns rule order and was observed ignoring a requested `sequence`
    # on a personal mailbox (POST and PATCH both stored the service's own value).
    # Returning what was asked for next to what landed means an ignored request
    # is visible instead of reading as a success.
    if sequence is not None and response["sequence"] != sequence:
        response["sequence_requested"] = sequence
    return response


async def update_rule(
    graph_client: Any,
    rule_id: str,
    *,
    config: Config,
    name: str | None = None,
    is_enabled: bool | None = None,
    sequence: int | None = None,
    move_to_folder: str | None = None,
    mark_as_read: bool | None = None,
) -> dict:
    """Update an existing inbox rule (partial patch — only provided fields change).

    ``move_to_folder`` accepts display names, well-known names, or Graph IDs.
    Pass is_enabled=False to disable a rule without deleting it. The rule's
    other actions survive: Graph replaces the ``actions`` object rather than
    merging it, so the current actions are read and sent back with the
    requested field overridden.

    Re-enabling (``is_enabled=True``) reads the rule first: switching a
    forwarding or redirecting rule back on requires the ``mail_send`` category,
    because it resumes mailing a third party for as long as the rule exists.

    Ordering is gated too, because it can switch forwarding on without any
    forwarding argument being supplied: a forwarding rule behind a
    ``stop_processing_rules`` rule never runs, so moving that forwarding rule
    earlier, or moving / disabling the stop rule, starts delivering mail. Any
    sequence change to a forwarding rule, and any change that unmasks a later
    forwarding rule, requires ``mail_send``.

    Overlapping updates of the same rule are serialized inside this server, so
    two calls in flight from one session cannot undo each other. Across
    processes the last write wins: Graph's ``messageRules`` endpoint returns no
    ``@odata.etag`` (checked against the live endpoint), so there is no
    ``If-Match`` to reject a stale write with.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_update_rule")
    rule_id = validate_graph_id(rule_id)

    from msgraph.generated.models.message_rule import MessageRule

    builder = graph_client.me.mail_folders.by_mail_folder_id(
        "inbox"
    ).message_rules.by_message_rule_id(rule_id)

    patch = MessageRule()
    if all(v is None for v in (name, is_enabled, sequence, move_to_folder, mark_as_read)):
        # An empty PATCH reached Graph as `{}` without reading the rule first,
        # so none of the gates above ran - a no-op call that still returned
        # "updated" while an errored forwarder sat unexamined. Refuse it: there
        # is nothing to change, and the read the caller skipped is the point.
        raise ToolInputError(
            "No fields to update: provide at least one of name, is_enabled, "
            "sequence, move_to_folder, or mark_as_read."
        )
    if name is not None:
        if not name.strip():
            # A blank name stripped to "", and Graph accepted it as
            # `displayName: ""` — a rule the user cannot identify afterwards.
            raise ToolInputError("A rule name cannot be blank.")
        patch.display_name = name.strip()
    if is_enabled is not None:
        patch.is_enabled = is_enabled
    if sequence is not None:
        if sequence < 1:
            raise ToolInputError("sequence must be a positive integer")
        patch.sequence = sequence

    # One read covers both reasons to look at the rule: re-pointing needs the
    # existing actions to send back, and re-enabling needs them to see whether
    # this rule emails somebody else.
    needs_current = (
        move_to_folder is not None
        or mark_as_read is not None
        or is_enabled is True
        # Reordering can switch forwarding on without touching the rule that
        # does it: moving a forwarding rule ahead of the stop rule that was
        # hiding it, or moving / disabling that stop rule back behind it.
        or sequence is not None
        or is_enabled is False
        # A rename reaches the repair gate too: it is a change to an errored
        # forwarding rule, which puts the rule back into service.
        or name is not None
    )

    async with _rule_lock(rule_id):
        current = await builder.get() if needs_current else None
        current_actions = current.actions if current else None

        if is_enabled is True and _forwards_offsite(current_actions):
            check_sends_mail(config, "outlook_update_rule", "re-enable a rule that forwards mail")

        # What the rule will be *after* this write. Both gates below read it:
        # a call that reorders and disables in one go leaves the rule off, so it
        # cannot deliver and demanding `mail_send` would be over-reach — while a
        # call that reorders without disabling is a start.
        enabled_after = is_enabled if is_enabled is not None else _is_on(current)

        # Only an enabled rule can be delivering: reordering a disabled one is
        # bookkeeping, and gating it would demand mail_send for mail that is not
        # moving — the same over-reach as gating an already-disabled stop rule.
        if sequence is not None and enabled_after and _forwards_offsite(current_actions):
            check_sends_mail(
                config,
                "outlook_update_rule",
                "move a forwarding rule in the run order, ahead of the rule hiding it",
            )

        # Whether Exchange still evaluates a rule marked `has_error` is not
        # established, so this gate assumes the worse case for its direction: the
        # repair puts the rule back into service, so it is a start. That repair
        # needs only `mail_rules` without this check. A *disabled* rule is not
        # covered here: a repair does not switch one on, and gating it would
        # demand mail_send for bookkeeping.
        # What matters is whether the rule is enabled *after* this write, not
        # what it is now: repairing a rule the caller is also disabling leaves it
        # off, so it cannot send, and asking for `mail_send` would be over-reach.
        repairing_a_stopped_forwarder = (
            current is not None
            and enabled_after
            and _is_broken(current)
            and _forwards_offsite(current_actions)
            and (
                move_to_folder is not None
                or mark_as_read is not None
                or sequence is not None
                or name is not None
            )
        )
        if repairing_a_stopped_forwarder:
            check_sends_mail(
                config,
                "outlook_update_rule",
                "repair a rule that forwards mail, putting it back into service",
            )

        if (is_enabled is False or sequence is not None) and _stops_processing(current_actions):
            rules, complete = await _rules_in_inbox(graph_client)
            # `not complete`: the list was cut short, so the rule after this one
            # cannot be ruled out. Refusing costs a permission check; the other
            # way leaves a forwarder running unseen.
            if not complete or _masks_a_forwarding_rule(current, rules):
                check_sends_mail(
                    config,
                    "outlook_update_rule",
                    "let a later forwarding rule run by changing the rule that stops processing",
                )

        if move_to_folder is not None or mark_as_read is not None:
            actions = _copy_actions(current.actions if current else None)
            if move_to_folder is not None:
                actions.move_to_folder = await resolve_folder_id(graph_client, move_to_folder)
            if mark_as_read is not None:
                actions.mark_as_read = mark_as_read
            patch.actions = actions

        updated = await builder.patch(patch)
    response = {
        "status": "updated",
        "id": updated.id if updated else rule_id,
        "display_name": sanitize_output(
            (updated.display_name if updated else None) or (name or "")
        ),
        "is_enabled": updated.is_enabled if updated else is_enabled,
        "sequence": updated.sequence if updated else sequence,
    }
    # A `sequence`-only update that Graph ignores would otherwise answer as if it
    # worked while the rule never moved. Report the request beside the result.
    if sequence is not None and response["sequence"] != sequence:
        response["sequence_requested"] = sequence
    return response


async def delete_rule(graph_client: Any, rule_id: str, *, config: Config) -> dict:
    """Delete an inbox rule by ID. Irreversible.

    Deleting a rule that halts processing can start a later forwarding rule, so
    the same ``mail_send`` check that guards creating or re-enabling one applies
    here when the rule being deleted is what was keeping the forwarding off.
    """
    check_permission(config, CATEGORY_MAIL_RULES, "outlook_delete_rule")
    rule_id = validate_graph_id(rule_id)

    builder = graph_client.me.mail_folders.by_mail_folder_id(
        "inbox"
    ).message_rules.by_message_rule_id(rule_id)

    async with _rule_lock(rule_id):
        # Both calls can answer 404, and either one can be the first a retry
        # meets: the SDK raises `ODataError` (an `APIError`) here, not one of
        # this package's errors, so the status has to be read off the SDK's own
        # attribute. A retry after a delete that landed is indistinguishable from
        # an id that never existed — Graph has no delete receipt — so the honest
        # answer is that the rule is not there, which is the state the caller
        # asked for.
        try:
            current = await builder.get()
        except APIError as exc:
            if getattr(exc, "response_status_code", None) != 404:
                raise
            _rule_locks.pop(rule_id, None)
            return {"status": "not_found", "id": rule_id}

        if _stops_processing(getattr(current, "actions", None)):
            rules, complete = await _rules_in_inbox(graph_client)
            if not complete or _masks_a_forwarding_rule(current, rules):
                check_sends_mail(
                    config,
                    "outlook_delete_rule",
                    "let a later forwarding rule run by deleting the rule that stops processing",
                )
        try:
            await builder.delete()
        except APIError as exc:
            if getattr(exc, "response_status_code", None) != 404:
                raise
            _rule_locks.pop(rule_id, None)
            return {"status": "not_found", "id": rule_id}

    _rule_locks.pop(rule_id, None)
    return {"status": "deleted", "id": rule_id}
