"""`idempotentHint` is a promise a client can act on, so it is opt-in per tool.

A client that believes a tool is idempotent may retry it after a timeout. On a
tool that is not, that is a second email sent, a second invitation issued, or a
404 on an ID that no longer exists. Absent is safe; wrong is not — so the hint
goes on a pinned list of tools whose operation is a PATCH of an absolute value,
and the list is guarded rather than inferred.

Deliberately excluded, each for a checked reason:

* every ``send`` / ``reply`` / ``forward`` — a repeat sends again
* ``outlook_rsvp`` — sends a fresh response each time
* ``outlook_update_event`` — with ``attendees`` it re-issues invitations
* ``outlook_move_message`` / ``outlook_copy_message`` — Exchange mints a new ID
  on the move, so the same arguments do not address the same message twice
* every ``create`` / ``attach`` — a repeat duplicates
* every ``delete`` — the second call finds nothing there
* ``outlook_complete_task`` — writes ``completedDateTime = now``, so a repeat
  changes the stored value even though the status does not move
"""

import pytest
from mcp.client import Client

from outlook_mcp import toolsets
from outlook_mcp.server import mcp

EXPECTED_IDEMPOTENT = {
    "outlook_flag_message",
    "outlook_mark_read",
    "outlook_categorize_message",
    "outlook_rename_folder",
    "outlook_set_inbox_override",
    "outlook_switch_account",
    "outlook_download_attachment",
    # Same audit as their mail twins: an absolute-value PATCH
    # (checklist check-off/rename) and a download that overwrites a fixed
    # path with the same bytes.
    "outlook_download_task_attachment",
    "outlook_update_checklist_item",
}


def test_the_idempotent_set_is_exactly_what_we_audited():
    """Drift guard: adding a tool here is a claim that needs the same audit."""
    assert toolsets.IDEMPOTENT == EXPECTED_IDEMPOTENT


def test_no_tool_is_both_destructive_and_idempotent():
    assert not (toolsets.IDEMPOTENT & toolsets.DESTRUCTIVE)


def test_read_only_tools_do_not_carry_the_hint():
    """It only says anything about a tool that writes; on a read it is noise."""
    assert not (toolsets.IDEMPOTENT & toolsets.READ_ONLY)


@pytest.mark.parametrize(
    "name",
    [
        "outlook_send_message",
        "outlook_reply",
        "outlook_forward",
        "outlook_rsvp",
        "outlook_update_event",
        "outlook_move_message",
        "outlook_copy_message",
        "outlook_create_event",
        "outlook_create_draft",
        "outlook_attach_to_draft",
        "outlook_delete_message",
        "outlook_complete_task",
    ],
)
def test_tools_that_repeat_badly_are_not_marked(name):
    """Each of these does something extra on a second identical call."""
    assert name not in toolsets.IDEMPOTENT


def test_every_idempotent_name_is_a_registered_tool():
    registered = set(mcp._tool_manager._tools)
    assert toolsets.IDEMPOTENT <= registered


@pytest.mark.asyncio
async def test_the_hint_reaches_the_wire():
    async with Client(mcp) as client:
        tools = {t.name: t for t in (await client.list_tools()).tools}

    for name in EXPECTED_IDEMPOTENT:
        assert tools[name].annotations.idempotent_hint is True, name

    assert tools["outlook_send_message"].annotations.idempotent_hint is not True


def test_open_world_hint_is_left_to_the_default():
    """It is true by default in the schema, so stating it costs tokens and says nothing.

    Every one of the 62 reaches Microsoft Graph, so the value would be `true` on
    all of them — which is exactly what a client already assumes when the field
    is absent. On a surface measured at ~8.6k tokens a turn, correct-but-inert
    metadata is not free.
    """
    from outlook_mcp.toolsets import annotation_for

    assert annotation_for("outlook_list_inbox").open_world_hint is None
    assert annotation_for("outlook_send_message").open_world_hint is None
