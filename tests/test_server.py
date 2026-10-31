"""Tests for MCP server tool registration."""

import os
from unittest.mock import MagicMock, patch

import pytest

from outlook_mcp.errors import GraphAPIError, ToolInputError
from outlook_mcp.server import _rule_write_hint, _wrap_tool_errors, mcp

EXPECTED_TOOLS = [
    # Auth (3)
    "outlook_auth_status",
    # Mail read (6)
    "outlook_list_inbox",
    "outlook_read_message",
    "outlook_read_messages",
    "outlook_search_mail",
    "outlook_list_folders",
    "outlook_list_inbox_delta",
    # Mail write (3)
    "outlook_send_message",
    "outlook_reply",
    "outlook_forward",
    # Mail triage (9)
    "outlook_move_message",
    "outlook_delete_message",
    "outlook_flag_message",
    "outlook_categorize_message",
    "outlook_mark_read",
    "outlook_reclassify_message",
    "outlook_list_inbox_overrides",
    "outlook_set_inbox_override",
    "outlook_delete_inbox_override",
    # Calendar read (3)
    "outlook_list_events",
    "outlook_get_event",
    "outlook_find_event_by_key",
    "outlook_list_events_delta",
    # Calendar write (4)
    "outlook_create_event",
    "outlook_update_event",
    "outlook_delete_event",
    "outlook_rsvp",
    "outlook_attach_to_event",
    "outlook_remove_event_attachment",
    # ── Tier 2 ──────────────────────────────────────────
    # Contacts (7)
    "outlook_list_contacts",
    "outlook_search_contacts",
    "outlook_get_contact",
    "outlook_create_contact",
    "outlook_update_contact",
    "outlook_delete_contact",
    "outlook_list_contacts_delta",
    # Digest (1)
    "outlook_changes_since",
    # To Do (14)
    "outlook_list_task_lists",
    "outlook_list_tasks",
    "outlook_get_task",
    "outlook_create_task",
    "outlook_update_task",
    "outlook_complete_task",
    "outlook_delete_task",
    "outlook_add_checklist_item",
    "outlook_update_checklist_item",
    "outlook_delete_checklist_item",
    "outlook_list_task_attachments",
    "outlook_download_task_attachment",
    "outlook_upload_task_attachment",
    "outlook_delete_task_attachment",
    # Mail drafts (5)
    "outlook_list_drafts",
    "outlook_create_draft",
    "outlook_update_draft",
    "outlook_send_draft",
    "outlook_delete_draft",
    # Mail attachments (5)
    "outlook_list_attachments",
    "outlook_download_attachment",
    "outlook_send_with_attachments",
    "outlook_attach_to_draft",
    "outlook_remove_draft_attachment",
    # Mail folders (3)
    "outlook_create_folder",
    "outlook_rename_folder",
    "outlook_delete_folder",
    # Mail thread (2)
    "outlook_list_thread",
    "outlook_copy_message",
    # Batch (1)
    "outlook_batch_triage",
    # User (2)
    "outlook_whoami",
    "outlook_list_calendars",
    # Admin (2)
    "outlook_list_categories",
    "outlook_get_mail_tips",
    # Documents (1) — main's opt-in `documents` group
    "outlook_read_attachment",
    # Mail rules (4) — the opt-in `rules` group
    "outlook_list_rules",
    "outlook_create_rule",
    "outlook_update_rule",
    "outlook_delete_rule",
    # Original MIME (1) — the opt-in `mime` group
    "outlook_export_message",
]


RULE_TOOLS = {
    "outlook_list_rules",
    "outlook_create_rule",
    "outlook_update_rule",
    "outlook_delete_rule",
}


def _documents_are_loaded() -> bool:
    """The `documents` group is opt-in: only when `OUTLOOK_MCP_TOOLSETS` names it."""
    return "outlook_read_attachment" in set(mcp._tool_manager._tools)


def _rules_are_loaded() -> bool:
    """The `rules` group is opt-in: `OUTLOOK_MCP_TOOLSETS` or `mailbox_settings_write`.

    Which surface this process registers therefore depends on the environment it
    was imported in, and both answers are correct — 71 by default (the surface a
    stock install pays for), 75 once rules are asked for.
    """
    return bool(set(mcp._tool_manager._tools) & RULE_TOOLS)


def _expected_tool_names() -> set[str]:
    """The default names, plus each opt-in group only when this process loaded it.

    `documents`, `rules` and `mime` are the off-by-default groups.
    Reading the loaded surface instead of the environment keeps the count honest
    when `mailbox_settings_write` turns rules on without naming the group.
    """
    expected = set(EXPECTED_TOOLS)
    if not _documents_are_loaded():
        expected.discard("outlook_read_attachment")
    if not _rules_are_loaded():
        expected -= RULE_TOOLS
    if "outlook_export_message" not in mcp._tool_manager._tools:
        expected.discard("outlook_export_message")
    return expected


def test_tool_count():
    """71 tools by default, up to 77 with all opt-in groups loaded."""
    registered = set(mcp._tool_manager._tools.keys())
    assert len(registered) == len(_expected_tool_names())


def test_all_tools_registered():
    """Every expected tool name is registered, minus the group that is off."""
    registered = set(mcp._tool_manager._tools.keys())
    assert registered == _expected_tool_names()


def test_no_unexpected_tools():
    """No extra tools beyond the expected set."""
    registered = set(mcp._tool_manager._tools.keys())
    expected = _expected_tool_names()
    extra = registered - expected
    assert not extra, f"Unexpected tools registered: {extra}"


def test_server_metadata():
    """Server has correct name."""
    assert mcp.name == "outlook-mcp"


@pytest.mark.parametrize(
    ("selection", "enabled"),
    [("", False), ("mail,attachments", False), ("mail,attachments,documents", True)],
)
def test_document_instructions_only_reach_clients_with_the_reader(selection, enabled):
    import json
    import subprocess
    import sys

    probe = """import asyncio,json
from mcp.client import Client
from outlook_mcp.server import mcp
async def main():
    async with Client(mcp) as client:
        tools=(await client.list_tools()).tools
        payload={"instructions":client.instructions or "", "tools":[t.name for t in tools]}
        print(json.dumps(payload))
asyncio.run(main())
"""
    result = subprocess.run(
        [sys.executable, "-c", probe],
        env={**os.environ, "OUTLOOK_MCP_TOOLSETS": selection},
        capture_output=True,
        text=True,
        timeout=60,
        check=True,
    )
    received = json.loads(result.stdout.strip().splitlines()[-1])
    assert ("outlook_read_attachment" in received["instructions"]) is enabled
    assert ("outlook_read_attachment" in received["tools"]) is enabled


def test_tools_have_descriptions():
    """Every registered tool has a non-empty description."""
    for name, tool in mcp._tool_manager._tools.items():
        assert tool.description, f"Tool {name} has no description"


# ── Error-wrapper end-to-end ──────────────────────────────────────────


def _make_odata_error(status_code: int, code: str, message: str):
    """Build a Graph SDK ODataError fixture (mirrors test_error_wrapper.py)."""
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    inner = MainError()
    inner.code = code
    inner.message = message

    err = ODataError()
    err.response_status_code = status_code
    err.message = message
    err.error = inner
    return err


@pytest.mark.asyncio
async def test_wrap_tool_errors_converts_graph_sdk_error():
    """A tool decorated with _wrap_tool_errors converts ODataError -> GraphAPIError."""

    @_wrap_tool_errors
    async def fake_tool():
        raise _make_odata_error(403, "ErrorAccessDenied", "no access")

    with pytest.raises(GraphAPIError) as exc_info:
        await fake_tool()

    assert exc_info.value.status_code == 403
    assert exc_info.value.error_code == "ErrorAccessDenied"
    assert exc_info.value.action is not None
    assert "ROADMAP" in exc_info.value.action


@pytest.mark.asyncio
async def test_wrap_tool_errors_passes_through_outlook_mcp_error():
    """OutlookMCPError subclasses must NOT be rewrapped (already structured)."""
    from outlook_mcp.errors import ReadOnlyError

    @_wrap_tool_errors
    async def fake_tool():
        raise ReadOnlyError("outlook_send_message")

    with pytest.raises(ReadOnlyError):
        await fake_tool()


@pytest.mark.asyncio
async def test_wrap_tool_errors_passes_through_value_error():
    """Validation errors stay as ValueError so callers see the original message."""

    @_wrap_tool_errors
    async def fake_tool():
        raise ValueError("Invalid datetime: 'oops'")

    with pytest.raises(ValueError, match="Invalid datetime"):
        await fake_tool()


@pytest.mark.asyncio
async def test_wrap_tool_errors_passes_through_unknown_exception():
    """Truly unexpected errors are bubbled unchanged (not wrapped)."""

    class WeirdError(Exception):
        pass

    @_wrap_tool_errors
    async def fake_tool():
        raise WeirdError("surprise")

    with pytest.raises(WeirdError, match="surprise"):
        await fake_tool()


@pytest.mark.asyncio
async def test_outlook_read_messages_wires_through_to_impl():
    """End-to-end: the new bulk-read tool is callable via the decorated wrapper."""
    fake_ctx = MagicMock()
    fake_ctx.request_context.lifespan_context = {
        "auth": MagicMock(),
        "config": MagicMock(),
    }

    expected = {
        "messages": [{"id": "AAA=", "subject": "x"}],
        "failures": [],
        "requested": 1,
        "succeeded": 1,
        "failed": 0,
    }

    from outlook_mcp import server as server_mod

    async def fake_read_messages(client, message_ids, **kwargs):
        assert message_ids == ["AAA="]
        return expected

    with (
        patch.object(server_mod, "_get_graph_client", return_value=MagicMock()),
        patch.object(server_mod.mail_read, "read_messages", side_effect=fake_read_messages),
    ):
        result = await server_mod.outlook_read_messages(fake_ctx, ["AAA="])

    assert result == expected


@pytest.mark.asyncio
async def test_outlook_changes_since_returns_top_level_shape():
    """End-to-end: the digest wrapper returns mail/events/contacts/delta_tokens/window."""
    fake_ctx = MagicMock()
    fake_ctx.request_context.lifespan_context = {
        "auth": MagicMock(),
        "config": MagicMock(),
    }

    fake_response = {
        "mail": {
            "new_count": 0,
            "modified_count": 0,
            "removed_count": 0,
            "urgent_flagged": [],
            "by_sender": {},
        },
        "events": {"new": [], "modified": [], "cancelled": []},
        "contacts": {"new_count": 0, "modified_count": 0, "removed_count": 0},
        "delta_tokens": {"mail": "m", "events": "e", "contacts": "c"},
        "window": {"from": "2026-05-21T00:00:00Z", "to": "2026-05-22T00:00:00Z"},
    }

    from outlook_mcp import server as server_mod

    async def fake_changes_since(client, tokens, hours):
        return fake_response

    with (
        patch.object(server_mod, "_get_graph_client", return_value=MagicMock()),
        patch.object(server_mod.digest, "changes_since", side_effect=fake_changes_since),
    ):
        result = await server_mod.outlook_changes_since(fake_ctx)

    assert set(result.keys()) >= {"mail", "events", "contacts", "delta_tokens", "window"}


@pytest.mark.asyncio
async def test_wrap_tool_errors_end_to_end_via_mocked_implementation():
    """End-to-end: mock the underlying impl to raise ODataError; the decorated
    server-level tool function hands back GraphAPIError, not the raw SDK shape.
    """
    fake_ctx = MagicMock()
    fake_ctx.request_context.lifespan_context = {
        "auth": MagicMock(),
        "config": MagicMock(),
    }

    odata = _make_odata_error(429, "TooManyRequests", "slow down")

    from outlook_mcp import server as server_mod

    # Stub the graph-client factory so we don't try to build a real Kiota auth
    # provider from a MagicMock credential.
    with (
        patch.object(server_mod, "_get_graph_client", return_value=MagicMock()),
        patch.object(server_mod.mail_read, "list_inbox", side_effect=odata),
    ):
        with pytest.raises(GraphAPIError) as exc_info:
            await server_mod.outlook_list_inbox(fake_ctx)

    assert exc_info.value.status_code == 429
    assert exc_info.value.error_code == "TooManyRequests"
    assert exc_info.value.action is not None
    assert "retry" in exc_info.value.action.lower()


class TestUpdateContactArgumentOrder:
    """`server.py` hands `update_contact` its arguments positionally, and `company`,
    `title` and `job_title` are three neighbouring `str | None` slots (#173). Swapped
    in one file only, a new employer would land in the name prefix and show up in the
    display name, the #119 bug again. As in the calendar tests above, these run the
    real handler behind the real wrapper and assert on the patch that reaches Graph,
    one value at a time, since a test passing all three would still pass a swap.
    """

    @staticmethod
    async def _update(**fields):
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        builder = MagicMock()
        builder.patch = AsyncMock()
        sdk = MagicMock()
        sdk.me.contacts.by_contact_id.return_value = builder
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_update_contact(MagicMock(), contact_id="AAMkAG123=", **fields)
        return builder.patch.call_args[0][0]

    @pytest.mark.asyncio
    async def test_company_alone_lands_in_company_name(self):
        contact = await self._update(company="Contoso Ltd")
        assert contact.company_name == "Contoso Ltd"
        assert contact.title is None
        assert contact.job_title is None

    @pytest.mark.asyncio
    async def test_title_alone_lands_in_the_name_prefix(self):
        contact = await self._update(title="Dr.")
        assert contact.title == "Dr."
        assert contact.company_name is None
        assert contact.job_title is None

    @pytest.mark.asyncio
    async def test_job_title_alone_lands_in_job_title(self):
        contact = await self._update(job_title="Engineer")
        assert contact.job_title == "Engineer"
        assert contact.company_name is None
        assert contact.title is None

    @pytest.mark.asyncio
    async def test_middle_name_alone_lands_in_middle_name(self):
        contact = await self._update(middle_name="Quincy")
        assert contact.middle_name == "Quincy"
        assert contact.personal_notes is None
        assert contact.given_name is None

    @pytest.mark.asyncio
    async def test_personal_notes_alone_land_in_personal_notes(self):
        contact = await self._update(personal_notes="Met at the fair")
        assert contact.personal_notes == "Met at the fair"
        assert contact.middle_name is None

    @pytest.mark.asyncio
    async def test_categories_alone_land_in_categories(self):
        contact = await self._update(categories=["Suppliers"])
        assert contact.categories == ["Suppliers"]
        assert contact.personal_notes is None


class TestCreateContactArgumentOrder:
    """The same positional hand-off on create: `middle_name` and `personal_notes`
    are neighbouring `str | None` slots after `birthday`, and a swap would file a
    note as a name, or a name as a note, with nothing raised."""

    @staticmethod
    async def _create(**fields):
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk = MagicMock()
        sdk.me.contacts.post = AsyncMock(return_value=MagicMock(id="C1"))
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_contact(MagicMock(), first_name="Jane", **fields)
        return sdk.me.contacts.post.call_args[0][0]

    @pytest.mark.asyncio
    async def test_middle_name_alone_lands_in_middle_name(self):
        contact = await self._create(middle_name="Quincy")
        assert contact.middle_name == "Quincy"
        assert contact.personal_notes is None

    @pytest.mark.asyncio
    async def test_personal_notes_alone_land_in_personal_notes(self):
        contact = await self._create(personal_notes="Met at the fair")
        assert contact.personal_notes == "Met at the fair"
        assert contact.middle_name is None

    @pytest.mark.asyncio
    async def test_categories_alone_land_in_categories(self):
        contact = await self._create(categories=["Suppliers"])
        assert contact.categories == ["Suppliers"]
        assert contact.personal_notes is None


class TestCalendarWriteArgumentOrder:
    """`server.py` hands both calendar write tools their arguments positionally.

    A tool-level test cannot see a mis-ordered positional: it patches the
    implementation and asserts on what it was handed, which is the same thing
    twice. These run the *real* handler behind the *real* server wrapper and
    assert on the Graph model that comes out the other end, so a `show_as`
    landing in `recurrence`'s slot fails here rather than at a user's mailbox.

    `timezone` and `show_as` are the pair that matters now. Both arrived as a
    bare `str | None` appended to `create_event`, from two branches that did not
    see each other — #76 and this one — and the signature and `server.py`'s
    positional call are edited in different files. Swap them in one place only
    and nothing fails until a caller supplies a value: `show_as="tentative"`
    alone becomes `Invalid timezone: tentative`, and `timezone="Europe/London"`
    alone becomes `Invalid show_as`. Each direction gets its own test, because a
    single test that passes both values would still pass if both were swapped.
    """

    @staticmethod
    def _ctx():
        from outlook_mcp.config import Config

        ctx = MagicMock()
        ctx.request_context.lifespan_context = {
            "auth": MagicMock(),
            "config": Config(client_id="test"),
        }
        return ctx

    @pytest.mark.asyncio
    async def test_create_event_arguments_land_in_the_right_slots(self):
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk = AsyncMock()
        sdk.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="Review"))
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                location="Room 101",
                timezone="America/New_York",
                show_as="tentative",
            )

        event = sdk.me.events.post.call_args[0][0]
        assert event.subject == "Review"
        assert event.location.display_name == "Room 101"
        assert event.show_as.value == "tentative"
        assert event.start.time_zone == "America/New_York"
        assert event.recurrence is None

    @pytest.mark.asyncio
    async def test_show_as_alone_is_not_read_as_a_timezone(self):
        """The exact call the slot swap breaks, with `timezone` left out.

        If `show_as` were passed into `timezone`'s slot this raises
        `Invalid timezone: tentative` rather than creating anything, and the
        anchor silently stops falling back to the configured zone.
        """
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk = AsyncMock()
        sdk.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="Review"))
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                show_as="tentative",
            )

        event = sdk.me.events.post.call_args[0][0]
        assert event.show_as.value == "tentative"
        # Config default, i.e. the fallback really was reached.
        assert event.start.time_zone == "UTC"

    @pytest.mark.asyncio
    async def test_timezone_alone_is_not_read_as_show_as(self):
        """The mirror of the above, which the other direction of the swap breaks.

        `show_as` must stay unset: assigning it a zone name would raise
        `Invalid show_as`, and a `show_as` we never asked for reaching the wire
        would overwrite the event's free/busy status.
        """
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk = AsyncMock()
        sdk.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="Review"))
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                timezone="Europe/London",
            )

        event = sdk.me.events.post.call_args[0][0]
        assert event.start.time_zone == "Europe/London"
        assert event.show_as is None

    @staticmethod
    def _sdk_with_a_second_calendar():
        from unittest.mock import AsyncMock

        cal = MagicMock()
        cal.name, cal.id, cal.can_edit = "Kids", "KIDS77=", True
        sdk = MagicMock()
        sdk.me.calendars.get = AsyncMock(return_value=MagicMock(value=[cal], odata_next_link=None))
        sdk.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="Review"))
        secondary = MagicMock()
        secondary.events.post = AsyncMock(return_value=MagicMock(id="E2", subject="Review"))
        sdk.me.calendars.by_calendar_id = MagicMock(return_value=secondary)
        return sdk, secondary

    @pytest.mark.asyncio
    async def test_calendar_alone_is_not_read_as_show_as(self):
        """`calendar` is appended after `show_as`, the slot a swap would put it in.

        Swapped, `calendar="Kids"` raises `Invalid show_as: Kids`, and only when
        a value is supplied.
        """
        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk, secondary = self._sdk_with_a_second_calendar()
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                calendar="Kids",
            )

        sdk.me.calendars.by_calendar_id.assert_called_once_with("KIDS77=")
        event = secondary.events.post.call_args[0][0]
        assert event.show_as is None
        sdk.me.events.post.assert_not_called()

    @pytest.mark.asyncio
    async def test_show_as_beside_calendar_is_not_read_as_a_calendar(self):
        """The other direction: swapped, `show_as` is looked up as a calendar name."""
        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk, secondary = self._sdk_with_a_second_calendar()
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                show_as="tentative",
            )

        sdk.me.calendars.get.assert_not_called()
        event = sdk.me.events.post.call_args[0][0]
        assert event.show_as.value == "tentative"

    @pytest.mark.asyncio
    async def test_key_alone_is_not_read_as_show_as(self):
        """`key` is appended after `show_as`; swapped, it raises `Invalid show_as`."""
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        sdk = AsyncMock()
        sdk.me.events.post = AsyncMock(return_value=MagicMock(id="E1", subject="Review"))
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_create_event(
                self._ctx(),
                subject="Review",
                start="2026-09-07T12:30:00Z",
                end="2026-09-07T13:00:00Z",
                key="booking:ABC123",
            )

        event = sdk.me.events.post.call_args[0][0]
        assert event.show_as is None
        assert event.single_value_extended_properties[0].value == "booking:ABC123"

    @pytest.mark.asyncio
    async def test_update_key_alone_is_not_read_as_a_timezone(self):
        """`key` is appended after `timezone`; swapped, the call is refused for
        lacking start and end."""
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        sdk = MagicMock()
        sdk.me.events.by_event_id = MagicMock(return_value=builder)
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_update_event(
                self._ctx(), event_id="AAMkAG123=", key="booking:ABC123"
            )

        event = builder.patch.call_args[0][0]
        assert event.single_value_extended_properties[0].value == "booking:ABC123"
        assert event.start is None

    @pytest.mark.asyncio
    async def test_update_event_arguments_land_in_the_right_slots(self):
        from unittest.mock import AsyncMock

        from outlook_mcp import server as server_mod
        from outlook_mcp.config import Config

        builder = MagicMock()
        builder.patch = AsyncMock(return_value=MagicMock(id="AAMkAG123="))
        sdk = MagicMock()
        sdk.me.events.by_event_id = MagicMock(return_value=builder)
        client = MagicMock()
        client.sdk_client = sdk

        with (
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=Config(client_id="test")),
        ):
            await server_mod.outlook_update_event(
                self._ctx(),
                event_id="AAMkAG123=",
                subject="Review",
                show_as="oof",
            )

        event = builder.patch.call_args[0][0]
        assert event.subject == "Review"
        assert event.show_as.value == "oof"
        # The slot next to show_as in the call, and the one a swap would hit.
        assert event.is_all_day is None


def test_request_urls_are_not_logged():
    """httpx logs every request URL at INFO, and those URLs carry search terms,
    filter addresses and message ids. The MCP SDK sets the root logger to INFO
    by default, which some clients write to a log file. README promises that
    recipient addresses are never logged.

    In a fresh process, because pytest's own log capture leaves the root logger
    with handlers before the server is imported, and `logging.basicConfig`
    does nothing when handlers exist — this has to see what a real start does.
    """
    import os
    import subprocess
    import sys

    probe = (
        "import logging\n"
        "import outlook_mcp.server\n"
        "print(logging.getLogger('httpx').isEnabledFor(logging.INFO))\n"
    )
    out = subprocess.run(
        [sys.executable, "-c", probe],
        env=dict(os.environ),
        capture_output=True,
        text=True,
        check=True,
    )
    assert out.stdout.strip().splitlines()[-1] == "False"


def test_tool_groups_names_every_real_tool_and_nothing_else():
    """A ghost entry in `TOOL_GROUPS` registers nothing but passes every guard.

    The docs test reads the group table, so an entry for a tool that does not
    exist satisfies it, the README row check and the tool count at once.
    """
    from outlook_mcp import server as server_module
    from outlook_mcp import toolsets

    registered = {
        name
        for name in dir(server_module)
        if name.startswith("outlook_") and callable(getattr(server_module, name))
    }

    assert set(toolsets.TOOL_GROUPS) == registered


def test_a_403_on_a_rule_write_points_at_the_settings_consent():
    """The generic 403 hint says "not supported for this account type", which is
    wrong for a rule write — the endpoint works, the consent does not cover it.
    """
    from outlook_mcp.errors import LEAVE_SETTINGS_TO_THE_USER, RULE_WRITE_NEEDS_SETTINGS_WRITE

    err = _rule_write_hint(
        GraphAPIError(403, "ErrorAccessDenied", "Access denied"),
        "outlook_create_rule",
        settings_write_on=False,
    )

    # The exact constant, not just phrases: the flag-on text contains every
    # phrase checked here, so a hint that always picked it passed this test.
    assert err.action == RULE_WRITE_NEEDS_SETTINGS_WRITE
    assert "mailbox_settings_write" in err.action
    assert LEAVE_SETTINGS_TO_THE_USER in err.action
    # The text cannot see which cause it is, so it must offer both rather than
    # assert the consent as the cause. A hint that says "this sign-in consented
    # the read scope only" states a fact the server cannot observe.
    assert "Either" in err.action
    assert "mail-flow policy" in err.action


def test_a_403_on_a_rule_write_with_the_setting_on_is_not_told_to_set_it():
    """With the flag on, the sign-in already asked for the scope, so the refusal
    is either a token from before it or a policy — and the advice has to say so
    instead of asserting a cause it cannot see."""
    from outlook_mcp.errors import RULE_WRITE_STILL_DENIED

    err = _rule_write_hint(
        GraphAPIError(403, "ErrorAccessDenied", "Forwarding blocked by policy"),
        "outlook_update_rule",
        settings_write_on=True,
    )

    assert err.action == RULE_WRITE_STILL_DENIED
    assert "policy" in err.action
    # Not "only its administrator can change that": a personal account has no
    # administrator, so the hint has to say that case has no fix at all.
    assert "Either" in err.action
    assert "personal account has no administrator" in err.action


def test_the_real_config_reader_turns_the_rule_tools_on(tmp_path):
    """Every other test patches `_mailbox_settings_write_is_on`, so a reader that
    always returned False passed all of them. This runs the real one in a fresh
    interpreter: a config that asks for the scope, and no `OUTLOOK_MCP_TOOLSETS`.
    """
    import json as _json
    import subprocess
    import sys

    (tmp_path / "config.json").write_text(
        _json.dumps({"client_id": "test-client", "mailbox_settings_write": True}),
        encoding="utf-8",
    )
    code = """
import json
from outlook_mcp import toolsets
from outlook_mcp.server import mcp
names = list(mcp._tool_manager._tools)
rules = [n for n in names if toolsets.TOOL_GROUPS.get(n) == "rules"]
print(json.dumps({"rules": len(rules), "count": len(names)}))
"""
    env = {**os.environ, "OUTLOOK_MCP_CONFIG_DIR": str(tmp_path)}
    env.pop("OUTLOOK_MCP_TOOLSETS", None)
    env.pop("PYTHONPATH", None)
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        env=env,
        timeout=120,
    )

    assert result.returncode == 0, result.stderr
    measured = _json.loads(result.stdout.strip().splitlines()[-1])
    assert measured["rules"] == 4, result.stdout
    assert measured["count"] == 75, result.stdout


@pytest.mark.parametrize(
    "tool_name", ["outlook_create_rule", "outlook_update_rule", "outlook_delete_rule"]
)
def test_every_rule_write_tool_gets_the_hint(tool_name):
    """Dropping one from `_RULE_WRITE_TOOLS` used to be invisible."""
    err = _rule_write_hint(
        GraphAPIError(403, "ErrorAccessDenied", "Access denied"),
        tool_name,
        settings_write_on=False,
    )

    assert "mailbox_settings_write" in (err.action or "")


def test_a_403_on_another_tool_keeps_the_dead_ends_hint():
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    from outlook_mcp.errors import wrap_graph_error

    inner = MainError()
    inner.code = "ErrorAccessDenied"
    inner.message = "Access is denied."
    failed = ODataError()
    failed.response_status_code = 403
    failed.error = inner

    kept = _rule_write_hint(wrap_graph_error(failed), "outlook_list_folders")

    assert "dead-ends" in (kept.action or "")


async def test_the_wrapper_applies_the_rule_hint_to_a_graph_api_error(monkeypatch):
    """`call_without_replay` raises `GraphAPIError`, and the wrapper used to pass
    that class through untouched — so create never got the hint update and delete
    did. This is the path create actually takes."""
    import outlook_mcp.server as server

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: False)

    async def outlook_create_rule():
        raise GraphAPIError(403, "ErrorAccessDenied", "Access denied.")

    wrapped = _wrap_tool_errors(outlook_create_rule)

    with pytest.raises(GraphAPIError) as excinfo:
        await wrapped()

    assert "mailbox_settings_write" in (excinfo.value.action or "")
    assert "dead-ends" not in (excinfo.value.action or "")
    # The exact flag-off text: the flag-on constant also mentions
    # `mailbox_settings_write`, so a hint that always picked it passed this test.
    from outlook_mcp.errors import RULE_WRITE_NEEDS_SETTINGS_WRITE

    assert excinfo.value.action == RULE_WRITE_NEEDS_SETTINGS_WRITE


async def test_the_wrapper_keeps_the_sdk_error_as_the_cause(monkeypatch):
    """`from None` wiped `__cause__` for every tool that raises its own error.

    The server-side traceback needs the SDK error underneath; the hint must not
    cost it.
    """
    import outlook_mcp.server as server
    from outlook_mcp.errors import RULE_WRITE_NEEDS_SETTINGS_WRITE

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: False)
    sdk_error = RuntimeError("the SDK said no")

    async def outlook_create_rule():
        raise GraphAPIError(403, "ErrorAccessDenied", "Access denied.") from sdk_error

    wrapped = _wrap_tool_errors(outlook_create_rule)

    with pytest.raises(GraphAPIError) as excinfo:
        await wrapped()

    assert excinfo.value.__cause__ is sdk_error
    assert excinfo.value.action == RULE_WRITE_NEEDS_SETTINGS_WRITE


async def test_the_wrapper_still_translates_an_sdk_error_for_a_rule_write(monkeypatch):
    """The other path, for the tools that raise straight from the SDK."""
    from kiota_abstractions.api_error import APIError

    import outlook_mcp.server as server

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: False)

    async def outlook_delete_rule():
        failed = APIError()
        failed.response_status_code = 403
        raise failed

    wrapped = _wrap_tool_errors(outlook_delete_rule)

    with pytest.raises(GraphAPIError) as excinfo:
        await wrapped()

    assert "mailbox_settings_write" in (excinfo.value.action or "")


@pytest.mark.parametrize("settings_write_on", [False, True])
def test_a_403_on_listing_rules_names_both_causes(settings_write_on):
    """Listing needs only the read scope the default consent asks for, so the write
    advice does not fit, and the generic hint sends the agent to a dead-ends page
    that still lists rules. A live personal mailbox answered this 403."""
    from outlook_mcp.errors import RULE_LIST_DENIED

    err = _rule_write_hint(
        GraphAPIError(403, "ErrorAccessDenied", "Access is denied."),
        "outlook_list_rules",
        settings_write_on=settings_write_on,
    )

    assert err.action == RULE_LIST_DENIED
    assert "dead-ends" not in err.action


async def test_the_wrapper_gives_a_refused_rule_listing_the_listing_hint(monkeypatch):
    """`list_rules` raises straight from the SDK, so the hint has to reach that path."""
    from kiota_abstractions.api_error import APIError

    import outlook_mcp.server as server
    from outlook_mcp.errors import RULE_LIST_DENIED

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: False)

    async def outlook_list_rules():
        failed = APIError()
        failed.response_status_code = 403
        raise failed

    with pytest.raises(GraphAPIError) as excinfo:
        await _wrap_tool_errors(outlook_list_rules)()

    assert excinfo.value.action == RULE_LIST_DENIED


async def test_the_wrapper_uses_the_denied_hint_when_the_setting_is_on(monkeypatch):
    """The three wrapper tests above all patch the flag to False, so a wrapper
    that always returned the "set the flag" text would pass all of them. With
    the flag on, the hint must be the denied text, never the consent advice."""
    import outlook_mcp.server as server
    from outlook_mcp.errors import RULE_WRITE_STILL_DENIED

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: True)

    async def outlook_create_rule():
        raise GraphAPIError(403, "ErrorAccessDenied", "Forwarding blocked by policy")

    wrapped = _wrap_tool_errors(outlook_create_rule)

    with pytest.raises(GraphAPIError) as excinfo:
        await wrapped()

    assert excinfo.value.action == RULE_WRITE_STILL_DENIED
    assert "mailbox_settings_write: true" not in excinfo.value.action


async def test_an_unrelated_tool_error_is_still_passed_through_untouched(monkeypatch):
    import outlook_mcp.server as server

    monkeypatch.setattr(server, "_mailbox_settings_write_is_on", lambda: False)

    async def outlook_list_folders():
        raise ToolInputError("no such folder")

    wrapped = _wrap_tool_errors(outlook_list_folders)

    with pytest.raises(ToolInputError, match="no such folder"):
        await wrapped()
