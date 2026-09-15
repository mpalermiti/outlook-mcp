"""Tests for MCP server tool registration."""

from unittest.mock import MagicMock, patch

import pytest

from outlook_mcp.errors import GraphAPIError
from outlook_mcp.server import _wrap_tool_errors, mcp

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
    "outlook_list_events_delta",
    # Calendar write (4)
    "outlook_create_event",
    "outlook_update_event",
    "outlook_delete_event",
    "outlook_rsvp",
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
    # Multi-account (2)
    "outlook_list_accounts",
    "outlook_switch_account",
]


def test_tool_count():
    """All 70 tools are registered (auth is CLI-only now)."""
    registered = set(mcp._tool_manager._tools.keys())
    assert len(registered) == 70


def test_all_tools_registered():
    """Every expected tool name is registered on the server."""
    registered = set(mcp._tool_manager._tools.keys())
    for name in EXPECTED_TOOLS:
        assert name in registered, f"Missing tool: {name}"


def test_no_unexpected_tools():
    """No extra tools beyond the expected set."""
    registered = set(mcp._tool_manager._tools.keys())
    expected = set(EXPECTED_TOOLS)
    extra = registered - expected
    assert not extra, f"Unexpected tools registered: {extra}"


def test_server_metadata():
    """Server has correct name."""
    assert mcp.name == "outlook-mcp"


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
