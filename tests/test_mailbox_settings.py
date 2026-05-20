"""Tests for mailbox settings tools: get_timezone, set_timezone, get/set auto-reply."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from msgraph.generated.models.automatic_replies_status import AutomaticRepliesStatus
from msgraph.generated.models.external_audience_scope import ExternalAudienceScope

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.mailbox_settings import (
    get_auto_reply,
    get_timezone,
    set_auto_reply,
    set_timezone,
)

_CFG = Config(client_id="test")
_CFG_RO = Config(client_id="test", read_only=True)
_CFG_ALLOW_TZ = Config(client_id="test", allow_categories=["mailbox_settings"])
_CFG_ALLOW_OTHER = Config(client_id="test", allow_categories=["mail_send"])


class TestGetTimezone:
    async def test_returns_timezone_string(self):
        """get_timezone returns the time_zone field from MailboxSettings."""
        settings = MagicMock()
        settings.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_timezone(mock_client)
        assert result == {"timezone": "America/Los_Angeles"}
        mock_client.me.mailbox_settings.get.assert_called_once()

    async def test_returns_empty_string_when_none(self):
        """get_timezone returns empty string when time_zone is None."""
        settings = MagicMock()
        settings.time_zone = None

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_timezone(mock_client)
        assert result == {"timezone": ""}


class TestSetTimezone:
    async def test_patches_mailbox_settings_and_returns_updated(self):
        """set_timezone PATCHes with MailboxSettings(time_zone=...) and returns echoed value."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG)
        assert result == {"status": "updated", "timezone": "America/Los_Angeles"}

        mock_client.me.mailbox_settings.patch.assert_called_once()
        body = mock_client.me.mailbox_settings.patch.call_args.args[0]
        assert body.time_zone == "America/Los_Angeles"

    async def test_falls_back_to_input_when_echo_missing(self):
        """set_timezone returns input value when patch returns None or no time_zone."""
        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_timezone(mock_client, timezone="Pacific Standard Time", config=_CFG)
        assert result == {"status": "updated", "timezone": "Pacific Standard Time"}

    async def test_rejects_empty_string(self):
        """set_timezone raises ValueError on empty string."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await set_timezone(mock_client, timezone="", config=_CFG)

    async def test_rejects_whitespace_only(self):
        """set_timezone raises ValueError on whitespace-only string."""
        mock_client = MagicMock()
        with pytest.raises(ValueError):
            await set_timezone(mock_client, timezone="   ", config=_CFG)

    @pytest.mark.asyncio
    async def test_rejects_control_characters(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock()
        with pytest.raises(ValueError, match="control"):
            await set_timezone(client, "America/\x00Los_Angeles", config=_CFG)

    async def test_raises_read_only(self):
        """set_timezone raises ReadOnlyError when config.read_only=True."""
        mock_client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG_RO)

    async def test_raises_permission_denied_when_not_in_whitelist(self):
        """set_timezone raises PermissionDeniedError when category not in allow_categories."""
        mock_client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG_ALLOW_OTHER)

    async def test_succeeds_when_category_in_whitelist(self):
        """set_timezone succeeds when 'mailbox_settings' is in allow_categories."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(
            mock_client, timezone="America/Los_Angeles", config=_CFG_ALLOW_TZ
        )
        assert result["status"] == "updated"
        assert result["timezone"] == "America/Los_Angeles"

    async def test_succeeds_with_empty_allow_categories(self):
        """set_timezone succeeds when allow_categories is empty (fully open)."""
        echoed = MagicMock()
        echoed.time_zone = "America/Los_Angeles"

        mock_client = MagicMock()
        mock_client.me.mailbox_settings.patch = AsyncMock(return_value=echoed)

        result = await set_timezone(mock_client, timezone="America/Los_Angeles", config=_CFG)
        assert result["status"] == "updated"


def _mk_dttz(date_time: str, time_zone: str) -> MagicMock:
    """Build a mock DateTimeTimeZone-shaped object."""
    obj = MagicMock()
    obj.date_time = date_time
    obj.time_zone = time_zone
    return obj


class TestGetAutoReply:
    async def test_returns_disabled_when_settings_have_no_auto_reply(self):
        """No automatic_replies_setting at all → status='disabled', empty fields, audience='all'."""
        settings = MagicMock()
        settings.automatic_replies_setting = None

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result == {
            "status": "disabled",
            "internal_message": "",
            "external_message": "",
            "external_audience": "all",
            "scheduled_start": "",
            "scheduled_end": "",
        }

    async def test_parses_always_enabled_with_messages_and_contacts_only(self):
        """alwaysEnabled + contactsOnly + both messages echoed."""
        ar = MagicMock()
        ar.status = AutomaticRepliesStatus.AlwaysEnabled
        ar.external_audience = ExternalAudienceScope.ContactsOnly
        ar.internal_reply_message = "Out of office"
        ar.external_reply_message = "Out, will reply later"
        ar.scheduled_start_date_time = None
        ar.scheduled_end_date_time = None

        settings = MagicMock()
        settings.automatic_replies_setting = ar

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result["status"] == "always"
        assert result["external_audience"] == "contacts_only"
        assert result["internal_message"] == "Out of office"
        assert result["external_message"] == "Out, will reply later"
        assert result["scheduled_start"] == ""
        assert result["scheduled_end"] == ""

    async def test_parses_scheduled_with_windows_tz_to_utc(self):
        """scheduled status converts DateTimeTimeZone to UTC ISO 8601 string."""
        ar = MagicMock()
        ar.status = AutomaticRepliesStatus.Scheduled
        ar.external_audience = ExternalAudienceScope.All
        ar.internal_reply_message = "Away"
        ar.external_reply_message = "Away ext"
        # 2026-06-01 08:00 in Pacific = 15:00 UTC (PDT, UTC-7 in summer)
        # However "Pacific Standard Time" is a Windows zone name; we expect the
        # fallback path (unrecognized name) to emit "<date_time>+00:00".
        ar.scheduled_start_date_time = _mk_dttz(
            "2026-06-01T08:00:00.0000000", "Pacific Standard Time"
        )
        ar.scheduled_end_date_time = _mk_dttz(
            "2026-06-05T17:00:00.0000000", "Pacific Standard Time"
        )

        settings = MagicMock()
        settings.automatic_replies_setting = ar

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result["status"] == "scheduled"
        assert result["external_audience"] == "all"
        # Unrecognized Windows zone → fallback uses raw string interpreted as UTC.
        assert result["scheduled_start"].startswith("2026-06-01T08:00:00")
        assert result["scheduled_start"].endswith("Z") or result["scheduled_start"].endswith(
            "+00:00"
        )
        assert result["scheduled_end"].startswith("2026-06-05T17:00:00")

    async def test_parses_scheduled_with_iana_tz_converts_to_utc(self):
        """IANA TZ name is recognized via zoneinfo and converted to UTC."""
        ar = MagicMock()
        ar.status = AutomaticRepliesStatus.Scheduled
        ar.external_audience = ExternalAudienceScope.None_
        ar.internal_reply_message = "Away"
        ar.external_reply_message = "Away ext"
        ar.scheduled_start_date_time = _mk_dttz(
            "2026-06-01T08:00:00.0000000", "America/Los_Angeles"
        )
        ar.scheduled_end_date_time = _mk_dttz("2026-06-05T17:00:00.0000000", "America/Los_Angeles")

        settings = MagicMock()
        settings.automatic_replies_setting = ar

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result["status"] == "scheduled"
        assert result["external_audience"] == "none"
        # 2026-06-01 08:00 PDT (UTC-7) → 15:00 UTC
        assert result["scheduled_start"] == "2026-06-01T15:00:00Z"
        # 2026-06-05 17:00 PDT → 00:00 next day UTC
        assert result["scheduled_end"] == "2026-06-06T00:00:00Z"

    async def test_sanitizes_control_chars_preserves_newlines(self):
        """Reply messages have control chars stripped but newlines preserved."""
        ar = MagicMock()
        ar.status = AutomaticRepliesStatus.AlwaysEnabled
        ar.external_audience = ExternalAudienceScope.All
        ar.internal_reply_message = "Out of office\nBack Monday\x00"
        ar.external_reply_message = "Ext\nMsg\x1b[31m"
        ar.scheduled_start_date_time = None
        ar.scheduled_end_date_time = None

        settings = MagicMock()
        settings.automatic_replies_setting = ar

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result["internal_message"] == "Out of office\nBack Monday"
        assert result["external_message"] == "Ext\nMsg"

    async def test_handles_none_enums_with_defaults(self):
        """None status/audience fall back to 'disabled'/'all'."""
        ar = MagicMock()
        ar.status = None
        ar.external_audience = None
        ar.internal_reply_message = None
        ar.external_reply_message = None
        ar.scheduled_start_date_time = None
        ar.scheduled_end_date_time = None

        settings = MagicMock()
        settings.automatic_replies_setting = ar

        client = MagicMock()
        client.me.mailbox_settings.get = AsyncMock(return_value=settings)

        result = await get_auto_reply(client)
        assert result == {
            "status": "disabled",
            "internal_message": "",
            "external_message": "",
            "external_audience": "all",
            "scheduled_start": "",
            "scheduled_end": "",
        }


class TestSetAutoReply:
    async def test_disabled_sends_minimal_body(self):
        """status='disabled' patches with status=Disabled and no scheduled fields."""
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_auto_reply(client, status="disabled", config=_CFG)
        assert result == {
            "status": "updated",
            "auto_reply_status": "disabled",
            "external_audience": "all",
            "scheduled_start": "",
            "scheduled_end": "",
        }

        client.me.mailbox_settings.patch.assert_called_once()
        body = client.me.mailbox_settings.patch.call_args.args[0]
        ar = body.automatic_replies_setting
        assert ar.status == AutomaticRepliesStatus.Disabled
        assert ar.scheduled_start_date_time is None
        assert ar.scheduled_end_date_time is None

    async def test_always_mirrors_internal_to_external_when_none(self):
        """When external_message is None it mirrors internal_message."""
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        await set_auto_reply(client, status="always", internal_message="OOO", config=_CFG)

        body = client.me.mailbox_settings.patch.call_args.args[0]
        ar = body.automatic_replies_setting
        assert ar.status == AutomaticRepliesStatus.AlwaysEnabled
        assert ar.internal_reply_message == "OOO"
        assert ar.external_reply_message == "OOO"

    async def test_always_uses_distinct_external_message(self):
        """external_message overrides mirror when provided."""
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        await set_auto_reply(
            client,
            status="always",
            internal_message="OOO",
            external_message="Out",
            config=_CFG,
        )

        body = client.me.mailbox_settings.patch.call_args.args[0]
        ar = body.automatic_replies_setting
        assert ar.internal_reply_message == "OOO"
        assert ar.external_reply_message == "Out"

    async def test_scheduled_missing_start_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="start"):
            await set_auto_reply(
                client,
                status="scheduled",
                internal_message="Away",
                end="2026-06-05T00:00:00Z",
                config=_CFG,
            )

    async def test_scheduled_missing_end_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="end"):
            await set_auto_reply(
                client,
                status="scheduled",
                internal_message="Away",
                start="2026-06-01T00:00:00Z",
                config=_CFG,
            )

    async def test_scheduled_builds_datetimetimezone_in_utc(self):
        """Scheduled fields are DateTimeTimeZone with normalized UTC and tz=UTC."""
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_auto_reply(
            client,
            status="scheduled",
            internal_message="Away",
            start="2026-06-01T08:00:00-07:00",
            end="2026-06-05T17:00:00-07:00",
            config=_CFG,
        )

        body = client.me.mailbox_settings.patch.call_args.args[0]
        ar = body.automatic_replies_setting
        assert ar.status == AutomaticRepliesStatus.Scheduled
        assert ar.scheduled_start_date_time.date_time == "2026-06-01T15:00:00Z"
        assert ar.scheduled_start_date_time.time_zone == "UTC"
        assert ar.scheduled_end_date_time.date_time == "2026-06-06T00:00:00Z"
        assert ar.scheduled_end_date_time.time_zone == "UTC"

        assert result["scheduled_start"] == "2026-06-01T15:00:00Z"
        assert result["scheduled_end"] == "2026-06-06T00:00:00Z"
        assert result["auto_reply_status"] == "scheduled"

    async def test_always_with_empty_internal_message_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="internal_message"):
            await set_auto_reply(client, status="always", internal_message="", config=_CFG)

    async def test_always_with_whitespace_internal_message_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="internal_message"):
            await set_auto_reply(client, status="always", internal_message="   ", config=_CFG)

    async def test_control_char_in_message_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="control"):
            await set_auto_reply(client, status="always", internal_message="x\x00y", config=_CFG)

    async def test_invalid_status_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="status"):
            await set_auto_reply(client, status="garbage", config=_CFG)

    async def test_invalid_external_audience_raises(self):
        client = MagicMock()
        with pytest.raises(ValueError, match="external_audience"):
            await set_auto_reply(
                client,
                status="disabled",
                external_audience="enemies",
                config=_CFG,
            )

    async def test_audience_none_maps_to_enum(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        await set_auto_reply(
            client,
            status="always",
            internal_message="OOO",
            external_audience="none",
            config=_CFG,
        )
        body = client.me.mailbox_settings.patch.call_args.args[0]
        assert body.automatic_replies_setting.external_audience == ExternalAudienceScope.None_

    async def test_audience_contacts_only_maps_to_enum(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        await set_auto_reply(
            client,
            status="always",
            internal_message="OOO",
            external_audience="contacts_only",
            config=_CFG,
        )
        body = client.me.mailbox_settings.patch.call_args.args[0]
        assert (
            body.automatic_replies_setting.external_audience == ExternalAudienceScope.ContactsOnly
        )

    async def test_raises_read_only(self):
        client = MagicMock()
        with pytest.raises(ReadOnlyError):
            await set_auto_reply(client, status="disabled", config=_CFG_RO)

    async def test_raises_permission_denied_when_not_in_whitelist(self):
        client = MagicMock()
        with pytest.raises(PermissionDeniedError):
            await set_auto_reply(client, status="disabled", config=_CFG_ALLOW_OTHER)

    async def test_succeeds_when_category_in_whitelist(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_auto_reply(client, status="disabled", config=_CFG_ALLOW_TZ)
        assert result["status"] == "updated"
        assert result["auto_reply_status"] == "disabled"

    async def test_disabled_return_contract(self):
        client = MagicMock()
        client.me.mailbox_settings.patch = AsyncMock(return_value=None)

        result = await set_auto_reply(client, status="disabled", config=_CFG)
        assert result == {
            "status": "updated",
            "auto_reply_status": "disabled",
            "external_audience": "all",
            "scheduled_start": "",
            "scheduled_end": "",
        }
