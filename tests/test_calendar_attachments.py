"""Tests for calendar event attachments: attach, remove, and the listing on get_event."""

import base64
import json
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import PermissionDeniedError, ReadOnlyError
from outlook_mcp.tools.calendar_attachments import attach_to_event, remove_event_attachment
from outlook_mcp.tools.calendar_read import get_event

_LARGE = 3 * 1024 * 1024 + 1


def _cfg(tmp_path, **kwargs):
    return Config(client_id="test", attachments_dir=str(tmp_path), **kwargs)


def _client(event=None):
    """A client whose `/me/events/{id}` answers ``event`` and accepts attachments."""
    client = MagicMock()
    builder = client.me.events.by_event_id.return_value
    builder.get = AsyncMock(return_value=event if event is not None else MagicMock(attendees=[]))
    builder.attachments.post = AsyncMock(return_value=MagicMock(id="ATT1="))
    builder.attachments.create_upload_session.post = AsyncMock(
        return_value=MagicMock(upload_url="https://outlook.office.com/upload/session")
    )
    builder.attachments.by_attachment_id.return_value.delete = AsyncMock(return_value=None)
    return client, builder


def _meeting():
    return MagicMock(attendees=[MagicMock()])


class TestAttachToEvent:
    async def test_a_small_file_is_posted_inline(self, tmp_path):
        (tmp_path / "ticket.pdf").write_bytes(b"%PDF-1.7 synthetic")
        client, builder = _client()

        result = await attach_to_event(client, "AAMkEVT=", ["ticket.pdf"], config=_cfg(tmp_path))

        client.me.events.by_event_id.assert_called_with("AAMkEVT=")
        posted = builder.attachments.post.call_args[0][0]
        assert posted.name == "ticket.pdf"
        assert posted.content_bytes == b"%PDF-1.7 synthetic"
        assert posted.content_type == "application/pdf"
        assert posted.odata_type == "#microsoft.graph.fileAttachment"
        assert result == {
            "status": "attached",
            "event_id": "AAMkEVT=",
            "attachment_count": 1,
            "attachment_ids": ["ATT1="],
        }

    async def test_a_large_file_takes_the_event_upload_session(self, tmp_path):
        big = tmp_path / "boarding.pdf"
        with open(big, "wb") as f:
            f.truncate(_LARGE)
        client, builder = _client()

        with patch(
            "outlook_mcp.tools.calendar_attachments._upload_large_file", new_callable=AsyncMock
        ) as upload:
            result = await attach_to_event(client, "AAMkEVT=", [str(big)], config=_cfg(tmp_path))

        body = builder.attachments.create_upload_session.post.call_args[0][0]
        # The event's own request-body class, not the message one.
        assert type(body).__module__.startswith("msgraph.generated.users.item.events.")
        item = body.attachment_item
        assert (item.name, item.size, item.content_type) == (
            "boarding.pdf",
            _LARGE,
            "application/pdf",
        )
        assert item.attachment_type.value == "file"
        upload.assert_awaited_once_with(
            "https://outlook.office.com/upload/session", str(big.resolve()), _LARGE
        )
        builder.attachments.post.assert_not_called()
        # No ID yet for an upload-session file: the shared helper does not read
        # the final chunk's response (#125).
        assert result["attachment_count"] == 1
        assert result["attachment_ids"] == []

    async def test_small_and_large_files_together(self, tmp_path):
        (tmp_path / "note.txt").write_bytes(b"seat 12A")
        with open(tmp_path / "big.bin", "wb") as f:
            f.truncate(_LARGE)
        client, builder = _client()

        with patch(
            "outlook_mcp.tools.calendar_attachments._upload_large_file", new_callable=AsyncMock
        ):
            result = await attach_to_event(
                client, "AAMkEVT=", ["note.txt", "big.bin"], config=_cfg(tmp_path)
            )

        assert result["attachment_count"] == 2
        assert result["attachment_ids"] == ["ATT1="]
        builder.attachments.post.assert_called_once()
        builder.attachments.create_upload_session.post.assert_called_once()

    async def test_exactly_three_megabytes_is_still_inline(self, tmp_path):
        """The threshold is the mail tools' own, `size > 3 MiB` takes the session."""
        with open(tmp_path / "edge.bin", "wb") as f:
            f.truncate(_LARGE - 1)
        client, builder = _client()

        await attach_to_event(client, "AAMkEVT=", ["edge.bin"], config=_cfg(tmp_path))

        builder.attachments.post.assert_called_once()
        builder.attachments.create_upload_session.post.assert_not_called()

    async def test_a_path_outside_attachments_dir_is_refused(self, tmp_path):
        outside = tmp_path / "outside"
        outside.mkdir()
        (outside / "secret.txt").write_bytes(b"x")
        fence = tmp_path / "attachments"
        client, builder = _client()

        with pytest.raises(ValueError, match="outside the permitted directory"):
            await attach_to_event(
                client, "AAMkEVT=", [str(outside / "secret.txt")], config=_cfg(fence)
            )

        builder.attachments.post.assert_not_called()

    async def test_a_missing_file_is_a_refusal_the_model_can_read(self, tmp_path):
        """A ValueError, not FileNotFoundError: an OSError reaches the model as a
        bare "Error executing tool"."""
        client, builder = _client()

        with pytest.raises(ValueError, match="Not a file") as exc:
            await attach_to_event(client, "AAMkEVT=", ["nope.pdf"], config=_cfg(tmp_path))

        assert not isinstance(exc.value, OSError)
        builder.attachments.post.assert_not_called()

    async def test_a_directory_is_refused(self, tmp_path):
        (tmp_path / "folder").mkdir()
        client, builder = _client()

        with pytest.raises(ValueError, match="Not a file"):
            await attach_to_event(client, "AAMkEVT=", ["folder"], config=_cfg(tmp_path))

    async def test_one_bad_path_attaches_nothing(self, tmp_path):
        (tmp_path / "good.txt").write_bytes(b"ok")
        client, builder = _client()

        with pytest.raises(ValueError):
            await attach_to_event(
                client, "AAMkEVT=", ["good.txt", "missing.txt"], config=_cfg(tmp_path)
            )

        builder.attachments.post.assert_not_called()

    async def test_no_files_is_refused_rather_than_answered_attached(self, tmp_path):
        client, builder = _client()

        with pytest.raises(ValueError, match="at least one"):
            await attach_to_event(client, "AAMkEVT=", [], config=_cfg(tmp_path))

        client.me.events.by_event_id.assert_not_called()

    async def test_the_event_id_is_validated_first(self, tmp_path):
        client, builder = _client()

        with pytest.raises(ValueError, match="invalid characters"):
            await attach_to_event(
                client, "../users/someone/events", ["x.txt"], config=_cfg(tmp_path)
            )

        client.me.events.by_event_id.assert_not_called()

    async def test_read_only_is_refused(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client()

        with pytest.raises(ReadOnlyError):
            await attach_to_event(
                client, "AAMkEVT=", ["x.txt"], config=_cfg(tmp_path, read_only=True)
            )

        builder.attachments.post.assert_not_called()

    async def test_it_needs_calendar_write(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client()

        with pytest.raises(PermissionDeniedError, match="calendar_write"):
            await attach_to_event(
                client,
                "AAMkEVT=",
                ["x.txt"],
                config=_cfg(tmp_path, allow_categories=["mail_drafts"]),
            )


class TestAttachingToAMeetingNeedsMailSend:
    """A file on an event with attendees reaches them with the next update, so
    under a policy that withholds `mail_send` it is refused, as rewording is."""

    async def test_an_event_with_attendees_is_refused_under_calendar_only(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client(_meeting())

        with pytest.raises(PermissionDeniedError) as exc:
            await attach_to_event(
                client,
                "AAMkEVT=",
                ["x.txt"],
                config=_cfg(tmp_path, allow_categories=["calendar_write"]),
            )

        assert "mail_send" in str(exc.value)
        assert "attendees" in str(exc.value)
        builder.attachments.post.assert_not_called()

    async def test_an_event_without_attendees_stays_calendar_only(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client(MagicMock(attendees=[]))

        result = await attach_to_event(
            client,
            "AAMkEVT=",
            ["x.txt"],
            config=_cfg(tmp_path, allow_categories=["calendar_write"]),
        )

        assert result["status"] == "attached"

    async def test_it_passes_once_mail_send_is_allowed(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client(_meeting())

        await attach_to_event(
            client,
            "AAMkEVT=",
            ["x.txt"],
            config=_cfg(tmp_path, allow_categories=["calendar_write", "mail_send"]),
        )

        builder.get.assert_not_called()
        builder.attachments.post.assert_called_once()

    async def test_an_unrestricted_server_pays_no_extra_read(self, tmp_path):
        (tmp_path / "x.txt").write_bytes(b"x")
        client, builder = _client(_meeting())

        await attach_to_event(client, "AAMkEVT=", ["x.txt"], config=_cfg(tmp_path))

        builder.get.assert_not_called()

    async def test_the_paths_are_checked_before_the_read(self, tmp_path):
        client, builder = _client(_meeting())

        with pytest.raises(ValueError):
            await attach_to_event(
                client,
                "AAMkEVT=",
                ["missing.txt"],
                config=_cfg(tmp_path, allow_categories=["calendar_write"]),
            )

        builder.get.assert_not_called()


class TestRemoveEventAttachment:
    async def test_it_deletes_that_attachment(self, tmp_path):
        client, builder = _client()

        result = await remove_event_attachment(client, "AAMkEVT=", "ATT1=", config=_cfg(tmp_path))

        client.me.events.by_event_id.assert_called_with("AAMkEVT=")
        builder.attachments.by_attachment_id.assert_called_with("ATT1=")
        builder.attachments.by_attachment_id.return_value.delete.assert_awaited_once()
        assert result == {"status": "removed", "event_id": "AAMkEVT=", "attachment_id": "ATT1="}

    @pytest.mark.parametrize(
        ("event_id", "attachment_id"), [("bad id!", "ATT1="), ("AAMkEVT=", "../x")]
    )
    async def test_both_ids_are_validated(self, tmp_path, event_id, attachment_id):
        client, builder = _client()

        with pytest.raises(ValueError, match="invalid characters"):
            await remove_event_attachment(client, event_id, attachment_id, config=_cfg(tmp_path))

        builder.attachments.by_attachment_id.return_value.delete.assert_not_called()

    async def test_read_only_is_refused(self, tmp_path):
        client, builder = _client()

        with pytest.raises(ReadOnlyError):
            await remove_event_attachment(
                client, "AAMkEVT=", "ATT1=", config=_cfg(tmp_path, read_only=True)
            )

        builder.attachments.by_attachment_id.return_value.delete.assert_not_called()


def _attachment(att_id, name, size, content_type, is_inline=False):
    att = MagicMock()
    att.id, att.name, att.size, att.content_type, att.is_inline = (
        att_id,
        name,
        size,
        content_type,
        is_inline,
    )
    return att


def _event(has_attachments):
    from tests.test_calendar_read import _make_mock_event

    event = _make_mock_event()
    event.has_attachments = has_attachments
    return event


class TestGetEventListsAttachments:
    @staticmethod
    def _client(event, pages):
        client = MagicMock()
        builder = client.me.events.by_event_id.return_value
        builder.get = AsyncMock(return_value=event)
        first, *rest = pages
        builder.attachments.get = AsyncMock(return_value=first)
        builder.attachments.with_url.return_value.get = AsyncMock(side_effect=rest or [None])
        return client, builder

    async def test_an_event_without_files_costs_no_second_request(self):
        client, builder = self._client(_event(False), [MagicMock(value=[], odata_next_link=None)])

        result = await get_event(client, "AAMkEVT=")

        assert result["attachments"] == []
        builder.attachments.get.assert_not_called()

    async def test_the_files_are_listed_by_metadata_only(self):
        page = MagicMock(
            value=[_attachment("ATT1=", "ticket.pdf", 2048, "application/pdf")],
            odata_next_link=None,
        )
        client, builder = self._client(_event(True), [page])

        result = await get_event(client, "AAMkEVT=")

        assert result["attachments"] == [
            {
                "id": "ATT1=",
                "name": "ticket.pdf",
                "size": 2048,
                "content_type": "application/pdf",
                "is_inline": False,
            }
        ]
        config = builder.attachments.get.call_args.kwargs["request_configuration"]
        # Without it Graph returns every file's contentBytes, base64.
        assert config.query_parameters.select == ["id", "name", "contentType", "size", "isInline"]

    async def test_a_second_page_is_followed(self):
        link = "https://graph.microsoft.com/v1.0/me/events/AAMkEVT=/attachments?$skiptoken=x"
        pages = [
            MagicMock(
                value=[_attachment("A1=", "one.pdf", 1, "application/pdf")], odata_next_link=link
            ),
            MagicMock(
                value=[_attachment("A2=", "two.pdf", 2, "application/pdf")], odata_next_link=None
            ),
        ]
        client, builder = self._client(_event(True), pages)

        result = await get_event(client, "AAMkEVT=")

        assert [a["id"] for a in result["attachments"]] == ["A1=", "A2="]
        builder.attachments.with_url.assert_called_once_with(link)

    async def test_a_sender_chosen_name_is_sanitised(self):
        page = MagicMock(
            value=[_attachment("A1=", "inv\x1b[2Joice\r\n.pdf", 1, "application/pdf\r\nX: y")],
            odata_next_link=None,
        )
        client, _ = self._client(_event(True), [page])

        (listed,) = (await get_event(client, "AAMkEVT="))["attachments"]

        assert "\x1b" not in listed["name"] and "\n" not in listed["name"]
        assert "\n" not in listed["content_type"]


class TestEventAttachmentsOnTheWire:
    """What the real SDK pipeline sends, over a mock transport.

    A mock of the SDK asserts which builder was called; this asserts the paths
    and bodies that reach Graph, including the casing of the upload-session
    body, which the event builder's own request-body class decides.
    """

    @staticmethod
    def _client(answers):
        import httpx
        from azure.core.credentials import AccessToken
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as _gra
        from msgraph_core import GraphClientFactory

        class _Credential:
            def get_token(self, *args, **kwargs):
                return AccessToken("not-a-real-token", 9_999_999_999)

        sent: list[httpx.Request] = []

        def handle(request: httpx.Request) -> httpx.Response:
            sent.append(request)
            for (method, suffix), (status, body) in answers.items():
                if request.method == method and request.url.path.endswith(suffix):
                    return httpx.Response(status, json=body) if body else httpx.Response(status)
            raise AssertionError(f"unexpected request: {request.method} {request.url}")

        http_client = GraphClientFactory.create_with_default_middleware(
            options=_gra.options,
            client=httpx.AsyncClient(transport=httpx.MockTransport(handle)),
        )
        auth = AzureIdentityAuthenticationProvider(
            _Credential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=_gra.GraphRequestAdapter(auth, http_client)), sent

    async def test_a_small_file_is_one_post_to_the_events_attachments(self, tmp_path):
        (tmp_path / "ticket.pdf").write_bytes(b"%PDF synthetic")
        client, sent = self._client(
            {("POST", "/attachments"): (201, {"id": "ATT1=", "name": "ticket.pdf"})}
        )

        await attach_to_event(client, "AAMkEVT=", ["ticket.pdf"], config=_cfg(tmp_path))

        (request,) = sent
        assert (request.method, request.url.path) == (
            "POST",
            "/v1.0/me/events/AAMkEVT=/attachments",
        )
        body = json.loads(request.content)
        assert body["@odata.type"] == "#microsoft.graph.fileAttachment"
        assert body["name"] == "ticket.pdf"
        assert base64.b64decode(body["contentBytes"]) == b"%PDF synthetic"
        assert body["contentType"] == "application/pdf"

    async def test_a_large_file_opens_the_events_upload_session(self, tmp_path):
        with open(tmp_path / "big.pdf", "wb") as f:
            f.truncate(_LARGE)
        client, sent = self._client(
            {
                ("POST", "/attachments/createUploadSession"): (
                    201,
                    {
                        "uploadUrl": "https://outlook.office.com/upload/s",
                        "nextExpectedRanges": ["0-"],
                    },
                )
            }
        )

        with patch(
            "outlook_mcp.tools.calendar_attachments._upload_large_file", new_callable=AsyncMock
        ) as upload:
            await attach_to_event(client, "AAMkEVT=", ["big.pdf"], config=_cfg(tmp_path))

        (request,) = sent
        assert request.url.path == "/v1.0/me/events/AAMkEVT=/attachments/createUploadSession"
        assert json.loads(request.content) == {
            "AttachmentItem": {
                "attachmentType": "file",
                "contentType": "application/pdf",
                "name": "big.pdf",
                "size": _LARGE,
            }
        }
        assert upload.await_args.args[0] == "https://outlook.office.com/upload/s"

    async def test_removal_is_one_delete(self, tmp_path):
        client, sent = self._client({("DELETE", "/attachments/ATT1="): (204, None)})

        await remove_event_attachment(client, "AAMkEVT=", "ATT1=", config=_cfg(tmp_path))

        (request,) = sent
        assert (request.method, request.url.path) == (
            "DELETE",
            "/v1.0/me/events/AAMkEVT=/attachments/ATT1=",
        )

    async def test_the_listing_selects_metadata_only(self):
        client, sent = self._client(
            {
                ("GET", "/me/events/AAMkEVT="): (200, {"id": "AAMkEVT=", "hasAttachments": True}),
                ("GET", "/attachments"): (
                    200,
                    {
                        "value": [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                "id": "A1=",
                                "name": "t.pdf",
                                "size": 3,
                                "contentType": "application/pdf",
                                "isInline": False,
                            }
                        ]
                    },
                ),
            }
        )

        result = await get_event(client, "AAMkEVT=")

        read, listing = sent
        assert read.url.path == "/v1.0/me/events/AAMkEVT="
        assert listing.url.path == "/v1.0/me/events/AAMkEVT=/attachments"
        assert listing.url.params["$select"] == "id,name,contentType,size,isInline"
        assert result["attachments"][0]["name"] == "t.pdf"


class TestTheServerWrappers:
    """`server.py` hands both tools their arguments positionally; run the real
    wrappers over the real handlers and look at what reaches the builder."""

    @staticmethod
    def _patched(sdk, config):
        from outlook_mcp import server as server_mod

        client = MagicMock()
        client.sdk_client = sdk
        return (
            server_mod,
            patch.object(server_mod, "_get_graph_client", return_value=client),
            patch.object(server_mod, "_get_config", return_value=config),
        )

    async def test_attach(self, tmp_path):
        (tmp_path / "ticket.pdf").write_bytes(b"x")
        sdk, builder = _client()
        server_mod, graph, cfg = self._patched(sdk, _cfg(tmp_path))

        with graph, cfg:
            result = await server_mod.outlook_attach_to_event(
                MagicMock(), event_id="AAMkEVT=", attachment_paths=["ticket.pdf"]
            )

        sdk.me.events.by_event_id.assert_called_with("AAMkEVT=")
        assert builder.attachments.post.call_args[0][0].name == "ticket.pdf"
        assert result["attachment_ids"] == ["ATT1="]

    async def test_remove(self, tmp_path):
        sdk, builder = _client()
        server_mod, graph, cfg = self._patched(sdk, _cfg(tmp_path))

        with graph, cfg:
            await server_mod.outlook_remove_event_attachment(
                MagicMock(), event_id="AAMkEVT=", attachment_id="ATT1="
            )

        sdk.me.events.by_event_id.assert_called_with("AAMkEVT=")
        builder.attachments.by_attachment_id.assert_called_with("ATT1=")
