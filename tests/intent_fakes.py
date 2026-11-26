"""A recorded Graph fake for the intent tests: one handler for the raw path and the SDK.

``FakeGraph`` answers every request an intent call makes, the SDK's included (a
real ``GraphServiceClient`` built over the same ``httpx.MockTransport``), and
records what was sent. It proves logic, not Graph behaviour: what Graph does with
a query is the live guards' job (``test_live_intent_queries.py``).

Counting follows the contracts: every top-level request is one Graph operation,
except a ``$batch`` envelope, which counts as its sub-requests.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any
from urllib.parse import unquote

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"
APPLIED = 'IdType="ImmutableId"'


@dataclass
class Seen:
    """One request as Graph received it."""

    method: str
    path: str
    params: dict[str, str]
    headers: dict[str, str]
    body: Any = None
    in_batch: bool = False

    @property
    def search(self) -> str | None:
        return self.params.get("$search")

    @property
    def filter(self) -> str | None:
        return self.params.get("$filter")


Handler = Callable[[Seen, re.Match], Any]


@dataclass
class Route:
    method: str
    pattern: re.Pattern
    handler: Handler


@dataclass
class FakeGraph:
    """Routes by method and path; a handler returns a body, or ``(status, body)``."""

    routes: list[Route] = field(default_factory=list)
    seen: list[Seen] = field(default_factory=list)
    # What `Preference-Applied` says on id-bearing replies: the id family Graph granted.
    preference_applied: str | None = APPLIED

    def on(self, path: str, handler: Handler, method: str = "GET") -> None:
        self.routes.insert(0, Route(method, re.compile(rf"^{path}$", re.IGNORECASE), handler))

    # ── Counting ────────────────────────────────────────────────────────

    @property
    def operations(self) -> list[Seen]:
        """Every counted Graph operation: top-level requests and batch sub-requests."""
        return [s for s in self.seen if not (s.path.endswith("/$batch") and not s.in_batch)]

    @property
    def counted(self) -> int:
        return len(self.operations)

    def searches(self) -> list[str]:
        return [s.search for s in self.operations if s.search]

    def filters(self) -> list[str]:
        return [s.filter for s in self.operations if s.filter]

    # ── Transport ───────────────────────────────────────────────────────

    def handle(self, request: httpx.Request) -> httpx.Response:
        path = unquote(request.url.path).removeprefix("/v1.0")
        params = dict(request.url.params)
        headers = {k.lower(): v for k, v in request.headers.items()}
        body = json.loads(request.content) if request.content else None
        if request.method == "POST" and path == "/$batch":
            self.seen.append(Seen("POST", path, params, headers, body))
            replies = [self._sub(sub) for sub in body["requests"]]
            return httpx.Response(200, json={"responses": replies})
        status, payload, reply_headers = self._dispatch(
            Seen(request.method, path, params, headers, body)
        )
        return httpx.Response(status, json=payload, headers=reply_headers)

    def _sub(self, sub: dict) -> dict:
        url = httpx.URL(GRAPH + sub["url"])
        seen = Seen(
            sub["method"],
            unquote(url.path).removeprefix("/v1.0"),
            dict(url.params),
            {k.lower(): v for k, v in (sub.get("headers") or {}).items()},
            sub.get("body"),
            in_batch=True,
        )
        status, payload, headers = self._dispatch(seen)
        return {"id": sub["id"], "status": status, "headers": headers, "body": payload}

    def _dispatch(self, seen: Seen) -> tuple[int, Any, dict]:
        self.seen.append(seen)
        for route in self.routes:
            match = route.pattern.match(seen.path)
            if match and route.method == seen.method:
                answer = route.handler(seen, match)
                status, payload = answer if isinstance(answer, tuple) else (200, answer)
                headers = {}
                if (
                    200 <= status < 300
                    and self.preference_applied
                    and "immutableid" in seen.headers.get("prefer", "").lower()
                ):
                    headers["Preference-Applied"] = self.preference_applied
                return status, payload, headers
        return 404, _error("ErrorItemNotFound", f"no fake route for {seen.method} {seen.path}"), {}

    # ── Clients ─────────────────────────────────────────────────────────

    def http_client(self) -> httpx.AsyncClient:
        return httpx.AsyncClient(transport=httpx.MockTransport(self.handle))

    def graph_client(self) -> SimpleNamespace:
        """A ``GraphClient`` stand-in: a real SDK client and a credential, both fake."""
        return SimpleNamespace(sdk_client=self._sdk_client(), credential=FakeCredential())

    def _sdk_client(self):
        from kiota_authentication_azure.azure_identity_authentication_provider import (
            AzureIdentityAuthenticationProvider,
        )
        from msgraph import GraphServiceClient
        from msgraph import graph_request_adapter as gra
        from msgraph_core import GraphClientFactory

        http_client = GraphClientFactory.create_with_default_middleware(
            client=self.http_client(), options=getattr(gra, "options", None)
        )
        auth = AzureIdentityAuthenticationProvider(
            FakeCredential(), scopes=["https://graph.microsoft.com/.default"]
        )
        return GraphServiceClient(request_adapter=gra.GraphRequestAdapter(auth, http_client))


class FakeCredential:
    def get_token(self, *args, **kwargs):
        from azure.core.credentials import AccessToken

        return AccessToken("not-a-real-token", 9_999_999_999)


def _error(code: str, message: str) -> dict:
    return {"error": {"code": code, "message": message}}


def error(status: int, code: str, message: str = "refused") -> tuple[int, dict]:
    return status, _error(code, message)


def page(items: list[dict], next_link: str | None = None) -> dict:
    body: dict = {"value": items}
    if next_link:
        body["@odata.nextLink"] = next_link
    return body


# ── Graph-shaped rows ─────────────────────────────────────────────────────

ME = "me@example.com"


def addr(address: str, name: str | None = None) -> dict:
    return {"emailAddress": {"address": address, "name": name if name is not None else address}}


def letter(
    to: list,
    *,
    sent: str = "2026-11-01T10:00:00Z",
    cc: list | None = None,
    bcc: list | None = None,
    sender: str = ME,
    odata_type: str | None = None,
    conversation: str = "conv-1",
) -> dict:
    """A Sent Items row as Q1 selects it.

    ``to``, ``cc`` and ``bcc`` take addresses or ``(address, name)`` pairs.
    """

    def people(entries):
        return [addr(*e) if isinstance(e, tuple) else addr(e) for e in entries or []]

    row = {
        "from": addr(sender, "Me"),
        "toRecipients": people(to),
        "ccRecipients": people(cc),
        "bccRecipients": people(bcc),
        "sentDateTime": sent,
        "conversationId": conversation,
    }
    if odata_type:
        row["@odata.type"] = odata_type
    return row


def meeting_response(to: str, *, sent: str = "2026-11-02T10:00:00Z") -> dict:
    """An accept sent to an organizer: a meeting message, not a letter."""
    return letter([to], sent=sent, odata_type="#microsoft.graph.eventMessageResponse")
