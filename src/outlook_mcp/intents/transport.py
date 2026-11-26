"""The intent tools' counted path to Graph: raw reads with budget, limiter and id family.

``IntentGraph`` is how an intent call reads anything that carries an Outlook id
(messages, events, contacts, attachments). Raw httpx rather than the SDK, for
three reasons:

- **Ids.** Each request carries ``Prefer: IdType="ImmutableId"`` and each reply's
  ``Preference-Applied`` is read (``ids``); a ``$batch`` carries the header on the
  envelope and on every sub-request.
- **Annotations.** The sent table needs each item's ``@odata.type`` to drop
  meeting messages, which the SDK turns into model classes.
- **Counting.** Every request spends from the call's budget before it leaves and
  holds the process limiter while it is out (``budget``).

The SDK helpers the intent tools reuse as they are (folders, calendars, To Do
lists) go through ``counted``, which spends and holds the same way. Their ids are
not affected by the preference (containers, and To Do), so they need no header.

Every URL that receives the token is checked against Graph's host first, and
throttled replies are retried with ``Retry-After`` (``throttle``), as on the
repo's other raw paths.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from typing import Any
from urllib.parse import parse_qsl, quote, urlsplit

import httpx

from outlook_mcp.errors import GraphAPIError, graph_json, wrap_http_status_error
from outlook_mcp.intents import budget as _budget
from outlook_mcp.intents import ids as _ids
from outlook_mcp.intents.budget import RequestBudget, WeightedLimiter, batch_units
from outlook_mcp.intents.ids import IdFamily, combine_prefer
from outlook_mcp.throttle import retry_throttled_subrequests, send_with_retry
from outlook_mcp.tools._delta import require_graph_url

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_TOKEN_SCOPE = "https://graph.microsoft.com/.default"
MAX_BATCH = 20  # Graph's limit on sub-requests per `$batch`
_TIMEOUT_S = 30.0

# The query parameters whose values `understood.queries` echoes (§2.1): the
# `$filter` and KQL strings sent, which hold only addresses, dates and the agent's words.
_ECHOED = ("$filter", "$search")


def make_client() -> httpx.AsyncClient:
    """The httpx client for one intent call. Tests replace this with a mock transport."""
    return httpx.AsyncClient(timeout=_TIMEOUT_S)


def build_url(path: str, params: Mapping[str, Any] | None = None) -> str:
    """``path`` with its query: keys as written (``$filter``), values percent-encoded."""
    if not params:
        return path
    query = "&".join(f"{key}={quote(str(value), safe='')}" for key, value in params.items())
    return f"{path}?{query}"


def sub_error(reply: dict) -> GraphAPIError:
    """A failed ``$batch`` sub-response as the ``GraphAPIError`` the SDK path would raise."""
    body = reply.get("body") if isinstance(reply.get("body"), dict) else {}
    err = body.get("error") if isinstance(body.get("error"), dict) else {}
    status = reply.get("status") if isinstance(reply.get("status"), int) else 0
    return GraphAPIError(
        status,
        str(err.get("code") or "UnknownError"),
        str(err.get("message") or f"Graph API error (status {status})"),
    )


def _header(headers: Any, name: str) -> str | None:
    """A header from an httpx reply or a ``$batch`` sub-response's plain dict."""
    if isinstance(headers, httpx.Headers):
        return headers.get(name)
    if isinstance(headers, Mapping):
        for key, value in headers.items():
            if isinstance(key, str) and key.lower() == name.lower():
                return value if isinstance(value, str) else None
    return None


class IntentGraph:
    """Counted raw Graph reads for one intent call. Use as ``async with``."""

    def __init__(
        self,
        credential: Any,
        budget: RequestBudget,
        *,
        id_family: IdFamily | None = None,
        limiter: WeightedLimiter | None = None,
    ) -> None:
        self.credential = credential
        self.budget = budget
        # The process-wide family and limiter unless given, looked up per call.
        self.id_family = id_family or _ids.ID_FAMILY
        self.limiter = limiter or _budget.LIMITER
        # Every `$filter` and `$search` value sent, in order, for `understood.queries`.
        self.queries: list[str] = []
        self._client: httpx.AsyncClient | None = None
        self._token: str | None = None

    async def __aenter__(self) -> IntentGraph:
        self._client = make_client()
        return self

    async def __aexit__(self, *exc_info: object) -> None:
        if self._client is not None:
            await self._client.aclose()
            self._client = None

    def _headers(self) -> dict[str, str]:
        if self._token is None:
            self._token = self.credential.get_token(GRAPH_TOKEN_SCOPE).token
        return {
            "Authorization": f"Bearer {self._token}",
            "Accept": "application/json",
            "Accept-Encoding": "gzip",
            "Prefer": combine_prefer(None),
        }

    def _note(self, url: str) -> None:
        for key, value in parse_qsl(urlsplit(url).query, keep_blank_values=True):
            if key in _ECHOED and value not in self.queries:
                self.queries.append(value)

    async def _send(self, method: str, url: str, *, units: int, content: str | None = None):
        assert self._client is not None, "IntentGraph is used inside `async with`"
        headers = self._headers()
        if content is not None:
            headers["Content-Type"] = "application/json"
        async with self.limiter.hold(units):
            response = await send_with_retry(
                self._client, method, url, headers=headers, content=content
            )
        try:
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            raise wrap_http_status_error(exc) from None
        return response

    async def get(
        self, path: str, params: Mapping[str, Any] | None = None, *, ids: bool = True
    ) -> dict:
        """One counted GET of ``path`` (``/me/...``). ``ids`` says whether the reply carries
        Outlook ids, and so whether its ``Preference-Applied`` speaks for the id family."""
        url = build_url(path, params)
        return await self._get_url(GRAPH_BASE + url, ids=ids)

    async def follow(self, next_link: str, *, ids: bool = True) -> dict:
        """One counted GET of an ``@odata.nextLink``, after checking it is a Graph URL."""
        return await self._get_url(require_graph_url(next_link, source="@odata.nextLink"), ids=ids)

    async def _get_url(self, url: str, *, ids: bool) -> dict:
        self.budget.spend(1)
        self._note(url)
        response = await self._send("GET", url, units=1)
        if ids:
            self.id_family.observe(response.headers.get("Preference-Applied"))
        return graph_json(response)

    async def batch(self, urls: list[str], *, ids: bool = True) -> list[dict]:
        """GET each relative URL in unordered ``$batch`` envelopes of at most 20.

        Every sub-request is counted before the first envelope leaves, so a batch
        that would pass the budget sends nothing. Replies come back in input
        order as ``{status, headers, body}``; a failed sub-request is returned,
        not raised, and ``sub_error`` reads it.
        """
        if not urls:
            return []
        self.budget.spend(len(urls))
        for url in urls:
            self._note(url)
        replies: list[dict] = []
        for first in range(0, len(urls), MAX_BATCH):
            chunk = urls[first : first + MAX_BATCH]
            requests = [
                {
                    "id": str(n),
                    "method": "GET",
                    "url": url,
                    "headers": {"Prefer": combine_prefer(None)},
                }
                for n, url in enumerate(chunk)
            ]

            async def post(reqs: list[dict], units: int = batch_units(ordered=False)) -> dict:
                response = await self._send(
                    "POST",
                    GRAPH_BASE + "/$batch",
                    units=units,
                    content=json.dumps({"requests": reqs}),
                )
                return graph_json(response)

            by_id = await retry_throttled_subrequests(post, requests)
            for n in range(len(chunk)):
                reply = by_id.get(str(n)) or {
                    "status": 0,
                    "headers": {},
                    "body": {"error": {"code": "NoResponse", "message": "no sub-response"}},
                }
                status = reply.get("status")
                if ids and isinstance(status, int) and 200 <= status < 300:
                    self.id_family.observe(_header(reply.get("headers"), "Preference-Applied"))
                replies.append(reply)
        return replies


class _Counted:
    """An SDK client, or any builder under it, whose requests spend and hold the limiter.

    The SDK's fluent builders end every request in an awaited method (``get`` here;
    the others for completeness), so wrapping attribute access and calls is enough
    to see each request before it goes out. What a request returns is not wrapped.
    """

    _REQUESTS = frozenset({"get", "post", "patch", "put", "delete"})

    def __init__(self, target: Any, budget: RequestBudget, limiter: WeightedLimiter) -> None:
        self._target = target
        self._budget = budget
        self._limiter = limiter

    def __getattr__(self, name: str) -> Any:
        attr = getattr(self._target, name)
        if name in self._REQUESTS and callable(attr):

            async def request(*args: Any, **kwargs: Any) -> Any:
                self._budget.spend(1)
                async with self._limiter.hold(1):
                    return await attr(*args, **kwargs)

            return request
        return _Counted(attr, self._budget, self._limiter)

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        return _Counted(self._target(*args, **kwargs), self._budget, self._limiter)


def counted(sdk_client: Any, budget: RequestBudget, limiter: WeightedLimiter | None = None) -> Any:
    """``sdk_client`` with every request counted against ``budget`` and held on ``limiter``
    (the process limiter unless given)."""
    return _Counted(sdk_client, budget, limiter or _budget.LIMITER)
