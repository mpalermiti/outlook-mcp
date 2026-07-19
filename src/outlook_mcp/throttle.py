"""Retry-After honoring for the raw-httpx Microsoft Graph paths.

The Graph SDK path retries 429/503 via kiota's ``RetryHandler``. The delta and
``$batch`` tools deliberately bypass the SDK and hit Graph with raw httpx
(``_delta.fetch_delta_pages``, ``mail_read.read_messages``), so they don't
inherit that behavior. These helpers restore it:

- ``send_with_retry`` — retry a throttled *request* (envelope 429/503).
- ``retry_throttled_subrequests`` — retry throttled *sub-requests* inside a
  ``$batch``. Graph returns HTTP 200 for the batch envelope even when
  individual sub-responses are 429/503, so a naive reader records those as
  permanent failures. This re-issues a smaller ``$batch`` of just the
  throttled sub-requests, honoring each sub-response's ``Retry-After``.

Graph also enforces a global 130,000-requests / 10-seconds app-wide ceiling on
top of per-mailbox limits, so direct callers must back off on their own.
"""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from typing import Any

# Status codes Graph uses to signal "back off and retry".
_RETRYABLE = (429, 503)

# Fallback wait (seconds) when a throttled response carries no Retry-After.
# Grows exponentially per attempt (base * 2**attempt).
_DEFAULT_BACKOFF = 2.0


def parse_retry_after(headers: dict[str, Any], default: float) -> float:
    """Return the ``Retry-After`` value in seconds, or ``default``.

    Graph sends ``Retry-After`` as an integer number of seconds. Header dicts
    from a ``$batch`` sub-response are plain JSON dicts (case as sent), so the
    lookup is case-insensitive. A missing or non-numeric value falls back to
    ``default``.
    """
    if not headers:
        return default
    value = None
    for key, val in headers.items():
        if key.lower() == "retry-after":
            value = val
            break
    if value is None:
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


async def send_with_retry(
    client: Any,
    method: str,
    url: str,
    *,
    headers: dict[str, str],
    content: Any = None,
    max_retries: int = 3,
    default_backoff: float = _DEFAULT_BACKOFF,
    sleep: Callable[[float], Awaitable[None]] = asyncio.sleep,
) -> Any:
    """Send an httpx request, retrying on 429/503 with Retry-After / backoff.

    Returns the final ``httpx.Response`` — the last one received, whether it
    succeeded or exhausted ``max_retries`` still throttled. Non-throttle
    responses (including other 4xx/5xx) are returned immediately for the
    caller's normal ``raise_for_status`` handling.

    Dispatches to the client's method-specific function (``client.get`` /
    ``client.post``) rather than ``client.request`` so it matches how the
    delta/batch call sites already invoke httpx.
    """
    send = getattr(client, method.lower())
    call_kwargs: dict[str, Any] = {"headers": headers}
    if content is not None:
        call_kwargs["content"] = content
    attempt = 0
    while True:
        resp = await send(url, **call_kwargs)
        if resp.status_code in _RETRYABLE and attempt < max_retries:
            wait = parse_retry_after(
                dict(resp.headers), default_backoff * (2**attempt)
            )
            await sleep(wait)
            attempt += 1
            continue
        return resp


async def retry_throttled_subrequests(
    post_batch: Callable[[list[dict]], Awaitable[dict]],
    requests: list[dict],
    *,
    max_retries: int = 3,
    default_backoff: float = _DEFAULT_BACKOFF,
    sleep: Callable[[float], Awaitable[None]] = asyncio.sleep,
) -> dict[str, dict]:
    """Run a ``$batch`` and retry any throttled sub-requests.

    ``post_batch`` takes a list of sub-request dicts (each with a string
    ``"id"``) and returns the parsed ``$batch`` payload (``{"responses": [...]}``).
    Sub-responses with status 429/503 are re-sent in a smaller batch, honoring
    the largest ``Retry-After`` in the throttled set, up to ``max_retries``.

    Returns a ``{id -> sub-response}`` dict. A sub-request still throttled after
    ``max_retries`` is recorded with its final 429/503 response (bounded, not
    infinite). Sub-requests Graph never answered are simply absent — the caller
    handles the missing id.
    """
    merged: dict[str, dict] = {}
    pending = list(requests)
    attempt = 0

    while pending:
        payload = await post_batch(pending)
        subs = {str(s.get("id")): s for s in (payload.get("responses") or [])}

        retry_next: list[dict] = []
        max_wait = 0.0
        for req in pending:
            rid = str(req["id"])
            sub = subs.get(rid)
            if sub is None:
                continue  # no response for this id; caller handles the gap
            status = sub.get("status", 0)
            if status in _RETRYABLE and attempt < max_retries:
                wait = parse_retry_after(
                    sub.get("headers") or {}, default_backoff * (2**attempt)
                )
                max_wait = max(max_wait, wait)
                retry_next.append(req)
            else:
                merged[rid] = sub

        if not retry_next:
            break
        attempt += 1
        await sleep(max_wait)
        pending = retry_next

    return merged
