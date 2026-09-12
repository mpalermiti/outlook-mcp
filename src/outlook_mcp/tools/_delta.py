"""Shared helpers for Graph ``$delta`` queries.

These tools deliberately bypass the msgraph SDK and hit Graph's delta
endpoints with raw httpx. The SDK's typed delta builders rebuild URL
templates from query-parameter dataclasses (so a deltaLink passed through
them gets its ``$deltatoken`` stripped) and discard the ``@removed``
annotation on returned items, which delta-query callers need to detect
tombstones. Raw httpx is also easier to wire to the per-endpoint quirks
documented inline:

- ``/me/calendarView/delta`` requires ``startDateTime`` + ``endDateTime``
  on the first call but rejects ``$top``; you must pass
  ``Prefer: odata.maxpagesize=N`` instead.
- ``/me/contacts/delta`` accepts no query string at all on the first call;
  ``Prefer: odata.maxpagesize=N`` is again how you size the page.
- ``/me/mailFolders/{id}/messages/delta`` accepts ``$top``.

Tokens are *not* persisted by outlook-mcp — that's the caller's job (an
agent decides where it wants to store its watermark). We pass the raw
``@odata.deltaLink`` / ``@odata.nextLink`` URL through as the opaque
cursor string.

That last point is why ``require_graph_url`` exists: because the cursor
lives outside this process and comes back in as a request URL carrying the
mailbox bearer token, it is untrusted input, and every URL that would
receive the token is checked against ``graph.microsoft.com`` first.
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlsplit

import httpx

from outlook_mcp.errors import UntrustedURLError
from outlook_mcp.throttle import send_with_retry

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
GRAPH_TOKEN_SCOPE = "https://graph.microsoft.com/.default"

# The only host that may ever receive a Graph bearer token.
GRAPH_HOST = "graph.microsoft.com"

# Control characters and spaces: the set that different URL parsers disagree
# about. A real Graph cursor contains none of them.
_FORBIDDEN_URL_CHARS = re.compile(r"[\x00-\x20\x7f-\xa0]")

# Safety cap multiplier — bound a single tool call to at most this many
# items even when Graph keeps handing us more ``@odata.nextLink`` pages
# inside one delta-sync round. Callers continue by passing the returned
# delta_token (a nextLink) back in.
PAGE_SIZE_CAP_MULTIPLIER = 4


def _bearer_token(credential: Any) -> str:
    """Mint a Graph access token from an azure-identity credential."""
    tok = credential.get_token(GRAPH_TOKEN_SCOPE)
    return tok.token


def require_graph_url(url: str, *, source: str) -> str:
    """Return ``url`` if it is an https Graph URL, else refuse.

    Every URL this module requests carries the mailbox bearer token, and two of
    them arrive from outside: the caller's ``delta_token`` and the
    ``@odata.nextLink`` read back out of a response body. Neither is trustworthy
    — an agent that reads mail can be told what cursor to use — so the host is
    checked before the token is attached, not after.

    Parsed, not string-matched, for the same reason
    ``resolve_attachment_path`` resolves instead of comparing substrings: a
    ``startswith`` test on the Graph prefix passes
    ``https://graph.microsoft.com@evil.example/`` (whose real host is
    ``evil.example``) and ``https://graph.microsoft.com.evil.example/``.

    Parsing alone is not enough either, because parsers disagree. ``urlsplit``
    deletes tab/CR/LF before parsing while an HTTP client does not, so the two
    can read different hosts out of one string. The characters that cause the
    disagreement are refused outright, and the comparison is against the whole
    ``netloc`` rather than ``hostname`` so userinfo and an explicit port — the
    other two classic sources of parser differentials — are refused with it.
    Graph emits neither in a deltaLink.

    ``source`` names where the URL came from so the refusal says which cursor to
    throw away.
    """
    candidate = (url or "").strip()
    if not candidate:
        raise UntrustedURLError(source, url)

    # ``urlsplit`` silently deletes tab, CR and LF before parsing, so
    # ``https://evil.example\t@graph.microsoft.com/x`` reads as Graph here and
    # as something else to an HTTP client. Rather than try to agree with every
    # parser, refuse the characters that make them disagree.
    if _FORBIDDEN_URL_CHARS.search(candidate):
        raise UntrustedURLError(source, url)

    try:
        parsed = urlsplit(candidate)
    except ValueError as exc:
        # A host urlsplit cannot parse is not a host we are willing to send a
        # token to.
        raise UntrustedURLError(source, url) from exc

    # ``netloc``, not ``hostname``: equality also rules out userinfo and an
    # explicit port, neither of which Graph puts in a deltaLink, and both of
    # which are classic ways to make two parsers read different hosts.
    if parsed.scheme.lower() != "https" or parsed.netloc.lower() != GRAPH_HOST:
        raise UntrustedURLError(source, url)

    # Return what was checked, not what was passed in — validating one string
    # and sending another is how a check gets bypassed.
    return candidate


def format_delta_item(raw: dict, normal_formatter) -> dict:
    """Map a raw Graph delta-response item to outlook-mcp's wire shape.

    Items annotated with ``@removed`` are tombstones — Graph returns only
    the ``id`` (and the annotation), no other fields. We collapse those
    to ``{id, is_deleted: True}`` so callers don't have to special-case
    sparse rows.

    Live items go through ``normal_formatter`` (which expects the dict
    shape Graph sends — *not* an SDK object) and get ``is_deleted: False``
    appended.
    """
    if "@removed" in raw:
        return {"id": raw.get("id"), "is_deleted": True}
    out = normal_formatter(raw)
    out["is_deleted"] = False
    return out


async def fetch_delta_pages(
    credential: Any,
    *,
    initial_url: str,
    delta_token: str | None,
    page_size: int,
    headers: dict[str, str] | None = None,
    timeout: float = 30.0,
) -> tuple[list[dict], str | None, bool]:
    """Walk Graph delta pages until a ``deltaLink`` or the safety cap.

    Behavior:

    - When ``delta_token`` is provided it's used as the request URL
      verbatim — it's already a full Graph URL (either a deltaLink that
      starts a new sync round or a nextLink that resumes an in-progress
      one). ``initial_url`` is ignored in that case.
    - When ``delta_token`` is ``None`` we hit ``initial_url`` (the
      caller-built first-sync URL with whatever query params and
      Prefer headers the resource accepts).
    - We auto-follow ``@odata.nextLink`` inside a single tool call up to
      ``page_size * PAGE_SIZE_CAP_MULTIPLIER`` items so an agent isn't
      forced to chain four follow-up calls just to drain Graph's
      pagination on a large initial snapshot.
    - We stop early on the cap and surface ``has_more=True`` plus the
      nextLink as the returned token so the caller can resume.
    - We stop on a ``deltaLink`` and surface ``has_more=False`` plus the
      deltaLink as the returned token — the caller stores that and uses
      it for the *next* sync round.
    - If Graph returns neither (rare but valid: zero changes plus no
      deltaLink, which means the previous token is still valid), we
      return ``has_more=False`` and ``delta_token=None`` so the caller
      keeps using their old token.

    Returns ``(raw_items, next_token, has_more)``. ``raw_items`` is the
    accumulated list of dicts straight from Graph; the per-resource
    wrapper is responsible for mapping them through ``format_delta_item``
    with the right per-item formatter.
    """
    if page_size < 1:
        page_size = 1
    cap = page_size * PAGE_SIZE_CAP_MULTIPLIER

    # Check the host before minting the token, so a poisoned cursor costs a
    # refusal rather than a request.
    url: str = (
        require_graph_url(delta_token, source="delta_token")
        if delta_token
        else require_graph_url(initial_url, source="initial_url")
    )
    base_headers = {
        "Authorization": f"Bearer {_bearer_token(credential)}",
        "Accept": "application/json",
        # httpx decompresses transparently; delta pages are large and highly
        # compressible, and polling agents fetch them constantly.
        "Accept-Encoding": "gzip",
    }
    if headers:
        base_headers.update(headers)

    collected: list[dict] = []
    next_token: str | None = None
    has_more = False

    async with httpx.AsyncClient(timeout=timeout) as client:
        while True:
            # Retry the delta GET on 429/503 (this raw-httpx path bypasses the
            # SDK's kiota RetryHandler, so it must honor Retry-After itself).
            r = await send_with_retry(client, "GET", url, headers=base_headers)
            r.raise_for_status()
            body = r.json()

            page_items = body.get("value") or []
            collected.extend(page_items)

            next_link = body.get("@odata.nextLink")
            delta_link = body.get("@odata.deltaLink")

            if delta_link:
                # Reached the end of this sync round. The deltaLink is the
                # cursor for the *next* round — validated before we return it so
                # a poisoned link is never stored by the caller and replayed.
                next_token = require_graph_url(delta_link, source="@odata.deltaLink")
                has_more = False
                break

            if next_link:
                next_link = require_graph_url(next_link, source="@odata.nextLink")
                if len(collected) >= cap:
                    # Hit the per-call cap mid-sync. Hand the nextLink back
                    # so the caller resumes from where we stopped.
                    next_token = next_link
                    has_more = True
                    break
                url = next_link
                continue

            # Neither link present — possible on the no-changes path when
            # Graph echoes the old window without a new deltaLink. Caller
            # should reuse their previous token.
            next_token = None
            has_more = False
            break

    return collected, next_token, has_more
