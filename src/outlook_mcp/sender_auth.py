"""Did Microsoft authenticate the domain in a message's From line? (#130)

The From line is whatever the sender wrote. Microsoft checks SPF, DKIM and DMARC
when a message arrives and stamps its verdict in an ``Authentication-Results``
header (RFC 8601) above every header the sender wrote. Graph lists
``internetMessageHeaders`` topmost first, so the first ``Authentication-Results``
is Microsoft's, and anything below it may be the sender's own text. This reads
that one header and nothing else.

Pure: no Graph, no I/O. ``tools/mail_read.py`` hands it the headers and the From
address of a message read with ``include_sender_auth=True``.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

# Microsoft's inbound stamp usually has no authserv-id at all, which RFC 8601
# requires: it starts directly with `spf=`. Some paths name one of these.
_MICROSOFT_AUTHSERV_ID = "mx.microsoft.com"
_MICROSOFT_AUTHSERV_SUFFIX = ".protection.outlook.com"

# Microsoft documents `dmarc` as pass | fail | bestguesspass | none and `compauth`
# as pass | fail | softpass | none. `softfail` is SPF's word, refused here too in
# case a result ever uses it.
_FAILED = frozenset({"fail", "softfail"})

_METHOD = re.compile(r"^[a-z][a-z0-9_-]*(?:/[0-9]+)?$")
_RESULT = re.compile(r"^[a-z][a-z0-9_-]*$")
# `ptype.property`, plus `reason` and Microsoft's own `action` on the dmarc clause.
_PROPERTY = re.compile(r"^[a-z][a-z0-9_.-]*$")
_KNOWN_RESULTS = frozenset(
    {
        "none",
        "pass",
        "fail",
        "softfail",
        "neutral",
        "policy",
        "temperror",
        "permerror",
        "bestguesspass",
        "softpass",
    }
)

NO_FROM = "no From address"
NO_HEADER = "no Microsoft Authentication-Results header"
UNREADABLE = "unreadable Authentication-Results header"
NOT_ALIGNED = "no pass aligned with the From domain"


@dataclass(frozen=True)
class _Result:
    method: str
    result: str
    props: dict[str, tuple[str, ...]]


def sender_authentication(headers: Iterable[tuple[str, str]], from_address: str) -> dict:
    """Microsoft's verdict on the From domain, as ``sender_authentication``.

    ``headers`` are ``(name, value)`` pairs in Graph's order, topmost first.
    ``authenticated`` is true only for a pass aligned with the From domain in
    Microsoft's own header: ``dmarc=pass`` for it, ``dkim=pass`` signed by it or
    a subdomain of it, or ``compauth=pass`` where every ``header.from`` is it. A
    failing ``dmarc`` or ``compauth`` refuses any pass, and a header that states
    either twice is refused whole: Microsoft states each once, so a second is
    text a sender got into the header.
    """
    _, at, domain = (from_address or "").strip().lower().rpartition("@")
    if not at or not domain:
        return _verdict(None, NO_FROM)

    results, problem = _microsoft_results(headers)
    if problem:
        return _verdict(domain, problem)

    for method in ("dmarc", "compauth"):
        if sum(r.method == method for r in results) > 1:
            return _verdict(domain, f"{method} stated twice")

    dmarc = _stated(results, "dmarc")
    compauth = _stated(results, "compauth")
    found = {"dkim": _dkim(results, domain), "dmarc": dmarc, "compauth": compauth}

    for method, verdict in (("dmarc", dmarc), ("compauth", compauth)):
        if verdict in _FAILED:
            return _verdict(domain, f"{method}={verdict}", **found)

    # `dmarc` and `compauth` judge the From domain as Microsoft read it, which it
    # writes once, as `header.from` in the dmarc clause; compauth names no domain
    # of its own. Both bind to the From line through it: every `header.from` the
    # header states must be the From domain, and there must be one.
    header_from = [v for r in results for v in r.props.get("header.from", ())]
    from_aligned = bool(header_from) and all(v == domain for v in header_from)

    if dmarc == "pass" and from_aligned:
        return _verdict(domain, "dmarc=pass", authenticated=True, **found)
    if found["dkim"] == "pass":
        return _verdict(domain, "aligned dkim=pass", authenticated=True, **found)
    if compauth == "pass" and from_aligned:
        return _verdict(domain, "compauth=pass", authenticated=True, **found)
    return _verdict(domain, NOT_ALIGNED, **found)


def _verdict(
    domain: str | None,
    reason: str,
    *,
    authenticated: bool = False,
    dkim: str | None = None,
    dmarc: str | None = None,
    compauth: str | None = None,
) -> dict:
    return {
        "authenticated": authenticated,
        "from_domain": domain,
        "dkim": dkim,
        "dmarc": dmarc,
        "compauth": compauth,
        "reason": reason,
    }


def _stated(results: list[_Result], method: str) -> str | None:
    return next((r.result for r in results if r.method == method), None)


def _signer_aligned(value: str, domain: str) -> bool:
    # `header.d` is a domain, `header.i` an identity (`[local]@domain`). The
    # signer must be the From domain or a true subdomain of it: `xcontoso.com`
    # and `contoso.com.example.net` are registrable by anyone.
    signer = value.rpartition("@")[2]
    return signer == domain or signer.endswith("." + domain)


def _dkim(results: list[_Result], domain: str) -> str | None:
    """The DKIM result for the From domain's own signature, preferring a pass.

    A pass and a domain must come from one clause: ``dkim=pass header.d=other``
    beside ``dkim=fail header.d=<From domain>`` is that domain's failure.
    """
    own = [
        r.result
        for r in results
        if r.method == "dkim"
        and any(
            _signer_aligned(v, domain)
            for key in ("header.d", "header.i")
            for v in r.props.get(key, ())
        )
    ]
    return "pass" if "pass" in own else next(iter(own), None)


def _microsoft_results(headers: Iterable[tuple[str, str]]) -> tuple[list[_Result], str | None]:
    """The method results of the topmost Authentication-Results, or why there are none.

    Every clause after the authserv-id must be a well-formed method result:
    Microsoft's own header always is, and one that is not was not written by
    Microsoft alone.
    """
    value = next(
        (v for n, v in headers if n.strip().casefold() == "authentication-results"),
        None,
    )
    if value is None:
        return [], NO_HEADER
    clauses = _tokenize(value.lower())
    if not clauses:
        return [], UNREADABLE
    if _parse_result(clauses[0]) is None:
        # authserv-id [version]
        if len(clauses[0]) > 2:
            return [], UNREADABLE
        authserv_id = clauses[0][0]
        if not (
            authserv_id == _MICROSOFT_AUTHSERV_ID
            or authserv_id.endswith(_MICROSOFT_AUTHSERV_SUFFIX)
        ):
            return [], NO_HEADER
        clauses = clauses[1:]
    results = [_parse_result(c) for c in clauses]
    if any(r is None for r in results):
        return [], UNREADABLE
    return results, None


def _tokenize(value: str) -> list[list[str]] | None:
    """Split an Authentication-Results value into `;`-separated clauses of tokens.

    Comments (nested, with `\\` escapes) are dropped. A quoted string is part of
    its token, quotes removed, so `;`, `=` and spaces inside it are plain text.
    `=` is a token of its own. None for an unterminated comment or quoted string.
    """
    clauses: list[list[str]] = [[]]
    word: list[str] | None = None  # None: no token open
    quoted = False
    depth = 0
    chars = iter(value)
    for ch in chars:
        if quoted:
            if ch == "\\":
                word.append(next(chars, ""))
            elif ch == '"':
                quoted = False
            else:
                word.append(ch)
            continue
        if depth:
            if ch == "\\":
                next(chars, "")
            elif ch in "()":
                depth += 1 if ch == "(" else -1
            continue
        if ch == '"':
            # Method and result are keywords, never quoted strings. Property
            # values may be quoted; preserve the existing handling for those.
            if len(clauses[-1]) < 3:
                return None
            if word is None:
                word = []
            quoted = True
            continue
        if ch not in "();=" and not ch.isspace():
            if word is None:
                word = []
            word.append(ch)
            continue
        if word is not None:
            clauses[-1].append("".join(word))
            word = None
        if ch == "(":
            depth = 1
        elif ch == ";":
            clauses.append([])
        elif ch == "=":
            clauses[-1].append("=")
    if quoted or depth:
        return None
    if word is not None:
        clauses[-1].append("".join(word))
    return [c for c in clauses if c]


def _parse_result(tokens: list[str]) -> _Result | None:
    """`method=result [property=value ...]`, else None."""
    if len(tokens) < 3 or tokens[1] != "=":
        return None
    method, result = tokens[0], tokens[2]
    if not (_METHOD.fullmatch(method) and _RESULT.fullmatch(result)):
        return None
    if result not in _KNOWN_RESULTS:
        result = "unknown"
    props: dict[str, list[str]] = {}
    index = 3
    while index < len(tokens):
        if not (
            index + 1 < len(tokens)
            and tokens[index + 1] == "="
            and _PROPERTY.fullmatch(tokens[index])
        ):
            return None
        # A property with an empty value (`smtp.mailfrom=;`) is two tokens.
        has_value = (
            index + 2 < len(tokens)
            and tokens[index + 2] != "="
            and not (index + 3 < len(tokens) and tokens[index + 3] == "=")
        )
        props.setdefault(tokens[index], []).append(tokens[index + 2] if has_value else "")
        index += 3 if has_value else 2
    return _Result(method.split("/", 1)[0], result, {k: tuple(v) for k, v in props.items()})
