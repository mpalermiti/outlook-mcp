"""The intent answer envelope (contracts §1): builders, vocabularies and the size cap.

Every intent tool returns one envelope, built here as plain dicts and never
through a separate validation model (the house rule against unwired I/O
layers). The rules this module holds for every tool:

- **Other people's text.** Every free-text value that came from Graph sits under
  an item's ``others_wrote``, capped per key and sanitized (§1.3). Nothing else
  carries free text, so an agent can tell information from instructions.
- **Closed vocabularies.** ``why`` and ``warnings`` codes come from fixed lists,
  and their parameters are ints, bools, times, ids, addresses or domains (§1.6).
- **``next``.** Suggested calls name only loaded tools, and their arguments are
  ids, addresses, enums or bools, never text (§1.7).
- **Size.** The serialized result is capped and trimmed the same way every time
  (§1.5).
"""

from __future__ import annotations

import copy
import json
import re
from collections.abc import Iterable
from datetime import datetime, timezone
from typing import Any

from outlook_mcp.sender_auth import NO_FROM, NO_HEADER, UNREADABLE
from outlook_mcp.validation import sanitize_output, validate_email, validate_graph_id

ENVELOPE_VERSION = 1

# The serialized cap, about 6,000 proxy tokens. `plan`, `receipt` and `cursor`
# have their own limits and are not counted here.
SIZE_CAP = 24_000
_UNCAPPED_KEYS = ("plan", "receipt", "cursor")

# `outcome` per tool (§1.2). Only the tools built so far.
# `incomplete` for find is a deviation from the contracts' four outcomes: a search that
# did not finish, or was capped, and matched nothing in what it read. It is never
# `nothing_found`, which only a finished search may say (review S4).
# The read tools answer `answered`, and `incomplete` when a section was skipped for
# budget or time, or failed: the same deviation, so an answer that could not finish
# never reads as an empty one.
OUTCOMES = {
    "outlook_find": frozenset({"exact", "single", "candidates", "nothing_found", "incomplete"}),
    "outlook_agenda": frozenset({"answered", "incomplete"}),
    "outlook_open_loops": frozenset({"answered", "incomplete", "candidates"}),
    "outlook_catch_up": frozenset({"answered", "incomplete"}),
    "outlook_brief": frozenset({"answered", "incomplete", "candidates"}),
}

SECTION_STATUSES = frozenset({"ok", "failed", "skipped_budget", "skipped_time"})
FAILURE_REASONS = frozenset({"graph_error", "budget", "time", "not_supported"})

# §1.3: the keys allowed under `others_wrote`, each with its cap in code points.
OTHERS_WROTE_CAPS = {
    "subject": 120,
    "preview": 160,
    "name": 64,
    "location": 80,
    "title": 120,
    "file_name": 120,
    "body_preview": 160,
    "organizer_name": 64,
    "address_text": 254,
    # A saved contact's details (brief's `who`, §12.5: "contact details go under
    # others_wrote"): its owner typed them, but anyone the agent saved could have.
    "company": 80,
    "job_title": 80,
    "phones": 120,
    "birthday": 10,
}

# §1.6: exhaustive. A code outside these lists is a bug, refused where it is built.
WHY_CODES = frozenset(
    {
        # People
        "exact_address",
        "exact_name",
        "first_name",
        "name_prefix",
        "name_contains",
        "written_to",
        "saved_contact",
        # Mail
        "you_are_on_to",
        "you_are_on_cc",
        "only_recipient",
        "sender_written_to",
        "sender_saved_contact",
        "no_reply_from_you",
        "reply_from_them",
        "you_replied",
        "flag_due",
        "flag_overdue",
        "high_importance",
        "has_attachments",
        "search_match",
        "filter_match",
        "you_wrote_in_thread",
        # Events
        "organizer_you",
        "organizer_written_to",
        "not_responded",
        "overlaps",
        "today",
        "changed_since_cursor",
        "cancelled",
        "meeting_with_them",
        # Tasks
        "due_today",
        "due_in_window",
        "overdue",
        "linked_to_message",
    }
)
WARNING_CODES = frozenset(
    {
        # People
        "lookalike_name",
        "lookalike_domain",
        "lookalike_address",
        "idn_domain",
        "invalid_address",
        "mixed_script_name",
        "name_contains_other_address",
        "first_time",
        "received_only",
        # Mail
        "sender_unverified",
        "sender_failed",
        "bulk_sender",
        "auto_generated",
        # Events
        "stranger_invite",
        "shows_free",
        "declined",
        "all_day",
        "no_new_time_proposals",
        # Search quality
        "search_ranked",
        "index_lag_possible",
        "may_repeat",
    }
)

# §2.2: a candidate carrying any of these is never `single`, and is left out of
# every bulk set.
BLOCKING_WARNINGS = frozenset(
    {
        "lookalike_name",
        "lookalike_domain",
        "lookalike_address",
        "idn_domain",
        "mixed_script_name",
        "name_contains_other_address",
        "invalid_address",
        "sender_failed",
        "stranger_invite",
    }
)

_UTC_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
_DOMAIN_RE = re.compile(r"^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+$")
_PERSON_ID_PREFIX = "person:"

# Verdicts that say nothing about the sender either way: no header to read.
_SAYS_NOTHING = frozenset({NO_FROM, NO_HEADER, UNREADABLE})


def utc(moment: datetime) -> str:
    """An aware datetime as the envelope's UTC form, ``YYYY-MM-DDTHH:MM:SSZ``."""
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def is_domain(value: str) -> bool:
    """An ASCII domain name: lowercase labels of letters, digits and inner hyphens."""
    return bool(_DOMAIN_RE.fullmatch(value))


def is_safe_value(value: Any) -> bool:
    """A value allowed outside ``others_wrote``: never free text (§1.3, §1.7).

    Ints and bools; a UTC time; a Graph id (which covers the enum words, since
    they share its alphabet); an address; a domain; a person id.
    """
    if isinstance(value, bool | int) or value is None:
        return True
    if isinstance(value, list | tuple):
        # A list of safe values: `answer`'s `event_ids` (§1.7).
        return all(isinstance(v, str) and is_safe_value(v) for v in value)
    if not isinstance(value, str):
        return False
    if _UTC_RE.fullmatch(value) or is_domain(value):
        return True
    if value.startswith(_PERSON_ID_PREFIX):
        value = value[len(_PERSON_ID_PREFIX) :]
    for check in (validate_graph_id, validate_email):
        try:
            check(value)
            return True
        except ValueError:
            continue
    return False


def cut(text: str, limit: int) -> str:
    """``text`` cut to ``limit`` code points, ending in ``…`` when anything was cut."""
    return text if len(text) <= limit else text[: limit - 1] + "…"


def others_wrote(**fields: str | None) -> dict[str, str]:
    """Other people's text for one item: sanitized, one line, capped per key.

    ``None`` values are left out. A key the contract does not name raises
    ``KeyError``: a new free-text field needs its cap decided first.
    """
    wrote: dict[str, str] = {}
    for key, value in fields.items():
        limit = OTHERS_WROTE_CAPS[key]
        if value is None:
            continue
        wrote[key] = cut(sanitize_output(str(value)), limit)
    return wrote


def _code(vocabulary: frozenset[str], code: str, params: dict[str, Any]) -> dict:
    if code not in vocabulary:
        raise KeyError(f"{code!r} is not in the closed vocabulary")
    for name, value in params.items():
        if not is_safe_value(value):
            raise ValueError(f"{code}.{name} must be an int, bool, time, id, address or domain")
    return {"code": code, **params}


def why(code: str, **params: Any) -> dict:
    """A positive reason, from ``WHY_CODES``."""
    return _code(WHY_CODES, code, params)


def warn(code: str, **params: Any) -> dict:
    """A risk flag, from ``WARNING_CODES``."""
    return _code(WARNING_CODES, code, params)


def has_blocking_warning(item: dict) -> bool:
    return any(w["code"] in BLOCKING_WARNINGS for w in item.get("warnings", []))


def sender_auth_result(verdict: dict | None) -> dict:
    """``SenderAuth`` (§1.4) from ``sender_auth.sender_authentication``'s verdict.

    ``None`` means the headers were not read. A failing method, or no pass aligned
    with the From domain, is ``failed``; a missing or unreadable header is
    ``unverified``, which says nothing either way.
    """
    if verdict is None:
        return {"result": "not_checked", "from_domain": None}
    domain = verdict.get("from_domain")
    domain = domain if isinstance(domain, str) and is_domain(domain) else None
    if verdict.get("authenticated") is True:
        result = "verified"
    elif verdict.get("reason") in _SAYS_NOTHING:
        result = "unverified"
    else:
        # A failing method, no aligned pass, or a header that states a method twice
        # (text a sender got into it): each is a reason to distrust the From line.
        result = "failed"
    return {"result": result, "from_domain": domain}


def section(
    name: str,
    items: list[dict],
    *,
    cap: int,
    status: str = "ok",
    total: int | None = None,
    total_exact: bool = True,
    total_display: str | None = None,
) -> dict:
    """One answer section: at most ``cap`` items, with the count before the cap."""
    if status not in SECTION_STATUSES:
        raise ValueError(f"unknown section status {status!r}")
    counted = len(items) if total is None and status == "ok" else total
    if total_display is None:
        total_display = "" if counted is None else str(counted)
    return {
        "name": name,
        "status": status,
        "items": list(items[:cap]),
        "total": counted,
        "total_exact": total_exact,
        "total_display": total_display,
        "cap": cap,
    }


def failure(name: str, reason: str, *, code: str | None = None, message: str = "") -> dict:
    """A section's ``Failure``: why it has no answer, in Graph's words (≤ 300 characters)."""
    if reason not in FAILURE_REASONS:
        raise ValueError(f"unknown failure reason {reason!r}")
    return {
        "section": name,
        "reason": reason,
        "code": code,
        "message": cut(sanitize_output(message), 300),
    }


def suggestion(template: str, tool: str, args: dict[str, Any], needs: Iterable[str] = ()) -> dict:
    """One suggested next call. Every argument must be an id, address, enum or bool."""
    for name, value in args.items():
        if not is_safe_value(value):
            raise ValueError(f"next.{template}.{name} must not carry free text")
    return {"template": template, "tool": tool, "args": dict(args), "needs": list(needs)}


def build_next(suggestions: Iterable[dict], loaded_tools: frozenset[str]) -> list[dict]:
    """At most 5 suggestions, in order, dropping any whose tool this process doesn't load."""
    return [s for s in suggestions if s["tool"] in loaded_tools][:5]


def build_envelope(
    *,
    tool: str,
    outcome: str,
    as_of: datetime,
    understood: dict,
    sections: list[dict],
    failed_sections: list[dict],
    graph_requests: int,
    next: list[dict],
    capped: bool = False,
) -> dict:
    """The envelope (§1.1), trimmed to the size cap.

    ``capped``: a list was cut at its cap, so the answer may be missing matches. It
    sets ``partial`` too, beyond the contracts' rule (failed sections and trims).
    """
    if outcome not in OUTCOMES[tool]:
        raise ValueError(f"{tool} cannot answer {outcome!r}")
    env = {
        "v": ENVELOPE_VERSION,
        "tool": tool,
        "outcome": outcome,
        "as_of": utc(as_of),
        "understood": understood,
        "sections": sections,
        "partial": capped or bool(failed_sections) or any(s["status"] != "ok" for s in sections),
        "failed_sections": failed_sections,
        "graph_requests": graph_requests,
        "next": next,
        "trimmed": False,
    }
    return trim_to_cap(env)


def _size(env: dict) -> int:
    counted = {k: v for k, v in env.items() if k not in _UNCAPPED_KEYS}
    return len(json.dumps(counted, ensure_ascii=False, separators=(",", ":")))


def _all_others_wrote(node: Any) -> Iterable[dict]:
    """Every ``others_wrote`` dict under ``node``, nested items included."""
    if isinstance(node, dict):
        for key, value in node.items():
            if key == "others_wrote" and isinstance(value, dict):
                yield value
            else:
                yield from _all_others_wrote(value)
    elif isinstance(node, list):
        for value in node:
            yield from _all_others_wrote(value)


def trim_to_cap(env: dict, cap: int = SIZE_CAP) -> dict:
    """``env`` trimmed to ``cap`` serialized characters, stopping as soon as it fits (§1.5).

    1. Remove items from the last section, then the one before, keeping one each.
    2. Cut every preview to 80 characters.
    3. Drop previews entirely.
    4. Remove items down to none, again from the last section first.
    ``total`` stays as it was; any trim sets ``trimmed`` and ``partial``.
    """
    if _size(env) <= cap:
        return env
    env = copy.deepcopy(env)
    env["trimmed"] = True
    env["partial"] = True

    def drop_items(keep: int) -> bool:
        for sec in reversed(env["sections"]):
            while len(sec["items"]) > keep:
                sec["items"].pop()
                if _size(env) <= cap:
                    return True
        return False

    def previews(edit) -> bool:
        for wrote in _all_others_wrote(env["sections"]):
            for key in ("preview", "body_preview"):
                if key in wrote:
                    edit(wrote, key)
        return _size(env) <= cap

    def shorten(wrote: dict, key: str) -> None:
        wrote[key] = cut(wrote[key], 80)

    def remove(wrote: dict, key: str) -> None:
        del wrote[key]

    if drop_items(keep=1) or previews(shorten) or previews(remove):
        return env
    drop_items(keep=0)
    return env
