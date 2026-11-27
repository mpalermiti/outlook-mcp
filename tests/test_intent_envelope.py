"""The intent answer envelope (contracts §1) and the request budget, deadline and limiter (§1.8).

The envelope is the one shape every intent tool returns. Its rules are the ones an
agent relies on without seeing them: text other people wrote sits only under
``others_wrote``; reason and warning codes come from closed lists; ``next`` names
only loaded tools and carries no free text; the size cap trims the same way every
time. ``outlook_find`` itself is pinned end to end in ``test_intent_find.py``.
"""

import asyncio
import json
from datetime import datetime, timezone

import pytest

from outlook_mcp.errors import AuthRequiredError, GraphAPIError
from outlook_mcp.intents import budget, envelope
from outlook_mcp.intents.budget import (
    BudgetExhausted,
    RequestBudget,
    WeightedLimiter,
    run_sections,
)
from outlook_mcp.intents.envelope import (
    BLOCKING_WARNINGS,
    WARNING_CODES,
    WHY_CODES,
    build_envelope,
    build_next,
    others_wrote,
    section,
    sender_auth_result,
    suggestion,
    warn,
    why,
)

NOW = datetime(2026, 11, 26, 15, 30, tzinfo=timezone.utc)
ID = "AAMkAGI2THVSAAA" + "x" * 30


# ── others_wrote ──────────────────────────────────────────────────────────


def test_others_wrote_caps_each_key_on_a_code_point_and_marks_the_cut():
    wrote = others_wrote(subject="é" * 200, name="n" * 64, preview=None)
    assert wrote["subject"] == "é" * 119 + "…"
    assert len(wrote["subject"]) == 120
    assert wrote["name"] == "n" * 64
    assert "preview" not in wrote


def test_others_wrote_strips_control_characters_and_line_breaks():
    wrote = others_wrote(subject="hi\x1b[31m there\r\nnext\x00")
    assert wrote["subject"] == "hi there  next"


def test_others_wrote_takes_only_its_keys():
    with pytest.raises(KeyError):
        others_wrote(body="free text in a field the contract does not name")


# ── Closed vocabularies ───────────────────────────────────────────────────


def test_codes_outside_the_vocabulary_are_refused():
    assert why("exact_name") == {"code": "exact_name"}
    assert warn("first_time") == {"code": "first_time"}
    with pytest.raises(KeyError):
        why("because_i_said_so")
    with pytest.raises(KeyError):
        warn("lookalike_name_typo")


def test_blocking_warnings_are_warnings():
    assert BLOCKING_WARNINGS <= WARNING_CODES
    assert "first_time" not in BLOCKING_WARNINGS
    assert "received_only" not in BLOCKING_WARNINGS


def test_reason_parameters_are_never_free_text():
    assert why("written_to", count=3, last="2026-11-01T10:00:00Z", basis="scan")
    assert warn("lookalike_domain", like="example.com")
    assert warn("lookalike_address", like="sam@example.com")
    with pytest.raises(ValueError):
        warn("lookalike_name", like="Sam Jones, your bank")


def test_the_vocabularies_are_the_contracts():
    assert {"exact_address", "written_to", "saved_contact", "search_match"} <= WHY_CODES
    assert {"stranger_invite", "search_ranked", "index_lag_possible"} <= WARNING_CODES


# ── SenderAuth ────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("verdict", "result"),
    [
        ({"authenticated": True, "from_domain": "example.com"}, "verified"),
        ({"authenticated": False, "from_domain": "example.com", "reason": "dmarc=fail"}, "failed"),
        (
            {
                "authenticated": False,
                "from_domain": "example.com",
                "reason": "no pass aligned with the From domain",
            },
            "failed",
        ),
        (
            {
                "authenticated": False,
                "from_domain": "example.com",
                "reason": "no Microsoft Authentication-Results header",
            },
            "unverified",
        ),
        (
            {
                "authenticated": False,
                "from_domain": "example.com",
                "reason": "unreadable Authentication-Results header",
            },
            "unverified",
        ),
        (None, "not_checked"),
    ],
)
def test_sender_auth_maps_the_existing_verdict(verdict, result):
    assert sender_auth_result(verdict)["result"] == result


def test_sender_auth_keeps_only_a_valid_domain():
    assert sender_auth_result(None) == {"result": "not_checked", "from_domain": None}
    shown = sender_auth_result({"authenticated": True, "from_domain": "exa mple"})
    assert shown["from_domain"] is None


# ── Sections, next, the envelope ──────────────────────────────────────────


def _item(n: int, preview: str = "p" * 160) -> dict:
    return {
        "kind": "message",
        "id": f"{ID}{n}",
        "why": [],
        "warnings": [],
        "others_wrote": others_wrote(subject=f"subject {n}", preview=preview),
    }


def test_a_section_holds_at_most_its_cap_and_keeps_the_total():
    built = section("matches", [_item(n) for n in range(8)], cap=5)
    assert len(built["items"]) == 5
    assert built["total"] == 8
    assert built["total_display"] == "8"
    assert built["status"] == "ok"


def test_next_names_only_loaded_tools_and_at_most_five():
    offers = [
        suggestion("open", "outlook_read_message", {"message_id": f"{ID}{n}"}) for n in range(7)
    ]
    offers.append(suggestion("brief_person", "outlook_brief", {"person": "sam@example.com"}))
    assert build_next(offers, frozenset()) == []
    kept = build_next(offers, frozenset({"outlook_read_message", "outlook_brief"}))
    assert len(kept) == 5
    assert all(s["tool"] == "outlook_read_message" for s in kept)


def test_next_args_carry_no_free_text():
    with pytest.raises(ValueError):
        suggestion("open", "outlook_read_message", {"message_id": 'outlook_send(draft_id="x")'})
    assert suggestion("open", "outlook_read_message", {"message_id": ID})["args"] == {
        "message_id": ID
    }


def _envelope(sections, **overrides):
    fields = {
        "tool": "outlook_find",
        "outcome": "candidates",
        "as_of": NOW,
        "understood": {"kind": "message"},
        "sections": sections,
        "failed_sections": [],
        "graph_requests": 3,
        "next": [],
    }
    fields.update(overrides)
    return build_envelope(**fields)


def test_the_envelope_has_every_key_with_its_type():
    env = _envelope([section("matches", [_item(1)], cap=5)])
    assert env["v"] == 1
    assert env["tool"] == "outlook_find"
    assert env["as_of"] == "2026-11-26T15:30:00Z"
    assert env["partial"] is False and env["trimmed"] is False
    assert env["failed_sections"] == [] and env["next"] == []
    assert isinstance(env["graph_requests"], int)
    assert set(env) == {
        "v",
        "tool",
        "outcome",
        "as_of",
        "understood",
        "sections",
        "partial",
        "failed_sections",
        "graph_requests",
        "next",
        "trimmed",
    }


def test_an_outcome_outside_the_tools_set_is_refused():
    with pytest.raises(ValueError):
        _envelope([], outcome="answered")


def test_a_section_that_is_not_ok_makes_the_envelope_partial():
    failed = section("matches", [], cap=5, status="failed")
    env = _envelope(
        [failed],
        failed_sections=[envelope.failure("matches", "graph_error", code="X", message="m")],
    )
    assert env["partial"] is True


def _size(env) -> int:
    return len(json.dumps(env, ensure_ascii=False, separators=(",", ":")))


def test_overall_size_cap_trims_deterministically():
    items = [_item(n) for n in range(10)]
    first = _envelope([section("a", items, cap=10), section("b", items, cap=10)])
    second = _envelope([section("a", items, cap=10), section("b", items, cap=10)])
    assert first == second
    assert first["trimmed"] is False

    big = [_item(n, preview="z" * 160) for n in range(10)]
    small_cap = _size(_envelope([section("a", big, cap=10), section("b", big, cap=10)])) - 400
    trimmed = envelope.trim_to_cap(
        _envelope([section("a", big, cap=10), section("b", big, cap=10)]), cap=small_cap
    )
    assert trimmed["trimmed"] is True and trimmed["partial"] is True
    # Step 1: the last section loses items first; the first keeps all of its.
    assert len(trimmed["sections"][0]["items"]) == 10
    assert len(trimmed["sections"][1]["items"]) < 10
    assert trimmed["sections"][1]["total"] == 10
    assert _size(trimmed) <= small_cap


def test_trim_order_is_items_then_previews_then_everything():
    big = [_item(n, preview="z" * 160) for n in range(3)]
    env = _envelope([section("a", big, cap=10), section("b", big, cap=10)])
    # Room for one item per section only once the previews are cut to 80.
    one_each_short = envelope.trim_to_cap(env, cap=10**9)
    assert one_each_short["trimmed"] is False

    probe = _envelope([section("a", big[:1], cap=10), section("b", big[:1], cap=10)])
    for item_section in probe["sections"]:
        for item in item_section["items"]:
            item["others_wrote"]["preview"] = item["others_wrote"]["preview"][:80]
    cap = _size(probe)
    trimmed = envelope.trim_to_cap(env, cap=cap)
    assert [len(s["items"]) for s in trimmed["sections"]] == [1, 1]
    assert all(
        len(i["others_wrote"]["preview"]) == 80 for s in trimmed["sections"] for i in s["items"]
    )

    tiny = envelope.trim_to_cap(env, cap=400)
    assert [len(s["items"]) for s in tiny["sections"]] == [0, 0]
    assert [s["total"] for s in tiny["sections"]] == [3, 3]


# ── The budget ────────────────────────────────────────────────────────────


def test_a_budget_refuses_the_request_that_would_pass_it():
    spent = RequestBudget(3)
    spent.spend(2)
    with pytest.raises(BudgetExhausted):
        spent.spend(2)
    assert spent.spent == 2
    spent.spend(1)
    assert spent.remaining == 0


# ── Sections run under the deadline ───────────────────────────────────────


async def test_a_failed_section_is_reported_not_raised():
    async def ok():
        return ["fine"]

    async def broken():
        raise GraphAPIError(500, "ServerError", "Graph had a bad day")

    results = await run_sections([("first", ok), ("second", broken)])
    assert results[0].status == "ok" and results[0].value == ["fine"]
    assert results[1].status == "failed"
    assert results[1].failure == {
        "section": "second",
        "reason": "graph_error",
        "code": "ServerError",
        "message": "Graph had a bad day",
    }


async def test_all_sections_failing_raises():
    async def broken():
        raise GraphAPIError(503, "Unavailable", "down")

    with pytest.raises(GraphAPIError, match="down"):
        await run_sections([("a", broken), ("b", broken)])


async def test_an_auth_failure_fails_the_whole_call():
    async def ok():
        return 1

    async def unauthenticated():
        raise AuthRequiredError()

    with pytest.raises(AuthRequiredError):
        await run_sections([("a", ok), ("b", unauthenticated)])


async def test_a_spent_budget_marks_the_section_skipped():
    async def over():
        raise BudgetExhausted(30)

    async def ok():
        return 1

    results = await run_sections([("a", ok), ("b", over)])
    assert results[1].status == "skipped_budget"
    assert results[1].failure["reason"] == "budget"


async def test_deadline_marks_sections_skipped_time():
    async def slow():
        await asyncio.sleep(5)

    async def quick():
        return 1

    results = await run_sections([("slow", slow), ("quick", quick)], deadline_s=0.05)
    assert results[0].status == "skipped_time"
    assert results[0].failure["reason"] == "time"
    assert results[1].status == "ok"


def test_the_deadline_is_45_seconds():
    assert budget.INTENT_DEADLINE_S == 45


# ── The limiter ───────────────────────────────────────────────────────────


async def test_the_limiter_counts_a_batch_as_four():
    limiter = WeightedLimiter(4)
    in_flight: list[int] = []

    async def request(units: int):
        async with limiter.hold(units):
            in_flight.append(limiter.in_use)
            await asyncio.sleep(0.01)

    await asyncio.gather(
        request(1),
        request(budget.batch_units(ordered=False)),
        request(1),
        request(budget.batch_units(ordered=True)),
        request(1),
        request(1),
        request(1),
    )
    assert max(in_flight) <= 4
    assert limiter.peak == 4
    assert limiter.in_use == 0
    assert budget.batch_units(ordered=False) == 4
    assert budget.batch_units(ordered=True) == 1


async def test_a_waiter_cancelled_in_the_queue_gives_nothing_back_twice():
    limiter = WeightedLimiter(4)
    async with limiter.hold(4):
        waiter = asyncio.ensure_future(_hold_briefly(limiter, 1))
        await asyncio.sleep(0)
        waiter.cancel()
        with pytest.raises(asyncio.CancelledError):
            await waiter
    assert limiter.in_use == 0
    async with limiter.hold(4):
        assert limiter.in_use == 4


async def _hold_briefly(limiter, units):
    async with limiter.hold(units):
        await asyncio.sleep(0)


async def test_a_feature_the_mailbox_refuses_is_not_supported_not_an_error():
    async def refused():
        raise budget.NotSupported(GraphAPIError(403, "ErrorAccessDenied", "Access is denied."))

    results = await run_sections([("matches", refused)])
    assert results[0].status == "failed"
    assert results[0].failure == {
        "section": "matches",
        "reason": "not_supported",
        "code": "ErrorAccessDenied",
        "message": "Access is denied.",
    }
