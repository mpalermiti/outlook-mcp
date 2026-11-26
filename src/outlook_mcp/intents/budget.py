"""Request budget, deadline and the weighted limiter for intent tools (contracts §1.8).

Three limits keep an intent call bounded:

- **Budget.** Each call gets a ``RequestBudget`` sized to its tool's ceiling.
  Every counted Graph request spends from it *before* it goes out: one per SDK
  call or raw page, one per ``$batch`` sub-request. A spend past the limit raises
  ``BudgetExhausted``, and the section that hit it is reported as skipped. kiota's
  own retries are not counted.
- **Deadline.** Sections run under ``INTENT_DEADLINE_S``; MCP clients commonly
  give up at 60 seconds. A section still running is cancelled and reported.
- **Limiter.** Outlook allows four concurrent requests per app and mailbox, and a
  ``$batch`` already sends four of its sub-requests at a time, so one unordered
  batch fills every slot. ``LIMITER`` holds 4 units for the process.
"""

from __future__ import annotations

import asyncio
import contextlib
from collections import deque
from collections.abc import AsyncIterator, Awaitable, Callable
from dataclasses import dataclass
from typing import Any

from outlook_mcp.errors import GraphAPIError

# Read and plan tools' deadline, in seconds.
INTENT_DEADLINE_S = 45

# Counted-request ceilings per tool (§1.8). Only the tools built so far.
CEILINGS = {
    "outlook_find": 30,
    "outlook_agenda": 30,
    "outlook_open_loops": 40,
    "outlook_catch_up": 40,
}

# Units on the limiter: one request, an unordered `$batch` (all four slots), or a
# `$batch` chained with dependsOn, whose sub-requests Graph sends one at a time.
LIMITER_UNITS = 4


def batch_units(*, ordered: bool) -> int:
    """Limiter units a ``$batch`` holds: every slot unless it is chained."""
    return 1 if ordered else LIMITER_UNITS


class BudgetExhausted(Exception):  # noqa: N818 - the contracts' name; never reaches a client
    """A counted request would pass the call's budget. Internal: never reaches a client.

    ``run_sections`` turns it into a ``skipped_budget`` section; anything that
    spends outside a section driver catches it itself.
    """

    def __init__(self, limit: int):
        super().__init__(f"request budget of {limit} spent")
        self.limit = limit


class NotSupported(Exception):  # noqa: N818 - a section outcome, not a failure
    """A section the mailbox refuses as a feature, such as inbox rules on some personal
    accounts (a 403). Reported as ``not_supported``, never raised to the client."""

    def __init__(self, error: GraphAPIError):
        super().__init__(error.message)
        self.error = error


class RequestBudget:
    """The counted Graph requests one intent call may make."""

    def __init__(self, limit: int):
        self.limit = limit
        self.spent = 0

    @property
    def remaining(self) -> int:
        return self.limit - self.spent

    def spend(self, n: int = 1) -> None:
        """Count ``n`` requests about to go out, or refuse all of them."""
        if self.spent + n > self.limit:
            raise BudgetExhausted(self.limit)
        self.spent += n


class WeightedLimiter:
    """At most ``units`` request units in flight, granted first come, first served.

    A request waits behind every earlier waiter, so a batch that needs all four
    units is not starved by single requests arriving after it.
    """

    def __init__(self, units: int = LIMITER_UNITS):
        self.units = units
        self.in_use = 0
        self.peak = 0
        self._waiters: deque[tuple[int, asyncio.Future]] = deque()

    @contextlib.asynccontextmanager
    async def hold(self, units: int) -> AsyncIterator[None]:
        """Hold ``units`` (at most all of them) for the duration of the block."""
        units = min(units, self.units)
        if self._waiters or self.in_use + units > self.units:
            granted = asyncio.get_running_loop().create_future()
            self._waiters.append((units, granted))
            try:
                await granted
            except BaseException:
                if granted.done() and not granted.cancelled():
                    # Granted, then cancelled before it could run: give the units back.
                    self._release(units)
                else:
                    self._wake()
                raise
        else:
            self.in_use += units
        self.peak = max(self.peak, self.in_use)
        try:
            yield
        finally:
            self._release(units)

    def _release(self, units: int) -> None:
        self.in_use -= units
        self._wake()

    def _wake(self) -> None:
        while self._waiters:
            units, granted = self._waiters[0]
            if granted.done():  # cancelled while queued
                self._waiters.popleft()
                continue
            if self.in_use + units > self.units:
                return
            self._waiters.popleft()
            self.in_use += units
            granted.set_result(None)


# Process-wide: every intent request goes through it. Short-lived action state,
# not mailbox data (the "no local caching" carve-out in contracts §0).
LIMITER = WeightedLimiter()


@dataclass
class SectionResult:
    """One section's outcome: its value when ``status`` is ``ok``, else a ``Failure``."""

    name: str
    status: str
    value: Any = None
    failure: dict | None = None


def _failure(name: str, reason: str, code: str | None, message: str) -> dict:
    from outlook_mcp.intents.envelope import failure

    return failure(name, reason, code=code, message=message)


async def run_sections(
    sections: list[tuple[str, Callable[[], Awaitable[Any]]]],
    *,
    deadline_s: float | None = None,
) -> list[SectionResult]:
    """Run each section concurrently under the deadline, and report rather than raise.

    Error routing (§1.8): a ``GraphAPIError`` inside a section becomes a
    ``Failure``; a spent budget or the deadline marks the section skipped; any
    other error (sign-in, read-only, permission, bad input) fails the whole call.
    If every section failed with a ``GraphAPIError``, the first is raised.
    """
    limit = INTENT_DEADLINE_S if deadline_s is None else deadline_s
    tasks = [asyncio.ensure_future(run()) for _, run in sections]
    done, pending = await asyncio.wait(tasks, timeout=limit)
    for task in pending:
        task.cancel()
    if pending:
        await asyncio.gather(*pending, return_exceptions=True)

    results: list[SectionResult] = []
    graph_errors: list[GraphAPIError] = []
    for (name, _), task in zip(sections, tasks, strict=True):
        if task in pending:
            results.append(
                SectionResult(
                    name,
                    "skipped_time",
                    failure=_failure(name, "time", None, f"Stopped at the {limit:g}-second limit."),
                )
            )
            continue
        error = task.exception()
        if error is None:
            results.append(SectionResult(name, "ok", value=task.result()))
        elif isinstance(error, BudgetExhausted):
            results.append(
                SectionResult(
                    name,
                    "skipped_budget",
                    failure=_failure(
                        name, "budget", None, f"Stopped at the {error.limit}-request limit."
                    ),
                )
            )
        elif isinstance(error, NotSupported):
            results.append(
                SectionResult(
                    name,
                    "failed",
                    failure=_failure(
                        name, "not_supported", error.error.error_code, error.error.message
                    ),
                )
            )
        elif isinstance(error, GraphAPIError):
            graph_errors.append(error)
            results.append(
                SectionResult(
                    name,
                    "failed",
                    failure=_failure(name, "graph_error", error.error_code, error.message),
                )
            )
        else:
            for other in tasks:
                if other.done() and not other.cancelled():
                    other.exception()  # retrieved, so asyncio does not log it
            raise error
    if sections and len(graph_errors) == len(sections):
        raise graph_errors[0]
    return results
