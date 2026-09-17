"""The account-type guard for consumer-only live assertions, tested offline.

``consumer_skip_reason`` decides whether a mailbox can answer a claim about
personal Microsoft accounts. The fixture that wraps it needs a real token, so
the decision lives in a pure function and the interesting cases are pinned here
with nothing live involved.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from tests.conftest import MSA_CONSUMER_TENANT, consumer_skip_reason


@dataclass
class _Record:
    """Just the field the decision reads, shaped like an AuthenticationRecord."""

    tenant_id: str


def test_a_personal_account_can_answer_the_claim():
    assert consumer_skip_reason(_Record(MSA_CONSUMER_TENANT)) is None


def test_a_work_account_is_skipped_rather_than_failed():
    """The case that reached two contributor PRs as a red test.

    A work or school account may legitimately support what the assertion says
    is unsupported, so failing there is a false report about Graph, not a
    finding.
    """
    reason = consumer_skip_reason(_Record("72f988bf-86f1-41af-91ab-2d7cd011db47"))
    assert reason is not None
    assert "Work/school account" in reason


def test_the_skip_names_the_tenant_it_actually_found():
    """A skip that does not say why it skipped reads as coverage.

    That was the whole failure mode here: the test looked like it was asserting
    something about consumer mailboxes and, on the machines where it ran, was
    asserting something else entirely.
    """
    reason = consumer_skip_reason(_Record("72f988bf-86f1-41af-91ab-2d7cd011db47"))
    assert "72f988bf-86f1-41af-91ab-2d7cd011db47" in reason
    assert MSA_CONSUMER_TENANT in reason


def test_no_auth_record_is_skipped_not_assumed_to_be_consumer():
    """Unknown must not default to 'yes, run it'.

    Defaulting to consumer would reintroduce the bug on exactly the hosts that
    cannot tell — the assertion would run and fail for a reason that has nothing
    to do with the behaviour under test.
    """
    reason = consumer_skip_reason(None)
    assert reason is not None
    assert "account type is unknown" in reason


def test_every_consumer_claim_in_the_live_tier_requests_the_guard():
    """The wiring itself, which no offline run would otherwise touch.

    The live tests are deselected by default, so a fixture name that was
    misspelt — or quietly dropped in a later edit — would not fail anything
    until someone ran the write tier on a work account, which is the exact
    situation this guard exists to prevent. So match on the claim rather than
    on a hardcoded test name: any live test that says in its name or docstring
    that it is about personal or consumer accounts has to request the fixture.
    """
    import ast

    live_dir = Path(__file__).parent
    offenders: list[str] = []

    for path in sorted(live_dir.glob("test_live_*.py")):
        tree = ast.parse(path.read_text())
        for node in ast.walk(tree):
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            if not node.name.startswith("test_"):
                continue
            claim = f"{node.name} {ast.get_docstring(node) or ''}".lower()
            if "personal account" not in claim and "consumer mailbox" not in claim:
                continue
            args = {a.arg for a in node.args.args}
            if "consumer_mailbox_only" not in args:
                offenders.append(f"{path.name}::{node.name}")

    assert not offenders, (
        f"These live tests make a claim about personal/consumer accounts but do not "
        f"request `consumer_mailbox_only`, so they will fail on a contributor's "
        f"work account for a reason that has nothing to do with the claim: {offenders}"
    )
