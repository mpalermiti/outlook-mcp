"""Regression tests for scripts/preflight.py.

The preflight script's entire purpose is to catch the v1.7.0-class bug:
a tool ships against a Graph endpoint that's documented as supported
but actually returns 403 / 501 / "not supported" on the target account
type (personal Outlook.com). These tests lock in the classifier so a
future refactor can't silently weaken the release-gate.
"""

from __future__ import annotations

import importlib.util
import pathlib

_PREFLIGHT_PATH = (
    pathlib.Path(__file__).resolve().parent.parent / "scripts" / "preflight.py"
)
_spec = importlib.util.spec_from_file_location("preflight", _PREFLIGHT_PATH)
preflight = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(preflight)


class TestClassify:
    def test_200_is_ok(self):
        assert preflight.classify(200) == "OK"

    def test_204_is_ok(self):
        assert preflight.classify(204) == "OK"

    def test_403_is_fail_v170_regression(self):
        """v1.7.0 regression: /me/mailboxSettings returned 403
        ErrorAccessDenied on personal Microsoft accounts. The preflight
        MUST flag 403 as FAIL so this class of bug blocks the release
        ritual instead of shipping to PyPI / ClawHub / MCP registry."""
        assert preflight.classify(403) == "FAIL"

    def test_501_is_fail(self):
        """Some Graph paths return 501 NotImplemented for unsupported
        account types — same release-blocker family as 403."""
        assert preflight.classify(501) == "FAIL"

    def test_400_is_skip_query_shape(self):
        assert preflight.classify(400) == "SKIP"

    def test_401_is_skip_token_issue(self):
        assert preflight.classify(401) == "SKIP"

    def test_404_is_skip_resource_id(self):
        assert preflight.classify(404) == "SKIP"

    def test_429_is_skip_rate_limit(self):
        assert preflight.classify(429) == "SKIP"

    def test_500_is_skip_transient_server_error(self):
        assert preflight.classify(500) == "SKIP"

    def test_502_is_skip_transient_server_error(self):
        assert preflight.classify(502) == "SKIP"


class TestEndpointsList:
    def test_mailbox_settings_endpoints_not_present(self):
        """v1.7.0 regression guard: /me/mailboxSettings/* tools were
        yanked in 1.7.1 because Graph doesn't support the resource on
        personal accounts. If a future commit re-adds an endpoint under
        that path to ENDPOINTS, this test fails so we revisit the
        decision (see ROADMAP "Investigated and not viable")."""
        for path, _label in preflight.ENDPOINTS:
            assert "mailboxSettings" not in path, (
                f"Endpoint {path!r} is in the preflight list, but "
                "/me/mailboxSettings/* is not supported on personal "
                "accounts (see ROADMAP and CHANGELOG 1.7.1). Re-investigate "
                "before re-adding."
            )

    def test_includes_inference_overrides(self):
        """Phase 3 of v1.7.0 added the override CRUD tools; the preflight
        must continue exercising that endpoint."""
        paths = [p for p, _ in preflight.ENDPOINTS]
        assert any("inferenceClassification/overrides" in p for p in paths)


class _Resp:
    def __init__(self, status_code, body=None):
        self.status_code = status_code
        self._body = body or {}

    def json(self):
        return self._body


class TestTodoTaskFamilies:
    """The To Do checklist/attachment endpoints live under a real task id,
    so their probe rows are provisioned at run time (see
    ``_todo_task_rows``). These tests lock in that the families stay
    covered — the attachment write itself is a POST and stays out of the
    read-only script, but the GETs see the same 403/501 refusal signal."""

    def _wire(self, monkeypatch, responses):
        def fake_get(url, headers=None, timeout=None):
            return responses[url]

        monkeypatch.setattr(preflight.httpx, "get", fake_get)

    def test_checklist_attachments_and_value_rows_are_provisioned(self, monkeypatch):
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=1": _Resp(200, {"value": [{"id": "L1"}]}),
                f"{base}/L1/tasks?$top=1": _Resp(200, {"value": [{"id": "T1"}]}),
                f"{base}/L1/tasks/T1/attachments?$top=1": _Resp(
                    200, {"value": [{"id": "A1"}]}
                ),
            },
        )

        rows = preflight._todo_task_rows({"Authorization": "Bearer x"})

        paths = [p for p, _ in rows]
        assert any("checklistItems" in p for p in paths), paths
        assert any(p.endswith("/attachments?$top=1") for p in paths), paths
        assert any("$value" in p for p in paths), paths

    def test_no_attachment_means_no_value_row_not_a_broken_one(self, monkeypatch):
        """A $value row without a real attachment id would 404 for the wrong
        reason; the row is skipped instead."""
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=1": _Resp(200, {"value": [{"id": "L1"}]}),
                f"{base}/L1/tasks?$top=1": _Resp(200, {"value": [{"id": "T1"}]}),
                f"{base}/L1/tasks/T1/attachments?$top=1": _Resp(200, {"value": []}),
            },
        )

        rows = preflight._todo_task_rows({"Authorization": "Bearer x"})

        paths = [p for p, _ in rows]
        assert any("checklistItems" in p for p in paths)
        assert not any("$value" in p for p in paths)

    def test_no_task_provisioning_degrades_to_no_rows(self, monkeypatch):
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(monkeypatch, {f"{base}?$top=1": _Resp(200, {"value": []})})

        assert preflight._todo_task_rows({"Authorization": "Bearer x"}) == []
