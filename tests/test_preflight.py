"""Regression tests for scripts/preflight.py.

The preflight script's entire purpose is to catch the v1.7.0-class bug:
a tool ships against a Graph endpoint that's documented as supported
but actually returns 403 / 501 / "not supported" on the target account
type (personal Outlook.com). These tests lock in the classifier so a
future refactor can't silently weaken the release-gate.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import pathlib
from urllib.parse import urljoin

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
    read-only script, but the GETs see the same 403/501 refusal signal —
    and that anything the run could NOT probe is reported rather than
    silently counted as checked."""

    def _wire(self, monkeypatch, responses):
        def fake_get(url, headers=None, timeout=None):
            if url not in responses:
                raise AssertionError(f"unexpected GET {url}")
            return responses[url]

        monkeypatch.setattr(preflight.httpx, "get", fake_get)

    def _lists(self, value):
        return _Resp(200, {"value": value})

    def test_provisioning_targets_the_default_list_not_the_first_row(self, monkeypatch):
        """?$top=1 on /me/todo/lists grabs whichever list Graph happens to
        return first — the tools resolve defaultList+isOwner (falling back
        to the first list), and the preflight must harvest from the list a
        tool call would actually hit."""
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=100": self._lists(
                    [
                        {"id": "SHARED", "isOwner": False, "wellknownListName": "none"},
                        {"id": "L1", "isOwner": True, "wellknownListName": "defaultList"},
                    ]
                ),
                f"{base}/L1/tasks?$top=10": self._lists([{"id": "T1"}]),
                f"{base}/L1/tasks/T1/attachments?$top=1": self._lists([]),
            },
        )

        rows, not_probed, failures = preflight._todo_task_rows(
            {"Authorization": "Bearer x"}
        )

        assert failures == []
        paths = [p for p, _ in rows]
        assert any("L1/tasks/T1/checklistItems" in p for p in paths), paths
        assert not any("SHARED" in p for p in paths), paths
        # No attachment anywhere: entity GET + $value are named as NOT probed
        # rather than quietly missing.
        assert len(not_probed) == 1
        assert "entity GET" in not_probed[0]

    def test_first_list_is_the_fallback_when_no_default_list(self, monkeypatch):
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=100": self._lists(
                    [{"id": "ONLY", "isOwner": True, "wellknownListName": "none"}]
                ),
                f"{base}/ONLY/tasks?$top=10": self._lists([{"id": "T1"}]),
                f"{base}/ONLY/tasks/T1/attachments?$top=1": self._lists([]),
            },
        )

        rows, not_probed, _ = preflight._todo_task_rows({"Authorization": "Bearer x"})

        assert any("/ONLY/tasks/T1/checklistItems" in p for p, _ in rows)
        assert not_probed  # no attachment to probe the entity GET with

    def test_entity_get_is_probed_with_content_bytes_in_the_body(self, monkeypatch, capsys):
        """The download tool reads contentBytes off the entity — a 200 whose
        body lacks it is the response shape the tool refuses, so the probe
        looks at the body, not just the status."""
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=100": self._lists(
                    [{"id": "L1", "isOwner": True, "wellknownListName": "defaultList"}]
                ),
                f"{base}/L1/tasks?$top=10": self._lists([{"id": "T1"}]),
                f"{base}/L1/tasks/T1/attachments?$top=1": self._lists([{"id": "A1"}]),
                f"{base}/L1/tasks/T1/attachments/A1": _Resp(
                    200, {"id": "A1", "size": 5, "contentBytes": "aGVsbG8="}
                ),
            },
        )

        rows, not_probed, failures = preflight._todo_task_rows(
            {"Authorization": "Bearer x"}
        )

        assert not_probed == []
        assert failures == []
        # $value stays as the extra legacy row alongside the entity probe.
        assert any("$value" in p for p, _ in rows)
        out = capsys.readouterr().out
        assert "To Do attachment entity" in out
        assert "contentBytes=yes" in out

    def test_entity_get_without_content_bytes_is_flagged_not_passed(self, monkeypatch, capsys):
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=100": self._lists(
                    [{"id": "L1", "isOwner": True, "wellknownListName": "defaultList"}]
                ),
                f"{base}/L1/tasks?$top=10": self._lists([{"id": "T1"}]),
                f"{base}/L1/tasks/T1/attachments?$top=1": self._lists([{"id": "A1"}]),
                f"{base}/L1/tasks/T1/attachments/A1": _Resp(
                    200, {"id": "A1", "size": 1234}
                ),
            },
        )

        _, not_probed, failures = preflight._todo_task_rows({"Authorization": "Bearer x"})

        assert failures == []
        out = capsys.readouterr().out
        assert "contentBytes=NO" in out
        assert "NOTE" in out

    def test_attachment_on_a_later_task_widens_the_probe(self, monkeypatch):
        """Tasks 1-5 are checked for an attachment, so the entity probe
        survives an attachment-free first task."""
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(
            monkeypatch,
            {
                f"{base}?$top=100": self._lists(
                    [{"id": "L1", "isOwner": True, "wellknownListName": "defaultList"}]
                ),
                f"{base}/L1/tasks?$top=10": self._lists(
                    [{"id": "T1"}, {"id": "T2"}, {"id": "T3"}]
                ),
                f"{base}/L1/tasks/T1/attachments?$top=1": self._lists([]),
                f"{base}/L1/tasks/T2/attachments?$top=1": self._lists([{"id": "A2"}]),
                f"{base}/L1/tasks/T3/attachments?$top=1": self._lists([]),
                f"{base}/L1/tasks/T2/attachments/A2": _Resp(
                    200, {"id": "A2", "size": 1, "contentBytes": "AA=="}
                ),
            },
        )

        rows, not_probed, _ = preflight._todo_task_rows({"Authorization": "Bearer x"})

        assert not_probed == []
        assert any("/T2/attachments/A2/$value" in p for p, _ in rows)

    def test_no_lists_reports_every_family_not_probed(self, monkeypatch):
        base = "https://graph.microsoft.com/v1.0/me/todo/lists"
        self._wire(monkeypatch, {f"{base}?$top=100": self._lists([])})

        rows, not_probed, failures = preflight._todo_task_rows(
            {"Authorization": "Bearer x"}
        )

        assert rows == []
        assert failures == []
        assert len(not_probed) == 3  # checklist, attachments listing, entity GET
        assert all("could not provision" in n for n in not_probed)


class TestRunSummary:
    """The v1.7.0 lesson applied to gaps: "All endpoints reachable. Safe to
    tag." printed over families the run never probed is how an unsupported
    endpoint ships. The summary now counts reachable vs not-probed."""

    def _run_with(self, monkeypatch, responses, todo_result):
        monkeypatch.setattr(preflight, "fetch_token", lambda: "tok")

        def fake_get(url, headers=None, timeout=None):
            return responses[url]

        monkeypatch.setattr(preflight.httpx, "get", fake_get)
        monkeypatch.setattr(preflight, "_todo_task_rows", lambda headers: todo_result)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            code = preflight.run()
        return code, buf.getvalue()

    def test_all_probed_and_reachable_says_safe_to_tag(self, monkeypatch):
        rows = [("me/todo/lists/L1/tasks/T1/checklistItems?$top=1", "To Do checklist items")]
        code, out = self._run_with(
            monkeypatch,
            {urljoin(preflight.GRAPH_BASE, p): _Resp(200, {}) for p, _ in
             preflight.ENDPOINTS + rows},
            (rows, [], []),
        )

        assert code == 0
        assert "Safe to tag" in out
        assert "NOT probed" not in out

    def test_unprobed_families_are_counted_not_silent(self, monkeypatch):
        rows = [("me/todo/lists/L1/tasks/T1/attachments?$top=1", "To Do task attachments")]
        not_probed = ["To Do attachment entity GET + $value (no attachment)"]

        code, out = self._run_with(
            monkeypatch,
            {urljoin(preflight.GRAPH_BASE, p): _Resp(200, {}) for p, _ in
             preflight.ENDPOINTS + rows},
            (rows, not_probed, []),
        )

        assert code == 0
        assert "Safe to tag" not in out
        assert f"{len(preflight.ENDPOINTS) + len(rows)} probed endpoints reachable" in out
        assert "1 NOT probed" in out
        assert "no attachment" in out
