"""``scripts/live-tests.sh`` refuses every account tier while the env file's stop file exists.

When Microsoft locks a test account, or blocks its sending, every further request
is more of what got it locked, a sign-in check included. Whatever else uses the
accounts writes a stop file then, and the env file can name it
(``OUTLOOK_MCP_LIVE_STOP_FILE``). These tests run the real script with a stand-in
``docker`` on ``PATH`` that records each call and fails, so nothing is built, no
volume is read and no account is reached: a refused run must exit before its first
``docker`` call, and a run let through reaches it.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRIPT = REPO / "scripts" / "live-tests.sh"

pytestmark = pytest.mark.skipif(sys.platform == "win32", reason="the runner is a bash script")

# Every tier that reaches an account. `offline` reaches none.
ACCOUNT_TIERS = [
    ["preflight"],
    ["live"],
    ["integration"],
    ["live_write"],
    ["--sends", "1", "live_send"],
    ["status"],
    ["leftovers"],
    ["clean"],
]


def _env_file(tmp_path: Path, stop_file: str | None) -> Path:
    lines = [
        "OUTLOOK_MCP_TEST_ACCOUNT=main@example.com",
        "OUTLOOK_MCP_TEST_PAIR=pair@example.com",
        "OUTLOOK_MCP_LIVE_VOLUME=stop-file-test-main",
        "OUTLOOK_MCP_LIVE_PAIR_VOLUME=stop-file-test-pair",
    ]
    if stop_file is not None:
        lines.append(f"OUTLOOK_MCP_LIVE_STOP_FILE={stop_file}")
    path = tmp_path / "live-tests.env"
    path.write_text("\n".join(lines) + "\n")
    return path


def _run(tmp_path: Path, env_file: Path, *args: str) -> tuple[subprocess.CompletedProcess, str]:
    """The script with a ``docker`` that logs its arguments and fails."""
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir(exist_ok=True)
    calls = tmp_path / "docker-calls"
    docker = bin_dir / "docker"
    docker.write_text(f'#!/bin/sh\necho "$*" >> "{calls}"\nexit 1\n')
    docker.chmod(0o755)
    env = {k: v for k, v in os.environ.items() if k != "OUTLOOK_MCP_LIVE_ENV_FILE"}
    env["PATH"] = f"{bin_dir}{os.pathsep}{env.get('PATH', '')}"
    result = subprocess.run(
        ["bash", str(SCRIPT), "--env-file", str(env_file), *args],
        capture_output=True,
        text=True,
        env=env,
        timeout=60,
    )
    return result, calls.read_text() if calls.exists() else ""


@pytest.mark.parametrize("tier", ACCOUNT_TIERS, ids=lambda t: t[-1])
def test_an_existing_stop_file_refuses_before_any_docker_call(tmp_path, tier):
    stop = tmp_path / "LOCKED"
    stop.write_text("account locked\n")
    result, calls = _run(tmp_path, _env_file(tmp_path, str(stop)), *tier)
    assert result.returncode == 2
    assert "stop file" in result.stderr
    assert "outlook.com" in result.stderr  # says how to get out of it
    assert str(stop) not in result.stderr + result.stdout  # values are never printed
    assert calls == ""


@pytest.mark.parametrize("tier", ACCOUNT_TIERS, ids=lambda t: t[-1])
def test_a_named_stop_file_that_does_not_exist_lets_the_run_through(tmp_path, tier):
    result, calls = _run(tmp_path, _env_file(tmp_path, str(tmp_path / "LOCKED")), *tier)
    assert "stop file" not in result.stderr
    # The next step after the gate: the volumes are checked, here by the stand-in.
    assert calls.startswith("volume inspect")


def test_no_stop_file_setting_changes_nothing(tmp_path):
    result, calls = _run(tmp_path, _env_file(tmp_path, None), "live")
    assert "stop file" not in result.stderr
    assert calls.startswith("volume inspect")


def test_a_relative_stop_file_path_is_refused(tmp_path):
    result, calls = _run(tmp_path, _env_file(tmp_path, "state/LOCKED"), "live")
    assert result.returncode == 2
    assert "absolute path" in result.stderr
    assert calls == ""


def test_offline_needs_no_account_so_is_not_stopped(tmp_path):
    stop = tmp_path / "LOCKED"
    stop.write_text("account locked\n")
    result, calls = _run(tmp_path, _env_file(tmp_path, str(stop)), "offline")
    assert "stop file" not in result.stderr
    assert calls.startswith("build")
