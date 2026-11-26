"""The live-tests image takes only what it needs from the checkout.

``scripts/live-tests.sh`` builds ``docker/live-tests.Dockerfile`` from the working
tree, which can hold files that must never leave it: the registry publishing
tokens ``.gitignore`` names, an env file, ``.git``, other worktrees. Two
allow-lists keep them out, and these tests keep the two in step: the ignore file
starts by excluding everything and re-includes by name, and the Dockerfile copies
by name.

The matcher below is a simplified reading of Docker's ignore rules (last matching
line wins; a pattern also matches everything under a directory it matches), which
is all these plain patterns need.
"""

from __future__ import annotations

import fnmatch
import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DOCKERFILE = REPO / "docker" / "live-tests.Dockerfile"
IGNORE = REPO / "docker" / "live-tests.Dockerfile.dockerignore"


def _rules() -> list[str]:
    return [
        line.strip()
        for line in IGNORE.read_text("utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]


def _matches(pattern: str, path: str) -> bool:
    pattern = pattern.strip("/")
    parts = path.split("/")
    if pattern.startswith("**/"):
        return any(_matches(pattern[3:], "/".join(parts[i:])) for i in range(len(parts)))
    return any(fnmatch.fnmatchcase("/".join(parts[:i]), pattern) for i in range(1, len(parts) + 1))


def _in_context(path: str) -> bool:
    included = True
    for rule in _rules():
        negated = rule.startswith("!")
        if _matches(rule.lstrip("!"), path):
            included = negated
    return included


def _reincluded() -> set[str]:
    return {rule[1:].strip("/") for rule in _rules() if rule.startswith("!")}


def _copied() -> set[str]:
    sources: set[str] = set()
    for line in DOCKERFILE.read_text("utf-8").splitlines():
        found = re.match(r"\s*COPY\s+(?!--from)(.+)$", line)
        if found:
            *names, _destination = found.group(1).split()
            sources |= {n.strip("/") for n in names}
    return sources


def test_the_ignore_file_starts_by_excluding_everything():
    assert _rules()[0] == "*"


def test_the_image_copies_by_name_and_never_the_whole_checkout():
    assert "." not in _copied()
    assert _copied() <= _reincluded(), _copied() - _reincluded()


def test_nothing_the_repo_keeps_out_of_git_is_let_back_in():
    ignored = [
        line.strip().strip("/")
        for line in (REPO / ".gitignore").read_text("utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    ]
    clashes = {
        (name, pattern)
        for name in _reincluded()
        for pattern in ignored
        if fnmatch.fnmatchcase(name, pattern)
    }
    assert not clashes


def test_secrets_and_checkout_state_stay_out_of_the_build():
    for path in (
        ".mcpregistry_github_token",
        ".mcpregistry_registry_token",
        ".env",
        "live-tests.env",
        "tests/.env",
        "scripts/live-tests.env",
        "tests/.mcpregistry_x",
        ".git/config",
        ".worktrees/other/tests/conftest.py",
        ".venv/bin/python",
        "dist/outlook_graph_mcp-1.0.0.tar.gz",
        "agent/MISSION.md",
        "src/outlook_mcp/__pycache__/server.cpython-313.pyc",
    ):
        assert not _in_context(path), path


def test_what_the_image_needs_is_in_the_build():
    for path in (
        "pyproject.toml",
        "uv.lock",
        "src/outlook_mcp/server.py",
        "tests/conftest.py",
        "scripts/preflight.py",
        ".github/workflows/ci.yml",
        "CHANGELOG.md",
        "server.json",
    ):
        assert _in_context(path), path
