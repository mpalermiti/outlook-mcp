"""The release pipeline runs only code it can name exactly.

A `uses: owner/action@v4` line trusts a tag the action's owner can move at any
time, and the job it runs in can publish to PyPI. Every action is pinned to a
full commit SHA, with the version it was resolved from as a trailing comment so
Dependabot (`.github/dependabot.yml`) can keep the pins current.

`publish.yml` also keeps the PyPI upload in a job of its own: the job that runs
the test suite and every dev dependency gets no OIDC token, so nothing in it
can publish.

Parsed by indentation rather than YAML: the suite has no YAML library, and the
workflows' shape (two-space jobs, `steps:` lists) is regular enough to read.
"""

import re
from pathlib import Path

WORKFLOWS = Path(__file__).resolve().parent.parent / ".github" / "workflows"
USES = re.compile(r"^\s*(?:-\s+)?uses:\s*(\S+)(.*)$")
PINNED = re.compile(r"^[\w.-]+/[\w./-]+@[0-9a-f]{40}$")
VERSION_COMMENT = re.compile(r"^\s+#\s*v\d+\.\d+\.\d+\s*$")


def _workflow_files() -> list[Path]:
    files = sorted(WORKFLOWS.glob("*.yml")) + sorted(WORKFLOWS.glob("*.yaml"))
    assert files, f"no workflows found under {WORKFLOWS}"
    return files


def _jobs(path: Path) -> dict[str, str]:
    """Split a workflow into {job_id: job_text} by the two-space job keys."""
    jobs: dict[str, list[str]] = {}
    current = None
    in_jobs = False
    for line in path.read_text().splitlines():
        if line.startswith("jobs:"):
            in_jobs = True
            continue
        if not in_jobs:
            continue
        if line and not line.startswith(" ") and not line.startswith("#"):
            break  # next top-level key
        m = re.match(r"^  ([\w-]+):\s*$", line)
        if m:
            current = m.group(1)
            jobs[current] = []
        elif current is not None:
            jobs[current].append(line)
    return {k: "\n".join(v) for k, v in jobs.items()}


def test_every_action_is_pinned_to_a_commit_sha_with_its_version():
    offenders = []
    for path in _workflow_files():
        for n, line in enumerate(path.read_text().splitlines(), 1):
            m = USES.match(line)
            if not m:
                continue
            ref, rest = m.groups()
            if ref.startswith("./"):
                continue  # local action: lives in this repo, nothing to pin
            if not PINNED.match(ref) or not VERSION_COMMENT.match(rest):
                offenders.append(f"{path.name}:{n}: {line.strip()}")
    assert not offenders, "pin each action as `owner/action@<40-hex sha> # vX.Y.Z`:\n" + "\n".join(
        offenders
    )


def test_dependabot_keeps_the_action_pins_current_and_nothing_else():
    config = WORKFLOWS.parent / "dependabot.yml"
    text = config.read_text()
    ecosystems = re.findall(r"package-ecosystem:\s*\"?([\w-]+)\"?", text)
    # Python dependencies are bounded by hand, with a comment naming the
    # upstream fix behind each cap — a bot bump would bypass that.
    assert ecosystems == ["github-actions"], ecosystems


def test_only_the_pypi_job_uploads_and_only_publishing_jobs_get_a_token():
    jobs = _jobs(WORKFLOWS / "publish.yml")
    uploaders = [name for name, body in jobs.items() if "uv publish" in body]
    assert uploaders == ["pypi"], uploaders

    with_token = sorted(name for name, body in jobs.items() if "id-token: write" in body)
    assert with_token == ["pypi", "registry"], with_token

    # The PyPI job publishes what the build job built: it runs no checkout, so
    # no repo code, and no test or build step.
    pypi = jobs["pypi"]
    for forbidden in ("actions/checkout", "pytest", "uv sync", "uv build", "run: scripts/"):
        assert forbidden not in pypi, f"the pypi job must only publish, found {forbidden!r}"
    assert re.search(r"needs:\s*build\b", pypi)
    # A cache restored into a job holding a publish token is a poisoning path.
    assert re.search(r"enable-cache:\s*false", pypi)


def test_no_publish_job_names_an_environment_pypi_does_not_expect():
    # PyPI's trusted-publisher entry for this project names no environment
    # (RELEASING.md appendix; every release run so far had none). Adding one
    # here without changing PyPI to match breaks the next release, so it is a
    # deliberate two-sided change, not something to slip into a workflow edit.
    for name, body in _jobs(WORKFLOWS / "publish.yml").items():
        assert not re.search(r"^\s+environment:", body, re.M), name
