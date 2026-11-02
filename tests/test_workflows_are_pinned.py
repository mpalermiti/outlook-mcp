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


def _environment(job: str) -> str | None:
    """A job's `environment:` value, unquoted and without its comment.

    None when the job names no environment. Any `environment:` key counts,
    whatever follows it; the mapping form reads its `name:`.
    """
    m = re.search(r"^    environment:(.*)$", job, re.M)
    if not m:
        return None
    value = re.sub(r"\s+#.*$", "", m.group(1)).strip()
    if not value:
        name = re.search(r"^    environment:\n(?:      .*\n)*?      name:(.*)$", job, re.M)
        value = re.sub(r"\s+#.*$", "", name.group(1)).strip() if name else ""
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        value = value[1:-1]
    return value


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


def test_dependabot_moves_action_pins_and_the_lock_but_never_a_bound():
    config = WORKFLOWS.parent / "dependabot.yml"
    text = config.read_text()
    blocks = {
        match.group(1): match.group(0)
        for match in re.finditer(
            r"- package-ecosystem:\s*\"?([\w-]+)\"?.*?(?=\n\s*- package-ecosystem:|\Z)", text, re.S
        )
    }
    assert list(blocks) == ["github-actions", "uv"], list(blocks)
    # Python dependencies are bounded by hand in pyproject.toml, with a comment
    # naming the upstream fix behind each cap. `lockfile-only` lets Dependabot move
    # uv.lock inside those bounds; any other strategy edits the bounds themselves.
    strategy = re.search(r"^\s+versioning-strategy:\s*\"?([\w-]+)\"?\s*$", blocks["uv"], re.M)
    assert strategy and strategy.group(1) == "lockfile-only", blocks["uv"]


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


def test_only_the_pypi_job_runs_in_the_pypi_environment():
    # The `pypi` GitHub environment is limited to release tags, and PyPI's
    # trusted-publisher entry names the same environment (RELEASING.md
    # appendix). PyPI then refuses a token from any job outside it, so
    # `registry`, which also holds `id-token: write`, cannot upload to PyPI,
    # and a dispatch from a branch cannot publish. A blank environment on
    # PyPI's side accepts any job, so naming it here first cannot break a
    # release; PyPI's entry is changed after one release has gone through.
    named = {
        name: env
        for name, body in _jobs(WORKFLOWS / "publish.yml").items()
        if (env := _environment(body)) is not None
    }
    assert named == {"pypi": "pypi"}, named


def test_the_environment_reader_sees_every_spelling():
    # Any `environment:` key on a job counts, however its value is written: a
    # spelling the reader skips would let `registry` run in `pypi` unnoticed.
    job = "  registry:\n    runs-on: ubuntu-latest\n{}    steps: []\n"
    for line, expected in [
        ("    environment: pypi # release tags only\n", "pypi"),
        ('    environment: "pypi"\n', "pypi"),
        ("    environment: 'pypi'\n", "pypi"),
        ("    environment: ${{ 'pypi' }}\n", "${{ 'pypi' }}"),
        ("    environment:\n      name: pypi\n      url: https://pypi.org\n", "pypi"),
        ("", None),
    ]:
        assert _environment(job.format(line)) == expected, line


def test_the_documented_publish_retry_names_the_release_ref():
    # A bare `gh workflow run publish.yml` dispatches on the default branch,
    # which the tag-only `pypi` environment refuses: the upload job fails and
    # `registry` is skipped, so a half-finished release could not be finished
    # the documented way.
    bare = re.compile(r"gh workflow run publish\.yml(?![^\n`]*--ref )")
    root = WORKFLOWS.parent.parent
    offenders = [
        f"{path.relative_to(root)}:{n}: {line.strip()}"
        for path in (root / "RELEASING.md", WORKFLOWS / "publish.yml")
        for n, line in enumerate(path.read_text().splitlines(), 1)
        if bare.search(line)
    ]
    assert not offenders, "retry with `--ref vX.Y.Z`:\n" + "\n".join(offenders)
