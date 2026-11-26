# The live tiers' container: this working tree, every extra, and a chosen Python.
#
# Built and run by scripts/live-tests.sh, which mounts a test account's signed-in
# settings at /config (and, for the tiers that need it, its pair's at /pair). The
# image itself holds no sign-in and names no account. Build context: the
# repository root, narrowed by live-tests.Dockerfile.dockerignore.
ARG PYTHON_VERSION=3.13
FROM python:${PYTHON_VERSION}-slim

# uv, pinned, so the container installs exactly what uv.lock says, as CI does.
RUN pip install --no-cache-dir uv==0.11.6

ENV UV_PROJECT_ENVIRONMENT=/venv \
    UV_PYTHON=/usr/local/bin/python3 \
    UV_PYTHON_DOWNLOADS=never \
    UV_LINK_MODE=copy \
    UV_NO_CACHE=1 \
    PATH=/venv/bin:$PATH \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /src

# Dependencies first, so a change to the code reuses this layer.
COPY pyproject.toml uv.lock README.md ./
RUN uv sync --locked --all-extras --no-install-project

# The rest, by name: an allow-list, so nothing else in a checkout (a publishing
# token, an env file, .git, other worktrees) can reach the image, whatever the
# ignore file or the builder does. tests/test_live_image.py holds the two lists
# together. The documents are read by the offline suite's docs checks.
COPY LICENSE CHANGELOG.md RELEASING.md ROADMAP.md SKILL.md server.json ./
COPY src ./src
COPY scripts ./scripts
COPY tests ./tests
COPY .github ./.github
RUN uv sync --locked --all-extras

# The mounted account is the settings directory and HOME: on Linux, azure-identity
# keeps its token cache under $HOME/.IdentityService.
ENV OUTLOOK_MCP_CONFIG_DIR=/config \
    HOME=/config
