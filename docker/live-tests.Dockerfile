# The live tiers' container: this working tree, every extra, and a chosen Python.
#
# Built and run by scripts/live-tests.sh, which mounts a test account's signed-in
# settings at /config and its pair's at /pair. The image itself holds no sign-in
# and names no account. Build context: the repository root.
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

COPY . .
RUN uv sync --locked --all-extras

# The mounted account is the settings directory and HOME: on Linux, azure-identity
# keeps its token cache under $HOME/.IdentityService.
ENV OUTLOOK_MCP_CONFIG_DIR=/config \
    HOME=/config
