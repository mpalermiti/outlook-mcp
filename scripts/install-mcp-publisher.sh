#!/usr/bin/env bash
# Install a pinned, checksum-verified mcp-publisher into the current directory.
#
# publish.yml runs mcp-publisher in the job that holds `id-token: write` — the
# job PyPI and the MCP registry trust to publish this package. It used to fetch
# whatever the registry repo had most recently released, with nothing checking
# what arrived, so a tampered or broken upstream release would have run there.
# Now the version is fixed and the tarball must match its SHA-256 before it is
# unpacked. CI runs this same script on every push, so a wrong pin fails there,
# not at release time.
#
# To move to a newer publisher: take the version from
# https://github.com/modelcontextprotocol/registry/releases, and the checksum
# from `gh api repos/modelcontextprotocol/registry/releases/tags/<version>`
# (the asset's `digest`) — then confirm it against a download of your own.
#
# Linux x86_64 only: that is the runner both workflows use.
set -euo pipefail

VERSION="v1.8.1"
SHA256="a06c9096dcb9727c13555b6be26c7effa707b01f06a4c561ba7a3635443cf2cc"
ASSET="mcp-publisher_linux_amd64.tar.gz"

curl -fsSL -o "$ASSET" \
  "https://github.com/modelcontextprotocol/registry/releases/download/${VERSION}/${ASSET}"
echo "${SHA256}  ${ASSET}" | sha256sum --check --strict -
tar xzf "$ASSET" mcp-publisher
rm "$ASSET"
echo "mcp-publisher ${VERSION} installed and verified"
