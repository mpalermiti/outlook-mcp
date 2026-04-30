"""Outlook MCP Server — Microsoft Outlook integration via Graph API."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("outlook-graph-mcp")
except PackageNotFoundError:
    __version__ = "0.0.0+unknown"
