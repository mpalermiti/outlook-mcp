"""FastMCP server for Microsoft Outlook."""

from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP


@asynccontextmanager
async def lifespan(server):
    """Initialize server state: config and auth manager."""
    yield {}


mcp = FastMCP(
    "outlook-mcp",
    version="0.1.0",
    description="MCP server for Microsoft Outlook via Microsoft Graph API",
    lifespan=lifespan,
)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
