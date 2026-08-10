from __future__ import annotations

import logging
import os
import sys

from mcp.server.fastmcp import FastMCP

from .connection import CATIAConnection
from .registry import ToolContext, load_tool_modules

logger = logging.getLogger("catia_mcp")


def _configure_logging() -> None:
    """Send server diagnostics to stderr, never to MCP protocol stdout."""
    if logger.handlers:
        return

    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(name)s] %(levelname)s: %(message)s"
        )
    )
    logger.addHandler(handler)

    level_name = os.environ.get("CATIA_MCP_LOG_LEVEL", "INFO").upper()
    logger.setLevel(getattr(logging, level_name, logging.INFO))
    logger.propagate = False


class CATIAMCPServer:
    """CATIA MCP server using the MCP stdio transport."""

    def __init__(self) -> None:
        self.catia = CATIAConnection()
        self._mcp: FastMCP | None = None
        self._registered_tools: list[str] = []

    def setup(self) -> None:
        """Initialize FastMCP and register tools without connecting to CATIA."""
        _configure_logging()

        self._mcp = FastMCP("CATIA MCP Server")
        ctx = ToolContext(conn=self.catia)
        self._registered_tools = load_tool_modules(self._mcp, ctx)

        # Do not connect to or launch CATIA here. Codex must be able to finish
        # the MCP initialize/list-tools handshake immediately. Individual tools
        # call conn.ensure_connected() or catia_start when CATIA is needed.
        logger.info(
            "CATIA MCP initialized with %d registered tool(s); "
            "CATIA connection deferred until first tool call.",
            len(self._registered_tools),
        )

    @property
    def mcp(self) -> FastMCP:
        """Return the initialized FastMCP instance."""
        if self._mcp is None:
            raise RuntimeError("Server not initialized. Call setup() first.")
        return self._mcp

    def run(self) -> None:
        """Start the MCP server; FastMCP uses stdio by default."""
        if self._mcp is None:
            try:
                self.setup()
            except Exception as exc:
                raise RuntimeError(f"Cannot run server: setup failed. {exc}") from exc

        logger.info("Starting CATIA MCP stdio server.")
        self.mcp.run()
