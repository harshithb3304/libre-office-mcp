"""Assemble the MCP server: settings -> tools -> ``MCPServer``."""

from __future__ import annotations

import logging
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from mcp.server.mcpserver import MCPServer

from libreoffice_mcp import __version__, tools
from libreoffice_mcp.config import Settings
from libreoffice_mcp.state import state
from libreoffice_mcp.tools import registry

log = logging.getLogger("libreoffice_mcp.app")

# Shown to the AI model when it connects. This is the cheapest place to teach it how to
# use the tools well, so keep it short and practical.
INSTRUCTIONS = """\
Create and edit documents with LibreOffice on the user's computer.

Workflow tips:
- Create with office_create_document (the file extension picks Writer/Calc/Impress), or
  work on an existing file. Relative paths resolve to the user's documents folder.
- Look before you edit: use the *_get_* / *_list_* tools to see the current structure and
  indexes, and refer to items by those indexes.
- For three or more changes to one document, use office_batch_edit: one load, one save, and
  nothing is saved if any step fails.
- Every edit is saved to the file immediately and atomically. Changes cannot be undone, so
  confirm with the user before destructive operations (deleting, replacing).
- Use office_convert_document to export PDF or convert between formats.
"""


@asynccontextmanager
async def _lifespan(_: MCPServer) -> AsyncIterator[None]:
    """Boot LibreOffice in the background at startup; shut it down cleanly at exit."""
    boot = state.office.start_in_background()
    try:
        yield
    finally:
        boot.cancel()
        await state.office.stop()


def build_server(settings: Settings) -> MCPServer:
    state.init(settings)
    tools.load_modules(settings.apps)
    mcp = MCPServer(
        "libreoffice",
        title="LibreOffice",
        instructions=INSTRUCTIONS,
        version=__version__,
        lifespan=_lifespan,
    )
    registry.register_all(mcp, settings)
    return mcp
