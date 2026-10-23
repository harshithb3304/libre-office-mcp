"""Logging configuration.

For the stdio transport, **stdout carries the MCP protocol**, so all logs go to
stderr (MCP clients capture it in their own log files). Logger names:

``libreoffice_mcp.*``   the server (tool calls, worker lifecycle)
``worker``              lines forwarded from the LibreOffice worker process
"""

from __future__ import annotations

import logging
import sys

FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def configure(level: str = "INFO") -> None:
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter(FORMAT, datefmt="%H:%M:%S"))
    root = logging.getLogger()
    root.handlers[:] = [handler]
    root.setLevel(level.upper())
    # Chatty libraries: keep them at WARNING unless the user asked for DEBUG.
    if level.upper() != "DEBUG":
        for noisy in ("httpx", "httpcore", "asyncio", "mcp"):
            logging.getLogger(noisy).setLevel(logging.WARNING)
