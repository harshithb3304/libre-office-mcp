"""MCP tool definitions (the part the AI model sees).

    tools/core.py       app-independent tools (create, convert, list, batch...)
    tools/writer/       Writer tools  (writer_*)
    tools/calc/         Calc tools    (calc_*)      [planned]
    tools/impress/      Impress tools (impress_*)   [planned]

A tool only describes the call (name, typed parameters, docstring) and forwards it to
the worker; the LibreOffice logic lives in ``libreoffice_mcp/worker/ops``.
See docs/ADDING_A_TOOL.md.
"""

from __future__ import annotations

import importlib
import logging

log = logging.getLogger("libreoffice_mcp.tools")


def load_modules(apps: tuple[str, ...]) -> None:
    """Import tool modules so their ``@tool`` decorators run."""
    importlib.import_module("libreoffice_mcp.tools.core")
    for app in apps:
        try:
            importlib.import_module(f"libreoffice_mcp.tools.{app}")
        except ModuleNotFoundError as exc:
            if exc.name != f"libreoffice_mcp.tools.{app}":
                raise
            log.warning("no tools implemented yet for app %r", app)
