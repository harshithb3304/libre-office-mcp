"""Worker operations, grouped by application.

Each operation is a function ``fn(session, **args)`` registered with
``@op("<app>.<name>")``. Importing a module registers its operations, so
``load()`` simply imports what is enabled.

    ops/common.py   ops that work on any document   (doc.*)
    ops/writer/     Writer-only operations          (writer.*)
    ops/calc/       Calc-only operations            (calc.*)      [planned]
    ops/impress/    Impress-only operations         (impress.*)   [planned]

To add an operation, see docs/ADDING_A_TOOL.md.
"""

from __future__ import annotations

import importlib
import logging

from libreoffice_mcp.worker.registry import OPS

log = logging.getLogger("worker.ops")

__all__ = ["OPS", "load"]


def load(apps: tuple[str, ...]) -> None:
    importlib.import_module("libreoffice_mcp.worker.ops.common")
    for app in apps:
        try:
            importlib.import_module(f"libreoffice_mcp.worker.ops.{app}")
        except ModuleNotFoundError as exc:
            if exc.name != f"libreoffice_mcp.worker.ops.{app}":
                raise
            log.warning("no operations implemented yet for app %r", app)
    log.info("loaded %d operations", len(OPS))
