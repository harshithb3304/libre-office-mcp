"""The ``@tool`` decorator and the logic that decides which tools get exposed.

A tool module just defines functions::

    @tool("writer", "edit", title="Add paragraph")
    async def writer_add_paragraph(path: DocPath, text: str) -> dict:
        '''Docstring = the description the AI model reads.'''
        return await edit(path, "writer.add_paragraph", text=text)

Importing the module records the function in ``TOOLS``. At startup,
:func:`register_all` filters ``TOOLS`` by the user's settings (apps, read-only,
include/exclude globs) and hands the survivors to the MCP SDK.
"""

from __future__ import annotations

import fnmatch
import functools
import logging
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Literal

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from libreoffice_mcp.config import Settings

log = logging.getLogger("libreoffice_mcp.tools")

#: read        never modifies anything
#: create      writes a new file
#: edit        modifies an existing document (non-destructive additions/changes)
#: destructive removes content (delete paragraph/slide/sheet, replace text...)
Mode = Literal["read", "create", "edit", "destructive"]

CORE = "core"


@dataclass(frozen=True)
class ToolDef:
    fn: Callable[..., Any]
    name: str
    app: str  # "core" | "writer" | "calc" | "impress"
    mode: Mode
    title: str


TOOLS: list[ToolDef] = []


def tool(app: str, mode: Mode = "read", title: str | None = None):
    """Register an async function as an MCP tool belonging to ``app``."""

    def deco(fn: Callable[..., Any]) -> Callable[..., Any]:
        if any(t.name == fn.__name__ for t in TOOLS):
            raise RuntimeError(f"duplicate tool name {fn.__name__!r}")
        TOOLS.append(ToolDef(fn, fn.__name__, app, mode, title or fn.__name__.replace("_", " ")))
        return fn

    return deco


def select(settings: Settings) -> list[ToolDef]:
    """Apply the user's settings to the full tool list."""
    chosen = []
    for t in TOOLS:
        if t.app != CORE and t.app not in settings.apps:
            continue
        if settings.read_only and t.mode != "read":
            continue
        if settings.include and not any(fnmatch.fnmatch(t.name, g) for g in settings.include):
            continue
        if any(fnmatch.fnmatch(t.name, g) for g in settings.exclude):
            continue
        chosen.append(t)
    return sorted(chosen, key=lambda t: t.name)  # deterministic order helps client caching


def _logged(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap a tool so every call logs its name and duration (signature is preserved)."""

    @functools.wraps(fn)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        started = time.monotonic()
        try:
            result = await fn(*args, **kwargs)
        except Exception:
            log.debug(
                "tool %s failed after %.0f ms", fn.__name__, (time.monotonic() - started) * 1000
            )
            raise
        log.debug("tool %s ok in %.0f ms", fn.__name__, (time.monotonic() - started) * 1000)
        return result

    return wrapper


def register_all(mcp: MCPServer, settings: Settings) -> list[ToolDef]:
    chosen = select(settings)
    for t in chosen:
        annotations = ToolAnnotations(
            title=t.title,
            read_only_hint=t.mode == "read",
            destructive_hint=t.mode == "destructive",
            idempotent_hint=t.mode == "read",
            open_world_hint=False,
        )
        mcp.add_tool(_logged(t.fn), name=t.name, title=t.title, annotations=annotations)
    log.info("exposing %d tools (apps: %s%s)", len(chosen), ", ".join(settings.apps),
             ", read-only" if settings.read_only else "")  # fmt: skip
    return chosen
