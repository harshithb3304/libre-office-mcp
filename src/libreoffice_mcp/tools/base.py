"""Helpers every tool uses to talk to the worker.

A tool is normally one line: ``return await edit(path, "writer.add_paragraph", text=text)``.

Convention: any operation argument whose name ends in ``_path`` (``image_path``,
``source_path``...) is a filesystem path. It is resolved and checked against the
configured roots here, so no tool can forget to.
"""

from __future__ import annotations

from typing import Any

from libreoffice_mcp import protocol
from libreoffice_mcp.state import state


def _prepare_args(args: dict[str, Any]) -> dict[str, Any]:
    """Drop ``None`` values (let the worker use its defaults) and resolve ``*_path`` args."""
    prepared: dict[str, Any] = {}
    for key, value in args.items():
        if value is None:
            continue
        if key.endswith("_path") and isinstance(value, str):
            value = str(state.paths.existing_file(value))
        prepared[key] = value
    return prepared


def make_ops(operations: list[tuple[str, dict[str, Any]]]) -> list[dict[str, Any]]:
    return [{"op": name, "args": _prepare_args(args)} for name, args in operations]


async def read(path: str, op: str, **args: Any) -> Any:
    """Run one read-only operation and return its result."""
    file = state.paths.supported_file(path)
    result = await state.office.call(protocol.KIND_READ, file, make_ops([(op, args)]))
    return result["results"][0]


async def edit(path: str, op: str, **args: Any) -> dict[str, Any]:
    """Run one modifying operation, save atomically, and return its result."""
    return await batch(path, [(op, args)])


async def batch(path: str, operations: list[tuple[str, dict[str, Any]]]) -> dict[str, Any]:
    """Run several operations on one open document: one load, one save, all-or-nothing.

    Returns the single operation's result dict (merged with ``file``), or
    ``{"file": ..., "results": [...]}`` for several.
    """
    file = state.paths.supported_file(path)
    reply = await state.office.call(protocol.KIND_EDIT, file, make_ops(operations))
    results = reply["results"]
    if len(results) == 1 and isinstance(results[0], dict):
        return {"file": str(file), **results[0]}
    return {"file": str(file), "results": results}
