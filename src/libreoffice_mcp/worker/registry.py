"""Registry of worker operations (``writer.add_paragraph`` and friends)."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from libreoffice_mcp import protocol


class OpError(Exception):
    """An operation failed in a way the caller should hear about verbatim."""

    def __init__(self, message: str, code: str = protocol.E_INVALID_ARGUMENT):
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass(frozen=True)
class OpDef:
    name: str
    fn: Callable[..., Any]
    #: Document kind this op requires ("writer", "calc", "impress") or None for any.
    doc_kind: str | None
    mutates: bool


OPS: dict[str, OpDef] = {}


def op(name: str, *, kind: str | None = None, mutates: bool = False):
    """Register ``fn(session, **args)`` as the implementation of operation ``name``."""

    def deco(fn: Callable[..., Any]) -> Callable[..., Any]:
        if name in OPS:
            raise RuntimeError(f"duplicate worker op {name!r}")
        OPS[name] = OpDef(name, fn, kind, mutates)
        return fn

    return deco
