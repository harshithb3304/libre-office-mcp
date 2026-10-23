"""Wire protocol between the MCP server and its LibreOffice workers.

Messages are single-line JSON objects terminated by ``\\n``, exchanged over the
worker's stdin/stdout. One worker handles one request at a time.

server -> worker
    {"id": 1, "type": "run", "kind": "read|edit|create|export", "path": "...",
     "target": "...", "ops": [{"op": "writer.add_paragraph", "args": {...}}],
     "options": {...}}
    {"id": 2, "type": "ping"}
    {"id": 3, "type": "shutdown"}

worker -> server
    {"type": "ready", "pid": 1, "office": "25.2.4.3", "python": "3.10.18"}
    {"id": 1, "ok": true, "result": {...}}
    {"id": 1, "ok": false, "error": {"code": "...", "message": "...", "op_index": 0}}

Everything here is stdlib-only because the worker runs on LibreOffice's Python.
"""

from __future__ import annotations

import json
from typing import Any

PROTOCOL_VERSION = 1

# Request kinds
KIND_READ = "read"  # open read-only, run read ops
KIND_EDIT = "edit"  # open a scratch copy, run ops, save atomically
KIND_CREATE = "create"  # new document from a blank template, run ops, save
KIND_EXPORT = "export"  # open read-only, store to another format

# Error codes (stable, surfaced to MCP clients)
E_INVALID_ARGUMENT = "invalid_argument"
E_NOT_FOUND = "not_found"
E_EXISTS = "already_exists"
E_UNSUPPORTED = "unsupported"
E_WRONG_DOC_TYPE = "wrong_document_type"
E_OFFICE = "office_error"
E_OFFICE_CRASHED = "office_crashed"
E_INTERNAL = "internal_error"


def encode(message: dict[str, Any]) -> bytes:
    """Serialise one message as a newline-terminated UTF-8 JSON line."""
    return (
        json.dumps(message, ensure_ascii=False, separators=(",", ":"), default=str).encode("utf-8")
        + b"\n"
    )


def decode(line: bytes | str) -> dict[str, Any]:
    """Parse one JSON line; raises ``ValueError`` on malformed or non-object input."""
    obj = json.loads(line)
    if not isinstance(obj, dict):
        raise ValueError("protocol message must be a JSON object")
    return obj
