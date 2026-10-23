"""Worker main loop: read JSON requests from stdin, answer on stdout.

stdout is the protocol channel, so we immediately point file descriptor 1 at
stderr; a stray ``print`` (or a noisy library) can then never corrupt the stream.
The server re-logs everything we write to stderr.

The worker exits when stdin closes, which is how it learns the MCP server went
away. ``finally`` then shuts soffice down, so no orphan office processes remain.
"""

from __future__ import annotations

import argparse
import logging
import os
import platform
import sys
import traceback
from typing import Any

from libreoffice_mcp import protocol
from libreoffice_mcp.worker import ops, runner
from libreoffice_mcp.worker.office import OfficeProcess, OfficeStartError
from libreoffice_mcp.worker.registry import OpError

log = logging.getLogger("worker")

EXIT_OFFICE_DIED = 3


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="libreoffice_mcp.worker")
    parser.add_argument("--soffice", required=True, help="path to the soffice executable")
    parser.add_argument("--apps", default="writer,calc,impress", help="comma-separated apps")
    parser.add_argument("--startup-timeout", type=float, default=90.0)
    parser.add_argument("--log-level", default="INFO")
    return parser.parse_args(argv)


def _error(request_id: Any, exc: OpError | Exception) -> dict[str, Any]:
    if isinstance(exc, OpError):
        error = {"code": exc.code, "message": exc.message}
        if hasattr(exc, "op_index"):
            error["op_index"] = exc.op_index
            error["op"] = exc.op_name
    else:
        error = {"code": protocol.E_INTERNAL, "message": f"{type(exc).__name__}: {exc}"}
    return {"id": request_id, "ok": False, "error": error}


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    logging.basicConfig(
        stream=sys.stderr,
        level=args.log_level.upper(),
        format="%(levelname)s %(name)s: %(message)s",
    )

    # Reserve the real stdout for the protocol; route accidental prints to stderr.
    channel = os.fdopen(os.dup(sys.stdout.fileno()), "wb", buffering=0)
    os.dup2(sys.stderr.fileno(), sys.stdout.fileno())

    def send(message: dict[str, Any]) -> None:
        channel.write(protocol.encode(message))

    ops.load(tuple(a for a in args.apps.split(",") if a))
    office = OfficeProcess(args.soffice, args.startup_timeout)
    try:
        try:
            office.start()
        except OfficeStartError as exc:
            log.error("%s", exc)
            send({"type": "fatal", "message": str(exc)})
            return 2
        send(
            {
                "type": "ready",
                "pid": os.getpid(),
                "office": office.version(),
                "python": platform.python_version(),
                "ops": len(ops.OPS),
            }
        )
        return _serve(office, send)
    finally:
        office.stop()


def _serve(office: OfficeProcess, send: Any) -> int:
    for line in sys.stdin.buffer:
        line = line.strip()
        if not line:
            continue
        request: dict[str, Any] = {}
        try:
            request = protocol.decode(line)
            kind = request.get("type")
            if kind == "shutdown":
                return 0
            if kind == "ping":
                send(
                    {
                        "id": request.get("id"),
                        "ok": True,
                        "result": {"healthy": office.is_healthy()},
                    }
                )
                continue
            if kind != "run":
                raise OpError(f"Unknown message type {kind!r}", protocol.E_INTERNAL)
            result = runner.execute(office, request)
            send({"id": request.get("id"), "ok": True, "result": result})
        except OpError as exc:
            send(_error(request.get("id"), exc))
        except Exception as exc:  # a bug or a UNO failure outside any operation
            log.error("request failed:\n%s", traceback.format_exc())
            send(_error(request.get("id"), exc))
        if not office.is_healthy():
            log.error("soffice is no longer healthy; exiting so the server can restart us")
            send({"type": "fatal", "code": protocol.E_OFFICE_CRASHED, "message": "soffice died"})
            return EXIT_OFFICE_DIED
    return 0  # stdin closed: the server is gone
