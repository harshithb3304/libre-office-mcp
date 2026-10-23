"""The server's connection to its LibreOffice worker process.

Big picture (see docs/ARCHITECTURE.md for the diagram)::

    MCP client  <-stdio/http->  this server  <-stdin/stdout JSON->  worker  <-UNO->  soffice

``OfficeClient`` owns the worker subprocess:

* **One request at a time.** A single ``asyncio.Lock`` serialises requests. LibreOffice
  is effectively single-threaded, so this costs nothing and removes every
  race condition between concurrent tool calls (two edits to the same file can
  never interleave - the second simply waits).
* **Lazy, background start.** LibreOffice takes 5-20 s to boot. We start it as soon
  as the server starts, but never block the MCP handshake on it; the first tool call
  waits for it if it is still starting.
* **Self-healing.** A crashed worker is restarted and the request retried once
  (safe, because the user's file is only replaced at the very end of a request).
  A request that exceeds ``op_timeout`` or is cancelled by the client kills the
  worker - the only reliable way to abort a stuck UNO call.
"""

from __future__ import annotations

import asyncio
import logging
import os
import signal
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from libreoffice_mcp import locate, protocol
from libreoffice_mcp.config import Settings
from libreoffice_mcp.errors import OfficeError

log = logging.getLogger("libreoffice_mcp.client")
worker_log = logging.getLogger("worker")

WORKER_SCRIPT = Path(__file__).parent / "worker" / "__main__.py"
_STREAM_LIMIT = 64 * 1024 * 1024  # a single JSON reply line may be large (document text)


class _WorkerGone(Exception):
    """The worker process exited (crashed, or soffice died)."""


class OfficeClient:
    def __init__(self, settings: Settings):
        self.settings = settings
        self._lock = asyncio.Lock()
        self._proc: asyncio.subprocess.Process | None = None
        self._stderr_task: asyncio.Task[None] | None = None
        self._next_id = 0
        self._started_at = 0.0
        self._paths: tuple[Path, Path] | None = None  # (soffice, python)
        self.worker_info: dict[str, Any] = {}
        self.requests_served = 0
        self.restarts = 0

    # -- public API ----------------------------------------------------------

    def start_in_background(self) -> asyncio.Task[None]:
        """Begin booting LibreOffice now so the first tool call is fast."""

        async def _boot() -> None:
            try:
                async with self._lock:
                    await self._ensure_running()
            except Exception as exc:  # reported again, with context, on the first real call
                log.error("background start failed: %s", exc)

        return asyncio.create_task(_boot())

    async def _ensure_running_for_tests(self) -> None:
        """Start the worker and wait until it is ready (tests only)."""
        async with self._lock:
            await self._ensure_running()

    async def call(
        self,
        kind: str,
        path: Path,
        ops: list[dict[str, Any]] | None = None,
        *,
        target: Path | None = None,
        options: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Run one request in the worker and return its result dict.

        Raises :class:`OfficeError` (message safe to show the model) on any failure.
        """
        request: dict[str, Any] = {
            "type": "run",
            "kind": kind,
            "path": str(path),
            "ops": ops or [],
            "options": options or {},
        }
        if target is not None:
            request["target"] = str(target)
        label = kind + (":" + ",".join(o["op"] for o in ops) if ops else "")
        started = time.monotonic()
        async with self._lock:
            try:
                result = await self._execute(request)
            except OfficeError as exc:
                log.warning("%s %s failed in %.0f ms: %s", label, path.name, _ms(started), exc)
                raise
        self.requests_served += 1
        log.info("%s %s ok in %.0f ms", label, path.name, _ms(started))
        return result

    async def stop(self) -> None:
        """Ask the worker to exit (it shuts soffice down), killing it if it lingers."""
        proc, self._proc = self._proc, None
        if proc is None or proc.returncode is not None:
            return
        try:
            proc.stdin.write(protocol.encode({"type": "shutdown"}))  # type: ignore[union-attr]
            await proc.stdin.drain()  # type: ignore[union-attr]
            await asyncio.wait_for(proc.wait(), timeout=15)
        except (TimeoutError, OSError):
            await self._kill(proc)
        if self._stderr_task:
            self._stderr_task.cancel()

    def status(self) -> dict[str, Any]:
        running = self._proc is not None and self._proc.returncode is None
        return {
            "running": running,
            "worker_pid": self._proc.pid if running and self._proc else None,
            "uptime_seconds": int(time.monotonic() - self._started_at) if running else 0,
            "requests_served": self.requests_served,
            "restarts": self.restarts,
            **{k: v for k, v in self.worker_info.items() if k in ("office", "python", "ops")},
        }

    # -- request execution ---------------------------------------------------

    async def _execute(self, request: dict[str, Any]) -> dict[str, Any]:
        timeout = self.settings.op_timeout
        for attempt in (1, 2):
            await self._ensure_running()
            self._next_id += 1
            request["id"] = self._next_id
            try:
                reply = await asyncio.wait_for(self._roundtrip(request), timeout)
            except TimeoutError:
                await self._kill()
                raise OfficeError(
                    f"The operation took longer than {timeout:.0f}s and was aborted. "
                    "The document was not modified. Try a smaller request.",
                    protocol.E_OFFICE,
                ) from None
            except asyncio.CancelledError:
                # The client gave up. The worker is mid-operation and cannot be
                # interrupted safely, so discard it; a fresh one starts on demand.
                await self._kill()
                raise
            except _WorkerGone:
                await self._kill()
                if attempt == 1:
                    log.warning("worker died during a request; restarting and retrying once")
                    continue
                raise OfficeError(
                    "LibreOffice crashed while processing this request (twice). "
                    "The file may be corrupt or use an unsupported feature.",
                    protocol.E_OFFICE_CRASHED,
                ) from None
            if reply.get("ok"):
                return reply["result"]
            raise self._to_error(reply.get("error") or {}, request)
        raise AssertionError("unreachable")  # pragma: no cover

    @staticmethod
    def _to_error(error: dict[str, Any], request: dict[str, Any]) -> OfficeError:
        message = str(error.get("message", "Unknown error"))
        if "op_index" in error and len(request.get("ops", [])) > 1:
            message += f" (failed at operation #{error['op_index'] + 1}: {error.get('op')}; nothing was saved)"
        return OfficeError(message, str(error.get("code", protocol.E_INTERNAL)))

    async def _roundtrip(self, request: dict[str, Any]) -> dict[str, Any]:
        proc = self._proc
        assert proc is not None and proc.stdin is not None
        try:
            proc.stdin.write(protocol.encode(request))
            await proc.stdin.drain()
        except (OSError, ConnectionError):
            raise _WorkerGone from None
        while True:
            message = await self._read_message(proc)
            if message.get("type") == "fatal":
                raise _WorkerGone
            if message.get("id") == request["id"]:
                return message
            log.debug("ignoring unexpected worker message: %s", message)

    @staticmethod
    async def _read_message(proc: asyncio.subprocess.Process) -> dict[str, Any]:
        assert proc.stdout is not None
        line = await proc.stdout.readline()
        if not line:
            raise _WorkerGone
        try:
            return protocol.decode(line)
        except ValueError:
            log.error("worker sent a malformed line: %r", line[:200])
            raise _WorkerGone from None

    # -- process management ---------------------------------------------------

    async def _ensure_running(self) -> None:
        """Start the worker if it is not running. Caller must hold ``self._lock``."""
        if self._proc is not None and self._proc.returncode is None:
            return
        if self._proc is not None:
            self.restarts += 1
            self._proc = None
        await self._spawn()

    async def _spawn(self) -> None:
        if self._paths is None:
            try:
                soffice = await asyncio.to_thread(locate.find_soffice, self.settings.soffice)
                python = await asyncio.to_thread(locate.find_python, soffice, self.settings.python)
            except locate.OfficeNotFoundError as exc:
                raise OfficeError(str(exc), protocol.E_OFFICE) from None
            self._paths = (soffice, python)
            log.info("using soffice=%s python=%s", soffice, python)
        soffice, python = self._paths

        cmd = [
            str(python), str(WORKER_SCRIPT),
            "--soffice", str(soffice),
            "--apps", ",".join(self.settings.apps),
            "--startup-timeout", str(self.settings.startup_timeout),
            "--log-level", self.settings.log_level,
        ]  # fmt: skip
        kwargs: dict[str, Any] = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        else:
            kwargs["start_new_session"] = True  # own process group, so _kill can reap the tree
        log.info("starting LibreOffice worker (first start can take 10-30 s)")
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            limit=_STREAM_LIMIT,
            **kwargs,
        )
        self._proc = proc
        self._stderr_task = asyncio.create_task(self._pump_stderr(proc))
        try:
            ready = await asyncio.wait_for(
                self._read_message(proc), timeout=self.settings.startup_timeout + 15
            )
        except (TimeoutError, _WorkerGone):
            await self._kill()
            raise OfficeError(
                "LibreOffice failed to start. Run `libreoffice-mcp doctor` and check the "
                "server log for details.",
                protocol.E_OFFICE,
            ) from None
        if ready.get("type") != "ready":
            await self._kill()
            raise OfficeError(
                f"LibreOffice failed to start: {ready.get('message', ready)}", protocol.E_OFFICE
            )
        self.worker_info = ready
        self._started_at = time.monotonic()
        log.info(
            "worker ready: LibreOffice %s, %d operations", ready.get("office"), ready.get("ops", 0)
        )

    async def _pump_stderr(self, proc: asyncio.subprocess.Process) -> None:
        """Forward worker stderr lines into our log (level taken from the line prefix)."""
        assert proc.stderr is not None
        levels = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
        }
        while True:
            line = await proc.stderr.readline()
            if not line:
                return
            text = line.decode("utf-8", "replace").rstrip()
            if not text:
                continue
            level_name, _, rest = text.partition(" ")
            if level_name in levels:  # "INFO worker.office: soffice ready" -> our logger
                worker_log.log(levels[level_name], rest)
            else:  # tracebacks, stray prints, soffice noise
                worker_log.info(text)

    async def _kill(self, proc: asyncio.subprocess.Process | None = None) -> None:
        """Forcefully end the worker and everything it started."""
        proc = proc or self._proc
        if proc is self._proc:
            self._proc = None
        if proc is None or proc.returncode is not None:
            return
        try:
            if sys.platform == "win32":
                # LibreOffice's python.exe is a launcher; /T kills the real worker under it.
                # The worker's kill-on-close job object then takes soffice.bin down too.
                await asyncio.to_thread(
                    subprocess.run,
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    capture_output=True,
                    check=False,
                )
            else:
                os.killpg(proc.pid, signal.SIGKILL)
        except (OSError, ProcessLookupError):
            pass
        try:
            await asyncio.wait_for(proc.wait(), timeout=10)
        except TimeoutError:
            log.error("worker pid %s did not exit after kill", proc.pid)


def _ms(started: float) -> float:
    return (time.monotonic() - started) * 1000
