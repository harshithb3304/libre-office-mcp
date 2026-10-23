"""Start, connect to, health-check and stop one headless ``soffice`` process.

The worker owns exactly one soffice. It is launched with:

* ``-env:UserInstallation`` -> a private, throw-away user profile, so nothing is shared with
  any other LibreOffice instance.
* ``--accept=pipe,name=...`` -> a UNO endpoint on a uniquely named pipe, so there are no TCP
  ports to collide on.

soffice deliberately stays in the *worker's* process group. The server starts the worker in
its own group and kills that whole group on a timeout or cancellation, which therefore takes
soffice down too (no orphans). On a normal shutdown the worker exits when the server closes
its stdin and then asks soffice to quit (see ``main.py``).
"""

from __future__ import annotations

import contextlib
import logging
import os
import shutil
import subprocess
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any

import uno

from libreoffice_mcp.worker.util import pv

log = logging.getLogger("worker.office")


class OfficeStartError(RuntimeError):
    pass


class OfficeProcess:
    def __init__(self, soffice: str, startup_timeout: float = 90.0):
        self.soffice = soffice
        self.startup_timeout = startup_timeout
        self._pipe = f"lomcp_{os.getpid()}_{uuid.uuid4().hex[:8]}"
        self._profile = Path(tempfile.mkdtemp(prefix="lomcp_profile_"))
        self._proc: subprocess.Popen[bytes] | None = None
        self._stderr_path = self._profile / "soffice.stderr.log"
        self.context: Any = None  # UNO component context of the running office
        self.desktop: Any = None  # com.sun.star.frame.Desktop

    # -- lifecycle ---------------------------------------------------------

    def start(self) -> None:
        args = [
            self.soffice,
            "--headless",
            "--invisible",
            "--norestore",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            f"-env:UserInstallation={uno.systemPathToFileUrl(str(self._profile))}",
            f"--accept=pipe,name={self._pipe};urp;",
        ]
        log.info("starting soffice (pipe=%s)", self._pipe)
        started = time.monotonic()
        with open(self._stderr_path, "wb") as err:
            self._proc = subprocess.Popen(
                args,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=err,
            )
        self._connect()
        log.info("soffice ready in %.1fs", time.monotonic() - started)

    def _connect(self) -> None:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        url = f"uno:pipe,name={self._pipe};urp;StarOffice.ComponentContext"
        deadline = time.monotonic() + self.startup_timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            assert self._proc is not None
            code = self._proc.poll()
            if code not in (None, 0):
                raise OfficeStartError(f"soffice exited with code {code}. {self._stderr_tail()}")
            try:
                self.context = resolver.resolve(url)
                self.desktop = self.context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.context
                )
                return
            except Exception as exc:  # NoConnectException until the pipe exists
                last_error = exc
                time.sleep(0.1)
        raise OfficeStartError(
            f"Timed out after {self.startup_timeout:.0f}s waiting for soffice "
            f"({last_error}). {self._stderr_tail()}"
        )

    def _stderr_tail(self) -> str:
        try:
            text = self._stderr_path.read_text(errors="replace").strip()
        except OSError:
            return ""
        return f"soffice stderr: {text[-500:]}" if text else ""

    def is_healthy(self) -> bool:
        """True while the UNO bridge still answers."""
        if self.desktop is None:
            return False
        try:
            self.desktop.getFrames()  # cheap remote call; raises once the office is gone
            return True
        except Exception:
            return False

    def version(self) -> str:
        try:
            provider = self.context.ServiceManager.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider", self.context
            )
            node = provider.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess",
                (pv("nodepath", "/org.openoffice.Setup/Product"),),
            )
            return str(node.getByName("ooSetupVersionAboutBox"))
        except Exception:
            return "unknown"

    def stop(self) -> None:
        """Ask soffice to quit; force-kill it if it lingers. Cleans up the profile."""
        if self.desktop is not None:
            with contextlib.suppress(Exception):  # the bridge drops while the office exits
                self.desktop.terminate()
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline and self.is_healthy():
                time.sleep(0.1)
            if self.is_healthy() and self._proc is not None:
                with contextlib.suppress(OSError):
                    self._proc.kill()
        self._proc = None
        self.desktop = None
        time.sleep(0.3)  # let soffice release profile files before deleting them
        shutil.rmtree(self._profile, ignore_errors=True)
