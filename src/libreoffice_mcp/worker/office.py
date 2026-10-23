"""Start, connect to, health-check and stop one headless ``soffice`` process.

The worker owns exactly one soffice. It is launched with:

* ``-env:UserInstallation`` -> a private, throw-away user profile. Without it,
  a second soffice (or the user's own LibreOffice window) shares the default
  profile and silently hands our request to the *other* instance.
* ``--accept=pipe,name=...`` -> a UNO endpoint on a uniquely named pipe, so there
  are no TCP ports to collide on.

Platform notes (learned the hard way):

* On Windows ``soffice.exe`` is a tiny launcher that starts ``soffice.bin`` and
  exits. So the process handle we get back is useless for health checks (we ask
  the UNO bridge instead) and the real office is *not* our child. To make sure it
  can never outlive the worker, the worker puts itself in a Windows Job Object
  with "kill on close" before launching: when the worker dies, so does the office.
* On Linux/macOS we start soffice in its own session/process group and kill the
  group if a graceful ``terminate()`` does not work.
"""

from __future__ import annotations

import contextlib
import logging
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any

import uno

from libreoffice_mcp.worker.util import pv

log = logging.getLogger("worker.office")

_CREATE_NO_WINDOW = 0x08000000  # Windows: don't flash a console

# Keep a reference so the job handle stays open for the life of the worker.
_job_handle: Any = None


class OfficeStartError(RuntimeError):
    pass


def _bind_children_to_worker_lifetime_windows() -> None:
    """Create a kill-on-close Job Object and put *this* process in it (Windows only).

    Children started afterwards (soffice.exe -> soffice.bin) inherit membership.
    When the last handle to the job closes - i.e. when this process exits for any
    reason, including being killed - Windows terminates every process in the job.
    """
    global _job_handle
    if sys.platform != "win32" or _job_handle is not None:
        return
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    # Declare handle types explicitly: the default (32-bit int) overflows on 64-bit Windows.
    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
    kernel32.GetCurrentProcess.restype = wintypes.HANDLE
    kernel32.GetCurrentProcess.argtypes = []
    kernel32.SetInformationJobObject.argtypes = [
        wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD,
    ]  # fmt: skip
    kernel32.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]

    class _BasicLimits(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_int64),
            ("PerJobUserTimeLimit", ctypes.c_int64),
            ("LimitFlags", wintypes.DWORD),
            ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t),
            ("ActiveProcessLimit", wintypes.DWORD),
            ("Affinity", ctypes.c_size_t),
            ("PriorityClass", wintypes.DWORD),
            ("SchedulingClass", wintypes.DWORD),
        ]

    class _IoCounters(ctypes.Structure):
        _fields_ = [(n, ctypes.c_uint64) for n in (
            "ReadOps", "WriteOps", "OtherOps", "ReadBytes", "WriteBytes", "OtherBytes")]  # fmt: skip

    class _ExtendedLimits(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", _BasicLimits),
            ("IoInfo", _IoCounters),
            ("ProcessMemoryLimit", ctypes.c_size_t),
            ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t),
            ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]

    job = kernel32.CreateJobObjectW(None, None)
    if not job:
        log.warning("CreateJobObject failed; soffice may outlive the worker if it is killed")
        return
    info = _ExtendedLimits()
    info.BasicLimitInformation.LimitFlags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    ok = kernel32.SetInformationJobObject(
        job,
        9,
        ctypes.byref(info),
        ctypes.sizeof(info),  # 9 = JobObjectExtendedLimitInformation
    )
    if not ok or not kernel32.AssignProcessToJobObject(job, kernel32.GetCurrentProcess()):
        log.warning(
            "could not bind soffice lifetime to the worker (error %s)", ctypes.get_last_error()
        )
        return
    _job_handle = job


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
        try:
            _bind_children_to_worker_lifetime_windows()
        except Exception:  # a safety net only; never block startup on it
            log.warning("could not create the kill-on-close job object", exc_info=True)
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
        kwargs: dict[str, Any] = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = _CREATE_NO_WINDOW
        else:
            kwargs["start_new_session"] = True  # own process group so we can kill the tree
        log.info("starting soffice (pipe=%s)", self._pipe)
        started = time.monotonic()
        with open(self._stderr_path, "wb") as err:
            self._proc = subprocess.Popen(
                args, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=err, **kwargs
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
            if code not in (None, 0):  # 0 = Windows launcher stub that handed off; keep polling
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
        """True while the UNO bridge still answers (see module docstring for why not poll())."""
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
        """Ask soffice to quit; force-kill (POSIX) if it lingers. Cleans up the profile."""
        if self.desktop is not None:
            with contextlib.suppress(Exception):  # the bridge drops while the office exits
                self.desktop.terminate()
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline and self.is_healthy():
                time.sleep(0.1)
            if self.is_healthy() and self._proc is not None and sys.platform != "win32":
                with contextlib.suppress(OSError):
                    os.killpg(self._proc.pid, signal.SIGKILL)
            # On Windows the kill-on-close job object finishes the job when we exit.
        self._proc = None
        self.desktop = None
        time.sleep(0.3)  # let soffice release profile files before deleting them
        shutil.rmtree(self._profile, ignore_errors=True)
