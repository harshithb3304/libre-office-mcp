"""Find the LibreOffice executable and a Python interpreter that can ``import uno``.

The worker that drives LibreOffice must run under a Python that has the ``uno`` bridge
module. In the Docker image that is the system ``/usr/bin/python3`` (package ``python3-uno``);
a virtualenv cannot see it, which is why the server and the worker are separate programs.

Stdlib-only so the ``doctor`` command can use it without importing the MCP SDK.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ENV_SOFFICE = "LIBREOFFICE_MCP_SOFFICE"
ENV_PYTHON = "LIBREOFFICE_MCP_PYTHON"

_SOFFICE_LOCATIONS = (
    Path("/usr/lib/libreoffice/program/soffice"),
    Path("/usr/lib64/libreoffice/program/soffice"),
    Path("/opt/libreoffice/program/soffice"),
)


class OfficeNotFoundError(RuntimeError):
    """LibreOffice (or a usable Python for it) could not be located."""


def find_soffice(override: str | None = None) -> Path:
    """Locate ``soffice``: explicit override, then env var, then PATH, then known install dirs."""
    explicit = override or os.environ.get(ENV_SOFFICE)
    if explicit:
        path = Path(explicit).expanduser()
        if path.is_file():
            return path
        raise OfficeNotFoundError(f"{ENV_SOFFICE} points to {path}, which is not a file")
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return Path(found)
    for candidate in _SOFFICE_LOCATIONS:
        if candidate.is_file():
            return candidate
    raise OfficeNotFoundError(
        "LibreOffice not found. Use the Docker image (it includes LibreOffice), or install it "
        f"and set {ENV_SOFFICE} to the full path of soffice."
    )


def can_import_uno(python: str | Path) -> bool:
    """True if ``python`` can ``import uno`` (a ~0.3 s subprocess probe)."""
    try:
        result = subprocess.run(
            [str(python), "-c", "import uno"],
            capture_output=True,
            timeout=30,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return result.returncode == 0


def find_python(soffice: Path, override: str | None = None) -> Path:
    """Locate a Python that can ``import uno`` (see module docstring)."""
    explicit = override or os.environ.get(ENV_PYTHON)
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise OfficeNotFoundError(f"{ENV_PYTHON} points to {path}, which is not a file")
        return path
    seen: set[Path] = set()
    for name in (shutil.which("python3"), "/usr/bin/python3", sys.executable):
        if not name:
            continue
        candidate = Path(name)
        if candidate in seen or not candidate.is_file():
            continue
        seen.add(candidate)
        if can_import_uno(candidate):
            return candidate
    raise OfficeNotFoundError(
        "No Python with the 'uno' module was found. Install the distro package "
        f"(`apt install python3-uno`) or set {ENV_PYTHON} to a Python that can `import uno`."
    )
