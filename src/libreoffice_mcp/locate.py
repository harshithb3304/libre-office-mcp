"""Find the LibreOffice executable and a Python interpreter that can ``import uno``.

Why two paths? ``soffice`` is the office program itself. The *worker* that drives
it must run under a Python that has the ``uno`` bridge module:

* Windows / macOS: LibreOffice bundles its own Python next to ``soffice``.
* Linux: ``uno`` comes from the distro package (``python3-uno``) and is
  importable from the system ``python3``, not from a virtualenv.

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


class OfficeNotFoundError(RuntimeError):
    """LibreOffice (or a usable Python for it) could not be located."""


def _soffice_candidates() -> list[Path]:
    if sys.platform == "win32":
        roots = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]  # noqa: SIM112
        names = ("LibreOffice", "Collabora Office")
        return [
            Path(root) / name / "program" / "soffice.exe"
            for root in roots
            if root
            for name in names
        ]
    if sys.platform == "darwin":
        return [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    return [
        Path("/usr/lib/libreoffice/program/soffice"),
        Path("/usr/lib64/libreoffice/program/soffice"),
        Path("/opt/libreoffice/program/soffice"),
        Path("/opt/collaboraoffice/program/soffice"),
        Path("/snap/libreoffice/current/lib/libreoffice/program/soffice"),
    ]


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
    for candidate in _soffice_candidates():
        if candidate.is_file():
            return candidate
    raise OfficeNotFoundError(
        "LibreOffice not found. Install it, or set "
        f"{ENV_SOFFICE} to the full path of soffice(.exe)."
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


def _python_candidates(soffice: Path) -> list[Path]:
    program = soffice.resolve().parent
    bundled = [program / "python.exe", program / "python", program.parent / "Resources" / "python"]
    system = [Path(p) for p in (shutil.which("python3"), "/usr/bin/python3", sys.executable) if p]
    return bundled + system


def find_python(soffice: Path, override: str | None = None) -> Path:
    """Locate a Python that can ``import uno`` (see module docstring)."""
    explicit = override or os.environ.get(ENV_PYTHON)
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise OfficeNotFoundError(f"{ENV_PYTHON} points to {path}, which is not a file")
        return path
    seen: set[Path] = set()
    for candidate in _python_candidates(soffice):
        if candidate in seen or not candidate.is_file():
            continue
        seen.add(candidate)
        if can_import_uno(candidate):
            return candidate
    hint = (
        "On Linux install the distro package (e.g. `apt install python3-uno`)."
        if sys.platform.startswith("linux")
        else "Expected LibreOffice's bundled Python next to soffice."
    )
    raise OfficeNotFoundError(
        f"No Python with the 'uno' module was found. {hint} "
        f"Or set {ENV_PYTHON} to a Python that can `import uno`."
    )
