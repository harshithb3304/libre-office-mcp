"""Turn the path strings the model sends into safe, absolute ``Path`` objects."""

from __future__ import annotations

from pathlib import Path

from libreoffice_mcp import formats, protocol
from libreoffice_mcp.errors import OfficeError, PathNotAllowedError


class PathPolicy:
    """Resolves model-supplied paths and (optionally) confines them to ``roots``.

    * ``~`` is expanded; relative paths are taken relative to ``default_dir``.
    * Symlinks and ``..`` are resolved *before* the roots check, so they cannot be
      used to escape a root.
    * With no roots configured everything is allowed - fine for a personal desktop
      setup, but set ``--root`` for anything shared.
    """

    def __init__(self, default_dir: Path, roots: tuple[Path, ...] = ()):
        self.default_dir = default_dir
        self.roots = tuple(r.resolve() for r in roots)

    def resolve(self, raw: str) -> Path:
        if not raw or not raw.strip():
            raise OfficeError("Path is empty.")
        path = Path(raw.strip()).expanduser()
        if not path.is_absolute():
            path = self.default_dir / path
        path = path.resolve()
        if self.roots and not any(path.is_relative_to(root) for root in self.roots):
            raise PathNotAllowedError(path, ", ".join(str(r) for r in self.roots))
        return path

    def existing_file(self, raw: str) -> Path:
        path = self.resolve(raw)
        if not path.is_file():
            hint = ""
            if not Path(raw).expanduser().is_absolute():
                hint = f" (relative paths are resolved against {self.default_dir})"
            raise OfficeError(f"File not found: {path}{hint}", protocol.E_NOT_FOUND)
        return path

    def existing_dir(self, raw: str) -> Path:
        path = self.resolve(raw)
        if not path.is_dir():
            raise OfficeError(f"Folder not found: {path}", protocol.E_NOT_FOUND)
        return path

    def supported_file(self, raw: str) -> Path:
        """An existing file whose extension LibreOffice can open for us."""
        path = self.existing_file(raw)
        ext = formats.ext_of(str(path))
        if ext not in formats.BY_EXT:
            raise OfficeError(
                f"Unsupported file type '.{ext}'. Supported: {', '.join(sorted(formats.BY_EXT))}",
                protocol.E_UNSUPPORTED,
            )
        return path
