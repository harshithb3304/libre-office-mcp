"""Runtime settings, read from command-line flags with environment-variable fallbacks.

Every flag has a ``LIBREOFFICE_MCP_*`` environment variable, which is handy in
MCP client configs (``"env": {...}``). A flag always wins over its variable.
"""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from pathlib import Path

from libreoffice_mcp import formats

ENV_PREFIX = "LIBREOFFICE_MCP_"


@dataclass(frozen=True)
class Settings:
    #: Which applications' tools to expose (``core`` tools are always on).
    apps: tuple[str, ...] = formats.APPS
    #: Glob patterns of tool names; if set, only matching tools are exposed.
    include: tuple[str, ...] = ()
    #: Glob patterns of tool names to hide.
    exclude: tuple[str, ...] = ()
    #: Expose only tools that never modify files.
    read_only: bool = False
    #: Where relative paths resolve. Defaults to ~/Documents.
    default_dir: Path = Path.home() / "Documents"
    #: If non-empty, every file the tools touch must live under one of these.
    roots: tuple[Path, ...] = ()
    #: Seconds a single operation may run before the worker is killed.
    op_timeout: float = 120.0
    #: Seconds to wait for LibreOffice to start.
    startup_timeout: float = 90.0
    soffice: str | None = None
    python: str | None = None
    log_level: str = "INFO"
    transport: str = "stdio"  # "stdio" | "http"
    host: str = "127.0.0.1"
    port: int = 8000


def _env(name: str, default: str | None = None) -> str | None:
    return os.environ.get(ENV_PREFIX + name, default)


def _csv(value: str | None) -> tuple[str, ...]:
    return tuple(p.strip() for p in (value or "").split(",") if p.strip())


def _bool(value: str | None) -> bool:
    return (value or "").strip().lower() in ("1", "true", "yes", "on")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="libreoffice-mcp",
        description="MCP server for creating and editing LibreOffice documents.",
    )
    p.add_argument("command", nargs="?", default="serve", choices=["serve", "doctor", "tools"],
                   help="serve (default): run the MCP server; doctor: check the installation; "
                        "tools: list the tools that would be exposed")  # fmt: skip
    p.add_argument("--apps", default=_env("APPS"),
                   help="comma-separated apps to expose: writer,calc,impress (default: all)")  # fmt: skip
    p.add_argument("--include", action="append", default=None, metavar="GLOB",
                   help="only expose tools matching this glob (repeatable), e.g. 'writer_get_*'")  # fmt: skip
    p.add_argument("--exclude", action="append", default=None, metavar="GLOB",
                   help="hide tools matching this glob (repeatable)")  # fmt: skip
    p.add_argument("--read-only", action="store_true", default=_bool(_env("READ_ONLY")),
                   help="expose only tools that cannot modify files")  # fmt: skip
    p.add_argument("--default-dir", default=_env("DEFAULT_DIR"),
                   help="folder for relative paths (default: ~/Documents)")  # fmt: skip
    p.add_argument("--root", action="append", default=None, metavar="DIR",
                   help="restrict all file access to this folder (repeatable). "
                        "Env: LIBREOFFICE_MCP_ROOTS, separated by the OS path separator")  # fmt: skip
    p.add_argument("--op-timeout", type=float, default=float(_env("OP_TIMEOUT", "120")),
                   help="seconds before a stuck operation is aborted (default 120)")  # fmt: skip
    p.add_argument("--startup-timeout", type=float, default=float(_env("STARTUP_TIMEOUT", "90")),
                   help="seconds to wait for LibreOffice to start (default 90)")  # fmt: skip
    p.add_argument("--soffice", default=None, help="path to soffice (env: LIBREOFFICE_MCP_SOFFICE)")
    p.add_argument("--python", default=None,
                   help="Python that can `import uno` (env: LIBREOFFICE_MCP_PYTHON)")  # fmt: skip
    p.add_argument("--log-level", default=_env("LOG_LEVEL", "INFO"),
                   choices=["DEBUG", "INFO", "WARNING", "ERROR"])  # fmt: skip
    p.add_argument("--transport", default=_env("TRANSPORT", "stdio"), choices=["stdio", "http"],
                   help="stdio (default, for desktop AI apps) or http (Streamable HTTP, local use)")  # fmt: skip
    p.add_argument("--host", default=_env("HOST", "127.0.0.1"), help="http transport bind address")
    p.add_argument(
        "--port", type=int, default=int(_env("PORT", "8000")), help="http transport port"
    )
    return p


def parse_settings(argv: list[str] | None = None) -> tuple[str, Settings]:
    """Parse CLI + environment into ``(command, Settings)``."""
    parser = build_parser()
    args = parser.parse_args(argv)

    apps = _csv(args.apps) or formats.APPS
    unknown = [a for a in apps if a not in formats.APPS]
    if unknown:
        parser.error(
            f"unknown app(s): {', '.join(unknown)} (choose from {', '.join(formats.APPS)})"
        )

    roots_raw = args.root or [r for r in (_env("ROOTS") or "").split(os.pathsep) if r]
    settings = Settings(
        apps=apps,
        include=tuple(args.include or _csv(_env("INCLUDE"))),
        exclude=tuple(args.exclude or _csv(_env("EXCLUDE"))),
        read_only=args.read_only,
        default_dir=Path(args.default_dir).expanduser()
        if args.default_dir
        else Settings.default_dir,
        roots=tuple(Path(r).expanduser().resolve() for r in roots_raw),
        op_timeout=args.op_timeout,
        startup_timeout=args.startup_timeout,
        soffice=args.soffice,
        python=args.python,
        log_level=args.log_level,
        transport=args.transport,
        host=args.host,
        port=args.port,
    )
    return args.command, settings
