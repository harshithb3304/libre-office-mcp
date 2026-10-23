"""Command-line entry point: ``libreoffice-mcp [serve|doctor|tools] [options]``."""

from __future__ import annotations

import sys

from libreoffice_mcp import __version__, locate, logging_setup
from libreoffice_mcp.config import Settings, parse_settings


def _serve(settings: Settings) -> int:
    from libreoffice_mcp.app import build_server

    mcp = build_server(settings)
    if settings.transport == "http":
        # Stateless + JSON responses: no sessions to track, simplest possible HTTP mode.
        mcp.run("streamable-http", host=settings.host, port=settings.port,
                json_response=True, stateless_http=True)  # fmt: skip
    else:
        mcp.run("stdio")
    return 0


def _doctor(settings: Settings) -> int:
    """Check that LibreOffice and its Python are found (the #1 source of setup problems)."""
    print(f"libreoffice-mcp {__version__} on {sys.platform}, Python {sys.version.split()[0]}")
    try:
        soffice = locate.find_soffice(settings.soffice)
        print(f"  [ok] soffice : {soffice}")
        python = locate.find_python(soffice, settings.python)
        print(f"  [ok] python  : {python}  (can import uno)")
    except locate.OfficeNotFoundError as exc:
        print(f"  [FAIL] {exc}")
        return 1
    print("Everything needed was found. Try: libreoffice-mcp tools")
    return 0


def _tools(settings: Settings) -> int:
    from libreoffice_mcp import tools
    from libreoffice_mcp.tools import registry

    tools.load_modules(settings.apps)
    chosen = registry.select(settings)
    width = max((len(t.name) for t in chosen), default=0)
    for t in chosen:
        print(f"{t.name:<{width}}  [{t.app}/{t.mode}]  {t.title}")
    print(f"\n{len(chosen)} tools")
    return 0


def main(argv: list[str] | None = None) -> int:
    command, settings = parse_settings(argv)
    logging_setup.configure(settings.log_level)
    if command == "doctor":
        return _doctor(settings)
    if command == "tools":
        return _tools(settings)
    return _serve(settings)


if __name__ == "__main__":
    sys.exit(main())
