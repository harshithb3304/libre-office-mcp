"""LibreOffice MCP: Writer, Calc and Impress automation for AI agents.

This package root must stay import-light and stdlib-only: the worker
(``libreoffice_mcp.worker``) runs inside LibreOffice's bundled Python, which
has no third-party packages installed.
"""

__version__ = "2.0.0a1"
