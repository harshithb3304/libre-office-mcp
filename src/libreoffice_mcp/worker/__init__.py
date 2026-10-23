"""LibreOffice worker: runs under LibreOffice's Python (stdlib + ``uno`` only).

Each worker owns exactly one headless ``soffice`` process and executes one
request at a time. The MCP server talks to it over stdin/stdout
(see :mod:`libreoffice_mcp.protocol`).
"""
