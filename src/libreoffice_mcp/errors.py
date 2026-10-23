"""Errors that are meant to be read by the AI model.

The MCP SDK shows the model the message of a :class:`ToolError` and hides the
text of any other exception (to avoid leaking internals). So every failure we
*anticipate* - bad path, wrong document type, unknown paragraph index - must be
raised as an ``OfficeError`` with a message that tells the model how to fix its
next call.
"""

from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError

from libreoffice_mcp import protocol


class OfficeError(ToolError):
    """A failure the model should see (and can usually recover from)."""

    def __init__(self, message: str, code: str = protocol.E_INVALID_ARGUMENT):
        super().__init__(message)
        self.code = code
        self.message = message


class PathNotAllowedError(OfficeError):
    def __init__(self, path: object, roots: object):
        super().__init__(
            f"Access to {path} is not allowed. Files must be inside: {roots}.",
            protocol.E_INVALID_ARGUMENT,
        )
