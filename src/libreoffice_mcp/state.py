"""Process-wide objects shared by all tools.

Tools are plain module-level functions (easy to read, easy to copy), so they reach
the office worker and the path policy through this one object instead of through
closures or parameters. It is filled in once, by :func:`init`, when the server starts.
"""

from __future__ import annotations

from libreoffice_mcp.client import OfficeClient
from libreoffice_mcp.config import Settings
from libreoffice_mcp.paths import PathPolicy


class AppState:
    settings: Settings
    paths: PathPolicy
    office: OfficeClient

    def init(self, settings: Settings) -> None:
        self.settings = settings
        self.paths = PathPolicy(settings.default_dir, settings.roots)
        self.office = OfficeClient(settings)


state = AppState()
