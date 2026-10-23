"""Shared fixtures. Tests marked ``office`` run against a real LibreOffice.

One LibreOffice is started for the whole test session (starting takes 10-20 s), and
each test works in its own temp folder.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

import pytest

from libreoffice_mcp import locate, protocol
from libreoffice_mcp.client import OfficeClient
from libreoffice_mcp.config import Settings


def _office_available() -> bool:
    try:
        locate.find_python(locate.find_soffice())
        return True
    except locate.OfficeNotFoundError:
        return False


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if _office_available():
        return
    skip = pytest.mark.skip(reason="LibreOffice (with Python-UNO) is not installed")
    for item in items:
        if "office" in item.keywords:
            item.add_marker(skip)


@pytest.fixture(scope="session")
async def office() -> AsyncIterator[OfficeClient]:
    client = OfficeClient(Settings(op_timeout=120, startup_timeout=120))
    await client._ensure_running_for_tests()
    try:
        yield client
    finally:
        await client.stop()


class Doc:
    """Small helper that runs operations against one file through the real worker."""

    def __init__(self, office: OfficeClient, path: Path):
        self.office, self.path = office, path

    async def create(self, *ops: tuple[str, dict[str, Any]]) -> list[Any]:
        reply = await self.office.call(
            protocol.KIND_CREATE, self.path, [{"op": n, "args": a} for n, a in ops],
            options={"overwrite": True},
        )  # fmt: skip
        return reply["results"]

    async def edit(self, *ops: tuple[str, dict[str, Any]]) -> list[Any]:
        reply = await self.office.call(
            protocol.KIND_EDIT, self.path, [{"op": n, "args": a} for n, a in ops]
        )
        return reply["results"]

    async def read(self, op: str, **args: Any) -> Any:
        reply = await self.office.call(protocol.KIND_READ, self.path, [{"op": op, "args": args}])
        return reply["results"][0]

    async def structure(self) -> list[tuple[str, str]]:
        """Compact (type, text) list of all blocks, for readable assertions."""
        info = await self.read("writer.get_structure", limit=1000, text_chars=200)
        return [(b["type"], b.get("text", "")) for b in info["blocks"]]


@pytest.fixture
async def doc(office: OfficeClient, tmp_path: Path) -> Doc:
    d = Doc(office, tmp_path / "test.odt")
    await d.create()
    return d
