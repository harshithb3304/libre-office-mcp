"""Process lifecycle: killing a worker must never leave an orphaned soffice behind."""

from __future__ import annotations

import asyncio
import os
from pathlib import Path

import pytest

from libreoffice_mcp import protocol
from libreoffice_mcp.client import OfficeClient
from libreoffice_mcp.config import Settings

pytestmark = pytest.mark.office


def group_is_empty(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)  # signal 0: only checks whether any process is in the group
    except ProcessLookupError:
        return True
    return False


async def test_kill_takes_worker_and_soffice_down_together():
    client = OfficeClient(Settings(startup_timeout=120))
    await client.wait_until_ready()
    group = client._proc.pid  # the worker leads its own process group; soffice shares it
    assert not group_is_empty(group)
    await client._kill()
    for _ in range(20):  # the kernel needs a moment to reap the processes
        if group_is_empty(group):
            break
        await asyncio.sleep(0.25)
    assert group_is_empty(group), "soffice survived the kill (orphan process)"


async def test_worker_restarts_after_being_killed(tmp_path: Path):
    client = OfficeClient(Settings(startup_timeout=120))
    try:
        await client.wait_until_ready()
        await client._kill()  # simulates a timeout / cancelled request
        target = tmp_path / "again.odt"
        reply = await client.call(protocol.KIND_CREATE, target, options={"overwrite": True})
        assert reply["saved"] and target.exists()
        assert client.restarts >= 1
    finally:
        await client.stop()
