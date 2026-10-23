"""HTTP transport: health endpoints, Host checking and a real MCP client over HTTP."""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
from collections.abc import Iterator

import httpx2
import pytest
from mcp import Client
from mcp.client.streamable_http import streamable_http_client

from libreoffice_mcp.config import Settings
from libreoffice_mcp.http_server import _security, is_loopback


def test_loopback_detection():
    assert all(is_loopback(h) for h in ("127.0.0.1", "localhost", "::1"))
    assert not any(is_loopback(h) for h in ("0.0.0.0", "10.0.0.5", "example.com"))


def test_allowed_hosts_always_include_localhost_and_add_configured_ones():
    hosts = _security(Settings(allowed_hosts=("mcp.example.com",))).allowed_hosts
    assert "localhost:*" in hosts and "mcp.example.com" in hosts and "mcp.example.com:*" in hosts


def test_star_disables_the_host_check():
    assert _security(Settings(allowed_hosts=("*",))).enable_dns_rebinding_protection is False


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="module")
def server() -> Iterator[str]:
    """A real server process serving HTTP on a free local port."""
    port = _free_port()
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "libreoffice_mcp",
            "--transport",
            "http",
            "--port",
            str(port),
            "--apps",
            "writer",
            "--log-level",
            "WARNING",
        ],
        env=os.environ.copy(),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    base = f"http://127.0.0.1:{port}"
    try:
        deadline = time.monotonic() + 30
        while time.monotonic() < deadline:
            try:
                if httpx2.get(f"{base}/healthz", timeout=1).status_code == 200:
                    break
            except httpx2.HTTPError:
                time.sleep(0.2)
        else:
            pytest.fail("HTTP server did not start")
        yield base
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()


def test_health_endpoints(server: str):
    assert httpx2.get(f"{server}/healthz").json() == {"status": "ok"}
    assert httpx2.get(f"{server}/readyz").status_code in (200, 503)  # 503 while LibreOffice boots


def test_foreign_host_header_is_rejected(server: str):
    r = httpx2.post(
        f"{server}/mcp",
        json={"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}},
        headers={"Host": "evil.example.com"},
    )
    assert r.status_code in (400, 403, 421)


async def test_a_real_mcp_client_can_list_tools_over_http(server: str):
    async with (
        httpx2.AsyncClient() as http,
        Client(streamable_http_client(f"{server}/mcp", http_client=http)) as client,
    ):
        names = {t.name for t in (await client.list_tools()).tools}
    assert "writer_add_paragraph" in names and "office_create_document" in names
    assert not any(n.startswith("calc_") for n in names)  # --apps writer was honoured
