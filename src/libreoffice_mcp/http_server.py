"""The optional HTTP transport (``--transport http``), used by Docker and hosted setups.

Desktop AI apps normally talk to the server over **stdio** and never touch this file.
HTTP mode serves MCP "Streamable HTTP" at ``/mcp`` and adds:

* ``/healthz`` (process is up) and ``/readyz`` (LibreOffice is up) for health checks;
* **Host header checking** (the SDK's DNS-rebinding protection): ``localhost`` is always
  accepted; if you reach the server through a real hostname, add it with ``--allowed-host``.

**There is no authentication.** Anyone who can reach the port can read and edit the files in
the workspace. The provided docker-compose file binds to 127.0.0.1 only; if you expose the
server publicly, put it behind something that authenticates (a reverse proxy, a VPN, a
private network).

The server is *stateless*: no MCP sessions are kept, so a restart loses nothing.
"""

from __future__ import annotations

import ipaddress
import logging
from typing import Any

import uvicorn
from mcp.server.mcpserver import MCPServer
from mcp.server.transport_security import TransportSecuritySettings
from starlette.requests import Request
from starlette.responses import JSONResponse, Response

from libreoffice_mcp.config import Settings
from libreoffice_mcp.state import state

log = logging.getLogger("libreoffice_mcp.http")


def is_loopback(host: str) -> bool:
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def register_health_routes(mcp: MCPServer) -> None:
    """Add ``/healthz`` and ``/readyz`` (they reveal nothing about documents)."""

    @mcp.custom_route("/healthz", methods=["GET"])
    async def healthz(_: Request) -> Response:
        return JSONResponse({"status": "ok"})

    @mcp.custom_route("/readyz", methods=["GET"])
    async def readyz(_: Request) -> Response:
        ready = bool(state.office.status()["ready"])
        return JSONResponse({"ready": ready}, status_code=200 if ready else 503)


def _security(settings: Settings) -> TransportSecuritySettings:
    if "*" in settings.allowed_hosts:
        return TransportSecuritySettings(enable_dns_rebinding_protection=False)
    hosts = ["127.0.0.1:*", "localhost:*", "[::1]:*", "127.0.0.1", "localhost"]
    for host in settings.allowed_hosts:
        hosts += [host, f"{host}:*"]
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True, allowed_hosts=hosts, allowed_origins=[]
    )


def build_http_app(mcp: MCPServer, settings: Settings) -> Any:
    return mcp.streamable_http_app(
        json_response=True,  # plain JSON replies: simplest for clients and proxies
        stateless_http=True,  # no sessions to track or lose
        transport_security=_security(settings),
    )


def run_http(mcp: MCPServer, settings: Settings) -> None:
    if not is_loopback(settings.host):
        log.warning(
            "listening on %s with NO authentication: only expose this to trusted networks",
            settings.host,
        )
    app = build_http_app(mcp, settings)
    log.info("serving MCP at http://%s:%d/mcp", settings.host, settings.port)
    # SIGTERM/SIGINT -> uvicorn stops -> app lifespan shuts LibreOffice down cleanly.
    uvicorn.run(
        app,
        host=settings.host,
        port=settings.port,
        log_level=settings.log_level.lower(),
        access_log=False,
        timeout_graceful_shutdown=20,
    )
