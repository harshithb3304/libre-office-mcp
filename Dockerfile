# syntax=docker/dockerfile:1
#
# LibreOffice MCP server image.
#
#   docker build -t libreoffice-mcp .                      # runtime image (default target)
#   docker build --target test -t libreoffice-mcp:test .   # image that runs the test suite
#
# Run it (stdio, for desktop AI apps):
#   docker run -i --rm -v /path/to/your/documents:/data libreoffice-mcp
# Run it (HTTP, for a long-lived container / hosted URL):
#   docker run -p 127.0.0.1:8000:8000 -v /path/to/your/documents:/data libreoffice-mcp --transport http
#
# Inside the container the AI can only see /data (your mounted folder).

# ---- base: LibreOffice + the server -------------------------------------------------
FROM ubuntu:24.04 AS base

ARG DEBIAN_FRONTEND=noninteractive
# - libreoffice-*   : the office apps (headless; no GUI packages are pulled in)
# - python3-uno     : the UNO bridge; the worker runs on the system python3 with it
# - fonts-*         : without fonts PDFs/previews look wrong. Carlito/Caladea are metric-
#                     compatible with Calibri/Cambria, so .docx files keep their layout.
# - tini            : PID 1 that reaps zombie processes and forwards stop signals to soffice
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-calc libreoffice-impress \
        python3 python3-uno ca-certificates tini \
        fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea \
    && rm -rf /var/lib/apt/lists/*

COPY --from=ghcr.io/astral-sh/uv:0.6 /uv /usr/local/bin/uv

# Use the system Python for the server too (never download one), keep the venv outside /app.
ENV UV_PYTHON=/usr/bin/python3 \
    UV_PYTHON_DOWNLOADS=never \
    UV_LINK_MODE=copy \
    UV_NO_CACHE=1 \
    UV_COMPILE_BYTECODE=1 \
    UV_PROJECT_ENVIRONMENT=/opt/venv

WORKDIR /app
# Dependencies first so this layer is cached until pyproject/uv.lock change.
COPY pyproject.toml uv.lock README.md LICENSE ./
RUN uv sync --frozen --no-dev --no-install-project
COPY src ./src
RUN uv sync --frozen --no-dev

# ---- test: the same image plus dev tools and the test suite ---------------------------
FROM base AS test
RUN uv sync --frozen
COPY tests ./tests
COPY scripts ./scripts
ENV PATH="/opt/venv/bin:$PATH" PYTEST_ADDOPTS="-p no:cacheprovider"
USER ubuntu
ENTRYPOINT ["tini", "--"]
CMD ["pytest", "-q"]

# ---- runtime: what you actually run (last stage = default build target) ---------------
FROM base AS runtime
RUN mkdir -p /data && chown ubuntu:ubuntu /data
ENV PATH="/opt/venv/bin:$PATH" \
    HOME=/home/ubuntu \
    LIBREOFFICE_MCP_PYTHON=/usr/bin/python3 \
    LIBREOFFICE_MCP_DEFAULT_DIR=/data \
    LIBREOFFICE_MCP_ROOTS=/data \
    LIBREOFFICE_MCP_HOST=0.0.0.0
USER ubuntu
WORKDIR /data
VOLUME /data
EXPOSE 8000
# No arguments = stdio transport. Add `--transport http` for a network server.
ENTRYPOINT ["tini", "--", "libreoffice-mcp"]
