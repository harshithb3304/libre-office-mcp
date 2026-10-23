# Contributing

Thanks for helping! The fastest way to understand the project is to read
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) (10 minutes) and then
[docs/ADDING_A_TOOL.md](docs/ADDING_A_TOOL.md).

## Setup

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
uv sync                               # creates .venv with the dev tools
uv run libreoffice-mcp doctor         # is LibreOffice found?
uv run pytest                         # unit tests always run; LibreOffice tests auto-skip if absent
```

You need LibreOffice installed to run most tests (they drive a real instance — mocks of UNO
would hide exactly the bugs that matter). On Linux install `libreoffice-writer python3-uno`.

## Workflow

1. Branch from `main`.
2. Make the change; add or update tests (`tests/`). Test failure *messages* as well as success.
3. `uv run ruff check . && uv run ruff format .` and `uv run pytest`.
4. If you touched a tool: `uv run python scripts/gen_tool_docs.py` and commit `docs/TOOLS.md`.
5. Open a pull request describing *what* and *why*.

## Code style

* Python 3.11+ for the server; **the worker (`worker/`) must stay Python 3.10-compatible and
  standard-library-only** — it runs inside LibreOffice's bundled Python.
* `ruff` enforces formatting and common mistakes (line length 100).
* Comments explain **why**, especially around UNO quirks. Docstrings of tools are read by AI
  models — write them as instructions to a smart colleague.
* Error messages are part of the product: they are what the AI sees. Say what is wrong and how
  to fix the next call.
* Keep modules small and single-purpose; match the existing layout (`tools/<app>` ↔
  `worker/ops/<app>`).

## Reporting bugs

Include your OS, LibreOffice version, the exact tool call and the log lines (see
[docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)). A minimal document that reproduces it helps a lot.
