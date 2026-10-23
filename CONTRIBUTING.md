# Contributing

Thanks for helping! The fastest way to understand the project is [AGENTS.md](AGENTS.md)
(what is where, rules that are easy to break), then [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
and [docs/ADDING_A_TOOL.md](docs/ADDING_A_TOOL.md).

## Setup

Docker is the only supported environment, and the tests run inside it (LibreOffice is in the image):

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker build --target test -t libreoffice-mcp:test .    # builds the image with the test suite
docker run --rm libreoffice-mcp:test                     # runs all tests (~30 s)
```

Re-run the two commands after changing code. For a quicker loop, mount your working copy:

```bash
docker run --rm -v "$(pwd)/src:/app/src" -v "$(pwd)/tests:/app/tests" libreoffice-mcp:test pytest tests/test_writer.py -k table -q
```

Linting and formatting don't need Linux: `uv sync && uv run ruff check . && uv run ruff format .`

## Workflow

1. Branch from `main`.
2. Make the change; add or update tests (`tests/`). Test failure *messages* as well as success.
3. Lint/format, then run the full test suite in Docker.
4. If you touched a tool, run `uv run python scripts/gen_tool_docs.py` and commit `docs/TOOLS.md`
   (it builds the tool list without LibreOffice).
5. Open a pull request describing *what* and *why*.

## Code style

* Python 3.11+ for the server; **the worker (`worker/`) stays standard-library-only and
  Python 3.10-compatible**.
* `ruff` enforces formatting and common mistakes (line length 100).
* Comments explain **why**, especially around UNO quirks. Tool docstrings are read by AI
  models — write them as instructions to a smart colleague.
* Error messages are part of the product: they are what the AI sees. Say what is wrong and how
  to fix the next call.
* Keep modules small and single-purpose; match the layout (`tools/<app>` ↔ `worker/ops/<app>`).

## Reporting bugs

Include the tool call, the log lines (see [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md))
and, if you can, a minimal document that reproduces it.
