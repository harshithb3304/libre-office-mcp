# AGENTS.md — guide for AI coding agents and new contributors

An MCP server that lets AI assistants create and edit LibreOffice/Office documents.
**Implemented: Writer (53 tools) + core `office_*` tools. Planned: Calc, then Impress.**
Humans: start with [README.md](README.md); this file is a map of *what is where* and the rules
that are easy to break.

## The architecture in 5 lines

```
AI client ──MCP (stdio | http)──▶ server (normal Python, `mcp` SDK v2)
                                    │  one JSON object per line on stdin/stdout
                                    ▼
                                  worker (LibreOffice's own Python: stdlib + `uno` only)
                                    │  UNO over a private named pipe
                                    ▼
                                  soffice --headless
```

* **Two Python programs** because only LibreOffice's Python can `import uno`
  (see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)).
* **One request at a time** (an `asyncio.Lock` in `client.py`). No pool, no sessions.
* **Every edit** = copy file to scratch → edit the copy → atomically replace the original.
  A batch is all-or-nothing.
* **Docker is the only supported way to run it** (`Dockerfile`, `docker-compose.yml`). The code
  assumes Linux (the image's Ubuntu); there is no Windows/macOS-specific code.

## Where things are

| You want to… | Look in |
|---|---|
| add/change what the AI **sees** (tool name, params, description) | `src/libreoffice_mcp/tools/<app>/` |
| add/change the **LibreOffice logic** (UNO calls) | `src/libreoffice_mcp/worker/ops/<app>/` |
| change how requests run (open → ops → save) | `worker/runner.py`, `worker/session.py` |
| change process start/stop, crash handling, timeouts | `client.py` (server side), `worker/office.py` + `worker/main.py` |
| change flags / env vars | `config.py` |
| change HTTP mode (health, host check) | `http_server.py` |
| change path rules / `--root` sandbox | `paths.py` |
| change what the AI is told on connect | `INSTRUCTIONS` in `app.py` |
| supported file types and LibreOffice filter names | `formats.py` |
| find LibreOffice / its Python | `locate.py` |
| container image | `Dockerfile`, `docker-compose.yml` |

Full file-by-file map: [docs/ARCHITECTURE.md § Source map](docs/ARCHITECTURE.md).

**Tool ↔ operation naming:** MCP tool `writer_add_paragraph` calls worker op
`writer.add_paragraph` (first `_` becomes `.`). `office_batch_edit` accepts op names directly.

## Adding a tool (the 80% case)

Follow [docs/ADDING_A_TOOL.md](docs/ADDING_A_TOOL.md). In short:

1. Operation in `worker/ops/<app>/…` with `@op("app.name", kind="app", mutates=True|False)`;
   signature `fn(session, **args)`; raise `OpError("message for the AI")` to fail.
2. Tool in `tools/<app>/…` with `@tool("app", "read|edit|destructive|create")`; body is one line
   (`return await edit(path, "app.name", ...)`).
3. Test in `tests/` (real LibreOffice), including the failure message.
4. `uv run python scripts/gen_tool_docs.py` to refresh `docs/TOOLS.md`.

`@op` and `@tool` only *register* a function in a dict (`OPS` / `TOOLS`); nothing is generated.

## Commands

```bash
docker build -t libreoffice-mcp .                      # build the image
docker build --target test -t libreoffice-mcp:test . && docker run --rm libreoffice-mcp:test   # ALL tests, in Linux
docker run --rm libreoffice-mcp:test pytest tests/test_unit.py -q        # fast subset
docker compose up -d --build                           # run the server (http://localhost:8000/mcp)
# Lint/format need no Linux:  uv sync && uv run ruff check . && uv run ruff format .
```

## Rules that are easy to break

1. **`worker/` must be stdlib-only and Python 3.10-compatible.** It runs inside LibreOffice's
   bundled Python (3.10 in LibreOffice 25.2). No third-party imports, no `match` statements, keep
   `from __future__ import annotations`.
2. **Never write to stdout** in server or worker code. For stdio, stdout *is* the MCP/worker
   protocol. Use `logging` (it goes to stderr). The worker re-points fd 1 at stderr for safety.
3. **Raise `OfficeError` (server) / `OpError` (worker) for expected failures** and write the message
   for the AI: what is wrong *and* how to fix the next call. The MCP SDK hides the text of any
   other exception from the model.
4. **Operation arguments named `*_path` are existing files** and are sandbox-checked
   automatically (`tools/base.py`). Paths nested inside structured args (like images in
   `writer_add_content` blocks) must be resolved by hand — see `tools/writer/blocks.py`.
5. **Don't save or close documents inside an operation.** The runner does it (and decides whether
   to save from `mutates` / `_noop`).
6. **Content is addressed by block index** (0-based; paragraph or table). Indexes shift after
   inserts/deletes — return new indexes from mutating ops.
7. **UNO lengths are 1/100 mm** (`helpers.mm()` converts). Enum properties need `uno.Enum(...)`.
8. **Test against real LibreOffice.** UNO has quirks no mock reproduces (examples already handled:
   paragraphs lack `setAllPropertiesToDefault`, footnotes must be inserted before setting text,
   merged cells break `getColumns().getCount()`, a failed range-comment insert leaves a ghost).
9. **Process groups matter.** The server starts the worker in its own group and `_kill()` kills the
   whole group; soffice must stay in the worker's group (don't give it `start_new_session`) or a
   killed worker leaves an orphan `soffice` behind.
10. Keep worker ops and tool signatures in sync (same parameter names). A test that calls the op
    through the MCP client (`tests/test_server.py`) catches drift.

## Style

* `ruff` (line length 100) is the formatter/linter; run it before committing.
* Comments explain *why*, especially around UNO and platform quirks. Tool docstrings are
  instructions to an AI model — write them that way.
* Small modules, one responsibility each; mirror `tools/<app>` ↔ `worker/ops/<app>`.
* Commit messages: imperative summary + a body saying *why*.

## Current status / next steps

* ✅ Foundation, Writer, Docker image/compose, HTTP transport (no auth by design — see
  `http_server.py` docstring), docs.
* ⏭ Calc (`calc_*`): sheets, ranges, formulas, formatting, charts, CSV.
* ⏭ Impress (`impress_*`): slides, text/shapes/images, notes.
* Ideas: render a page/slide to PNG so the AI can *see* results; MCP resource for document text;
  keep documents open between calls for very large files.
