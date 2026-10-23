# Architecture

This page explains **how the server works and why it is built this way**. If you only want to
add a tool, jump to [ADDING_A_TOOL.md](ADDING_A_TOOL.md).

## The one-minute version

```
 AI app (Claude Desktop, Claude Code, …)
        │  MCP over stdio (or local HTTP)
        ▼
┌───────────────────────────────┐
│  MCP SERVER   (your normal    │   src/libreoffice_mcp/
│  Python 3.11+, the `mcp` SDK) │   - declares the tools the AI sees
│                               │   - validates arguments, sandboxes paths
│   OfficeClient (client.py)    │   - runs ONE request at a time (a lock)
└───────────────┬───────────────┘
                │  one JSON object per line over the worker's stdin/stdout
                ▼
┌───────────────────────────────┐
│  WORKER  (LibreOffice's own   │   src/libreoffice_mcp/worker/
│  Python — the only one with   │   - runs the LibreOffice logic ("operations")
│  the `uno` module)            │   - works on a scratch copy of your file
└───────────────┬───────────────┘
                │  UNO bridge over a private named pipe
                ▼
        soffice --headless          (one LibreOffice process, private profile)
```

## Why two Python programs?

LibreOffice is controlled through **UNO**, exposed to Python by a module called `uno`. In the
Docker image that module comes from the distro package `python3-uno` and is importable only from
the **system** `/usr/bin/python3`, not from the virtualenv that holds the `mcp` SDK. So:

* the **server** runs in the virtualenv, speaks MCP, and knows nothing about UNO;
* the **worker** runs on the system Python, speaks UNO, and knows nothing about MCP;
* they talk over a tiny line-based JSON protocol ([`protocol.py`](../src/libreoffice_mcp/protocol.py)).

The worker only uses the standard library. `locate.py` finds the right interpreter (or use
`LIBREOFFICE_MCP_PYTHON`).

> **v1 did something similar** with three processes and TCP ports 2002/8765. v2 uses one worker
> and stdin/stdout instead: no ports to collide on, nothing to forget to start, and the worker
> dies with the server.

## What happens on a tool call

Example: the AI calls `writer_add_paragraph(path="report.docx", text="Hello")`.

1. **Tool** (`tools/writer/content.py`) — the SDK validates the arguments against the function
   signature. The function body is one line: `edit(path, "writer.add_paragraph", text=text, …)`.
2. **Helper** (`tools/base.py`) — resolves `path` (relative → default folder, `~`, sandbox check),
   builds the request `{"kind": "edit", "path": …, "ops": [{"op": "writer.add_paragraph", …}]}`.
3. **OfficeClient** (`client.py`) — waits for the lock, sends the line to the worker, waits for the
   reply with a timeout.
4. **Worker runner** (`worker/runner.py`) — copies the file to a private scratch folder, opens the
   *copy* in LibreOffice (hidden, macros disabled), looks up `writer.add_paragraph` in the
   operation registry and calls it with the open document.
5. **Operation** (`worker/ops/writer/content.py`) — the actual UNO code.
6. **Save** — if any operation changed the document, it is stored to the scratch copy, which then
   **atomically replaces** your file (`os.replace`). The document is closed, scratch is deleted.
7. The reply travels back; the SDK turns it into the tool result.

`office_batch_edit` (and `writer_add_content`) run *many* operations in steps 4–6 with a single
load and a single save — much faster, and **all-or-nothing**.

## Running it: stdio, HTTP, Docker

* **stdio** (default): the AI app launches the server and talks over stdin/stdout.
* **HTTP** (`--transport http`): MCP "Streamable HTTP" at `/mcp`, stateless, with `/healthz` and
  `/readyz`. No authentication — see `http_server.py` and [DOCKER.md](DOCKER.md).
* **Docker** packages LibreOffice + fonts + the server; the mounted `/data` folder is the only
  thing the AI can touch (`LIBREOFFICE_MCP_ROOTS=/data`). `tini` is PID 1 so stray processes are
  reaped and stop signals reach LibreOffice.

## Safety and reliability

| Concern | How it is handled | Where |
|---|---|---|
| Corrupt/half-written files | Edit a scratch copy; replace the original atomically only on success | `worker/session.py` |
| Failure in a multi-step batch | Nothing is saved unless every operation succeeded | `worker/runner.py` |
| Concurrent calls / races | One `asyncio.Lock`: requests are processed strictly one at a time. LibreOffice serialises UNO calls internally anyway, so this costs no real throughput | `client.py` |
| Lock files next to your documents | LibreOffice never opens your file — only the scratch copy | `worker/session.py` |
| Malicious macros | `MacroExecutionMode = NEVER_EXECUTE` on every load | `worker/session.py` |
| Path traversal | `..`, `~`, symlinks resolved first, then checked against `--root` | `paths.py` |
| Hung LibreOffice | Per-request timeout (default 120 s) → worker is killed, original file untouched, next call restarts it | `client.py` |
| Client cancels a request | The worker is killed (UNO calls can't be interrupted safely) and restarted on demand | `client.py` |
| LibreOffice crashes | Worker notices, exits; the server restarts it and **retries the request once** (safe because the original is untouched) | `client.py`, `worker/main.py` |
| Orphan `soffice` processes | stdin EOF → worker shuts soffice down. soffice stays in the worker's process group, and the server kills that whole group on timeout/cancel, so nothing survives | `worker/office.py`, `client.py` |
| Interference from the user's own LibreOffice | Private temporary profile (`-env:UserInstallation`) and a unique pipe name | `worker/office.py` |
| Slow first request | LibreOffice is started in the background when the server starts (a few seconds); the MCP handshake never waits for it | `app.py`, `client.py` |

## Design decisions (and what we chose *not* to do)

* **One worker, serial requests.** An office suite is effectively single-threaded and an AI agent
  rarely needs parallelism. A pool of workers would add complexity (file locking across workers,
  memory use, start-up cost ×N) for little gain. If you ever need throughput, run several servers
  on different folders.
* **Stateless, file-based tools.** Every call names the file. There are no "open document handles"
  to leak or go stale, which also matches the 2026 MCP specification's move to stateless requests.
* **Block addressing.** Paragraphs and tables are addressed by 0-based index
  (`writer_get_structure` lists them). This is more reliable for an AI than "the paragraph that
  says…". Indexes shift after inserts/deletes — tools return the new index to help.
* **Typed tools + a generic batch tool.** Each capability is a separate, well-typed tool (best for
  the model's accuracy); `office_batch_edit` composes them for speed.
* **Explicit module-level tool functions.** No decorators that generate hidden code: a tool is a
  plain function you can read top to bottom.

## Source map

```
src/libreoffice_mcp/
  cli.py            command line: serve | doctor | tools
  config.py         flags + LIBREOFFICE_MCP_* env vars → Settings
  app.py            builds the MCP server, start-up/shutdown, the instructions the AI reads
  state.py          the shared Settings / PathPolicy / OfficeClient used by tools
  client.py         OfficeClient: starts/talks to the worker, lock, timeouts, restarts
  http_server.py    optional HTTP transport: /healthz, /readyz, Host check (no auth by design)
  paths.py          path resolution and the --root sandbox
  errors.py         OfficeError: messages meant for the AI (a ToolError)
  protocol.py       worker message format  (stdlib only, shared with the worker)
  formats.py        file extension → document kind → LibreOffice filter names (shared)
  locate.py         finds soffice and a Python that can `import uno`
  logging_setup.py  logging to stderr (stdout carries the MCP protocol!)
  tools/            ── what the AI sees ──
    registry.py       @tool decorator; applies --apps / --read-only / --include / --exclude
    base.py           read() / edit() / batch(): the only code tools use to reach the worker
    types.py          shared parameter types with descriptions
    core.py           office_* tools
    writer/           writer_* tools (read, content, tables, objects) + typed content blocks
  worker/           ── runs inside LibreOffice's Python (stdlib + uno only) ──
    __main__.py       entry point (puts the package on sys.path itself)
    main.py           stdin/stdout loop; exits if soffice dies
    office.py         starts/stops soffice
    session.py        scratch copy, load, atomic replace, Session object given to operations
    runner.py         one request: open → run ops → save → close
    registry.py       @op decorator, OpError
    util.py           UNO helpers (PropertyValue, colours…)
    ops/              ── the LibreOffice logic ──
      common.py         doc.* operations (info, metadata)
      writer/           writer.* operations: helpers, read, content, formatting, tables, page, objects, review
tests/              unit tests (no LibreOffice), worker tests, MCP-client tests (real LibreOffice)
scripts/            gen_tool_docs.py
Dockerfile          runtime image (+ `test` target); docker-compose.yml: long-running HTTP server
```

## The worker protocol (for the curious)

One JSON object per line. Server → worker:
`{"id": 7, "type": "run", "kind": "edit", "path": "…", "ops": [{"op": "writer.add_paragraph", "args": {…}}], "options": {…}}`.
Worker → server: `{"id": 7, "ok": true, "result": {…}}` or `{"id": 7, "ok": false, "error": {"code": "…", "message": "…", "op_index": 0}}`.
Request kinds: `read`, `edit`, `create`, `export`. Full description in
[`protocol.py`](../src/libreoffice_mcp/protocol.py).

## Roadmap

1. ✅ Foundation + Writer (this branch)
2. Calc: sheets, ranges, formulas, formatting, charts, named ranges, CSV
3. Impress: slides, text/shapes/images, notes, layouts, transitions
4. Document preview as an image so the AI can *see* its result (PDF → PNG)
5. Upload/download tools so a *hosted* server is usable by remote users (today files live on the server's disk)
