# LibreOffice MCP 2.0

An [MCP](https://modelcontextprotocol.io) server that lets AI assistants (Claude Desktop, Claude
Code, and any other MCP client) **create, read and edit LibreOffice / Microsoft Office documents**
on your computer — no GUI, no copy-paste.

> Ask: *"Write a one-page project brief in `~/Documents/brief.docx` with a title, three sections, a
> risks table and page numbers in the footer, then export it to PDF."* — and the assistant builds it.

| | Status |
|---|---|
| **Writer** (text: `.odt` `.docx` `.doc` `.rtf` `.txt`) | ✅ 53 tools |
| **Calc** (spreadsheets: `.ods` `.xlsx` `.csv`) | 🔜 next |
| **Impress** (presentations: `.odp` `.pptx`) | 🔜 after Calc |
| Convert / export (PDF, docx↔odt, …), batch edits, document info | ✅ |

Ships as a **Docker image** (LibreOffice included), so it runs the same on Windows, macOS and Linux.

---

## Quick start

Everything (LibreOffice, its Python bridge, fonts) is inside the image; you only need Docker.

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker build -t libreoffice-mcp .
```

Then add it to your AI app. The folder you mount at `/data` is the **only** place the AI can read
and write; it refers to files by short paths like `report.docx`.

<details open>
<summary><b>Claude Desktop</b> — edit <code>claude_desktop_config.json</code></summary>

(Windows: `%APPDATA%\Claude\`, macOS: `~/Library/Application Support/Claude/`)

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "C:\Users\me\ai-workspace:/data", "libreoffice-mcp"]
    }
  }
}
```
Restart Claude Desktop. LibreOffice starts in the background (a couple of seconds); edits then take about 0.1–0.5 s.
</details>

<details>
<summary><b>Claude Code</b></summary>

```bash
claude mcp add libreoffice -- docker run -i --rm -v /path/to/workspace:/data libreoffice-mcp
```
</details>

<details>
<summary><b>Always-on HTTP server (docker compose)</b> — no start-up wait, one URL for any client</summary>

```bash
mkdir workspace && docker compose up -d --build     # endpoint: http://localhost:8000/mcp
claude mcp add --transport http libreoffice http://localhost:8000/mcp
```
There is **no authentication**: the compose file binds to `127.0.0.1` only. See
[docs/DOCKER.md](docs/DOCKER.md) for hosting it on a server and what to watch out for.
</details>

### Build and run it yourself with Docker

There is no published image: you build it once on your own machine (needs Docker, ~5 min the
first time because LibreOffice is downloaded; later builds are cached).

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp

# 1. Build the image
docker build -t libreoffice-mcp .

# 2a. Run it from an AI app (stdio): the app launches the container itself.
#     Use the "Claude Desktop" / "Claude Code" snippets above. To try it by hand:
docker run --rm libreoffice-mcp doctor

# 2b. ...or keep it running in the background (HTTP) with docker compose:
mkdir workspace                     # the folder the AI can read/write (shown as /data inside)
docker compose up -d --build        # builds the image if needed, then starts it
curl http://localhost:8000/healthz  # {"status":"ok"}
#   MCP endpoint for your client: http://localhost:8000/mcp
```

| Task | Command |
|---|---|
| See logs | `docker compose logs -f` |
| Stop | `docker compose down` |
| Update after `git pull` | `docker compose up -d --build` (or `docker build -t libreoffice-mcp .` for stdio) |
| Run the tests on Linux | `docker build --target test -t libreoffice-mcp:test . && docker run --rm libreoffice-mcp:test` |
| Remove everything | `docker compose down --rmi local` and delete the folder |

The compose file only listens on `127.0.0.1` and has **no authentication**; see
[docs/DOCKER.md](docs/DOCKER.md) before exposing it anywhere else.

---

## Choosing which tools the AI gets

More tools = more context the model must read. Expose only what you need:

| Goal | Flag |
|---|---|
| Only Word-processing tools | `--apps writer` |
| Writer + Calc, no presentations | `--apps writer,calc` |
| Let the AI *read* but never change files | `--read-only` |
| Hide risky tools | `--exclude "*delete*" --exclude "*replace*"` |
| Only a hand-picked set | `--include "writer_get_*" --include "writer_add_*"` |

By default **everything is enabled**. To narrow it down, add flags after the image name
(`… libreoffice-mcp --apps writer --read-only`); `/data` is already the root and the default
folder. `docker run --rm libreoffice-mcp tools --apps writer` prints exactly which tools a set of
flags exposes. Every flag also has a `LIBREOFFICE_MCP_*` environment variable.

`office_*` core tools (create, convert, info, batch edit…) are always available unless excluded.

---

## What the AI can do (Writer)

The full list with parameters is in **[docs/TOOLS.md](docs/TOOLS.md)** (generated from the code).

- **Read** — text (paged), structure (every paragraph/table with an index), outline, find, paragraph formatting, styles, tables, images, comments, bookmarks, links, footnotes, page setup, headers/footers, tracked changes
- **Write** — paragraphs, headings, bullet/numbered lists, whole sections in one call (`writer_add_content`), page breaks, merge another file
- **Format** — bold/italic/colour/font/size on specific text, paragraph alignment & spacing, create/modify styles
- **Tables** — create, edit cells, add/delete rows & columns, merge cells, borders, shading, header rows, column widths
- **Objects** — images (embedded, resizable, alt text), hyperlinks, bookmarks, footnotes, review comments, table of contents
- **Page** — paper size, orientation, margins, header/footer text with “Page N of M”
- **Review** — turn change tracking on/off, accept or reject all changes
- **Any document** — create, info, copy, **convert/export (PDF, docx, odt, html, txt…)**, and `office_batch_edit` (many edits, one load, one save, all-or-nothing)

Content is addressed by **block index** (0 = first paragraph or table). The AI calls
`writer_get_structure` to see the indexes, then edits by index — no fragile text matching.

---

## Safety

- **Your file is never half-written.** Each request edits a private copy and replaces your file atomically at the end. A crash, a timeout or a failing step in a batch leaves the original untouched.
- **No races.** Requests run one at a time, so two simultaneous edits to the same file never interleave (the second waits).
- **Macros never run**, even in documents that contain them.
- **Optional sandbox:** `--root` limits every read and write to chosen folders; `..` and symlink escapes are resolved first.
- **No lock files** are created next to your documents, and a stuck or cancelled request is killed cleanly (no leftover LibreOffice processes).
- There is **no undo**. Edits are saved immediately; keep backups of important files (or use `office_copy_document` first).

---

## Coming from v1?

v2 is a rewrite; tool names changed to be consistent (`office_*`, `writer_*`):

| v1 | v2 |
|---|---|
| `create_document` | `office_create_document` (type from file extension) |
| `open_text_document` | `writer_get_text` |
| `get_document_properties` | `office_get_document_info` |
| `list_documents` / `copy_document` | `office_list_documents` / `office_copy_document` |
| `add_text`, `add_paragraph` | `writer_add_paragraph` |
| `add_heading` | `writer_add_heading` |
| `add_table` / `format_table` | `writer_insert_table` / `writer_format_table` |
| `insert_image` / `insert_page_break` | `writer_insert_image` / `writer_insert_page_break` |
| `format_text` | `writer_format_text` |
| `search_replace_text` / `delete_text` | `writer_replace_text` (empty replacement deletes) |
| `delete_paragraph` | `writer_delete_blocks` |
| `create_custom_style` / `apply_document_style` | `writer_set_style` |

The old `libreoffice:{path}` MCP resource is not ported yet.

---

## Documentation

| | |
|---|---|
| [docs/DOCKER.md](docs/DOCKER.md) | Run with Docker, hosting, permissions |
| [AGENTS.md](AGENTS.md) | Repo map and rules for AI coding agents / new contributors |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | How it works, why it is built this way, file-by-file map |
| [docs/ADDING_A_TOOL.md](docs/ADDING_A_TOOL.md) | Step-by-step guide to add a tool (start here to contribute) |
| [docs/TOOLS.md](docs/TOOLS.md) | Every tool and parameter (generated) |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | Setup problems, logs, common errors |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Dev setup, tests, code style |

## Development

Tests run inside the Docker image (details in [CONTRIBUTING.md](CONTRIBUTING.md)):

```bash
docker build --target test -t libreoffice-mcp:test . && docker run --rm libreoffice-mcp:test
```

## Related projects

[nelson-mcp](https://github.com/quazardous/nelson-mcp) takes a different approach: an extension that
runs *inside* a running LibreOffice and edits the document you have open on screen. This project
works on files with no GUI (good for automation and batch work); pick the one that fits your workflow.

## License

[MIT](LICENSE)
