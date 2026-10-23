# LibreOffice MCP

An [MCP](https://modelcontextprotocol.io) server that lets AI assistants (Claude Code, Claude
Desktop, and any other MCP client) **create, read and edit documents with LibreOffice**:
text documents, spreadsheets and presentations — no GUI, no copy-paste.

| | Tools |
|---|---|
| **Writer** — text documents (`.odt` `.docx` `.rtf` `.txt`) | ✅ 55 |
| **Calc** — spreadsheets (`.ods` `.xlsx` `.csv`) | ✅ 51 |
| **Impress** — presentations (`.odp` `.pptx`) | ✅ 42 |
| Create, info, copy, **convert / export to PDF**, batch edits | ✅ 8 |

## Quick start

All you need is [Docker](https://www.docker.com/products/docker-desktop/). Three commands:

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker compose up -d --build          # builds the image and starts the server
claude mcp add --transport http libreoffice http://localhost:8000/mcp
```

That's it. **Every tool is enabled out of the box.** The first build takes a few minutes (it
downloads LibreOffice); after that the server starts in seconds. Your files live in the
`workspace/` folder next to `docker-compose.yml` (see [The workspace folder](#the-workspace-folder)).

### Try it — one write, one read

Ask your AI:

> *Create `notes.docx` with the heading "Groceries" and a bullet list of milk, eggs and bread.
> Then read the file back and tell me what is in it.*

It writes the document (`writer_add_content`), reads it back (`writer_get_text`), and
`workspace/notes.docx` appears on your computer. Spreadsheets and slides work the same way:
*"Make `budget.xlsx` with 4 months of income and expenses and a chart"*, or *"Make a 3-slide
deck about our Q3 plan and export it to PDF."*

### Using another MCP client

Any client that accepts an MCP server URL works. In the standard MCP JSON:

```json
{
  "mcpServers": {
    "libreoffice": { "url": "http://localhost:8000/mcp" }
  }
}
```

(Some clients also want `"type": "http"` next to the `url`. Where the file lives depends on the
client — Claude Desktop, Cursor, Windsurf, Cline and others each document it. A few clients, e.g.
VS Code, name the top-level key `servers` instead of `mcpServers`.)

### Prefer no background server?

Let your MCP client start the container itself instead. Build the image once, then give the
client this command (standard MCP JSON; on Windows double the backslashes in the folder path):

```bash
docker build -t libreoffice-mcp .
```

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "/path/to/my-workspace:/data", "libreoffice-mcp"]
    }
  }
}
```

For Claude Code: `claude mcp add libreoffice -- docker run -i --rm -v /path/to/my-workspace:/data libreoffice-mcp`.
`-i` is required; do **not** add `-t`.

### Everyday commands

| Task | Command |
|---|---|
| See logs | `docker compose logs -f` |
| Stop | `docker compose down` |
| Start again | `docker compose up -d` |
| Update after `git pull` | `docker compose up -d --build` |
| Use another workspace folder | `WORKSPACE=/path/to/folder docker compose up -d` |
| Run the tests (on Linux, inside Docker) | `docker build --target test -t libreoffice-mcp:test . && docker run --rm libreoffice-mcp:test` |

The server listens on `127.0.0.1` only and has **no authentication**, so just your computer can
reach it. Read [docs/DOCKER.md](docs/DOCKER.md) before exposing it further.

## How it works

* The **Docker container *is* the MCP server**: LibreOffice, its Python bridge and fonts are
  inside the image. Your AI app only needs to be told where to find it.
* The AI can only see **one folder** of yours — the workspace — which appears as `/data` inside
  the container.
* Every edit is done by LibreOffice itself and saved atomically.

---

## The workspace folder

The container cannot see your computer; it only sees the folder you mount. We call it the
**workspace**, and inside the container it appears as `/data`.

* **Relative paths** the AI uses (`report.docx`, `clients/acme.docx`) resolve to the workspace.
  It is *not* your terminal's current directory unless you mount it (`-v "$(pwd):/data"`).
* **Absolute paths** work only inside `/data`; anything else is refused. Paths from your own
  computer (`C:\Users\…`) don't exist in the container, so just give file names.
* It lives **on your disk**, not in the image or container. Stopping or deleting the container
  does not touch it, and the image does not grow as you create documents.
* To work on an existing file, copy it into the workspace first.
* **Close a file in other programs before the AI edits it.** Some programs lock open files (the
  write is then refused); others keep showing the old content until you reload.

More detail and permissions on Linux hosts: [docs/DOCKER.md](docs/DOCKER.md).

---

## Choosing which tools the AI gets

By default **everything is enabled**. More tools means more for the model to read, so you can
narrow it down by adding flags after the image name:

| Goal | Flag |
|---|---|
| Only text-document tools | `--apps writer` |
| Writer + Calc, no presentations | `--apps writer,calc` |
| Let the AI *read* but never change files | `--read-only` |
| Hide risky tools | `--exclude "*delete*" --exclude "*replace*"` |
| Only a hand-picked set | `--include "writer_get_*" --include "writer_add_*"` |

Example: `claude mcp add libreoffice -- docker run -i --rm -v /path/ws:/data libreoffice-mcp --apps writer`.
`docker run --rm libreoffice-mcp tools --apps writer` prints exactly which tools a set of flags
exposes. Every flag also has a `LIBREOFFICE_MCP_*` environment variable. The `office_*` core tools
(create, convert, info, batch edit…) stay available unless excluded.

---

## What the AI can do

The full list with parameters is in **[docs/TOOLS.md](docs/TOOLS.md)** (generated from the code).

**Writer**
- **Read** — text (paged), structure (every paragraph/table with an index), outline, find, paragraph formatting, styles, tables, images, comments, bookmarks, links, footnotes, page setup, headers/footers, tracked changes
- **Write** — paragraphs, headings, bullet/numbered lists, whole sections in one call (`writer_add_content`), page breaks, merge another file
- **Format** — bold/italic/colour/font/size on specific text, paragraph alignment & spacing, create/modify styles
- **Tables** — create, edit cells, add/delete rows & columns, merge cells, borders, shading, header rows, column widths
- **Objects** — images (embedded, resizable, alt text), hyperlinks, bookmarks, footnotes, review comments, table of contents
- **Page** — paper size, orientation, margins, header/footer text with “Page N of M”
- **Review** — turn change tracking on/off, accept or reject all changes

**Calc**
- **Build** — a whole table in one call (`calc_write_table`: header, rows, number formats, auto-width, filter), values and formulas (the usual `=SUM(A1,B1)` syntax works), fill/copy/sort/find & replace, insert/delete rows and columns, merge cells
- **Format** — fonts, colours, number formats, borders, column widths, conditional formatting and data bars, data validation
- **Charts and objects** — column/bar/line/pie/scatter/area charts, images, hyperlinks, cell comments
- **Workbook** — add/rename/move/copy/hide sheets, named ranges, freeze panes (xlsx), autofilter, CSV

**Impress**
- **Build a deck in one call** (`impress_add_slides`: title, bullets with levels, two columns, image, table, notes), or add/duplicate/move/delete slides one by one
- **Shapes and text** — text boxes, 37 shape types, lines and arrows, images, tables, fills/outlines/shadows, text formatting, hyperlinks, alignment and z-order
- **Look and feel** — whole-deck themes (`impress_apply_theme`), master styles, slide size, backgrounds, transitions, footers, speaker notes

**Any document** — create, info, copy, **convert/export (PDF, docx, odt, html, txt…)**, and
`office_batch_edit` (many edits, one load, one save, all-or-nothing).

Writer content is addressed by **block index** (0 = first paragraph or table): the AI calls
`writer_get_structure` to see the indexes, then edits by index — no fragile text matching.

---

## Safety

- **Your file is never half-written.** Each request edits a private copy and replaces your file atomically at the end. A crash, a timeout or a failing step in a batch leaves the original untouched.
- **No races.** Requests run one at a time, so two simultaneous edits to the same file never interleave (the second waits).
- **Macros never run**, even in documents that contain them.
- **The container only sees the workspace** (`/data`); paths outside it, `..` and symlink escapes are refused.
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
| [docs/DOCKER.md](docs/DOCKER.md) | Docker in detail: options, workspace, hosting, permissions |
| [AGENTS.md](AGENTS.md) | Repo map and rules for AI coding agents / new contributors |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | How it works, why it is built this way, file-by-file map |
| [docs/ADDING_A_TOOL.md](docs/ADDING_A_TOOL.md) | Step-by-step guide to add a tool (start here to contribute) |
| [docs/TOOLS.md](docs/TOOLS.md) | Every tool and parameter (generated) |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | Problems, logs, common errors |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Dev setup, tests, code style |


## License

[MIT](LICENSE)
