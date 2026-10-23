# LibreOffice MCP 2.0

An [MCP](https://modelcontextprotocol.io) server that lets AI assistants (Claude Desktop, Claude
Code, and any other MCP client) **create, read and edit documents with LibreOffice** (Writer, Calc, Impress)
— no GUI, no copy-paste.

> Ask: *"Write a one-page project brief as `brief.docx` with a title, three sections, a risks
> table and page numbers in the footer, then export it to PDF."* — and the assistant builds it.

| | Status |
|---|---|
| **Writer** (text: `.odt` `.docx` `.doc` `.rtf` `.txt`) | ✅ 53 tools |
| **Calc** (spreadsheets: `.ods` `.xlsx` `.csv`) | 🚧 in progress |
| **Impress** (presentations: `.odp` `.pptx`) | 🚧 in progress |
| Convert / export (PDF, docx↔odt, …), batch edits, document info | ✅ |

---

## How it works (read this first)

* The **Docker container *is* the MCP server.** LibreOffice, its Python bridge and fonts are all
  inside the image. You build the image once.
* Your AI app (Claude Desktop, Claude Code…) only needs to be **told how to reach it**
  (`claude mcp add …` or a few lines of JSON). It does not install anything else.
* The AI can only see **one folder** on your computer: the *workspace* you mount into the container
  as `/data`. See [The workspace folder](#the-workspace-folder).

## Quick start

### Step 1 — build the image (once)

Needs [Docker](https://www.docker.com/products/docker-desktop/). The first build takes a few
minutes (it downloads LibreOffice); later builds are cached.

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker build -t libreoffice-mcp .
docker run --rm libreoffice-mcp doctor      # optional check: should print [ok] twice
```

### Step 2 — connect your AI app (pick **one** option)

**Option A — the AI app starts the container itself** (simplest; nothing stays running).
The part after `-v` is `<a folder on your computer>:/data`.

<details open>
<summary><b>Claude Code</b></summary>

```bash
claude mcp add libreoffice -- docker run -i --rm -v /path/to/my-workspace:/data libreoffice-mcp
```
</details>

<details>
<summary><b>Claude Desktop</b> — add to <code>claude_desktop_config.json</code></summary>

(Windows: `%APPDATA%\Claude\`, macOS: `~/Library/Application Support/Claude/`)

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "C:\\Users\\me\\my-workspace:/data", "libreoffice-mcp"]
    }
  }
}
```
Restart Claude Desktop. (On Windows, double every backslash in JSON.)
</details>

**Option B — keep the server running in the background** (docker compose; instant responses, one
URL for any client). Start it, then just register its address:

```bash
docker compose up -d --build        # starts the server on http://localhost:8000/mcp
claude mcp add --transport http libreoffice http://localhost:8000/mcp
```

Compose uses the folder `./workspace` next to `docker-compose.yml` (created automatically). To
use a different folder: `WORKSPACE=/path/to/my-workspace docker compose up -d`
(or put `WORKSPACE=...` in a `.env` file next to the compose file).

There is **no authentication**; the port is published on `127.0.0.1` only, so just your
computer can reach it. Read [docs/DOCKER.md](docs/DOCKER.md) before exposing it further.

### Step 3 — try it

Ask your AI: *"Create `hello.docx` with a heading and a three-item bullet list."*
Then open `hello.docx` from your workspace folder.

### Everyday commands

| Task | Command |
|---|---|
| See logs (Option B) | `docker compose logs -f` |
| Stop (Option B) | `docker compose down` |
| Update after `git pull` | `docker build -t libreoffice-mcp .` (Option B: `docker compose up -d --build`) |
| Run the tests (on Linux, inside Docker) | `docker build --target test -t libreoffice-mcp:test . && docker run --rm libreoffice-mcp:test` |
| Remove the image | `docker rmi libreoffice-mcp` |

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
| Only Word-processing tools | `--apps writer` |
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

## Related projects

[nelson-mcp](https://github.com/quazardous/nelson-mcp) takes a different approach: an extension that
runs *inside* a running LibreOffice and edits the document you have open on screen. This project
works on files with no GUI (good for automation and batch work); pick the one that fits your workflow.

## License

[MIT](LICENSE)
