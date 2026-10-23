# Troubleshooting

Everything runs in Docker, so most problems are about the container, the mounted folder, or
the AI client's config.

## Is it working at all?

```bash
docker run --rm libreoffice-mcp doctor          # should print [ok] for soffice and python
docker compose ps                               # compose users: STATUS should be "healthy"
curl http://localhost:8000/readyz               # compose users: {"ready":true}
```

## Where are the logs?

* **compose:** `docker compose logs -f`
* **Claude Desktop (stdio):** its MCP log file, e.g. `mcp-server-libreoffice.log` in Claude's
  logs folder. **Claude Code:** `/mcp`, or run the same `docker run -i …` command by hand.

Everything goes to stderr (stdout carries the MCP protocol). Add `--log-level DEBUG` after the
image name for more detail. Each request logs one line with its duration, e.g.
`edit:writer.add_paragraph report.docx ok in 213 ms`; lines starting `worker` come from the
LibreOffice side.

## Common problems

**The AI app says the server failed to start / `docker run` exits immediately**
* Use `-i` and **not** `-t` in the `docker run` args.
* Is the image built? `docker images libreoffice-mcp`. Is Docker running?
* Run the exact command from your config in a terminal to see the error.

**“File not found” for a file you can see on your computer**
The AI only sees the folder mounted at `/data`. Put the file there (or mount the right folder)
and refer to it by a relative path like `report.docx`. Windows-style host paths don't exist in
the container.

**“Access to … is not allowed. Files must be inside: /data”**
The AI tried a path outside `/data`. Use relative paths or `/data/...`.

**Created files are owned by root / can't be edited (Linux hosts)**
The server runs as user `ubuntu` (uid 1000). Make the folder writable for it
(`chmod a+rwx workspace`) or run with `--user "$(id -u):$(id -g)"`.

**The AI can't save a file / my changes don't show up in Word or LibreOffice**
Close the document in the other app first. Word on Windows locks open files, so the AI's write
is refused ("file in use"); other apps keep showing their old copy until you reload the file.

**The first request is slow**
LibreOffice starts in the background when the container starts (about 2 s); the very first call
may wait for it. With `docker compose` it is already warm.

**“The operation took longer than 120s and was aborted”**
The request was too big or LibreOffice hung. The file was not modified. Try smaller steps, or
raise the limit with `--op-timeout 300` after the image name.

**HTTP 421 “Invalid Host header”** (compose / HTTP mode)
You reached the server through a hostname other than `localhost`. Add it:
`--allowed-host mcp.example.com`.

**“… needs a writer document, but x.ods is a calc document”**
You used a `writer_*` tool on a spreadsheet. (`calc_*` tools are not implemented yet.)

**The AI keeps editing the wrong paragraph**
Indexes shift after inserts/deletes. Tell it to call `writer_get_structure` again, or to use
`writer_add_content` / `office_batch_edit` so related edits happen in one go.

**Fonts look wrong in PDFs**
The image ships Liberation, DejaVu, Carlito and Caladea. For other fonts, extend the image
(`FROM libreoffice-mcp` + `apt-get install fonts-…`); for Chinese/Japanese/Korean install
`fonts-noto-cjk`.

## Still stuck?

Open an issue with: the output of `docker run --rm libreoffice-mcp doctor`, the tool call, and
the log lines around the failure (`office_status` shows the LibreOffice version).
