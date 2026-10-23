# Troubleshooting

First, run the built-in check:

```bash
uv run libreoffice-mcp doctor
```

It tells you whether LibreOffice and a Python that can `import uno` were found.

## Where are the logs?

All logs go to **stderr** (stdout carries the MCP protocol). Your MCP client captures them:

* **Claude Desktop** — `%APPDATA%\Claude\logs\mcp-server-libreoffice.log` (Windows),
  `~/Library/Logs/Claude/` (macOS).
* **Claude Code** — `claude mcp list` / `/mcp`; run the command by hand to see the output.

Set `--log-level DEBUG` (or `LIBREOFFICE_MCP_LOG_LEVEL=DEBUG`) for more detail. Lines starting
`worker` come from the LibreOffice side. Each request logs one line with its duration, e.g.
`edit:writer.add_paragraph report.docx ok in 213 ms`.

## Common problems

**“LibreOffice not found”**
Install LibreOffice, or point to it: `--soffice "C:\Program Files\LibreOffice\program\soffice.exe"`
(env `LIBREOFFICE_MCP_SOFFICE`).

**“No Python with the 'uno' module was found”**
* *Linux:* `sudo apt install python3-uno` (Debian/Ubuntu) or `sudo dnf install python3-libreoffice`
  (Fedora). Distro packages install `uno` for the **system** Python; a virtualenv can't see it, but
  the worker doesn't need the virtualenv — `locate.py` finds `/usr/bin/python3` by itself.
* *Windows/macOS:* the Python bundled next to `soffice` is used; if it's missing, reinstall
  LibreOffice with the default (full) options.
* Or set `--python` / `LIBREOFFICE_MCP_PYTHON` to any Python that can `import uno`.

**The first request is very slow / “LibreOffice failed to start”**
Starting LibreOffice takes 10–30 s (longer with antivirus scanning it). It starts in the
background when the server starts, so by the time you ask something it is usually ready. If it
times out, raise `--startup-timeout 180`. A startup failure usually logs `soffice stderr: …`.

**“The operation took longer than 120s and was aborted”**
The request was too large (huge document or enormous batch) or LibreOffice hung. The file was not
modified. Try smaller steps or raise `--op-timeout`.

**“Could not write …: the file is in use by another program”** (Windows)
Another program (Word, an indexer, antivirus) has the file open without sharing. Close it and retry.

**“Access to … is not allowed. Files must be inside: …”**
You started the server with `--root`. Put the file inside that folder or widen the root.

**Relative paths end up in the wrong folder**
Relative paths resolve against `--default-dir` (default `~/Documents`). Check `office_status`.

**“… needs a writer document, but x.ods is a calc document”**
You used a `writer_*` tool on a spreadsheet. (`calc_*` tools are not implemented yet.)

**The AI keeps editing the wrong paragraph**
Indexes shift after inserts/deletes. Tell it to call `writer_get_structure` again, or use
`writer_add_content` / `office_batch_edit` to do related edits in one go.

**Leftover `soffice` processes**
They shouldn't exist: they are tied to the worker, which is tied to the server. If you killed the
server with `kill -9` on Linux and see strays, `pkill -f lomcp_profile`.

## Still stuck?

Open an issue with: OS, LibreOffice version (`office_status` shows it), the output of
`libreoffice-mcp doctor`, and the log lines around the failure.
