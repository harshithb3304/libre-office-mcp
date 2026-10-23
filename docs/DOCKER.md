# Running with Docker (recommended)

Docker is the easiest way to run this on any OS: LibreOffice, its Python bridge and the fonts
are all inside the image, so there is nothing to install or configure on your machine except
Docker itself.

**The one thing to understand:** the AI can only see the folder you mount into the container at
`/data`. Everything else on your computer is invisible to it. You refer to files by short paths
(`report.docx`, `clients/acme.docx`); `/data` is the default location for relative paths.

## Build the image

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker build -t libreoffice-mcp .
```

The first build downloads LibreOffice (several hundred MB); later builds are cached.

## Option A — stdio (Claude Desktop, Claude Code, most desktop AI apps)

The AI app starts the container itself and talks to it over stdin/stdout.

**Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "docker",
      "args": ["run", "-i", "--rm",
               "-v", "C:\\Users\\me\\Documents\\ai-workspace:/data",
               "libreoffice-mcp"]
    }
  }
}
```

Use your own folder on the left of the colon (macOS/Linux: `/Users/me/ai-workspace:/data`).
`-i` is required (it keeps stdin open); do **not** add `-t`.

**Claude Code:**

```bash
claude mcp add libreoffice -- docker run -i --rm -v /path/to/workspace:/data libreoffice-mcp
```

Each time the AI app starts, a fresh container starts and LibreOffice boots in the background
(about 2 s on Linux/Docker; the first request waits for it). For zero wait, use Option B.

Expose fewer tools by adding flags after the image name:
`... libreoffice-mcp --apps writer` or `... libreoffice-mcp --read-only`.

## Option B — a long-running HTTP server (docker compose)

```bash
mkdir workspace                 # the folder the AI can read and edit
docker compose up -d --build
```

* MCP endpoint: `http://localhost:8000/mcp`
* Health: `http://localhost:8000/healthz` (process up), `/readyz` (LibreOffice ready)
* Your files: `./workspace` on the host = `/data` in the container.
* Stop: `docker compose down`. Logs: `docker compose logs -f`.

Connect a client that supports HTTP MCP servers, e.g. Claude Code:

```bash
claude mcp add --transport http libreoffice http://localhost:8000/mcp
```

LibreOffice stays warm, so requests are fast from the very first call. Settings go in the
`environment:` section of `docker-compose.yml` (`LIBREOFFICE_MCP_APPS`, `…_READ_ONLY`, …); every
command-line flag has a matching `LIBREOFFICE_MCP_*` variable. By default everything is enabled.

## ⚠️ There is no authentication

Anyone who can reach the HTTP port can read and edit everything in the workspace.
`docker-compose.yml` publishes the port on `127.0.0.1` only, so just your computer can use it.
Keep it that way, or put the server behind something that authenticates (a reverse proxy with
login, a VPN, a private network such as Tailscale).

## Hosting it somewhere (a URL)

The same image works on any host that runs containers: a small VPS with Docker, Fly.io,
Railway, Render, Google Cloud Run, AWS ECS/Fargate…

* **Not Vercel/Netlify/Lambda-style serverless.** They need short-lived functions and have no
  persistent disk; this server is a long-running process (LibreOffice takes seconds to start)
  that works on real files.
* Give it **~1 GB of RAM** (LibreOffice is the heavy part) and a **persistent volume** mounted
  at `/data`, otherwise files vanish when the container restarts.
* Run it with `--transport http` and reach it through a hostname: add that hostname with
  `LIBREOFFICE_MCP_ALLOWED_HOSTS=mcp.example.com` (otherwise requests get HTTP 421), or `*` if a
  proxy in front already controls the Host header.
* Put authentication in front (see the warning above).
* Remember what "files" means when hosted: the tools read and write files **on the server's
  disk**. A remote AI can create, edit and convert documents there, but there is no built-in way
  yet to upload a local file or download the result (a future feature).
* Scale = one container per workspace; requests inside a container are processed one at a time.

## Paths and permissions

* Inside the container use relative paths or `/data/...`. Windows-style paths from your host
  (`C:\Users\me\...`) don't exist in there.
* The server runs as user `ubuntu` (uid 1000). On a Linux host make sure the mounted folder is
  writable by it (`chmod a+rwx workspace`, or `docker run --user "$(id -u):$(id -g)" ...`).
  Docker Desktop (Windows/macOS) needs nothing special.
* All file access is confined to `/data` (`LIBREOFFICE_MCP_ROOTS=/data` is preset).

## Running the tests inside Linux

```bash
docker build --target test -t libreoffice-mcp:test .
docker run --rm libreoffice-mcp:test
```

## Troubleshooting

* *The first request is slow* — LibreOffice is still starting (a couple of seconds). With compose it is already warm.
* *`docker run` exits immediately in Claude Desktop* — check you used `-i` and not `-t`, and that
  the image name is right (`docker images`).
* *Files appear but are owned by root / can't be edited on Linux* — see "Paths and permissions".
* *Fonts look wrong in PDFs* — the image ships Liberation, DejaVu, Carlito and Caladea. Add other
  fonts with a derived image (`FROM libreoffice-mcp` + `apt-get install fonts-…`); for Chinese/
  Japanese/Korean text install `fonts-noto-cjk`.
* More in [TROUBLESHOOTING.md](TROUBLESHOOTING.md).
