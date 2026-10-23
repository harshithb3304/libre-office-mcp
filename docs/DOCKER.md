# Docker in detail

Docker is the only supported way to run this: LibreOffice, its Python bridge and the fonts are all
inside the image, so the only thing to install on your computer is Docker. The README has the
quick start; this page explains the details.

* The **container is the MCP server**. Your AI app just needs to be told how to reach it.
* The AI sees **one folder** of yours — the workspace — as `/data`.

## Build the image

```bash
git clone https://github.com/harshithb3304/libre-office-mcp.git
cd libre-office-mcp
docker build -t libreoffice-mcp .
```

The first build downloads LibreOffice (several hundred MB); later builds are cached. The image is
about 1 GB, a one-time cost: it does not grow when you create documents.

## The workspace folder

The workspace is a normal folder on your computer that you mount into the container as `/data`.
It is the *only* thing the AI can see.

| | |
|---|---|
| Where is it? | Whatever folder you mount. With compose: `./workspace` next to `docker-compose.yml`, or `WORKSPACE=/some/folder`. With `docker run`: the left side of `-v <folder>:/data`. |
| Is it cleared? | No. It is on your disk, outside the image and container. Stopping or removing the container leaves it untouched. |
| Relative paths | `report.docx` → `/data/report.docx`. This is the workspace, **not** your terminal's current directory (unless you mount it: `-v "$(pwd):/data"`). |
| Absolute paths | Only paths under `/data`. Anything else is refused. |
| Host paths | `C:\Users\me\file.docx` doesn't exist inside the container — copy the file into the workspace and use its name. |
| Existing files | Copy them into the workspace; the AI then edits them there. |
| Open files | Close a file in other programs before the AI edits it. Some programs lock open files (the write is then refused with "file in use"); others keep showing the old content until you reload. |
| Permissions | The server runs as user `ubuntu` (uid 1000). On a Linux host make the folder writable for it (`chmod a+rwx folder`) or run with `--user "$(id -u):$(id -g)"`. Docker Desktop (Windows/macOS) needs nothing special. |

## Option A — the AI app starts the container (stdio)

Nothing stays running; each AI session starts a fresh container and removes it afterwards.

**Claude Code:**

```bash
claude mcp add libreoffice -- docker run -i --rm -v /path/to/my-workspace:/data libreoffice-mcp
```

**Claude Desktop** (`claude_desktop_config.json`; double the backslashes on Windows):

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

`-i` is required (it keeps stdin open); do **not** add `-t`. LibreOffice boots in the background
in a couple of seconds, so the first request may wait briefly.

## Option B — an always-on HTTP server (docker compose)

```bash
docker compose up -d --build
claude mcp add --transport http libreoffice http://localhost:8000/mcp
```

* MCP endpoint: `http://localhost:8000/mcp`
* Health checks for Docker/monitoring (not needed by you): `/healthz` = the server process is
  up, `/readyz` = LibreOffice has started and can take requests.
* Logs: `docker compose logs -f`. Stop: `docker compose down`.

LibreOffice stays warm, so requests are fast from the very first call.

## Choosing tools and other settings

Everything is enabled by default. To change that, add flags after the image name (Option A) or in
the `command:` line of `docker-compose.yml` (Option B), e.g. `--apps writer` or `--read-only`.
Every flag has a matching `LIBREOFFICE_MCP_*` environment variable. The README lists the common ones.

## ⚠️ There is no authentication

Anyone who can reach the HTTP port can read and edit everything in the workspace.
`docker-compose.yml` publishes the port on `127.0.0.1` only, so just your computer can use it.
Keep it that way, or put the server behind something that authenticates (a reverse proxy with
login, a VPN, a private network such as Tailscale).

## Hosting it somewhere (a URL)

The same image works on any host that runs containers: a small VPS with Docker, Fly.io, Railway,
Render, Google Cloud Run, AWS ECS/Fargate…

* **Not Vercel/Netlify/Lambda-style serverless.** They need short-lived functions and have no
  persistent disk; this is a long-running process that works on real files.
* Give it **~1 GB of RAM** and a **persistent volume** mounted at `/data`, otherwise files vanish
  when the container restarts.
* Run it with `--transport http` and reach it through a hostname: add that hostname with
  `--allowed-host mcp.example.com` (otherwise requests get HTTP 421), or `*` if a proxy in front
  already controls the Host header.
* Put authentication in front (see above).
* When hosted, the tools read and write files **on the server's disk**. A remote AI can create,
  edit and convert documents there, but there is no built-in way yet to upload a local file or
  download the result.
* Requests inside one container are processed one at a time.

## Running the tests inside Linux

```bash
docker build --target test -t libreoffice-mcp:test .
docker run --rm libreoffice-mcp:test
```

## Troubleshooting

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md).
