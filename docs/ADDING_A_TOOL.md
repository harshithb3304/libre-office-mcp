# Adding a tool

A tool has **two halves** that live in different processes (see [ARCHITECTURE.md](ARCHITECTURE.md)
for why):

| Half | Where | Runs in | Knows about |
|---|---|---|---|
| **Operation** — the LibreOffice logic | `src/libreoffice_mcp/worker/ops/<app>/…` | LibreOffice's Python | UNO, documents |
| **Tool** — what the AI sees | `src/libreoffice_mcp/tools/<app>/…` | normal Python | MCP, argument schemas |

The tool forwards to the operation by name: tool `writer_delete_bookmark` →
operation `writer.delete_bookmark` (first `_` becomes `.`).

This guide walks through a real example from the code base: **deleting a bookmark**.

## 1. Write the operation (worker side)

`src/libreoffice_mcp/worker/ops/writer/objects.py`:

```python
@op("writer.delete_bookmark", kind="writer", mutates=True)
def delete_bookmark(s: Session, name: str) -> dict[str, Any]:
    bookmarks = s.doc.getBookmarks()
    if not bookmarks.hasByName(name):
        raise OpError(f"No bookmark named '{name}'. Use writer_get_bookmarks.", protocol.E_NOT_FOUND)
    bookmarks.getByName(name).dispose()
    return {"deleted_bookmark": name}
```

What each part means:

* `@op("writer.delete_bookmark", …)` registers the function under that name.
* `kind="writer"` — the runner refuses with a clear message if the file is not a Writer document
  (use `None` for operations that work on any document).
* `mutates=True` — the document is saved afterwards. Leave it `False` for read-only operations
  (they run on a read-only load and can never change the file).
* `s: Session` — the first parameter is always the session. `s.doc` is the open UNO document;
  `s.new("com.sun.star.…")` creates document objects; `s.service(…)` creates global services.
  **Never save or close the document yourself** — the runner does that.
* The other parameters are the arguments. Give them defaults where sensible.
* **Return a JSON-serialisable value** (usually a `dict`). It becomes the tool's result.
* To fail, raise `OpError("…")`. **Write the message for the AI**: say what is wrong *and* how to
  fix the next call (“…Use writer_get_bookmarks.”). Any other exception is reported as a generic
  failure with the exception type, which is much less helpful.
* If the operation turned out to change nothing, return `{"…": …, "_noop": True}` — the file is
  then not rewritten.

If you created a new file, make sure it is imported in the package `__init__.py`
(`worker/ops/writer/__init__.py`) — importing is what registers operations.

## 2. Write the tool (server side)

`src/libreoffice_mcp/tools/writer/objects.py`:

```python
@tool(W, "destructive", title="Delete bookmark")
async def writer_delete_bookmark(path: DocPath, name: str) -> dict[str, Any]:
    """Delete a bookmark (the marked text stays)."""
    return await edit(path, "writer.delete_bookmark", name=name)
```

* `@tool(app, mode, title=…)`:
  * `app`: `"writer"`, `"calc"`, `"impress"` or `"core"` (core is always on).
  * `mode`: `"read"` (never changes files), `"edit"` (adds/changes), `"destructive"` (removes
    content), `"create"` (writes a new file). The mode sets the MCP annotations
    (`readOnlyHint`, `destructiveHint`) and decides what `--read-only` hides.
* **The docstring is the description the AI reads.** Say what the tool does, when to use it, and
  any trap. The first sentence is used as the summary in `docs/TOOLS.md`.
* **The signature is the schema the AI sees.** Use real types, defaults, `Literal[...]` for choices,
  and `Annotated[..., Field(description=…)]` for anything not obvious (`types.py` has shared
  aliases such as `DocPath`, `BlockIndex`, `Position`, `Color`).
* The body is one line:
  * `read(path, "app.op", **args)` — read-only operation, returns its result;
  * `edit(path, "app.op", **args)` — one modifying operation; result is merged with `{"file": …}`;
  * `batch(path, [("app.op1", {…}), ("app.op2", {…})])` — several in one load/save.
* Pass `None` for omitted optional arguments — they are dropped before reaching the worker, so the
  operation's own defaults apply.
* **Arguments whose name ends in `_path`** (`image_path`, `source_path`) are treated as existing
  files: they are resolved and checked against `--root` automatically. Use a different name
  (`output_file`) for files that don't exist yet, and resolve them yourself with
  `state.paths.resolve(...)`.
* Make sure the module is imported in `tools/writer/__init__.py`.

Adding to an existing app (`writer`, `calc`, `impress`)? Put the new module next to its siblings and
import it from that app's `__init__.py` in both `tools/<app>/` and `worker/ops/<app>/`. The
`calc` and `impress` packages are good examples of a larger layout.

## 3. Test it

Worker-level test (`tests/test_writer.py`) — drives the real LibreOffice:

```python
async def test_bookmark_roundtrip(doc: Doc):
    await doc.edit(("writer.add_paragraph", {"text": "hi"}),
                   ("writer.add_bookmark", {"name": "b", "block_index": 0}))
    await doc.edit(("writer.delete_bookmark", {"name": "b"}))
    assert (await doc.read("writer.get_bookmarks"))["count"] == 0
```

Test the **failure message** too, not only the happy path:

```python
with pytest.raises(OfficeError, match="No bookmark named"):
    await doc.edit(("writer.delete_bookmark", {"name": "missing"}))
```

Run: `uv run pytest tests/test_writer.py -k bookmark`.
(Tests that need LibreOffice start it once per session; this takes 10–20 s.)

## 4. Regenerate the docs

```bash
uv run python scripts/gen_tool_docs.py     # rewrites docs/TOOLS.md
uv run ruff check . && uv run ruff format .
```

## Conventions checklist

* Names: `writer_<verb>_<noun>`; verbs `get_`, `list_` (read), `add_`, `insert_`, `set_`,
  `format_`, `delete_`, `replace_`.
* 0-based indexes everywhere (`block_index`, `table_index`, `image_index`, `comment_index`).
* Lengths: millimetres (`width_mm`), font sizes: points, colours: `"#RRGGBB"` or a basic name.
* Mutating operations should return something useful for the next call (new `block_index`,
  counts) rather than just “ok”.
* Prefer one operation that handles a whole job (`add_content`) over many chatty ones.
* Keep worker code Python 3.10-compatible and standard-library-only.
* UNO has surprises. Test against real LibreOffice, and put a comment where behaviour is
  non-obvious (see `Slot` in `worker/ops/writer/helpers.py` for examples).

## UNO tips

* The API reference: <https://api.libreoffice.org/docs/idl/ref/> (search a service name such as
  `TextTable`, then look at its properties).
* Lengths in UNO are in **1/100 mm** (helper `h.mm(value)` converts from mm).
* Properties with enum values need `uno.Enum("com.sun.star.…", "NAME")` — see `util.enum`.
* Text positions are `XTextRange`s; to insert content use a *cursor* created from a range
  (`text.createTextCursorByRange(range)`).
* Exploring is easiest with a scratch script run by LibreOffice's Python and `worker/office.py`'s
  `OfficeProcess` — see how the tests' fixtures start one.
