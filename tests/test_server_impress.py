"""MCP-level tests for the Impress tools: a real MCP client talking to our server.

These exercise what an AI app actually sees: tool listing, schemas, argument validation,
error messages, the path sandbox, concurrency and atomic batches.
"""

from __future__ import annotations

import ast
import asyncio
import inspect
import typing
import zipfile
from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

import pytest
from mcp import Client

from libreoffice_mcp import tools
from libreoffice_mcp.app import build_server
from libreoffice_mcp.config import Settings
from libreoffice_mcp.tools import registry
from libreoffice_mcp.tools.impress import types as impress_types
from tests.test_impress import make_png

office = pytest.mark.office


@pytest.fixture(scope="module")
async def env(tmp_path_factory: pytest.TempPathFactory) -> AsyncIterator[tuple[Client, Path]]:
    """One server + MCP client for the whole module (LibreOffice starts only once).

    The client lives in one long-lived task because the MCP client uses anyio task groups,
    which must be entered and exited from the same task.
    """
    root = tmp_path_factory.mktemp("mcp_impress_root")
    server = build_server(Settings(default_dir=root, roots=(root.resolve(),), op_timeout=120))
    ready: asyncio.Future[Client] = asyncio.get_running_loop().create_future()
    done = asyncio.Event()

    async def run_client() -> None:
        async with Client(server) as client:
            ready.set_result(client)
            await done.wait()

    task = asyncio.create_task(run_client())
    try:
        yield await ready, root
    finally:
        done.set()
        await task


async def call(client: Client, name: str, **args: Any) -> Any:
    return await client.call_tool(name, args)


def text_of(result: Any) -> str:
    return result.content[0].text


# -- static checks (no LibreOffice needed) ---------------------------------------------------------


def worker_op_signatures() -> dict[str, tuple[set[str], set[str]]]:
    """op name -> (all parameter names, required parameter names), read from the worker source.

    The worker modules import ``uno`` and cannot be imported by the server's Python, so the
    source is parsed instead.
    """
    root = Path(__file__).resolve().parents[1] / "src/libreoffice_mcp/worker/ops/impress"
    found: dict[str, tuple[set[str], set[str]]] = {}
    for path in root.glob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if not isinstance(node, ast.FunctionDef):
                continue
            for deco in node.decorator_list:
                if isinstance(deco, ast.Call) and getattr(deco.func, "id", "") == "op":
                    name = ast.literal_eval(deco.args[0])
                    args = node.args.args[1:]  # skip the session
                    defaults = len(node.args.defaults)
                    required = {a.arg for a in args[: len(args) - defaults]}
                    found[name] = ({a.arg for a in args}, required)
    return found


def impress_tools() -> list[registry.ToolDef]:
    tools.load_modules(("impress",))
    return [t for t in registry.TOOLS if t.app == "impress"]


def test_every_tool_matches_its_worker_operation():
    """Tool and operation signatures must not drift apart (AGENTS.md rule 10)."""
    ops = worker_op_signatures()
    mapped = set()
    for tool in impress_tools():
        op_name = tool.name.replace("_", ".", 1)
        assert op_name in ops, f"{tool.name} has no worker operation {op_name}"
        mapped.add(op_name)
        op_params, op_required = ops[op_name]
        tool_params = set(inspect.signature(tool.fn).parameters) - {"path"}
        assert tool_params <= op_params, (
            f"{tool.name} passes unknown args {tool_params - op_params}"
        )
        assert op_required <= tool_params, f"{tool.name} never sends {op_required - tool_params}"
    # add_slides / add_slide / ... all tools exist; every op is reachable through a tool
    assert set(ops) == mapped, f"ops without a tool: {set(ops) - mapped}"


def test_tool_inventory():
    names = {t.name for t in impress_tools()}
    assert len(names) >= 40
    assert all(n.startswith("impress_") for n in names)
    modes = {t.name: t.mode for t in impress_tools()}
    assert modes["impress_get_slide"] == "read"
    assert modes["impress_delete_slides"] == "destructive"
    assert modes["impress_delete_shapes"] == "destructive"
    assert modes["impress_add_slides"] == "edit"


def literal_args(annotation: Any) -> set[str]:
    """All string values of a (possibly Optional) ``Literal`` type."""
    out: set[str] = set()
    for arg in typing.get_args(annotation) or (annotation,):
        if typing.get_origin(arg) is typing.Literal:
            out |= set(typing.get_args(arg))
        elif isinstance(arg, str):
            out.add(arg)
    return out


@office
async def test_tool_literals_match_the_worker(env):
    """The layout / shape / theme choices offered to the model are what the worker accepts."""
    client, _ = env
    await call(client, "office_create_document", path="lits.odp")
    r = await call(client, "impress_list_masters", path="lits.odp")
    assert literal_args(impress_types.SlideLayout) == {
        x["name"] for x in r.structured_content["layouts"]
    }
    assert literal_args(impress_types.ThemeName) == set(r.structured_content["themes"])
    r = await call(client, "office_batch_edit", path="lits.odp", operations=[
        {"op": "impress.add_shape", "args": {"slide_index": 0, "shape": "blob", "x_mm": 0, "y_mm": 0, "width_mm": 1, "height_mm": 1}},
    ])  # fmt: skip
    valid = text_of(r).split("Valid shapes: ")[1].split(".")[0].split(", ")
    assert set(valid) == literal_args(impress_types.ShapeType)
    r = await call(client, "office_batch_edit", path="lits.odp", operations=[
        {"op": "impress.set_slide_transition", "args": {"effect": "nope"}},
    ])  # fmt: skip
    effects = text_of(r).split("Use one of: ")[1].rstrip(".").split(", ")
    assert set(effects) == literal_args(impress_types.TransitionEffect)


# -- through the MCP client ----------------------------------------------------------------------------


@office
async def test_tool_listing_and_annotations(env):
    client, _ = env
    tools_ = {t.name: t for t in (await client.list_tools()).tools}
    impress = {n: t for n, t in tools_.items() if n.startswith("impress_")}
    assert len(impress) >= 40
    assert impress["impress_get_slide"].annotations.read_only_hint is True
    assert impress["impress_add_slide"].annotations.read_only_hint is False
    assert impress["impress_delete_slides"].annotations.destructive_hint is True
    assert impress["impress_modify_table"].annotations.destructive_hint is True
    schema = impress["impress_add_slides"].input_schema
    assert "path" in schema["required"] and "slides" in schema["required"]
    # the typed slide block (with nested bullets / image / table) is visible to the model
    flat = str(schema)
    assert "right_bullets" in flat and "image_path" in flat and "title_content" in flat
    add_shape = impress["impress_add_shape"].input_schema
    assert "rounded_rectangle" in str(add_shape["properties"]["shape"])
    assert all(t.description for t in impress.values())


@office
async def test_full_deck_workflow_with_export(env, tmp_path: Path):
    client, root = env
    png = make_png(root / "logo.png", 120, 60)
    r = await call(client, "office_create_document", path="flow.pptx", title="Flow", author="Test")
    assert not r.is_error and r.structured_content["kind"] == "impress"

    r = await call(client, "impress_apply_theme", path="flow.pptx", theme="corporate")
    assert not r.is_error, text_of(r)
    r = await call(
        client, "impress_add_slides", path="flow.pptx",
        slides=[
            {"title": "Quarterly Review", "subtitle": "Finance", "notes": "Welcome"},
            {"title": "Highlights", "bullets": ["Revenue +12%", {"text": "EU +18%", "level": 1}, "Churn down"]},
            {"title": "Logo", "image": {"image_path": str(png), "alt_text": "logo"}},
            {"title": "Numbers", "table": {"data": [["Region", "Sales"], ["EU", 10], ["US", 20.5]]}},
        ],
    )  # fmt: skip
    assert not r.is_error, text_of(r)
    assert r.structured_content["slides_added"] == 4 and r.structured_content["slide_count"] == 4

    for name, args in [
        ("impress_add_shape", {"slide_index": 1, "shape": "star", "x_mm": 220, "y_mm": 100, "width_mm": 30, "height_mm": 30, "fill_color": "#FFC000"}),
        ("impress_add_text_box", {"slide_index": 1, "x_mm": 20, "y_mm": 120, "width_mm": 150, "text": "Source: ERP", "font_size": 14, "italic": True}),
        ("impress_set_notes", {"slide_index": 1, "text": "Pause for questions"}),
        ("impress_format_text", {"slide_index": 1, "shape_index": 1, "paragraph_index": 0, "bold": True}),
        ("impress_set_slide_transition", {"effect": "fade"}),
        ("impress_set_footer", {"footer_text": "Acme", "show_slide_number": True}),
    ]:  # fmt: skip
        r = await call(client, name, path="flow.pptx", **args)
        assert not r.is_error, (name, text_of(r))

    r = await call(client, "impress_get_overview", path="flow.pptx")
    assert [s["title"] for s in r.structured_content["slides"]] == [
        "Quarterly Review",
        "Highlights",
        "Logo",
        "Numbers",
    ]
    r = await call(client, "impress_get_text", path="flow.pptx")
    assert (
        "Revenue +12%" in r.structured_content["outline"]
        and "US | 20.5" in r.structured_content["outline"]
    )

    r = await call(client, "office_convert_document", path="flow.pptx", output_file="flow.pdf")
    assert not r.is_error, text_of(r)
    assert (root / "flow.pdf").read_bytes().startswith(b"%PDF")
    r = await call(client, "office_convert_document", path="flow.pptx", output_file="flow_copy.odp")
    assert not r.is_error, text_of(r)
    r = await call(
        client, "office_convert_document", path="flow_copy.odp", output_file="flow_back.pptx"
    )
    assert not r.is_error, text_of(r)
    with zipfile.ZipFile(root / "flow_back.pptx") as z:
        assert (
            sum(n.startswith("ppt/slides/slide") and n.endswith(".xml") for n in z.namelist()) == 4
        )
    r = await call(client, "impress_get_overview", path="flow_back.pptx")
    assert r.structured_content["slide_count"] == 4


@office
async def test_validation_errors_reach_the_model(env):
    client, _ = env
    await call(client, "office_create_document", path="val.odp", overwrite=True)
    r = await call(client, "impress_add_slide", path="val.odp", layout="fancy")
    assert r.is_error and "layout" in text_of(r)  # schema validation names the field
    r = await call(
        client, "impress_add_slides", path="val.odp", slides=[{"title": "x", "bullet": ["typo"]}]
    )
    assert r.is_error and "bullet" in text_of(r)  # extra="forbid" catches typos
    r = await call(
        client,
        "impress_set_shape_geometry",
        path="val.odp",
        slide_index=0,
        shape_index=0,
        rotation_deg=900,
    )
    assert r.is_error and "rotation_deg" in text_of(r)
    r = await call(client, "impress_get_slide", path="val.odp", slide_index=7)
    assert (
        r.is_error
        and "slide_index 7 is out of range" in text_of(r)
        and "impress_get_overview" in text_of(r)
    )
    await call(
        client, "impress_add_slide", path="val.odp", layout="title_only", title="Only a title"
    )
    r = await call(client, "impress_set_body", path="val.odp", slide_index=0, items=["x"])
    assert (
        r.is_error
        and "no content placeholder" in text_of(r)
        and "impress_set_slide_layout" in text_of(r)
    )


@office
async def test_sandbox_applies_to_image_paths_top_level_and_nested(env, tmp_path: Path):
    client, _ = env
    outside = make_png(tmp_path / "evil.png")
    await call(client, "office_create_document", path="sandbox.odp", overwrite=True)
    nested = await call(
        client,
        "impress_add_slides",
        path="sandbox.odp",
        slides=[{"title": "x", "image": {"image_path": str(outside)}}],
    )
    assert nested.is_error and "not allowed" in text_of(nested)
    top = await call(
        client, "impress_add_image", path="sandbox.odp", slide_index=0, image_path=str(outside)
    )
    assert top.is_error and "not allowed" in text_of(top)
    traversal = await call(client, "impress_get_overview", path="../escape.odp")
    assert traversal.is_error and "not allowed" in text_of(traversal)


@office
async def test_wrong_document_type_message(env):
    client, _ = env
    await call(client, "office_create_document", path="notes.odt", overwrite=True)
    r = await call(client, "impress_get_overview", path="notes.odt")
    assert r.is_error and "needs a impress document" in text_of(r)
    await call(client, "office_create_document", path="deck_for_writer_tool.odp", overwrite=True)
    r = await call(client, "writer_add_paragraph", path="deck_for_writer_tool.odp", text="x")
    assert r.is_error and "needs a writer document" in text_of(r)


@office
async def test_concurrent_edits_to_one_deck_all_land(env):
    """12 simultaneous slide additions must serialise: none lost, none duplicated."""
    client, _ = env
    await call(client, "office_create_document", path="race.odp", overwrite=True)
    results = await asyncio.gather(
        *[call(client, "impress_add_slide", path="race.odp", title=f"slide-{i}") for i in range(12)]
    )
    assert all(not r.is_error for r in results), [text_of(r) for r in results if r.is_error]
    r = await call(client, "impress_get_overview", path="race.odp")
    titles = [s["title"] for s in r.structured_content["slides"]]
    assert sorted(titles) == sorted(f"slide-{i}" for i in range(12))


@office
async def test_concurrent_reads_and_writes_across_decks(env):
    client, _ = env
    for name in ("one.odp", "two.odp"):
        await call(client, "office_create_document", path=name, overwrite=True)
    jobs = []
    for i in range(4):
        jobs.append(call(client, "impress_add_slide", path="one.odp", title=f"A{i}"))
        jobs.append(call(client, "impress_add_slide", path="two.odp", title=f"B{i}"))
        jobs.append(call(client, "impress_get_overview", path="one.odp"))
    results = await asyncio.gather(*jobs)
    assert all(not r.is_error for r in results)
    for name, prefix in (("one.odp", "A"), ("two.odp", "B")):
        r = await call(client, "impress_get_text", path=name)
        assert all(f"Slide {i}: {prefix}" in r.structured_content["outline"] for i in range(4))


@office
async def test_batch_edit_with_impress_ops_and_atomic_failure(env):
    client, root = env
    await call(client, "office_create_document", path="batch.odp", overwrite=True)
    r = await call(client, "office_batch_edit", path="batch.odp", operations=[
        {"op": "impress.add_slide", "args": {"title": "One", "bullets": ["a"]}},
        {"op": "impress.add_shape", "args": {"slide_index": 0, "shape": "ellipse", "x_mm": 10, "y_mm": 100, "width_mm": 30, "height_mm": 20}},
        {"op": "impress.set_notes", "args": {"slide_index": 0, "text": "batched"}},
    ])  # fmt: skip
    assert not r.is_error, text_of(r)
    before = (root / "batch.odp").read_bytes()
    r = await call(client, "office_batch_edit", path="batch.odp", operations=[
        {"op": "impress.add_slide", "args": {"title": "Two"}},
        {"op": "impress.delete_slides", "args": {"start_slide": 0, "end_slide": 99}},
    ])  # fmt: skip
    assert r.is_error and "nothing was saved" in text_of(r)
    assert (root / "batch.odp").read_bytes() == before
    r = await call(client, "impress_get_notes", path="batch.odp")
    assert r.structured_content["slides"][0]["notes"] == "batched"


@office
async def test_move_and_reorder_through_mcp(env):
    client, _ = env
    await call(client, "office_create_document", path="order.pptx", overwrite=True)
    await call(
        client,
        "impress_add_slides",
        path="order.pptx",
        slides=[{"title": t, "bullets": [t.lower()]} for t in "ABCD"],
    )
    r = await call(client, "impress_move_slide", path="order.pptx", from_index=3, to_index=0)
    assert not r.is_error, text_of(r)
    r = await call(client, "impress_get_overview", path="order.pptx")
    assert [s["title"] for s in r.structured_content["slides"]] == ["D", "A", "B", "C"]
    r = await call(client, "impress_move_slide", path="order.pptx", from_index=0, to_index=9)
    assert r.is_error and "to_index 9 is out of range" in text_of(r)


@office
async def test_no_stray_temp_or_lock_files_are_left_behind(env):
    _, root = env
    leftovers = [
        p.name for p in root.iterdir() if p.name.startswith((".", "~")) or p.suffix == ".tmp"
    ]
    assert leftovers == []
