"""MCP-level Calc tests: a real MCP client talking to our server, backed by real LibreOffice.

They check what an AI app actually sees for the calc_* tools: listing and schemas, argument
validation, error messages, the path sandbox, concurrency and batching.
"""

from __future__ import annotations

import asyncio
from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

import pytest
from mcp import Client

from libreoffice_mcp.app import build_server
from libreoffice_mcp.config import Settings
from tests.test_calc import make_png

pytestmark = pytest.mark.office


@pytest.fixture(scope="module")
async def env(tmp_path_factory: pytest.TempPathFactory) -> AsyncIterator[tuple[Client, Path]]:
    """One server + MCP client for the module; the client lives in one long-lived task (anyio)."""
    root = tmp_path_factory.mktemp("mcp_calc_root")
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


SALES = [["North", 10, 2.5], ["South", 4, 3], ["East", 7, 1.25], ["West", 1, 10]]


async def test_calc_tool_listing_and_annotations(env):
    client, _ = env
    tools = {t.name: t for t in (await client.list_tools()).tools}
    calc = {n: t for n, t in tools.items() if n.startswith("calc_")}
    assert len(calc) >= 45
    for name in ("calc_get_range", "calc_list_sheets", "calc_find_text", "calc_list_charts"):
        assert calc[name].annotations.read_only_hint is True, name
    for name in ("calc_set_range_values", "calc_write_table", "calc_add_chart"):
        assert calc[name].annotations.read_only_hint is False
        assert calc[name].annotations.destructive_hint is False
    for name in ("calc_clear_range", "calc_delete_rows", "calc_delete_sheet", "calc_replace_text"):
        assert calc[name].annotations.destructive_hint is True, name
    schema = calc["calc_write_table"].input_schema
    assert {"path", "headers", "rows"} <= set(schema["required"])
    assert "oneOf" in str(calc["calc_set_borders"].input_schema) or "enum" in str(
        calc["calc_set_borders"].input_schema
    )
    # every tool describes itself to the model
    assert [n for n, t in calc.items() if len(t.description or "") < 20] == []
    sort_schema = str(calc["calc_sort_range"].input_schema)
    assert "ascending" in sort_schema and "column" in sort_schema


async def test_read_only_mode_exposes_only_calc_read_tools():
    from libreoffice_mcp import tools
    from libreoffice_mcp.tools import registry

    settings = Settings(apps=("calc",), read_only=True)
    tools.load_modules(("calc",))
    names = {t.name for t in registry.select(settings)}
    assert "calc_get_range" in names and "calc_set_range_values" not in names
    assert not any(n.startswith("writer_") for n in names)


async def test_full_calc_workflow_create_write_format_chart_pdf(env):
    client, root = env
    r = await call(client, "office_create_document", path="flow.xlsx", title="Flow")
    assert not r.is_error and r.structured_content["kind"] == "calc"

    r = await call(
        client, "calc_write_table", path="flow.xlsx", headers=["Region", "Units", "Price", "Total"],
        rows=[[n, u, p, f"=B{i + 2}*C{i + 2}"] for i, (n, u, p) in enumerate(SALES)],
        column_formats={"Price": "$#,##0.00", "Total": "$#,##0.00"}, freeze_header=True,
        autofilter=True,
    )  # fmt: skip
    assert not r.is_error, text_of(r)
    assert r.structured_content["range"] == "A1:D5" and "formula_errors" not in r.structured_content

    r = await call(client, "calc_set_cell", path="flow.xlsx", cell="A6", value="Total")
    r = await call(client, "calc_set_cell", path="flow.xlsx", cell="D6", value="=SUM(D2:D5)")
    assert r.structured_content["value"] == "55.75" or r.structured_content["value"] == "$55.75"
    r = await call(
        client,
        "calc_format_range",
        path="flow.xlsx",
        range="A6:D6",
        bold=True,
        background="#FFF2CC",
    )
    assert not r.is_error, text_of(r)
    r = await call(
        client, "calc_add_conditional_format", path="flow.xlsx", range="D2:D5", rule="greater",
        value=20, background="#C6EFCE",
    )  # fmt: skip
    assert not r.is_error, text_of(r)
    r = await call(
        client, "calc_add_chart", path="flow.xlsx", data_range="A1:B5", chart_type="column",
        title="Units by region", x_axis_title="Region", y_axis_title="Units", anchor_cell="F2",
    )  # fmt: skip
    assert not r.is_error, text_of(r)

    r = await call(client, "calc_get_range", path="flow.xlsx", range="A1:D6", formatted=True)
    assert r.structured_content["formatted"][1][3] == "$25.00"
    assert r.structured_content["values"][5][3] == 55.75
    r = await call(client, "calc_list_charts", path="flow.xlsx")
    assert r.structured_content["count"] == 1

    r = await call(client, "office_convert_document", path="flow.xlsx", output_file="flow.pdf")
    assert not r.is_error, text_of(r)
    assert (root / "flow.pdf").read_bytes().startswith(b"%PDF")


async def test_validation_errors_reach_the_model_with_field_names(env):
    client, _ = env
    await call(client, "office_create_document", path="val.ods", overwrite=True)
    r = await call(
        client, "calc_add_conditional_format", path="val.ods", range="A1", rule="sparkly",
        bold=True,
    )  # fmt: skip
    assert r.is_error and "rule" in text_of(r)
    r = await call(client, "calc_set_range_values", path="val.ods", data="not a grid")
    assert r.is_error and "data" in text_of(r)
    r = await call(client, "calc_freeze_panes", path="val.ods", rows=-1)
    assert r.is_error and "rows" in text_of(r)


async def test_operation_errors_are_written_for_the_model(env):
    client, _ = env
    await call(client, "office_create_document", path="err.ods", overwrite=True)
    r = await call(client, "calc_get_range", path="err.ods", sheet="Nope")
    assert r.is_error and "No sheet named 'Nope'" in text_of(r) and "'Sheet1'" in text_of(r)
    r = await call(client, "calc_get_range", path="err.ods", range="B2:A")
    assert r.is_error and "A1 notation" in text_of(r)
    r = await call(client, "calc_set_range_values", path="err.ods", data=[["=FOO(1)"]])
    assert not r.is_error and r.structured_content["formula_errors"][0]["cell"] == "A1"


async def test_wrong_document_type_message(env):
    client, _ = env
    await call(client, "office_create_document", path="text.odt", overwrite=True)
    r = await call(client, "calc_list_sheets", path="text.odt")
    assert r.is_error and "needs a calc document" in text_of(r)


async def test_sandbox_applies_to_calc_paths(env, tmp_path: Path):
    client, root = env
    outside = str(root.parent / "outside.xlsx")
    r = await call(client, "calc_list_sheets", path=outside)
    assert r.is_error and "not allowed" in text_of(r)
    evil = make_png(tmp_path / "evil.png")  # a real image, but outside the allowed folder
    await call(client, "office_create_document", path="img.ods", overwrite=True)
    r = await call(client, "calc_insert_image", path="img.ods", image_path=str(evil))
    assert r.is_error and "not allowed" in text_of(r)
    inside = make_png(root / "ok.png")
    r = await call(client, "calc_insert_image", path="img.ods", image_path=str(inside), width_mm=30)
    assert not r.is_error, text_of(r)
    r = await call(client, "calc_list_images", path="img.ods")
    assert r.structured_content["count"] == 1


async def test_concurrent_edits_to_one_sheet_all_land(env):
    client, _ = env
    await call(client, "office_create_document", path="race.xlsx", overwrite=True)
    results = await asyncio.gather(
        *[
            call(client, "calc_set_cell", path="race.xlsx", cell=f"A{i + 1}", value=f"line-{i}")
            for i in range(12)
        ]
    )
    assert all(not r.is_error for r in results), [text_of(r) for r in results if r.is_error]
    r = await call(client, "calc_get_range", path="race.xlsx", range="A1:A12")
    assert sorted(v[0] for v in r.structured_content["values"]) == sorted(
        f"line-{i}" for i in range(12)
    )


async def test_concurrent_reads_and_writes_on_different_files(env):
    client, _ = env
    for name in ("p.ods", "q.ods"):
        await call(client, "office_create_document", path=name, overwrite=True)
    jobs = []
    for i in range(4):
        jobs.append(call(client, "calc_set_cell", path="p.ods", cell=f"A{i + 1}", value=i))
        jobs.append(call(client, "calc_set_cell", path="q.ods", cell=f"B{i + 1}", value=i * 10))
        jobs.append(call(client, "calc_list_sheets", path="p.ods"))
    assert all(not r.is_error for r in await asyncio.gather(*jobs))
    p = await call(client, "calc_get_range", path="p.ods")
    q = await call(client, "calc_get_range", path="q.ods")
    assert [r[0] for r in p.structured_content["values"]] == [0, 1, 2, 3]
    assert [r[0] for r in q.structured_content["values"]] == [0, 10, 20, 30]


async def test_batch_edit_accepts_calc_operations(env):
    client, _ = env
    await call(client, "office_create_document", path="batch.xlsx", overwrite=True)
    r = await call(
        client, "office_batch_edit", path="batch.xlsx",
        operations=[
            {"op": "calc.write_table", "args": {"headers": ["k", "v"], "rows": [["a", 1], ["b", 2]]}},
            {"op": "calc.add_sheet", "args": {"name": "Summary"}},
            {"op": "calc.set_cell", "args": {"sheet": "Summary", "cell": "A1", "value": "=SUM(Sheet1.B2:B3)"}},
            {"op": "calc.format_range", "args": {"sheet": "Summary", "range": "A1", "bold": True}},
        ],
    )  # fmt: skip
    assert not r.is_error, text_of(r)
    r = await call(client, "calc_get_range", path="batch.xlsx", sheet="Summary")
    assert r.structured_content["values"] == [[3]]


async def test_failed_batch_leaves_the_spreadsheet_untouched(env):
    client, root = env
    await call(client, "office_create_document", path="atomic.ods", overwrite=True)
    await call(client, "calc_set_cell", path="atomic.ods", cell="A1", value="original")
    before = (root / "atomic.ods").read_bytes()
    r = await call(
        client, "office_batch_edit", path="atomic.ods",
        operations=[
            {"op": "calc.replace_text", "args": {"search": "original", "replace": "changed"}},
            {"op": "calc.rename_sheet", "args": {"sheet": "Missing", "new_name": "X"}},
        ],
    )  # fmt: skip
    assert r.is_error and "nothing was saved" in text_of(r) and "No sheet named" in text_of(r)
    assert (root / "atomic.ods").read_bytes() == before


async def test_csv_through_the_tools(env):
    client, root = env
    r = await call(client, "office_create_document", path="export.csv", overwrite=True)
    assert not r.is_error and r.structured_content["kind"] == "calc"
    r = await call(
        client, "calc_write_table", path="export.csv", headers=["id", "note"],
        rows=[[1, "plain"], [2, "has, comma"], [3, 'has "quote"']], header_style=False,
        autofit=False,
    )  # fmt: skip
    assert not r.is_error, text_of(r)
    text = (root / "export.csv").read_text(encoding="utf-8")
    assert text.splitlines()[0] == "id,note"
    assert '"has, comma"' in text and '"has ""quote"""' in text
    r = await call(client, "calc_get_range", path="export.csv")
    assert r.structured_content["values"][3] == [3, 'has "quote"']


async def test_no_stray_temp_files_after_calc_work(env):
    _, root = env
    leftovers = [
        p.name for p in root.iterdir() if p.name.startswith((".", "~")) or p.suffix == ".tmp"
    ]
    assert leftovers == []
