"""MCP-level tests: a real MCP client talking to our server, backed by real LibreOffice.

These exercise what an AI app actually sees: tool listing, schemas, argument validation,
error messages, the path sandbox, and concurrent calls.
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

pytestmark = pytest.mark.office


@pytest.fixture(scope="module")
async def env(tmp_path_factory: pytest.TempPathFactory) -> AsyncIterator[tuple[Client, Path]]:
    """One server + MCP client for the whole module (LibreOffice starts only once).

    The client runs inside a single long-lived task: the MCP client uses anyio task groups,
    which must be entered and exited from the same task - something a pytest fixture's
    setup and teardown do not guarantee.
    """
    root = tmp_path_factory.mktemp("mcp_root")
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
    result = await client.call_tool(name, args)
    return result


def text_of(result: Any) -> str:
    return result.content[0].text


async def test_tool_listing_and_annotations(env):
    client, _ = env
    tools = {t.name: t for t in (await client.list_tools()).tools}
    assert len(tools) >= 60
    assert tools["writer_get_text"].annotations.read_only_hint is True
    assert tools["writer_add_paragraph"].annotations.read_only_hint is False
    assert tools["writer_delete_blocks"].annotations.destructive_hint is True
    schema = tools["writer_add_paragraph"].input_schema
    assert "path" in schema["required"] and "text" in schema["required"]
    # The discriminated content-block schema must be visible to the model.
    blocks = tools["writer_add_content"].input_schema["properties"]["blocks"]
    assert "oneOf" in str(blocks) or "anyOf" in str(blocks)


async def test_full_writer_workflow(env):
    client, root = env
    r = await call(client, "office_create_document", path="flow.docx", title="Flow", author="Test")
    assert not r.is_error and r.structured_content["kind"] == "writer"

    r = await call(
        client, "writer_add_content", path="flow.docx",
        blocks=[
            {"type": "heading", "text": "Quarterly Report", "level": 1},
            {"type": "paragraph", "text": "Revenue grew.", "alignment": "justify"},
            {"type": "table", "data": [["Region", "Sales"], ["EU", "10"], ["US", "20"]]},
            {"type": "bullets", "items": ["one", {"text": "nested", "level": 1}]},
        ],
    )  # fmt: skip
    assert not r.is_error, text_of(r)

    r = await call(client, "writer_get_structure", path="flow.docx")
    kinds = [b["type"] for b in r.structured_content["blocks"]]
    assert kinds == ["heading", "paragraph", "table", "list_item", "list_item"]

    r = await call(
        client, "office_batch_edit", path="flow.docx",
        operations=[
            {"op": "writer.replace_text", "args": {"search": "Revenue", "replace": "Sales"}},
            {"op": "writer.format_text", "args": {"find": "Sales", "bold": True}},
            {"op": "writer.add_comment", "args": {"comment": "check numbers", "block_index": 1}},
        ],
    )  # fmt: skip
    assert not r.is_error, text_of(r)

    r = await call(client, "writer_find_text", path="flow.docx", query="Sales grew")
    assert r.structured_content["total_matches"] == 1

    r = await call(client, "office_convert_document", path="flow.docx", output_file="flow.pdf")
    assert not r.is_error
    assert (root / "flow.pdf").read_bytes().startswith(b"%PDF")


async def test_validation_errors_reach_the_model_with_field_names(env):
    client, _ = env
    r = await call(client, "writer_add_heading", path="flow.docx", text="x", level=42)
    assert r.is_error and "level" in text_of(r)


async def test_missing_file_error_is_actionable(env):
    client, root = env
    r = await call(client, "writer_get_text", path="does-not-exist.docx")
    assert r.is_error
    assert "File not found" in text_of(r) and str(root) in text_of(r)


async def test_sandbox_blocks_paths_outside_root(env):
    client, root = env
    outside = str(root.parent / "outside.docx")  # absolute on every OS, and not under root
    for path in (outside, "../escape.docx"):
        r = await call(client, "office_create_document", path=path)
        assert r.is_error and "not allowed" in text_of(r), path


async def test_sandbox_applies_to_nested_image_paths(env, tmp_path: Path):
    client, _ = env
    outside = tmp_path / "evil.png"
    outside.write_bytes(b"x")
    await call(client, "office_create_document", path="img.docx", overwrite=True)
    r = await call(
        client, "writer_add_content", path="img.docx",
        blocks=[{"type": "image", "image_path": str(outside)}],
    )  # fmt: skip
    assert r.is_error and "not allowed" in text_of(r)


async def test_wrong_document_type_message(env):
    client, _ = env
    await call(client, "office_create_document", path="sheet.ods", overwrite=True)
    r = await call(client, "writer_add_paragraph", path="sheet.ods", text="x")
    assert r.is_error and "needs a writer document" in text_of(r)


async def test_existing_file_not_overwritten_by_default(env):
    client, _ = env
    await call(client, "office_create_document", path="keep.odt", overwrite=True)
    r = await call(client, "office_create_document", path="keep.odt")
    assert r.is_error and "already exists" in text_of(r)


async def test_concurrent_edits_to_one_document_all_land(env):
    """The race-condition test: 12 simultaneous edits must serialise, none may be lost."""
    client, _ = env
    await call(client, "office_create_document", path="race.odt", overwrite=True)
    results = await asyncio.gather(
        *[
            call(client, "writer_add_paragraph", path="race.odt", text=f"line-{i}")
            for i in range(12)
        ]
    )
    assert all(not r.is_error for r in results), [text_of(r) for r in results if r.is_error]
    r = await call(client, "writer_get_text", path="race.odt")
    lines = [ln for ln in r.structured_content["text"].split("\n") if ln]
    assert sorted(lines) == sorted(f"line-{i}" for i in range(12))  # all present, none duplicated


async def test_concurrent_reads_and_writes_across_documents(env):
    client, _ = env
    for name in ("a.odt", "b.odt"):
        await call(client, "office_create_document", path=name, overwrite=True)
    jobs = []
    for i in range(5):
        jobs.append(call(client, "writer_add_paragraph", path="a.odt", text=f"A{i}"))
        jobs.append(call(client, "writer_add_paragraph", path="b.odt", text=f"B{i}"))
        jobs.append(call(client, "writer_get_structure", path="a.odt"))
    results = await asyncio.gather(*jobs)
    assert all(not r.is_error for r in results)
    for name, prefix in (("a.odt", "A"), ("b.odt", "B")):
        text = (await call(client, "writer_get_text", path=name)).structured_content["text"]
        assert sorted(text.split("\n")) == sorted(f"{prefix}{i}" for i in range(5))


async def test_failed_batch_leaves_file_untouched(env):
    client, root = env
    await call(client, "office_create_document", path="atomic.odt", overwrite=True)
    await call(client, "writer_add_paragraph", path="atomic.odt", text="original")
    before = (root / "atomic.odt").read_bytes()
    r = await call(
        client, "office_batch_edit", path="atomic.odt",
        operations=[
            {"op": "writer.replace_text", "args": {"search": "original", "replace": "changed"}},
            {"op": "writer.edit_paragraph", "args": {"block_index": 99, "text": "boom"}},
        ],
    )  # fmt: skip
    assert r.is_error and "nothing was saved" in text_of(r)
    assert (root / "atomic.odt").read_bytes() == before


async def test_no_stray_temp_or_lock_files_are_left_behind(env):
    _, root = env
    leftovers = [
        p.name for p in root.iterdir() if p.name.startswith((".", "~")) or p.suffix == ".tmp"
    ]
    assert leftovers == []


async def test_status_reports_a_running_worker(env):
    client, _ = env
    r = await call(client, "office_status")
    assert r.structured_content["worker"]["running"] is True
    assert r.structured_content["worker"]["requests_served"] > 0
