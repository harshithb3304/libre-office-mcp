"""Writer operations, end to end against a real LibreOffice (worker + UNO).

Run just these with:  uv run pytest tests/test_writer.py -v
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

import pytest

from libreoffice_mcp.errors import OfficeError
from tests.conftest import Doc

pytestmark = pytest.mark.office


def make_png(path: Path, width: int = 60, height: int = 30) -> Path:
    """Write a solid red PNG using only the stdlib."""
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return (
            struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
        )

    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    return path


# -- content -------------------------------------------------------------------


async def test_new_document_is_one_empty_paragraph(doc: Doc):
    assert await doc.structure() == [("paragraph", "")]


async def test_add_paragraphs_reuses_empty_trailing_paragraph(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "One"}),
        ("writer.add_paragraph", {"text": "Two"}),
    )
    assert await doc.structure() == [("paragraph", "One"), ("paragraph", "Two")]


async def test_heading_levels_and_outline(doc: Doc):
    await doc.edit(
        ("writer.add_heading", {"text": "Title", "level": 1}),
        ("writer.add_paragraph", {"text": "body"}),
        ("writer.add_heading", {"text": "Sub", "level": 2}),
    )
    outline = await doc.read("writer.get_outline")
    assert [(h["level"], h["text"]) for h in outline["headings"]] == [(1, "Title"), (2, "Sub")]
    assert outline["headings"][1]["block_index"] == 2


async def test_positions_start_before_after(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "B"}),
        ("writer.add_paragraph", {"text": "D"}),
        ("writer.add_paragraph", {"text": "A", "position": "start"}),
        ("writer.add_paragraph", {"text": "C", "position": "after", "block_index": 1}),
        ("writer.add_paragraph", {"text": "Z", "position": "before", "block_index": 3}),
    )
    assert [t for _, t in await doc.structure()] == ["A", "B", "C", "Z", "D"]


async def test_new_paragraph_does_not_inherit_previous_formatting(doc: Doc):
    await doc.edit(
        ("writer.add_heading", {"text": "Head", "level": 1}),
        ("writer.add_paragraph", {"text": "Plain", "bold": None}),
    )
    info = await doc.read("writer.get_paragraph", block_index=1)
    assert info["style"] == "Standard"
    assert info["type"] == "paragraph"


async def test_add_content_builds_a_whole_document_in_order(doc: Doc, tmp_path: Path):
    png = make_png(tmp_path / "x.png")
    await doc.edit(
        (
            "writer.add_content",
            {
                "blocks": [
                    {"type": "heading", "text": "Report", "level": 1},
                    {"type": "paragraph", "text": "Intro", "alignment": "justify"},
                    {"type": "bullets", "items": ["a", {"text": "b", "level": 1}]},
                    {"type": "table", "data": [["H1", "H2"], ["1", "2"]]},
                    {"type": "paragraph", "text": "After table"},
                    {"type": "image", "image_path": str(png), "width_mm": 30},
                    {"type": "numbered", "items": ["x", "y"]},
                ]
            },
        )
    )
    kinds = [k for k, _ in await doc.structure()]
    assert kinds == [
        "heading", "paragraph", "list_item", "list_item", "table", "paragraph",
        "paragraph", "list_item", "list_item",
    ]  # fmt: skip
    info = await doc.read("writer.get_images")
    assert info["count"] == 1 and info["images"][0]["width_mm"] == 30.0


async def test_add_content_error_names_the_failing_block(doc: Doc):
    with pytest.raises(OfficeError, match=r"blocks\[1\].*missing required field"):
        await doc.edit(
            (
                "writer.add_content",
                {"blocks": [{"type": "paragraph", "text": "ok"}, {"type": "heading"}]},
            )
        )
    assert await doc.structure() == [("paragraph", "")]  # all-or-nothing: nothing saved


async def test_edit_and_delete_blocks(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "one"}),
        ("writer.add_paragraph", {"text": "two"}),
        ("writer.add_paragraph", {"text": "three"}),
        ("writer.edit_paragraph", {"block_index": 1, "text": "TWO"}),
        ("writer.delete_blocks", {"start_block": 0}),
    )
    assert [t for _, t in await doc.structure()] == ["TWO", "three"]


async def test_delete_blocks_range(doc: Doc):
    await doc.edit(*[("writer.add_paragraph", {"text": c}) for c in "abcde"])
    await doc.edit(("writer.delete_blocks", {"start_block": 1, "end_block": 3}))
    assert [t for _, t in await doc.structure()] == ["a", "e"]


async def test_out_of_range_index_gives_helpful_error(doc: Doc):
    with pytest.raises(OfficeError, match=r"out of range.*0\.\.0"):
        await doc.edit(("writer.edit_paragraph", {"block_index": 5, "text": "x"}))


async def test_replace_text_plain_regex_and_noop(doc: Doc):
    await doc.edit(("writer.add_paragraph", {"text": "cat bat Cat"}))
    r = await doc.edit(("writer.replace_text", {"search": "cat", "replace": "dog"}))
    assert r[0]["replacements"] == 2
    r = await doc.edit(("writer.replace_text", {"search": r"b.t", "replace": "rat", "regex": True}))
    assert r[0]["replacements"] == 1
    assert (await doc.structure())[0][1] == "dog rat dog"
    r = await doc.edit(("writer.replace_text", {"search": "zzz", "replace": "y"}))
    assert r[0]["replacements"] == 0


async def test_find_text_reports_block_and_context(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "alpha beta"}),
        ("writer.add_paragraph", {"text": "gamma beta"}),
    )
    found = await doc.read("writer.find_text", query="beta")
    assert found["total_matches"] == 2
    assert [m["block_index"] for m in found["matches"]] == [0, 1]
    assert found["matches"][0]["offset"] == 6


async def test_get_text_paginates(doc: Doc):
    await doc.edit(
        *[("writer.add_paragraph", {"text": f"para {i} " + "x" * 50}) for i in range(10)]
    )
    page = await doc.read("writer.get_text", max_chars=200)
    assert page["next_block"] is not None and page["blocks_returned"] < 10
    rest = await doc.read("writer.get_text", start_block=page["next_block"], max_chars=10_000)
    assert rest["next_block"] is None
    assert rest["blocks_returned"] + page["blocks_returned"] == 10


async def test_page_break_and_page_count(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "p1"}),
        ("writer.insert_page_break", {}),
        ("writer.add_paragraph", {"text": "p2"}),
    )
    assert [t for _, t in await doc.structure()] == ["p1", "p2"]  # no stray blank paragraph
    info = await doc.read("doc.info")
    assert info["statistics"]["pages"] == 2


# -- formatting ----------------------------------------------------------------


async def test_format_text_on_found_text_only(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "make this bold please"}),
        ("writer.format_text", {"find": "this", "bold": True, "color": "#FF0000"}),
    )
    runs = (await doc.read("writer.get_paragraph", block_index=0))["runs"]
    bold = [r["text"] for r in runs if r.get("bold")]
    assert bold == ["this"]
    assert next(r for r in runs if r["text"] == "this")["color"] == "#FF0000"


async def test_format_text_without_properties_is_an_error(doc: Doc):
    await doc.edit(("writer.add_paragraph", {"text": "x"}))
    with pytest.raises(OfficeError, match="No formatting requested"):
        await doc.edit(("writer.format_text", {"block_index": 0}))


async def test_format_paragraph_alignment_style_spacing(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "centered"}),
        (
            "writer.format_paragraph",
            {"block_index": 0, "alignment": "center", "style": "Heading 2", "space_after_mm": 5},
        ),
    )
    info = await doc.read("writer.get_paragraph", block_index=0)
    assert (
        info["alignment"] == "center"
        and info["style"] == "Heading 2"
        and info["space_after_mm"] == 5
    )


async def test_unknown_style_suggests_alternatives(doc: Doc):
    await doc.edit(("writer.add_paragraph", {"text": "x"}))
    with pytest.raises(OfficeError, match=r"Similar:.*Heading"):
        await doc.edit(("writer.format_paragraph", {"block_index": 0, "style": "Headin 1"}))


async def test_set_style_creates_then_updates_and_is_usable(doc: Doc):
    r = await doc.edit(
        ("writer.set_style", {"name": "Callout", "parent": "Standard", "bold": True, "size": 14})
    )
    assert r[0]["created"] is True
    r = await doc.edit(("writer.set_style", {"name": "Callout", "italic": True}))
    assert r[0]["created"] is False
    await doc.edit(("writer.add_paragraph", {"text": "hey", "style": "Callout"}))
    assert (await doc.read("writer.get_paragraph", block_index=0))["style"] == "Callout"


# -- tables --------------------------------------------------------------------


async def test_insert_table_and_read_back(doc: Doc):
    r = await doc.edit(
        ("writer.add_paragraph", {"text": "before"}),
        ("writer.insert_table", {"data": [["Name", "Qty"], ["Apple", "3"], ["Pear", "5"]]}),
        ("writer.add_paragraph", {"text": "after"}),
    )
    assert r[1]["rows"] == 3 and r[1]["columns"] == 2 and r[1]["table_index"] == 0
    kinds = [k for k, _ in await doc.structure()]
    assert kinds == ["paragraph", "table", "paragraph"]  # no stray empty paragraph around the table
    table = await doc.read("writer.get_table", table_index=0)
    assert table["cells"] == [["Name", "Qty"], ["Apple", "3"], ["Pear", "5"]]


async def test_table_in_the_middle_does_not_leave_blank_paragraphs(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "a"}),
        ("writer.add_paragraph", {"text": "b"}),
        ("writer.insert_table", {"data": [["x"]], "position": "after", "block_index": 0}),
    )
    assert [k for k, _ in await doc.structure()] == ["paragraph", "table", "paragraph"]


async def test_table_edit_rows_columns_merge_delete(doc: Doc):
    await doc.edit(("writer.insert_table", {"data": [["a", "b"], ["c", "d"]]}))
    await doc.edit(
        ("writer.add_table_rows", {"table_index": 0, "count": 2}),
        ("writer.add_table_columns", {"table_index": 0}),
        (
            "writer.set_table_cells",
            {"table_index": 0, "data": [["1", "2", "3"]], "start_cell": "A3"},
        ),
    )
    t = await doc.read("writer.get_table", table_index=0)
    assert (t["rows"], t["columns"]) == (4, 3)
    assert t["cells"][2] == ["1", "2", "3"]
    await doc.edit(
        ("writer.delete_table_rows", {"table_index": 0, "start_row": 3}),
        ("writer.delete_table_columns", {"table_index": 0, "start_column": 2}),
        ("writer.merge_table_cells", {"table_index": 0, "cell_range": "A1:B1"}),
    )
    t = await doc.read("writer.get_table", table_index=0)
    assert t["rows"] == 3  # (a merged first row makes LibreOffice report fewer columns)
    assert t["cells"][1] == ["c", "d"]
    await doc.edit(("writer.delete_table", {"table_index": 0}))
    assert (await doc.read("writer.get_tables"))["count"] == 0


async def test_format_table_and_cells(doc: Doc):
    await doc.edit(("writer.insert_table", {"data": [["h1", "h2"], ["x", "y"]]}))
    r = await doc.edit(
        ("writer.format_table", {"table_index": 0, "border_width_pt": 1, "border_color": "#336699",
                                 "background_color": "#EEEEEE", "alignment": "center", "column_widths": [1, 3]}),
        ("writer.format_table_cells", {"table_index": 0, "cell_range": "A2:B2", "bold": True, "alignment": "right"}),
    )  # fmt: skip
    assert r[0]["changed"] == ["borders", "background_color", "alignment", "column_widths"]
    assert r[1]["cells_formatted"] == 2


async def test_bad_cell_name_is_a_clear_error(doc: Doc):
    await doc.edit(("writer.insert_table", {"data": [["a"]]}))
    with pytest.raises(OfficeError, match="outside the table"):
        await doc.edit(
            ("writer.set_table_cells", {"table_index": 0, "data": [["z"]], "start_cell": "C9"})
        )


# -- lists ---------------------------------------------------------------------


async def test_lists_and_set_list(doc: Doc):
    await doc.edit(
        ("writer.add_list", {"items": ["a", "b"]}),
        ("writer.add_paragraph", {"text": "plain"}),
        ("writer.set_list", {"block_index": 2, "kind": "numbered"}),
        ("writer.set_list", {"block_index": 0, "end_block_index": 1, "kind": "none"}),
    )
    assert [k for k, _ in await doc.structure()] == ["paragraph", "paragraph", "list_item"]


# -- objects -------------------------------------------------------------------


async def test_images_insert_resize_delete(doc: Doc, tmp_path: Path):
    png = make_png(tmp_path / "p.png", 60, 30)
    r = await doc.edit(
        ("writer.insert_image", {"image_path": str(png), "width_mm": 40, "alt_text": "red box"})
    )
    assert r[0]["width_mm"] == 40.0 and r[0]["height_mm"] == 20.0  # aspect ratio kept
    await doc.edit(("writer.set_image", {"image_index": 0, "height_mm": 10}))
    img = (await doc.read("writer.get_images"))["images"][0]
    assert (img["width_mm"], img["height_mm"], img["alt_text"]) == (20.0, 10.0, "red box")
    await doc.edit(("writer.delete_image", {"image_index": 0}))
    assert (await doc.read("writer.get_images"))["count"] == 0


async def test_image_rejects_non_images(doc: Doc, tmp_path: Path):
    bad = tmp_path / "x.png"
    bad.write_text("not an image")
    with pytest.raises(OfficeError, match="not an image"):
        await doc.edit(("writer.insert_image", {"image_path": str(bad)}))


async def test_hyperlink_bookmark_footnote_comment(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "Visit our website for details"}),
        ("writer.add_hyperlink", {"url": "https://example.com", "find": "website"}),
        ("writer.add_bookmark", {"name": "intro", "block_index": 0}),
        ("writer.add_footnote", {"note_text": "See the FAQ.", "find": "details"}),
        ("writer.add_comment", {"comment": "Verify link", "find": "website", "author": "Bot"}),
    )
    links = await doc.read("writer.get_hyperlinks")
    (link,) = links["hyperlinks"]
    assert (link["block_index"], link["text"]) == (0, "website")
    assert link["url"].rstrip("/") == "https://example.com"  # LibreOffice normalises the slash
    assert [b["name"] for b in (await doc.read("writer.get_bookmarks"))["bookmarks"]] == ["intro"]
    assert (await doc.read("writer.get_footnotes"))["footnotes"][0]["text"] == "See the FAQ."
    comments = (await doc.read("writer.get_comments"))["comments"]
    assert len(comments) == 1, comments
    assert comments[0]["text"] == "Verify link" and comments[0]["author"] == "Bot"
    assert (await doc.structure())[0][1].startswith("Visit our website")  # text intact
    await doc.edit(
        ("writer.delete_comment", {"comment_index": 0}),
        ("writer.delete_bookmark", {"name": "intro"}),
    )
    assert (await doc.read("writer.get_comments"))["count"] == 0


async def test_insert_file_merges_another_document(doc: Doc, office, tmp_path: Path):
    other = Doc(office, tmp_path / "other.odt")
    await other.create(("writer.add_paragraph", {"text": "from other"}))
    await doc.edit(
        ("writer.add_paragraph", {"text": "mine"}),
        ("writer.insert_file", {"source_path": str(other.path)}),
    )
    texts = [t for _, t in await doc.structure()]
    assert "from other" in texts and texts[0] == "mine"


# -- page ----------------------------------------------------------------------


async def test_page_setup_landscape_letter_margins(doc: Doc):
    await doc.edit(
        ("writer.set_page_setup", {"paper": "Letter", "orientation": "landscape", "left_mm": 30})
    )
    setup = await doc.read("writer.get_page_setup")
    assert setup["landscape"] is True and setup["width_mm"] > setup["height_mm"]
    assert setup["margins_mm"]["left"] == 30


async def test_header_footer_with_page_numbers(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "x"}),
        (
            "writer.set_header_footer",
            {
                "header_text": "Acme Corp",
                "footer_text": "Confidential",
                "page_numbers": "page_of_total",
            },
        ),
    )
    hf = (await doc.read("writer.get_headers_footers"))["page_styles"][0]
    assert hf["header"] == "Acme Corp"
    assert (
        hf["footer"].startswith("Confidential")
        and "Page" in hf["footer"]
        and " of " in hf["footer"]
    )
    await doc.edit(("writer.set_header_footer", {"header_text": ""}))
    assert (await doc.read("writer.get_headers_footers"))["page_styles"][0]["header"] is None


async def test_table_of_contents(doc: Doc):
    await doc.edit(
        ("writer.add_heading", {"text": "Chapter One", "level": 1}),
        ("writer.add_paragraph", {"text": "text"}),
        ("writer.insert_toc", {"position": "start"}),
    )
    text = (await doc.read("writer.get_text"))["text"]
    assert "Table of Contents" in text and text.count("Chapter One") >= 2


# -- review --------------------------------------------------------------------


async def test_track_changes_cycle(doc: Doc):
    await doc.edit(
        ("writer.add_paragraph", {"text": "original"}),
        ("writer.set_track_changes", {"enabled": True}),
    )
    await doc.edit(("writer.replace_text", {"search": "original", "replace": "revised"}))
    tc = await doc.read("writer.get_track_changes")
    assert tc["tracking_enabled"] and tc["count"] > 0
    await doc.edit(("writer.reject_all_changes", {}))
    assert (await doc.structure())[0][1] == "original"
    await doc.edit(("writer.replace_text", {"search": "original", "replace": "final"}))
    await doc.edit(("writer.accept_all_changes", {}))
    assert (await doc.structure())[0][1] == "final"
    assert (await doc.read("writer.get_track_changes"))["count"] == 0


# -- formats -------------------------------------------------------------------


async def test_docx_round_trip_keeps_structure(office, tmp_path: Path):
    d = Doc(office, tmp_path / "t.docx")
    await d.create(
        ("writer.add_heading", {"text": "Title", "level": 1}),
        ("writer.add_paragraph", {"text": "Hello", "bold": True}),
        ("writer.insert_table", {"data": [["a", "b"], ["c", "d"]]}),
    )
    await d.edit(("writer.add_paragraph", {"text": "Edited later"}))
    structure = await d.structure()
    assert [k for k, _ in structure] == ["heading", "paragraph", "table", "paragraph"]
    assert structure[-1][1] == "Edited later"


async def test_wrong_document_type_is_explained(office, tmp_path: Path):
    sheet = Doc(office, tmp_path / "s.ods")
    await sheet.create()
    with pytest.raises(OfficeError, match=r"needs a writer document.*calc"):
        await sheet.edit(("writer.add_paragraph", {"text": "x"}))
