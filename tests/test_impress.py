"""Impress operations, end to end against a real LibreOffice (worker + UNO).

Run just these with:  uv run pytest tests/test_impress.py -v
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path
from typing import Any

import pytest

from libreoffice_mcp.client import OfficeClient
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


@pytest.fixture
async def deck(office: OfficeClient, tmp_path: Path) -> Doc:
    d = Doc(office, tmp_path / "test.odp")
    await d.create()
    return d


async def slide(deck: Doc, index: int) -> dict[str, Any]:
    return await deck.read("impress.get_slide", slide_index=index)


def shape_types(info: dict[str, Any]) -> list[str]:
    return [sh["type"] for sh in info["shapes"]]


# -- overview / reading -------------------------------------------------------------------


async def test_new_deck_overview(deck: Doc):
    info = await deck.read("impress.get_overview")
    assert info["slide_count"] == 1
    assert (info["slide_width_mm"], info["slide_height_mm"]) == (280.0, 157.5)
    assert info["slides"][0]["layout"] == "title_slide"
    assert info["masters"] == ["Default"]


async def test_add_slide_reuses_the_empty_first_slide(deck: Doc):
    (res,) = await deck.edit(("impress.add_slide", {"title": "Hello", "subtitle": "World"}))
    assert res["slide_index"] == 0 and res["reused_empty_first_slide"] is True
    assert res["slide_count"] == 1
    info = await slide(deck, 0)
    assert [s["text"] for s in info["shapes"]] == ["Hello", "World"]


async def test_add_slides_builds_a_deck_in_one_call(deck: Doc, tmp_path: Path):
    png = make_png(tmp_path / "pic.png")
    (res,) = await deck.edit(
        (
            "impress.add_slides",
            {
                "slides": [
                    {"title": "Deck", "subtitle": "Intro", "notes": "Say hi"},
                    {
                        "title": "Points",
                        "bullets": ["one", {"text": "nested", "level": 1}, "two"],
                        "notes": "Explain the points",
                    },
                    {"title": "Compare", "bullets": ["left"], "right_bullets": ["right"]},
                    {
                        "title": "Picture",
                        "image": {"image_path": str(png), "alt_text": "red box"},
                    },
                    {
                        "title": "Numbers",
                        "table": {"data": [["Region", "Sales"], ["EU", "10"], ["US", "20"]]},
                    },
                    {"title": "Both", "bullets": ["text"], "image": {"image_path": str(png)}},
                    {"layout": "blank"},
                ]
            },
        )
    )
    assert res["slides_added"] == 7 and res["slide_count"] == 7
    assert res["reused_empty_first_slide"] is True
    layouts = [s["layout"] for s in res["slides"]]
    assert layouts == [
        "title_slide",
        "title_content",
        "two_content",
        "title_only",
        "title_only",
        "title_content",
        "blank",
    ]
    overview = await deck.read("impress.get_overview")
    assert [s["title"] for s in overview["slides"][:6]] == [
        "Deck",
        "Points",
        "Compare",
        "Picture",
        "Numbers",
        "Both",
    ]
    points = await slide(deck, 1)
    body = next(s for s in points["shapes"] if s["type"] == "body")
    assert body["levels"] == [0, 1, 0]
    assert points["notes"] == "Explain the points"
    picture = await slide(deck, 3)
    image = next(s for s in picture["shapes"] if s["type"] == "image")
    assert image["alt_text"] == "red box"
    assert "body" not in shape_types(picture)  # the unused content area was replaced
    table = next(s for s in (await slide(deck, 4))["shapes"] if s["type"] == "table")
    assert table["rows"] == 3 and table["columns"] == 2
    both = await slide(deck, 5)
    body = next(s for s in both["shapes"] if s["type"] == "body")
    image = next(s for s in both["shapes"] if s["type"] == "image")
    assert body["x_mm"] + body["width_mm"] <= image["x_mm"] + 0.5  # side by side, no overlap


async def test_add_slides_error_names_the_slide(deck: Doc):
    with pytest.raises(OfficeError, match=r"slides\[1\].*no title placeholder"):
        await deck.edit(
            ("impress.add_slides", {"slides": [{"title": "ok"}, {"layout": "blank", "title": "x"}]})
        )
    assert (await deck.read("impress.get_overview"))["slide_count"] == 1  # nothing saved


async def test_unknown_layout_lists_valid_ones(deck: Doc):
    with pytest.raises(OfficeError, match=r"Unknown layout 'fancy'.*title_content"):
        await deck.edit(("impress.add_slide", {"layout": "fancy"}))


@pytest.mark.parametrize(
    ("layout", "expected"),
    [
        ("title_slide", ["title", "subtitle"]),
        ("title_content", ["title", "body"]),
        ("two_content", ["title", "body", "body"]),
        ("title_only", ["title"]),
        ("blank", []),
        ("centered_text", ["subtitle"]),
        ("four_content", ["title", "body", "body", "body", "body"]),
        ("six_content", ["title"] + ["body"] * 6),
    ],
)
async def test_layouts_create_expected_placeholders(deck: Doc, layout: str, expected: list[str]):
    await deck.edit(
        ("impress.add_slide", {"title": "first"}),  # takes the empty first slide
        ("impress.add_slide", {"layout": layout}),
    )
    assert shape_types(await slide(deck, 1)) == expected


async def test_slide_index_out_of_range_message(deck: Doc):
    with pytest.raises(OfficeError, match=r"slide_index 5 is out of range.*1 slide\(s\)"):
        await slide(deck, 5)
    with pytest.raises(OfficeError, match=r"shape_index 9 is out of range.*2 shape"):
        await deck.read("impress.get_shape", slide_index=0, shape_index=9)


async def test_wrong_document_type_is_reported(office: OfficeClient, tmp_path: Path):
    writer = Doc(office, tmp_path / "w.odt")
    await writer.create()
    with pytest.raises(OfficeError, match="needs a impress document"):
        await writer.read("impress.get_overview")


# -- slides ---------------------------------------------------------------------------------


async def make_deck(deck: Doc, titles: list[str]) -> None:
    await deck.edit(
        (
            "impress.add_slides",
            {"slides": [{"title": t, "bullets": [f"{t} point"]} for t in titles]},
        )
    )


async def titles(deck: Doc) -> list[str]:
    return [s["title"] for s in (await deck.read("impress.get_overview"))["slides"]]


async def test_delete_slides_and_cannot_delete_all(deck: Doc):
    await make_deck(deck, ["A", "B", "C", "D"])
    (res,) = await deck.edit(("impress.delete_slides", {"start_slide": 1, "end_slide": 2}))
    assert res["deleted"] == 2 and res["slide_count"] == 2
    assert await titles(deck) == ["A", "D"]
    with pytest.raises(OfficeError, match="at least one slide"):
        await deck.edit(("impress.delete_slides", {"start_slide": 0, "end_slide": 1}))
    with pytest.raises(OfficeError, match=r"end_slide 9 is out of range"):
        await deck.edit(("impress.delete_slides", {"start_slide": 0, "end_slide": 9}))


async def test_duplicate_slide_copies_content_and_notes(deck: Doc):
    await deck.edit(
        ("impress.add_slide", {"title": "Orig", "bullets": ["x"], "notes": "note"}),
        ("impress.add_slide", {"title": "Other"}),
    )
    (res,) = await deck.edit(("impress.duplicate_slide", {"slide_index": 0, "count": 2}))
    assert res["new_slide_indexes"] == [1, 2] and res["slide_count"] == 4
    assert await titles(deck) == ["Orig", "Orig", "Orig", "Other"]
    assert (await slide(deck, 2))["notes"] == "note"


async def test_move_slide_keeps_everything(deck: Doc):
    await deck.edit(
        ("impress.add_slide", {"title": "One", "bullets": ["a", {"text": "b", "level": 1}]}),
        (
            "impress.add_slide",
            {
                "layout": "title_only",
                "title": "Two",
                "notes": "two notes",
                "background_color": "#112233",
            },
        ),
        ("impress.add_slide", {"title": "Three"}),
        (
            "impress.add_text_box",
            {"slide_index": 1, "x_mm": 10, "y_mm": 100, "width_mm": 50, "text": "box"},
        ),
        ("impress.set_slide_transition", {"slide_index": 1, "effect": "fade"}),
        ("impress.set_slide_hidden", {"slide_index": 1}),
    )
    await deck.edit(("impress.move_slide", {"from_index": 1, "to_index": 2}))
    assert await titles(deck) == ["One", "Three", "Two"]
    moved = await slide(deck, 2)
    assert moved["notes"] == "two notes" and moved["hidden"] is True
    assert moved["background"] == "#112233" and moved["transition"]["effect"] == "fade"
    assert shape_types(moved) == ["title", "text_box"]  # no resurrected empty body placeholder
    assert moved["layout"] == "title_only"
    await deck.edit(("impress.move_slide", {"from_index": 2, "to_index": 0}))
    assert await titles(deck) == ["Two", "One", "Three"]
    first = await slide(deck, 1)
    body = next(s for s in first["shapes"] if s["type"] == "body")
    assert body["levels"] == [0, 1] and body["placeholder"] is True
    with pytest.raises(OfficeError, match="to_index 7 is out of range"):
        await deck.edit(("impress.move_slide", {"from_index": 0, "to_index": 7}))


async def test_move_slide_to_same_place_is_a_noop(deck: Doc):
    await make_deck(deck, ["A", "B"])
    before = deck.path.read_bytes()
    await deck.edit(("impress.move_slide", {"from_index": 1, "to_index": 1}))
    assert deck.path.read_bytes() == before


async def test_set_slide_layout_keeps_text(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "T", "bullets": ["keep me"]}))
    (res,) = await deck.edit(
        ("impress.set_slide_layout", {"slide_index": 0, "layout": "two_content"})
    )
    assert [s["type"] for s in res["shapes"]] == ["title", "body", "body"]
    info = await slide(deck, 0)
    assert [s.get("text") for s in info["shapes"]][:2] == ["T", "keep me"]
    (again,) = await deck.edit(
        ("impress.set_slide_layout", {"slide_index": 0, "layout": "two_content"})
    )
    assert again["layout"] == "two_content"


async def test_set_slide_size_scales_content(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "Wide"}))
    before = (await slide(deck, 0))["shapes"][0]
    await deck.edit(("impress.set_slide_size", {"preset": "4:3"}))
    overview = await deck.read("impress.get_overview")
    assert (overview["slide_width_mm"], overview["slide_height_mm"]) == (280.0, 210.0)
    after = (await slide(deck, 0))["shapes"][0]
    assert after["width_mm"] == before["width_mm"]  # same width, taller slide
    assert after["height_mm"] == pytest.approx(before["height_mm"] * 210 / 157.5, abs=0.2)
    await deck.edit(
        ("impress.set_slide_size", {"width_mm": 200, "height_mm": 100, "scale_content": False})
    )
    assert (await deck.read("impress.get_overview"))["slide_width_mm"] == 200.0
    with pytest.raises(OfficeError, match="Unknown preset"):
        await deck.edit(("impress.set_slide_size", {"preset": "square"}))
    with pytest.raises(OfficeError, match="both width_mm and height_mm"):
        await deck.edit(("impress.set_slide_size", {"width_mm": 100}))


async def test_hide_and_unhide_slide(deck: Doc):
    await make_deck(deck, ["A", "B"])
    await deck.edit(("impress.set_slide_hidden", {"slide_index": 1}))
    overview = await deck.read("impress.get_overview")
    assert overview["slides"][1]["hidden"] is True and "hidden" not in overview["slides"][0]
    await deck.edit(("impress.set_slide_hidden", {"slide_index": 1, "hidden": False}))
    assert "hidden" not in (await deck.read("impress.get_overview"))["slides"][1]
    before = deck.path.read_bytes()
    await deck.edit(("impress.set_slide_hidden", {"slide_index": 1, "hidden": False}))
    assert deck.path.read_bytes() == before  # already visible: nothing rewritten


async def test_slide_background_solid_gradient_reset(deck: Doc):
    await make_deck(deck, ["A", "B"])
    await deck.edit(("impress.set_slide_background", {"color": "#336699", "slide_index": 0}))
    assert (await slide(deck, 0))["background"] == "#336699"
    assert "background" not in await slide(deck, 1)
    await deck.edit(("impress.set_slide_background", {"color": "red"}))  # all slides
    assert (await slide(deck, 1))["background"] == "#FF0000"
    await deck.edit(
        (
            "impress.set_slide_background",
            {"color": "#000000", "gradient_end_color": "#FFFFFF", "slide_index": 0},
        )
    )
    await deck.edit(("impress.set_slide_background", {"reset": True}))
    assert "background" not in await slide(deck, 0)
    with pytest.raises(OfficeError, match="Give a colour"):
        await deck.edit(("impress.set_slide_background", {}))
    with pytest.raises(OfficeError, match="Invalid colour"):
        await deck.edit(("impress.set_slide_background", {"color": "not-a-colour"}))


@pytest.mark.parametrize("effect", ["fade", "dissolve", "wipe", "push", "cover", "random"])
async def test_transitions_roundtrip(deck: Doc, effect: str):
    await make_deck(deck, ["A"])
    await deck.edit(
        (
            "impress.set_slide_transition",
            {
                "slide_index": 0,
                "effect": effect,
                "direction": "from_right",
                "speed": "fast",
                "advance_after_seconds": 5,
            },
        )
    )
    info = (await slide(deck, 0))["transition"]
    assert info["effect"] == effect and info["advance_after_seconds"] == 5
    if effect in ("wipe", "push", "cover"):
        assert info["direction"] == "from_right"
    await deck.edit(
        (
            "impress.set_slide_transition",
            {"slide_index": 0, "effect": "none", "advance_after_seconds": 0},
        )
    )
    assert "transition" not in await slide(deck, 0)


async def test_transition_validation(deck: Doc):
    with pytest.raises(OfficeError, match="Unknown effect 'spin'"):
        await deck.edit(("impress.set_slide_transition", {"effect": "spin"}))
    with pytest.raises(OfficeError, match="Give an effect"):
        await deck.edit(("impress.set_slide_transition", {}))


async def test_footer_settings(deck: Doc):
    await make_deck(deck, ["A", "B"])
    await deck.edit(("impress.set_footer", {"footer_text": "Acme", "show_slide_number": True}))
    await deck.edit(("impress.set_footer", {"slide_index": 1, "footer_text": ""}))
    with pytest.raises(OfficeError, match="Give footer_text"):
        await deck.edit(("impress.set_footer", {}))


async def test_add_slide_at_index_positions(deck: Doc):
    await make_deck(deck, ["B", "D"])
    await deck.edit(("impress.add_slide", {"title": "C", "at_index": 1}))
    await deck.edit(("impress.add_slide", {"title": "A", "at_index": 0, "notes": "first"}))
    assert await titles(deck) == ["A", "B", "C", "D"]
    assert (await slide(deck, 0))["notes"] == "first" and "notes" not in await slide(deck, 1)
    await deck.edit(("impress.add_slide", {"title": "E", "at_index": 4}))
    assert await titles(deck) == ["A", "B", "C", "D", "E"]
    with pytest.raises(OfficeError, match=r"at_index 9 is out of range.*0\.\.5"):
        await deck.edit(("impress.add_slide", {"title": "X", "at_index": 9}))


# -- text and notes ---------------------------------------------------------------------------


async def test_set_title_and_subtitle_and_missing_placeholder(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "title_slide"}))
    await deck.edit(("impress.set_title", {"slide_index": 0, "text": "Hello", "subtitle": "World"}))
    assert [s["text"] for s in (await slide(deck, 0))["shapes"]] == ["Hello", "World"]
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    with pytest.raises(OfficeError, match=r"has no title placeholder.*impress_set_slide_layout"):
        await deck.edit(("impress.set_title", {"slide_index": 1, "text": "x"}))
    with pytest.raises(OfficeError, match="Give text"):
        await deck.edit(("impress.set_title", {"slide_index": 0}))


async def test_set_body_levels_and_content_index(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "two_content", "title": "T"}))
    await deck.edit(
        ("impress.set_body", {"slide_index": 0, "items": ["a", {"text": "b", "level": 2}, "c"]}),
        ("impress.set_body", {"slide_index": 0, "items": "right side", "content_index": 1}),
    )
    shapes = (await slide(deck, 0))["shapes"]
    bodies = [s for s in shapes if s["type"] == "body"]
    assert bodies[0]["levels"] == [0, 2, 0] and bodies[1]["text"] == "right side"
    detail = await deck.read("impress.get_shape", slide_index=0, shape_index=bodies[0]["index"])
    assert [p["level"] for p in detail["paragraphs"]] == [0, 2, 0]
    assert all(p["bullet"] for p in detail["paragraphs"])  # every paragraph has a bullet level
    with pytest.raises(OfficeError, match=r"content_index 5 is out of range.*2 content"):
        await deck.edit(
            ("impress.set_body", {"slide_index": 0, "items": ["x"], "content_index": 5})
        )
    with pytest.raises(OfficeError, match="level 12 is out of range"):
        await deck.edit(
            ("impress.set_body", {"slide_index": 0, "items": [{"text": "x", "level": 12}]})
        )


async def test_set_shape_text_variants(deck: Doc):
    await deck.edit(
        ("impress.add_slide", {"title": "T", "bullets": ["old"]}),
        (
            "impress.add_text_box",
            {"slide_index": 0, "x_mm": 10, "y_mm": 100, "width_mm": 60, "text": "box"},
        ),
    )
    await deck.edit(
        (
            "impress.set_shape_text",
            {"slide_index": 0, "shape_index": 1, "paragraphs": ["x", {"text": "y", "level": 1}]},
        ),
        ("impress.set_shape_text", {"slide_index": 0, "shape_index": 2, "text": "line1\nline2"}),
    )
    info = await slide(deck, 0)
    assert info["shapes"][1]["text"] == "x\ny" and info["shapes"][1]["levels"] == [0, 1]
    assert info["shapes"][2]["text"] == "line1\nline2"
    with pytest.raises(OfficeError, match="either 'text' or 'paragraphs'"):
        await deck.edit(("impress.set_shape_text", {"slide_index": 0, "shape_index": 1}))


async def test_speaker_notes_set_get_clear(deck: Doc):
    await make_deck(deck, ["A", "B", "C"])
    await deck.edit(
        ("impress.set_notes", {"slide_index": 0, "text": "intro notes"}),
        ("impress.set_notes", {"slide_index": 2, "text": "line1\nline2"}),
    )
    assert (await deck.read("impress.get_notes", slide_index=0))["notes"] == "intro notes"
    everything = await deck.read("impress.get_notes")
    assert [n["slide_index"] for n in everything["slides"]] == [0, 2]
    overview = await deck.read("impress.get_overview")
    assert overview["slides"][0]["has_notes"] is True and "has_notes" not in overview["slides"][1]
    await deck.edit(("impress.set_notes", {"slide_index": 0, "text": ""}))
    assert (await deck.read("impress.get_notes"))["count"] == 1
    before = deck.path.read_bytes()
    await deck.edit(("impress.set_notes", {"slide_index": 1, "text": ""}))  # already empty
    assert deck.path.read_bytes() == before


async def test_get_text_outline_and_paging(deck: Doc):
    await deck.edit(
        ("impress.add_slides", {"slides": [
            {"title": "Intro", "bullets": ["a", {"text": "b", "level": 1}], "notes": "say hi"},
            {"title": "Data", "table": {"data": [["k", "v"], ["x", 1]]}},
            {"title": "More", "bullets": ["z"]},
        ]}),
    )  # fmt: skip
    out = (await deck.read("impress.get_text"))["outline"]
    assert "Slide 0: Intro" in out and "  - a" in out and "    - b" in out
    assert "Notes: say hi" in out and "k | v" in out and "x | 1" in out
    assert "Notes" not in (await deck.read("impress.get_text", include_notes=False))["outline"]
    paged = await deck.read("impress.get_text", max_chars=100)
    assert paged["slides_returned"] < 3 and paged["next_start_slide"] == paged["slides_returned"]
    with pytest.raises(OfficeError, match="past the end"):
        await deck.read("impress.get_text", start_slide=9)


async def test_find_and_replace_text(deck: Doc):
    await deck.edit(
        ("impress.add_slides", {"slides": [
            {"title": "Hello World", "bullets": ["say hello", "bye"], "notes": "hello notes"},
            {"title": "Table", "table": {"data": [["hello", "x"]]}},
        ]}),
    )  # fmt: skip
    found = await deck.read("impress.find_text", query="HELLO")
    kinds = sorted((m["type"], m["slide_index"]) for m in found["matches"])
    assert kinds == [("body", 0), ("notes", 0), ("table", 1), ("title", 0)]
    assert (await deck.read("impress.find_text", query="HELLO", case_sensitive=True))[
        "total_matches"
    ] == 0
    (res,) = await deck.edit(("impress.replace_text", {"search": "hello", "replace": "howdy"}))
    assert res["replacements"] == 3  # title, bullet, notes (table cells are not searched)
    out = (await deck.read("impress.get_text"))["outline"]
    assert "howdy" in out and "Hello" not in out
    before = deck.path.read_bytes()
    await deck.edit(("impress.replace_text", {"search": "nothing-here", "replace": "x"}))
    assert deck.path.read_bytes() == before
    with pytest.raises(OfficeError, match="query must not be empty"):
        await deck.read("impress.find_text", query="")


# -- shapes -----------------------------------------------------------------------------------


async def test_add_text_box_formatting_and_validation(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    (res,) = await deck.edit(
        (
            "impress.add_text_box",
            {
                "slide_index": 0,
                "x_mm": 20,
                "y_mm": 30,
                "width_mm": 100,
                "text": "Styled",
                "font_size": 30,
                "bold": True,
                "color": "#FF0000",
                "alignment": "center",
                "fill_color": "#EEEEEE",
                "name": "note",
            },
        )
    )
    assert (res["x_mm"], res["y_mm"], res["width_mm"]) == (20.0, 30.0, 100.0)
    shape = await deck.read("impress.get_shape", slide_index=0, shape_index=res["shape_index"])
    para = shape["paragraphs"][0]
    assert (para["size_pt"], para["bold"], para["color"], para["alignment"]) == (
        30.0,
        True,
        "#FF0000",
        "center",
    )
    assert shape["name"] == "note" and shape["fill"] == "#EEEEEE"
    with pytest.raises(OfficeError, match="either 'text' or 'paragraphs'"):
        await deck.edit(
            ("impress.add_text_box", {"slide_index": 0, "x_mm": 0, "y_mm": 0, "width_mm": 10})
        )


async def test_add_text_box_as_bullet_list(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    await deck.edit(
        (
            "impress.add_text_box",
            {
                "slide_index": 0,
                "x_mm": 20,
                "y_mm": 30,
                "width_mm": 100,
                "bullets": "bullets",
                "paragraphs": ["one", {"text": "two", "level": 1}],
            },
        ),
        (
            "impress.add_text_box",
            {
                "slide_index": 0,
                "x_mm": 20,
                "y_mm": 80,
                "width_mm": 100,
                "bullets": "numbered",
                "paragraphs": ["first", "second"],
            },
        ),
    )
    for index in (0, 1):
        detail = await deck.read("impress.get_shape", slide_index=0, shape_index=index)
        assert all(p["bullet"] for p in detail["paragraphs"])


async def test_every_shape_type_can_be_added(deck: Doc):
    types = await shape_type_names(deck)
    assert len(types) >= 25
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    ops = [
        (
            "impress.add_shape",
            {"slide_index": 0, "shape": t, "x_mm": 10, "y_mm": 10, "width_mm": 30, "height_mm": 20},
        )
        for t in types
    ]
    results = await deck.edit(*ops)
    assert [r["shape_index"] for r in results] == list(range(len(types)))
    info = await slide(deck, 0)
    assert info["shape_count"] == len(types)
    assert next(s for s in info["shapes"] if s.get("geometry") == "ellipse")["type"] == "shape"


async def shape_type_names(deck: Doc) -> list[str]:
    with pytest.raises(OfficeError, match="Valid shapes: ") as err:
        await deck.edit(
            (
                "impress.add_shape",
                {
                    "slide_index": 0,
                    "shape": "blob",
                    "x_mm": 0,
                    "y_mm": 0,
                    "width_mm": 1,
                    "height_mm": 1,
                },
            )
        )
    return str(err.value).split("Valid shapes: ")[1].split(".")[0].split(", ")


async def test_add_shape_with_text_and_style(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    (res,) = await deck.edit(
        (
            "impress.add_shape",
            {
                "slide_index": 0,
                "shape": "rounded_rectangle",
                "x_mm": 20,
                "y_mm": 20,
                "width_mm": 80,
                "height_mm": 30,
                "text": "Plan",
                "fill_color": "#1F4E79",
                "no_line": True,
                "text_color": "#FFFFFF",
                "font_size": 24,
                "bold": True,
                "shadow": True,
                "transparency": 20,
            },
        )
    )
    shape = await deck.read("impress.get_shape", slide_index=0, shape_index=res["shape_index"])
    assert shape["fill"] == "#1F4E79" and "line" not in shape and shape["shadow"] is True
    assert shape["fill_transparency"] == 20 and shape["geometry"] == "round-rectangle"
    para = shape["paragraphs"][0]
    assert para["color"] == "#FFFFFF" and para["bold"] is True and para["alignment"] == "center"


async def test_add_line_and_arrow(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "blank"}))
    (res,) = await deck.edit(
        (
            "impress.add_line",
            {
                "slide_index": 0,
                "x1_mm": 100,
                "y1_mm": 80,
                "x2_mm": 20,
                "y2_mm": 20,
                "arrow_end": True,
                "color": "#FF0000",
                "dash": "dashed",
            },
        )
    )
    shape = (await slide(deck, 0))["shapes"][res["shape_index"]]
    assert shape["type"] == "line" and shape["line"] == "#FF0000"
    assert (shape["x_mm"], shape["y_mm"], shape["width_mm"], shape["height_mm"]) == (
        20.0,
        20.0,
        80.0,
        60.0,
    )
    with pytest.raises(OfficeError, match="identical"):
        await deck.edit(
            ("impress.add_line", {"slide_index": 0, "x1_mm": 5, "y1_mm": 5, "x2_mm": 5, "y2_mm": 5})
        )


async def test_add_image_sizing_and_errors(deck: Doc, tmp_path: Path):
    png = make_png(tmp_path / "wide.png", 200, 100)  # 2:1
    await deck.edit(("impress.add_slide", {"layout": "title_only", "title": "Pic"}))
    r1, r2, r3 = await deck.edit(
        ("impress.add_image", {"slide_index": 0, "image_path": str(png), "alt_text": "wide"}),
        (
            "impress.add_image",
            {"slide_index": 0, "image_path": str(png), "width_mm": 60, "x_mm": 10, "y_mm": 100},
        ),
        (
            "impress.add_image",
            {"slide_index": 0, "image_path": str(png), "height_mm": 20, "x_mm": 100, "y_mm": 100},
        ),
    )
    assert r1["width_mm"] / r1["height_mm"] == pytest.approx(2.0, rel=0.02)  # fitted, aspect kept
    assert r1["y_mm"] > 30  # below the title, inside the content area
    assert (r2["width_mm"], r2["height_mm"], r2["x_mm"]) == (60.0, 30.0, 10.0)
    assert (r3["width_mm"], r3["height_mm"]) == (40.0, 20.0)
    assert (await slide(deck, 0))["shapes"][r1["shape_index"]]["alt_text"] == "wide"
    bad = tmp_path / "not-an-image.png"
    bad.write_bytes(b"this is not a png")
    with pytest.raises(OfficeError, match="not an image LibreOffice can read"):
        await deck.edit(("impress.add_image", {"slide_index": 0, "image_path": str(bad)}))


# -- tables -----------------------------------------------------------------------------------


async def test_table_create_edit_modify(deck: Doc):
    await deck.edit(("impress.add_slide", {"layout": "title_only", "title": "T"}))
    (res,) = await deck.edit(
        (
            "impress.add_table",
            {
                "slide_index": 0,
                "data": [["Name", "Qty"], ["Apple", 3], ["Pear", 4.5]],
                "column_widths": [3, 1],
            },
        )
    )
    assert (res["rows"], res["columns"]) == (3, 2)
    index = res["shape_index"]
    grid = (await deck.read("impress.get_shape", slide_index=0, shape_index=index))["data"]
    assert grid == [["Name", "Qty"], ["Apple", "3"], ["Pear", "4.5"]]
    await deck.edit(
        (
            "impress.set_table_cells",
            {
                "slide_index": 0,
                "shape_index": index,
                "cells": [
                    {"row": 1, "column": 1, "text": "30", "bold": True, "fill": "#FFFF00"},
                    {"row": 0, "column": 0, "text": "Item", "alignment": "right"},
                ],
            },
        ),
        (
            "impress.modify_table",
            {"slide_index": 0, "shape_index": index, "action": "insert_rows", "index": 3},
        ),
        (
            "impress.modify_table",
            {"slide_index": 0, "shape_index": index, "action": "insert_columns", "index": 0},
        ),
    )
    grid = (await deck.read("impress.get_shape", slide_index=0, shape_index=index))["data"]
    assert len(grid) == 4 and len(grid[0]) == 3
    assert grid[0][1] == "Item" and grid[1][2] == "30" and grid[2][1] == "Pear"
    await deck.edit(
        (
            "impress.modify_table",
            {"slide_index": 0, "shape_index": index, "action": "delete_rows", "index": 3},
        )
    )
    summary = (await slide(deck, 0))["shapes"][index]
    assert summary["rows"] == 3 and summary["columns"] == 3 and len(summary["preview"]) == 3


async def test_table_errors(deck: Doc):
    await deck.edit(
        ("impress.add_slides", {"slides": [{"title": "T", "table": {"data": [["a", "b"]]}}]})
    )
    table_index = next(s["index"] for s in (await slide(deck, 0))["shapes"] if s["type"] == "table")
    with pytest.raises(
        OfficeError, match=r"row 5, column 0 is outside the table \(1 rows x 2 columns"
    ):
        await deck.edit(
            (
                "impress.set_table_cells",
                {
                    "slide_index": 0,
                    "shape_index": table_index,
                    "cells": [{"row": 5, "column": 0, "text": "x"}],
                },
            )
        )
    with pytest.raises(OfficeError, match="is a title, not a table"):
        await deck.edit(
            (
                "impress.set_table_cells",
                {"slide_index": 0, "shape_index": 0, "cells": [{"row": 0, "column": 0}]},
            )
        )
    with pytest.raises(OfficeError, match="at least one row and one column"):
        await deck.edit(
            (
                "impress.modify_table",
                {"slide_index": 0, "shape_index": table_index, "action": "delete_rows", "index": 0},
            )
        )
    with pytest.raises(OfficeError, match="column_widths has 3 entries but the table has 2"):
        await deck.edit(
            (
                "impress.add_table",
                {"slide_index": 0, "data": [["a", "b"]], "column_widths": [1, 1, 1]},
            )
        )
    with pytest.raises(OfficeError, match="non-empty list of rows"):
        await deck.edit(("impress.add_table", {"slide_index": 0, "data": []}))
    with pytest.raises(OfficeError, match="no text"):
        await deck.edit(
            ("impress.set_shape_text", {"slide_index": 0, "shape_index": table_index, "text": "x"})
        )


# -- arranging ---------------------------------------------------------------------------------


async def blank_with_boxes(deck: Doc, count: int = 3) -> None:
    await deck.edit(
        ("impress.add_slide", {"layout": "blank"}),
        *[
            (
                "impress.add_shape",
                {
                    "slide_index": 0,
                    "shape": "rectangle",
                    "x_mm": 10 + 50 * i + 7 * i,
                    "y_mm": 20 + 9 * i,
                    "width_mm": 30 + 5 * i,
                    "height_mm": 20,
                    "text": f"S{i}",
                },
            )
            for i in range(count)
        ],
    )


async def test_geometry_move_resize_rotate(deck: Doc):
    await blank_with_boxes(deck, 1)
    (res,) = await deck.edit(
        (
            "impress.set_shape_geometry",
            {
                "slide_index": 0,
                "shape_index": 0,
                "x_mm": 50,
                "y_mm": 60,
                "width_mm": 40,
                "height_mm": 10,
            },
        )
    )
    assert (res["x_mm"], res["y_mm"], res["width_mm"], res["height_mm"]) == (50.0, 60.0, 40.0, 10.0)
    await deck.edit(
        ("impress.set_shape_geometry", {"slide_index": 0, "shape_index": 0, "rotation_deg": 90})
    )
    shape = (await slide(deck, 0))["shapes"][0]
    assert shape["rotation_deg"] == pytest.approx(90.0, abs=0.5)
    await deck.edit(
        ("impress.set_shape_geometry", {"slide_index": 0, "shape_index": 0, "rotation_deg": 0})
    )
    assert "rotation_deg" not in (await slide(deck, 0))["shapes"][0]
    (res,) = await deck.edit(
        (
            "impress.set_shape_geometry",
            {"slide_index": 0, "shape_index": 0, "width_mm": 80, "keep_aspect_ratio": True},
        )
    )
    assert res["height_mm"] == 20.0  # 40 x 10 -> 80 x 20
    with pytest.raises(OfficeError, match="at least one of"):
        await deck.edit(("impress.set_shape_geometry", {"slide_index": 0, "shape_index": 0}))
    with pytest.raises(OfficeError, match="width_mm must be positive"):
        await deck.edit(
            ("impress.set_shape_geometry", {"slide_index": 0, "shape_index": 0, "width_mm": -5})
        )


async def test_z_order(deck: Doc):
    await blank_with_boxes(deck, 3)  # indexes 0, 1, 2 (bottom -> top), texts S0, S1, S2

    async def order() -> list[str]:
        return [s["text"] for s in (await slide(deck, 0))["shapes"]]

    (res,) = await deck.edit(
        ("impress.set_shape_order", {"slide_index": 0, "shape_index": 0, "action": "front"})
    )
    assert res["shape_index"] == 2 and await order() == ["S1", "S2", "S0"]
    (res,) = await deck.edit(
        ("impress.set_shape_order", {"slide_index": 0, "shape_index": 2, "action": "backward"})
    )
    assert res["shape_index"] == 1 and await order() == ["S1", "S0", "S2"]
    await deck.edit(
        ("impress.set_shape_order", {"slide_index": 0, "shape_index": 1, "action": "back"})
    )
    assert await order() == ["S0", "S1", "S2"]
    (res,) = await deck.edit(
        ("impress.set_shape_order", {"slide_index": 0, "shape_index": 0, "action": "forward"})
    )
    assert await order() == ["S1", "S0", "S2"]
    before = deck.path.read_bytes()
    await deck.edit(
        ("impress.set_shape_order", {"slide_index": 0, "shape_index": 2, "action": "front"})
    )
    assert deck.path.read_bytes() == before


async def test_delete_shapes(deck: Doc):
    await blank_with_boxes(deck, 4)
    (res,) = await deck.edit(
        ("impress.delete_shapes", {"slide_index": 0, "shape_indexes": [1, 3, 1]})
    )
    assert res["deleted"] == 2 and res["shapes_remaining"] == 2
    assert [s["text"] for s in (await slide(deck, 0))["shapes"]] == ["S0", "S2"]
    with pytest.raises(OfficeError, match=r"shape_index 7 is out of range.*2 shape\(s\)"):
        await deck.edit(("impress.delete_shapes", {"slide_index": 0, "shape_indexes": [7]}))
    with pytest.raises(OfficeError, match="must not be empty"):
        await deck.edit(("impress.delete_shapes", {"slide_index": 0, "shape_indexes": []}))


async def test_duplicate_shape_same_and_other_slide(deck: Doc):
    await blank_with_boxes(deck, 1)
    await deck.edit(
        ("impress.format_shape", {"slide_index": 0, "shape_index": 0, "fill_color": "#00FF00"}),
        ("impress.add_slide", {"layout": "blank"}),
    )
    (same,) = await deck.edit(
        ("impress.duplicate_shape", {"slide_index": 0, "shape_index": 0, "count": 2})
    )
    assert same["new_shape_indexes"] == [1, 2]
    shapes = (await slide(deck, 0))["shapes"]
    assert (
        shapes[1]["x_mm"] == shapes[0]["x_mm"] + 10 and shapes[2]["x_mm"] == shapes[0]["x_mm"] + 20
    )
    assert shapes[1]["text"] == "S0" and shapes[1]["fill"] == "#00FF00"
    (other,) = await deck.edit(
        ("impress.duplicate_shape", {"slide_index": 0, "shape_index": 0, "target_slide_index": 1})
    )
    assert other["slide_index"] == 1 and other["new_shape_indexes"] == [0]
    moved = (await slide(deck, 1))["shapes"][0]
    assert (moved["x_mm"], moved["y_mm"]) == (shapes[0]["x_mm"], shapes[0]["y_mm"])


async def test_align_and_distribute(deck: Doc):
    await blank_with_boxes(deck, 3)
    await deck.edit(
        ("impress.align_shapes", {"slide_index": 0, "shape_indexes": [0, 1, 2], "align": "top"})
    )
    tops = {s["y_mm"] for s in (await slide(deck, 0))["shapes"]}
    assert tops == {20.0}
    await deck.edit(
        (
            "impress.align_shapes",
            {"slide_index": 0, "shape_indexes": [0, 1, 2], "distribute": "horizontal"},
        )
    )
    s0, s1, s2 = (await slide(deck, 0))["shapes"]
    gap1 = s1["x_mm"] - (s0["x_mm"] + s0["width_mm"])
    gap2 = s2["x_mm"] - (s1["x_mm"] + s1["width_mm"])
    assert gap1 == pytest.approx(gap2, abs=0.2)
    await deck.edit(
        (
            "impress.align_shapes",
            {"slide_index": 0, "shape_indexes": [1], "align": "center", "relative_to": "slide"},
        )
    )
    mid = (await slide(deck, 0))["shapes"][1]
    assert mid["x_mm"] + mid["width_mm"] / 2 == pytest.approx(140.0, abs=0.2)
    with pytest.raises(OfficeError, match="at least three"):
        await deck.edit(
            (
                "impress.align_shapes",
                {"slide_index": 0, "shape_indexes": [0, 1], "distribute": "vertical"},
            )
        )
    with pytest.raises(OfficeError, match="Give align and/or distribute"):
        await deck.edit(("impress.align_shapes", {"slide_index": 0, "shape_indexes": [0, 1]}))


async def test_shape_description(deck: Doc, tmp_path: Path):
    png = make_png(tmp_path / "p.png")
    await deck.edit(
        ("impress.add_slide", {"layout": "blank"}),
        ("impress.add_image", {"slide_index": 0, "image_path": str(png)}),
    )
    await deck.edit(
        (
            "impress.set_shape_description",
            {"slide_index": 0, "shape_index": 0, "name": "logo", "alt_text": "Company logo"},
        )
    )
    shape = (await slide(deck, 0))["shapes"][0]
    assert shape["name"] == "logo" and shape["alt_text"] == "Company logo"
    with pytest.raises(OfficeError, match="Give a name and/or alt_text"):
        await deck.edit(("impress.set_shape_description", {"slide_index": 0, "shape_index": 0}))


# -- formatting ----------------------------------------------------------------------------------


async def test_format_text_whole_paragraph_and_find(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "T", "bullets": ["alpha beta", "gamma alpha"]}))
    await deck.edit(
        (
            "impress.format_text",
            {
                "slide_index": 0,
                "shape_index": 1,
                "paragraph_index": 1,
                "bold": True,
                "size": 40,
                "color": "#0000FF",
                "alignment": "right",
                "line_spacing_percent": 150,
            },
        ),
        (
            "impress.format_text",
            {
                "slide_index": 0,
                "shape_index": 1,
                "find": "alpha",
                "italic": True,
                "underline": True,
                "highlight_color": "#FFFF00",
            },
        ),
    )
    paras = (await deck.read("impress.get_shape", slide_index=0, shape_index=1))["paragraphs"]
    assert (paras[1]["bold"], paras[1]["size_pt"], paras[1]["color"], paras[1]["alignment"]) == (
        True,
        40.0,
        "#0000FF",
        "right",
    )
    assert paras[0]["bold"] is False and paras[0]["alignment"] == "left"
    # 'alpha' is italic + underlined in BOTH paragraphs, but only partly so; the paragraph-level
    # readback reports the first run's formatting, so check through the run-level API instead.
    (res,) = await deck.edit(
        (
            "impress.format_text",
            {"slide_index": 0, "shape_index": 1, "find": "alpha", "bold": False},
        )
    )
    assert res["occurrences_formatted"] == 2


async def test_format_text_bullets_and_levels(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "T", "bullets": ["a", "b", "c"]}))
    await deck.edit(
        (
            "impress.format_text",
            {"slide_index": 0, "shape_index": 1, "paragraph_index": 0, "bullets": "none"},
        ),
        (
            "impress.format_text",
            {"slide_index": 0, "shape_index": 1, "paragraph_index": 1, "level": 1},
        ),
        (
            "impress.format_text",
            {"slide_index": 0, "shape_index": 1, "paragraph_index": 2, "bullets": "numbered"},
        ),
    )
    paras = (await deck.read("impress.get_shape", slide_index=0, shape_index=1))["paragraphs"]
    assert [p["bullet"] for p in paras] == [False, True, True]
    assert [p["level"] for p in paras] == [0, 1, 0]
    await deck.edit(
        (
            "impress.format_text",
            {"slide_index": 0, "shape_index": 1, "paragraph_index": 0, "bullets": "bullets"},
        )
    )
    assert (await deck.read("impress.get_shape", slide_index=0, shape_index=1))["paragraphs"][0][
        "bullet"
    ] is True


async def test_format_text_errors(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "T", "bullets": ["a"]}))
    base = {"slide_index": 0, "shape_index": 1}
    with pytest.raises(OfficeError, match="Nothing to change"):
        await deck.edit(("impress.format_text", base))
    with pytest.raises(OfficeError, match="was not found in this shape"):
        await deck.edit(("impress.format_text", {**base, "find": "zzz", "bold": True}))
    with pytest.raises(OfficeError, match="only character formatting"):
        await deck.edit(("impress.format_text", {**base, "find": "a", "alignment": "left"}))
    with pytest.raises(OfficeError, match=r"paragraph_index 4 is out of range.*1 paragraph"):
        await deck.edit(("impress.format_text", {**base, "paragraph_index": 4, "bold": True}))
    with pytest.raises(OfficeError, match="alignment must be one of"):
        await deck.edit(("impress.format_text", {**base, "alignment": "diagonal"}))
    with pytest.raises(OfficeError, match="Invalid colour"):
        await deck.edit(("impress.format_text", {**base, "color": "bluish"}))
    await deck.edit(("impress.add_slide", {"layout": "title_only"}))
    with pytest.raises(OfficeError, match="no text yet"):
        await deck.edit(("impress.format_text", {"slide_index": 1, "shape_index": 0, "bold": True}))


async def test_format_shape_options_and_errors(deck: Doc):
    await blank_with_boxes(deck, 1)
    base = {"slide_index": 0, "shape_index": 0}
    await deck.edit(
        (
            "impress.format_shape",
            {
                **base,
                "fill_color": "#FF8800",
                "gradient_end_color": "#FFFFFF",
                "gradient_angle": 90,
                "line_color": "#000000",
                "line_width_mm": 2,
                "dash": "dashed",
                "shadow": True,
                "corner_radius_mm": 5,
                "text_anchor": "bottom",
                "margin_mm": 3,
            },
        ),
    )
    shape = await deck.read("impress.get_shape", **base)
    assert (
        shape["fill"] == "#FF8800" and shape["line"] == "#000000" and shape["line_width_mm"] == 2.0
    )
    await deck.edit(("impress.format_shape", {**base, "no_fill": True, "no_line": True}))
    shape = await deck.read("impress.get_shape", **base)
    assert "fill" not in shape and "line" not in shape
    with pytest.raises(OfficeError, match="not both"):
        await deck.edit(("impress.format_shape", {**base, "fill_color": "#fff", "no_fill": True}))
    with pytest.raises(OfficeError, match="Nothing to change"):
        await deck.edit(("impress.format_shape", base))
    with pytest.raises(OfficeError, match="dash must be"):
        await deck.edit(("impress.format_shape", {**base, "dash": "wavy"}))
    with pytest.raises(OfficeError, match="transparency is a percentage"):
        await deck.edit(("impress.format_shape", {**base, "transparency": 150}))
    await deck.edit(
        (
            "impress.add_shape",
            {
                "slide_index": 0,
                "shape": "ellipse",
                "x_mm": 1,
                "y_mm": 1,
                "width_mm": 5,
                "height_mm": 5,
            },
        )
    )
    with pytest.raises(OfficeError, match="corner_radius_mm only applies to plain rectangles"):
        await deck.edit(
            ("impress.format_shape", {"slide_index": 0, "shape_index": 1, "corner_radius_mm": 3})
        )


async def test_hyperlinks_on_text_and_shape(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "T", "bullets": ["Visit our site today"]}))
    await deck.edit(
        (
            "impress.set_hyperlink",
            {"slide_index": 0, "shape_index": 1, "url": "https://example.com", "find": "our site"},
        ),
    )
    assert (await slide(deck, 0))["shapes"][1]["text"] == "Visit our site today"
    await deck.edit(
        (
            "impress.set_hyperlink",
            {"slide_index": 0, "shape_index": 0, "url": "https://example.org"},
        )
    )
    assert (await slide(deck, 0))["shapes"][0]["link"].rstrip("/") == "https://example.org"
    await deck.edit(("impress.set_hyperlink", {"slide_index": 0, "shape_index": 0, "url": ""}))
    assert "link" not in (await slide(deck, 0))["shapes"][0]
    with pytest.raises(OfficeError, match="was not found"):
        await deck.edit(
            (
                "impress.set_hyperlink",
                {"slide_index": 0, "shape_index": 1, "url": "https://x.y", "find": "nope"},
            )
        )


# -- master slides ----------------------------------------------------------------------------------


async def test_list_masters_and_layout_catalogue(deck: Doc):
    info = await deck.read("impress.list_masters")
    assert (
        info["masters"][0]["name"] == "Default" and info["masters"][0]["title"]["size_pt"] == 44.0
    )
    assert {layout["name"] for layout in info["layouts"]} >= {
        "title_slide",
        "title_content",
        "blank",
    }
    assert "dark" in info["themes"]


async def test_set_master_style_changes_whole_deck(deck: Doc):
    await make_deck(deck, ["A", "B"])
    (res,) = await deck.edit(
        (
            "impress.set_master_style",
            {
                "background": "#101820",
                "title_font": "Arial",
                "title_size": 36,
                "title_color": "#FFCC00",
                "title_bold": True,
                "body_font": "Arial",
                "body_size": 20,
                "body_color": "#DDDDDD",
            },
        )
    )
    assert res["masters"] == ["Default"]
    master = (await deck.read("impress.list_masters"))["masters"][0]
    assert master["background"] == "#101820"
    assert master["title"] == {"font": "Arial", "size_pt": 36.0, "color": "#FFCC00"}
    assert master["body_level_1"]["size_pt"] == 20.0 and master["body_level_2"]["size_pt"] < 20.0
    await deck.edit(("impress.add_slide", {"title": "Later", "bullets": ["x"]}))  # inherits too
    with pytest.raises(OfficeError, match="Nothing to change"):
        await deck.edit(("impress.set_master_style", {}))


async def test_apply_theme_and_unknown_theme(deck: Doc):
    await make_deck(deck, ["A"])
    (res,) = await deck.edit(("impress.apply_theme", {"theme": "corporate"}))
    assert res["theme"] == "corporate"
    master = (await deck.read("impress.list_masters"))["masters"][0]
    assert master["title"]["color"] == "#1F4E79" and master["title"]["font"] == "Calibri"
    with pytest.raises(OfficeError, match=r"Unknown theme 'neon'.*corporate"):
        await deck.edit(("impress.apply_theme", {"theme": "neon"}))
    with pytest.raises(OfficeError, match="No master slide named 'Nope'"):
        await deck.edit(("impress.set_master_style", {"master_name": "Nope", "background": "#fff"}))


# -- file formats ------------------------------------------------------------------------------------


@pytest.mark.parametrize("ext", ["odp", "pptx", "ppt"])
async def test_content_survives_roundtrip_in_every_format(
    office: OfficeClient, tmp_path: Path, ext: str
):
    png = make_png(tmp_path / "r.png", 80, 40)
    d = Doc(office, tmp_path / f"deck.{ext}")
    await d.create()
    await d.edit(
        ("impress.add_slides", {"slides": [
            {"title": "Round trip", "subtitle": "sub", "notes": "title notes"},
            {"title": "Bullets", "bullets": ["one", {"text": "two", "level": 1}, "three"], "notes": "bullet notes"},
            {"title": "Table", "table": {"data": [["a", "b"], ["1", "2"]]}},
            {"title": "Picture", "image": {"image_path": str(png), "alt_text": "red"}},
        ]}),
        ("impress.add_shape", {"slide_index": 3, "shape": "ellipse", "x_mm": 10, "y_mm": 100, "width_mm": 40, "height_mm": 20,
                                "text": "oval", "fill_color": "#00AA00"}),
        ("impress.set_slide_transition", {"slide_index": 1, "effect": "push", "direction": "from_left"}),
        ("impress.set_slide_hidden", {"slide_index": 2}),
    )  # fmt: skip
    # a second, separate load: everything must come back from the file
    assert await titles(d) == ["Round trip", "Bullets", "Table", "Picture"]
    s1 = await slide(d, 1)
    body = next(s for s in s1["shapes"] if s["type"] == "body")
    assert body["text"] == "one\ntwo\nthree" and body["levels"] == [0, 1, 0]
    assert s1["notes"] == "bullet notes"
    assert (await d.read("impress.get_notes", slide_index=0))["notes"] == "title notes"
    table = next(s for s in (await slide(d, 2))["shapes"] if s["type"] == "table")
    assert table["preview"] == [["a", "b"], ["1", "2"]]
    s3 = await slide(d, 3)
    assert any(s["type"] == "image" for s in s3["shapes"])
    assert any(s.get("text") == "oval" for s in s3["shapes"])
    assert (await d.read("impress.get_overview"))["slides"][2].get("hidden") is True
    if ext != "ppt":  # the binary format keeps the transition too, but under another name
        assert s1["transition"]["effect"] == "push"


async def test_theme_survives_pptx_roundtrip_and_applies_to_new_slides(
    office: OfficeClient, tmp_path: Path
):
    d = Doc(office, tmp_path / "themed.pptx")
    await d.create()
    await d.edit(
        ("impress.apply_theme", {"theme": "dark"}),
        ("impress.add_slide", {"title": "One", "bullets": ["x"]}),
    )
    await d.edit(
        ("impress.add_slide", {"title": "Two", "bullets": ["y"]})
    )  # a later call, after a reload
    masters = (await d.read("impress.list_masters"))["masters"]
    assert all(m["background"] == "#1B1F2A" for m in masters)
    for index in (0, 1):
        title = next(s for s in (await slide(d, index))["shapes"] if s["type"] == "title")
        detail = await d.read("impress.get_shape", slide_index=index, shape_index=title["index"])
        assert (
            detail["paragraphs"][0]["color"] == "#FFFFFF"
            and detail["paragraphs"][0]["font"] == "Calibri"
        )


async def test_ops_on_wrong_type_and_batch_atomicity(deck: Doc):
    await make_deck(deck, ["A"])
    before = deck.path.read_bytes()
    with pytest.raises(OfficeError, match="failed at operation #2"):
        await deck.edit(
            ("impress.add_slide", {"title": "ok"}),
            ("impress.set_shape_text", {"slide_index": 5, "shape_index": 0, "text": "boom"}),
        )
    assert deck.path.read_bytes() == before


async def test_title_slide_layout_is_centred_and_titles_shrink(deck: Doc):
    await deck.edit(("impress.add_slide", {"title": "Opening", "subtitle": "Sub"}))
    title, subtitle = (await slide(deck, 0))["shapes"]
    assert title["y_mm"] > 30 and subtitle["y_mm"] > title["y_mm"] + title["height_mm"] - 1
    assert 20 < title["x_mm"] < 40 and title["width_mm"] > 200  # centred block, not edge to edge


async def test_get_shape_details_for_image_table_line(deck: Doc, tmp_path: Path):
    png = make_png(tmp_path / "d.png")
    await deck.edit(
        ("impress.add_slide", {"layout": "blank"}),
        ("impress.add_image", {"slide_index": 0, "image_path": str(png), "alt_text": "alt"}),
        ("impress.add_table", {"slide_index": 0, "data": [["a", "b"], ["c", "d"]], "y_mm": 100}),
        ("impress.add_line", {"slide_index": 0, "x1_mm": 1, "y1_mm": 1, "x2_mm": 50, "y2_mm": 1}),
    )
    image = await deck.read("impress.get_shape", slide_index=0, shape_index=0)
    assert image["type"] == "image" and image["alt_text"] == "alt" and "paragraphs" not in image
    table = await deck.read("impress.get_shape", slide_index=0, shape_index=1)
    assert table["data"] == [["a", "b"], ["c", "d"]] and "preview" not in table
    line = await deck.read("impress.get_shape", slide_index=0, shape_index=2)
    assert line["type"] == "line" and line["width_mm"] == 49.0


async def test_set_slide_master_and_unknown_master(deck: Doc):
    await make_deck(deck, ["A", "B"])
    (res,) = await deck.edit(
        ("impress.set_slide_master", {"master_name": "Default", "slide_index": 1})
    )
    assert res["slides_changed"] == 1
    with pytest.raises(OfficeError, match=r"No master slide named 'Zed'.*Default"):
        await deck.edit(("impress.set_slide_master", {"master_name": "Zed"}))
    with pytest.raises(OfficeError, match="Unknown slide field"):
        await deck.edit(("impress.add_slides", {"slides": [{"title": "x", "bogus": 1}]}))
    with pytest.raises(OfficeError, match="image or a table"):
        await deck.edit(
            (
                "impress.add_slides",
                {"slides": [{"image": {"image_path": "/x.png"}, "table": {"data": [["a"]]}}]},
            )
        )


async def test_add_slides_with_an_image_that_is_not_there_names_the_slide(deck: Doc):
    with pytest.raises(OfficeError, match=r"slides\[0\]"):
        await deck.edit(
            (
                "impress.add_slides",
                {"slides": [{"title": "x", "image": {"image_path": "/nope/missing.png"}}]},
            )
        )
    assert (await deck.read("impress.get_overview"))["slide_count"] == 1
