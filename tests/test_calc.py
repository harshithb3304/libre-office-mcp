"""Calc operations, end to end against a real LibreOffice (worker + UNO).

Run just these with:  uv run pytest tests/test_calc.py -v   (inside the Linux test image)
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path
from typing import Any

import pytest

from libreoffice_mcp import protocol
from libreoffice_mcp.client import OfficeClient
from libreoffice_mcp.errors import OfficeError

pytestmark = pytest.mark.office


class Book:
    """Runs calc operations against one spreadsheet file through the real worker."""

    def __init__(self, office: OfficeClient, path: Path):
        self.office, self.path = office, path

    async def create(self, *ops: tuple[str, dict[str, Any]]) -> list[Any]:
        reply = await self.office.call(
            protocol.KIND_CREATE, self.path, [{"op": n, "args": a} for n, a in ops],
            options={"overwrite": True},
        )  # fmt: skip
        return reply["results"]

    async def edit(self, *ops: tuple[str, dict[str, Any]]) -> list[Any]:
        reply = await self.office.call(
            protocol.KIND_EDIT, self.path, [{"op": n, "args": a} for n, a in ops]
        )
        return reply["results"]

    async def do(self, op: str, **args: Any) -> Any:
        """One editing op; returns its result."""
        return (await self.edit((op, args)))[0]

    async def read(self, op: str, **args: Any) -> Any:
        reply = await self.office.call(protocol.KIND_READ, self.path, [{"op": op, "args": args}])
        return reply["results"][0]

    async def values(self, range: str | None = None, sheet: str | None = None, **kw: Any) -> list:
        return (await self.read("calc.get_range", sheet=sheet, range=range, **kw))["values"]

    async def convert(self, target: Path) -> Path:
        await self.office.call(protocol.KIND_EXPORT, self.path, [], target=target)
        return target


@pytest.fixture
async def book(office: OfficeClient, tmp_path: Path) -> Book:
    b = Book(office, tmp_path / "test.ods")
    await b.create()
    return b


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


SALES = [
    ["Region", "Units", "Price"],
    ["North", 10, 2.5],
    ["South", 4, 3],
    ["East", 7, 1.25],
    ["West", 1, 10],
]


# -- values and formulas -------------------------------------------------------


async def test_new_spreadsheet_has_one_empty_sheet(book: Book):
    info = await book.read("calc.list_sheets")
    assert info["count"] == 1
    sheet = info["sheets"][0]
    assert sheet["name"] == "Sheet1" and sheet["used_range"] is None and sheet["visible"]


async def test_set_range_values_types_and_read_back(book: Book):
    result = await book.do(
        "calc.set_range_values",
        data=[["Name", "Qty", "OK", None], ["Apples", 3, True, "x"], ["Pears", 2.5, False, ""]],
    )
    assert result["range"] == "A1:D3" and result["cells_written"] == 12
    assert result["formulas"] == 2  # booleans are written as TRUE()/FALSE()
    values = await book.values()
    assert values[0] == ["Name", "Qty", "OK", ""]
    assert values[1][:2] == ["Apples", 3] and values[1][3] == "x"
    assert values[2][1] == 2.5
    assert values[1][2] == 1 and values[2][2] == 0  # TRUE / FALSE
    cell = await book.read("calc.get_cell_info", cell="C2")
    assert cell["displayed"] == "TRUE"


async def test_strings_stay_text_unless_parse_text(book: Book):
    await book.do("calc.set_range_values", data=[["007", "12", "2024-01-31", "50%"]])
    assert (await book.values())[0] == ["007", "12", "2024-01-31", "50%"]
    await book.do(
        "calc.set_range_values", data=[["007", "12", "2024-01-31", "50%"]], start_cell="A2",
        parse_text=True,
    )  # fmt: skip
    row = (await book.values("A2:D2"))[0]
    assert row[0] == 7 and row[1] == 12 and row[3] == 0.5
    assert isinstance(row[2], int | float) and row[2] > 40000  # a date serial number
    # typed-in dates and percentages get the format keyboard input would give them
    shown = await book.read("calc.get_range", range="A2:D2", formatted=True)
    assert shown["formatted"][0][2] == "2024-01-31" and shown["formatted"][0][3] == "50.00%"


async def test_formulas_accept_excel_and_libreoffice_syntax(book: Book):
    await book.do("calc.set_range_values", data=[[1, 2], [3, 4]])
    await book.do("calc.add_sheet", name="Other Sheet")
    await book.do("calc.set_range_values", sheet="Other Sheet", data=[[100]])
    result = await book.do(
        "calc.set_range_values", start_cell="D1",
        data=[
            ["=SUM(A1:B2)"],            # Excel commas would also work: see next
            ["=SUM(A1:A2,B1:B2)"],      # comma-separated arguments
            ["=A1&\",\"&B1"],           # comma inside a string literal is preserved
            ["='Other Sheet'!A1+1"],    # Excel sheet reference with quotes
            ["=IF(A1>0,\"yes, ok\",\"no\")"],
        ],
    )  # fmt: skip
    assert "formula_errors" not in result, result
    assert [r[0] for r in await book.values("D1:D5")] == [10, 10, "1,2", 101, "yes, ok"]
    info = await book.read("calc.get_range", range="D1:D2", include_formulas=True)
    assert info["formulas"]["D2"] == "=SUM(A1:A2;B1:B2)"  # LibreOffice syntax on the way out


async def test_formula_errors_are_reported(book: Book):
    result = await book.do("calc.set_range_values", data=[["=NOSUCHFUNCTION(1)", "=1/0", "=1+1"]])
    errors = {e["cell"]: e["error"] for e in result["formula_errors"]}
    assert set(errors) == {"A1", "B1"} and errors["B1"] == "#DIV/0!"
    assert "warning" in result


async def test_set_cell_returns_displayed_value(book: Book):
    await book.do("calc.set_range_values", data=[[2, 3]])
    result = await book.do("calc.set_cell", cell="C1", value="=A1*B1")
    assert result["value"] == "6" and result["cell"] == "C1"
    cleared = await book.do("calc.set_cell", cell="C1", value=None)
    assert cleared["value"] == ""


async def test_get_range_default_is_used_range_and_pages(book: Book):
    await book.do("calc.set_range_values", data=[[i, i * 2, i * 3] for i in range(10)])
    info = await book.read("calc.get_range", max_cells=12)
    assert info["range"] == "A1:C4" and info["truncated"] and info["next_range"] == "A5:C10"
    assert len(info["values"]) == 4
    rest = await book.read("calc.get_range", range=info["next_range"])
    assert rest["rows"] == 6 and "truncated" not in rest and rest["values"][0] == [4, 8, 12]
    whole = await book.read("calc.get_range", range="B:B")  # whole column is clipped to the data
    assert whole["range"] == "B1:B10"


async def test_get_range_formatted_shows_display_text(book: Book):
    await book.do("calc.set_range_values", data=[[0.256, 1234.5]])
    await book.do("calc.format_range", range="A1", number_format="0.0%")
    await book.do("calc.format_range", range="B1", number_format="#,##0.00")
    info = await book.read("calc.get_range", range="A1:B1", formatted=True)
    assert info["formatted"] == [["25.6%", "1,234.50"]] and info["values"] == [[0.256, 1234.5]]


async def test_write_table_builds_a_styled_table_in_one_call(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "table.xlsx")
    await book.create()
    result = await book.do(
        "calc.write_table",
        headers=["Region", "Units", "Price", "Revenue"],
        rows=[[r, u, p, f"=C{i + 3}*D{i + 3}"] for i, (r, u, p) in enumerate(SALES[1:])],
        start_cell="B2",
        column_formats={"Price": "$#,##0.00", "E": "#,##0.00"},
        autofilter=True,
        freeze_header=True,
    )
    assert result["range"] == "B2:E6" and result["header_row"] == 2 and result["data_rows"] == 4
    assert "formula_errors" not in result
    assert (await book.values("E3:E6")) == [[25], [12], [8.75], [10]]
    header = await book.read("calc.get_cell_info", cell="B2")
    assert header["font"]["bold"] and header["background"] == "#D9E2F3"
    price = await book.read("calc.get_cell_info", cell="D3")
    assert price["displayed"] == "$2.50"
    layout = await book.read("calc.get_sheet_layout")
    assert layout["auto_filter"] == "B2:E6" and layout["frozen_rows"] == 2
    assert layout["column_widths_mm"]["B"] < 40  # auto-fitted, not left at 'wide'


async def test_write_table_validates_input(book: Book):
    with pytest.raises(OfficeError, match="only 2 headers"):
        await book.do("calc.write_table", headers=["a", "b"], rows=[[1, 2, 3]])
    with pytest.raises(OfficeError, match="neither a header"):
        await book.do("calc.write_table", headers=["a"], rows=[[1]], column_formats={"zzz": "0"})


async def test_set_range_values_rejects_bad_data(book: Book):
    with pytest.raises(OfficeError, match="list of rows"):
        await book.do("calc.set_range_values", data=[1, 2, 3])
    with pytest.raises(OfficeError, match=r"data\[0\]\[1\] is a dict"):
        await book.do("calc.set_range_values", data=[[1, {"a": 1}]])


# -- addressing errors ---------------------------------------------------------


@pytest.mark.parametrize("bad", ["", "A", "1A", "A1:B", "Sheet1!A1", "A1:B2:C3", "ZZZZ1"])
async def test_bad_a1_notation_is_explained(book: Book, bad: str):
    with pytest.raises(OfficeError, match=r"Invalid cell range|outside the sheet"):
        await book.read("calc.get_range", range=bad)


async def test_unknown_sheet_lists_the_available_ones(book: Book):
    await book.do("calc.add_sheet", name="Data")
    with pytest.raises(OfficeError, match=r"No sheet named 'Nope'.*'Sheet1'.*'Data'"):
        await book.read("calc.get_range", sheet="Nope", range="A1")
    # case-insensitive fallback
    assert (await book.read("calc.get_range", sheet="data", range="A1"))["sheet"] == "Data"


async def test_out_of_range_cells(book: Book):
    with pytest.raises(OfficeError, match="outside the sheet"):
        await book.do("calc.set_cell", cell="A2000000", value=1)
    with pytest.raises(OfficeError, match="outside the sheet"):
        await book.do("calc.set_cell", cell="XFE1", value=1)


async def test_calc_ops_refuse_a_writer_document(office: OfficeClient, tmp_path: Path):
    doc = tmp_path / "t.odt"
    await office.call(protocol.KIND_CREATE, doc, [], options={"overwrite": True})
    with pytest.raises(OfficeError, match="needs a calc document"):
        await office.call(protocol.KIND_READ, doc, [{"op": "calc.list_sheets", "args": {}}])


# -- clear / fill / copy -------------------------------------------------------


async def test_clear_range_modes_and_noop(book: Book):
    await book.do("calc.set_range_values", data=SALES)
    await book.do("calc.format_range", range="A1:C1", bold=True, background="#FFFF00")
    await book.do("calc.clear_range", range="A1:C1", what="formats")
    info = await book.read("calc.get_cell_info", cell="A1")
    assert info["value"] == "Region" and not info["font"]["bold"] and info["background"] is None
    await book.do("calc.clear_range", range="B2:C5")
    assert (await book.values("A2:C2"))[0] == ["North", "", ""]
    await book.do("calc.clear_range", range="A1:C5", what="all")
    assert (await book.read("calc.list_sheets"))["sheets"][0]["used_range"] is None
    result = (await book.edit(("calc.clear_range", {"range": "A1:C5"})))[0]
    assert result["cleared"] == "contents"
    # noop: nothing changed => the file is not rewritten
    reply = await book.office.call(
        protocol.KIND_EDIT, book.path, [{"op": "calc.clear_range", "args": {"range": "A1:C5"}}]
    )
    assert reply["saved"] is False


async def test_fill_range_copy_autofill_and_series(book: Book):
    await book.do(
        "calc.set_range_values",
        data=[[5, "=A1*2", "Mon"], [None, None, None], [None, None, None], [None, None, None]],
    )
    await book.do("calc.fill_range", range="A1:B4", mode="copy")
    assert await book.values("A1:B4") == [[5, 10], [5, 10], [5, 10], [5, 10]]
    forms = await book.read("calc.get_range", range="B1:B4", include_formulas=True)
    assert forms["formulas"]["B4"] == "=A4*2"  # relative references adjusted
    await book.do("calc.fill_range", range="C1:C4", mode="autofill")
    assert [r[0] for r in await book.values("C1:C4")] == ["Mon", "Tue", "Wed", "Thu"]
    await book.do("calc.set_range_values", data=[[100]], start_cell="E1")
    await book.do("calc.fill_range", range="E1:E4", mode="linear", step=50)
    assert [r[0] for r in await book.values("E1:E4")] == [100, 150, 200, 250]
    await book.do("calc.fill_range", range="E1:H1", direction="right", mode="copy")
    assert (await book.values("E1:H1"))[0] == [100, 100, 100, 100]


async def test_fill_range_up_copy_and_errors(book: Book):
    await book.do("calc.set_range_values", data=[[None], [None], [None], ["last"]])
    await book.do("calc.fill_range", range="A1:A4", direction="up", mode="copy")
    assert [r[0] for r in await book.values("A1:A4")] == ["last"] * 4
    with pytest.raises(OfficeError, match="nothing to fill"):
        await book.do("calc.fill_range", range="A1")
    with pytest.raises(OfficeError, match="mode must be"):
        await book.do("calc.fill_range", range="A1:A3", mode="weird")


async def test_copy_and_move_range_across_sheets(book: Book):
    await book.do("calc.set_range_values", data=[[1, "=A1+1"], [3, "=A2+1"]])
    await book.do("calc.add_sheet", name="Target")
    result = await book.do(
        "calc.copy_range", source_range="A1:B2", target_cell="C3", target_sheet="Target"
    )
    assert result["to"] == "Target!C3:D4"
    assert await book.values("C3:D4", sheet="Target") == [[1, 2], [3, 4]]
    assert await book.values("A1:B2") == [[1, 2], [3, 4]]  # source untouched
    await book.do("calc.copy_range", source_range="A1:B2", target_cell="F1", move=True)
    assert await book.values("A1:B2") == [["", ""], ["", ""]]
    assert (await book.values("F1:F2")) == [[1], [3]]


# -- find / replace / sort -----------------------------------------------------


async def test_find_text_across_sheets(book: Book):
    await book.do("calc.set_range_values", data=SALES)
    await book.do("calc.add_sheet", name="Notes")
    await book.do("calc.set_range_values", sheet="Notes", data=[["North star", "=1+1"]])
    found = await book.read("calc.find_text", query="north")
    assert found["total_matches"] == 2
    assert {(m["sheet"], m["cell"]) for m in found["matches"]} == {
        ("Sheet1", "A2"),
        ("Notes", "A1"),
    }
    exact = await book.read("calc.find_text", query="north", whole_cell=True, sheet="Sheet1")
    assert exact["total_matches"] == 1
    num = await book.read("calc.find_text", query=r"^\d+$", regex=True, sheet="Sheet1")
    assert {m["cell"] for m in num["matches"]} == {"B2", "B3", "B4", "B5", "C3", "C5"}
    forms = await book.read("calc.find_text", query="=1+1", search_in="formulas", sheet="Notes")
    assert forms["matches"][0]["formula"] == "=1+1"
    with pytest.raises(OfficeError, match="Invalid regular expression"):
        await book.read("calc.find_text", query="(", regex=True)


async def test_replace_text_counts_cells_and_is_noop_when_nothing_matches(book: Book):
    await book.do("calc.set_range_values", data=SALES)
    result = await book.do("calc.replace_text", search="th", replace="TH")
    assert result["replaced_cells"] == 2  # North, South
    assert [r[0] for r in await book.values("A2:A3")] == ["NorTH", "SouTH"]
    reply = await book.office.call(
        protocol.KIND_EDIT, book.path,
        [{"op": "calc.replace_text", "args": {"search": "zzz", "replace": "y"}}],
    )  # fmt: skip
    assert reply["results"][0]["replaced_cells"] == 0 and reply["saved"] is False
    scoped = await book.do(
        "calc.replace_text", search="s", replace="S", sheet="Sheet1", range="A4:A5", match_case=True
    )
    assert scoped["replaced_cells"] == 2
    assert [r[0] for r in await book.values("A2:A5")] == ["NorTH", "SouTH", "EaSt", "WeSt"]
    with pytest.raises(OfficeError, match=r"Give .sheet."):
        await book.do("calc.replace_text", search="a", replace="b", range="A1:A3")


async def test_sort_range_by_header_name_and_descending(book: Book):
    await book.do("calc.set_range_values", data=SALES)
    await book.do("calc.sort_range", range="A1:C5", sort_by=[{"column": "Units"}])
    assert [r[0] for r in await book.values("A2:A5")] == ["West", "South", "East", "North"]
    await book.do("calc.sort_range", range="A1:C5", sort_by=[{"column": "C", "ascending": False}])
    assert [r[0] for r in await book.values("A2:A5")] == ["West", "South", "North", "East"]
    assert (await book.values("A1:C1"))[0] == ["Region", "Units", "Price"]  # header kept


async def test_sort_range_without_header_and_errors(book: Book):
    await book.do("calc.set_range_values", data=[[3], [1], [2]])
    await book.do("calc.sort_range", range="A1:A3", sort_by=[{"column": "A"}], has_header=False)
    assert [r[0] for r in await book.values("A1:A3")] == [1, 2, 3]
    with pytest.raises(OfficeError, match="not a column letter or a header"):
        await book.do("calc.sort_range", range="A1:A3", sort_by=[{"column": "nope"}])
    with pytest.raises(OfficeError, match="outside the range"):
        await book.do("calc.sort_range", range="A1:A3", sort_by=[{"column": "D"}], has_header=False)


# -- rows, columns, merge ------------------------------------------------------


async def test_insert_and_delete_rows_and_columns(book: Book):
    await book.do("calc.set_range_values", data=[["a", "b"], ["c", "d"], ["=A1&B1", "=A2&B2"]])
    await book.do("calc.insert_rows", before_row=2, count=2)
    assert [r[0] for r in await book.values("A1:A5")] == ["a", "", "", "c", "ab"]
    await book.do("calc.delete_rows", start_row=2, count=2)
    assert [r[0] for r in await book.values("A1:A3")] == ["a", "c", "ab"]
    await book.do("calc.insert_columns", before_column="B")
    assert (await book.values("A1:C1"))[0] == ["a", "", "b"]
    await book.do("calc.delete_columns", start_column="B")
    assert (await book.values("A1:B1"))[0] == ["a", "b"]
    assert (await book.values("A3:B3"))[0] == ["ab", "cd"]  # formulas still right
    with pytest.raises(OfficeError, match="outside the sheet"):
        await book.do("calc.delete_rows", start_row=0)
    with pytest.raises(OfficeError, match="Invalid column"):
        await book.do("calc.insert_columns", before_column="1A")


async def test_merge_and_unmerge(book: Book):
    await book.do("calc.set_range_values", data=[["Title", "extra", None]])
    result = await book.do("calc.merge_cells", range="A1:C1")
    assert "discarded" in result["warning"]
    info = await book.read("calc.get_cell_info", cell="B1")
    assert info["merged_range"] == "A1:C1"
    layout = await book.read("calc.get_sheet_layout")
    assert layout["merged_ranges"] == ["A1:C1"]
    again = await book.edit(("calc.merge_cells", {"range": "A1:C1"}))
    assert again[0]["merged"] is True
    await book.do("calc.unmerge_cells", range="A1:C1")
    assert "merged_range" not in await book.read("calc.get_cell_info", cell="B1")
    with pytest.raises(OfficeError, match="at least two cells"):
        await book.do("calc.merge_cells", range="A1")


# -- sheets --------------------------------------------------------------------


async def test_sheet_management(book: Book):
    await book.do("calc.add_sheet", name="Second")
    await book.do("calc.add_sheet", name="First", position=0)
    names = [s["name"] for s in (await book.read("calc.list_sheets"))["sheets"]]
    assert names == ["First", "Sheet1", "Second"]
    await book.do("calc.move_sheet", sheet="First", position=2)
    names = [s["name"] for s in (await book.read("calc.list_sheets"))["sheets"]]
    assert names == ["Sheet1", "Second", "First"]
    await book.do("calc.move_sheet", sheet="First", position=0)
    names = [s["name"] for s in (await book.read("calc.list_sheets"))["sheets"]]
    assert names == ["First", "Sheet1", "Second"]
    await book.do("calc.copy_sheet", sheet="Second", new_name="Copy", position=1)
    await book.do("calc.rename_sheet", sheet="Copy", new_name="Renamed")
    names = [s["name"] for s in (await book.read("calc.list_sheets"))["sheets"]]
    assert names == ["First", "Renamed", "Sheet1", "Second"]
    await book.do("calc.delete_sheet", sheet="Renamed")
    assert len((await book.read("calc.list_sheets"))["sheets"]) == 3


async def test_sheet_errors(book: Book):
    with pytest.raises(OfficeError, match="already exists"):
        await book.do("calc.add_sheet", name="sheet1")
    with pytest.raises(OfficeError, match="Invalid sheet name"):
        await book.do("calc.add_sheet", name="a/b")
    with pytest.raises(OfficeError, match="at least one sheet"):
        await book.do("calc.delete_sheet", sheet="Sheet1")
    with pytest.raises(OfficeError, match="out of range"):
        await book.do("calc.move_sheet", sheet="Sheet1", position=5)


async def test_rename_sheet_updates_formulas(book: Book):
    await book.do("calc.add_sheet", name="Data")
    await book.do("calc.set_range_values", sheet="Data", data=[[41]])
    await book.do("calc.set_range_values", data=[["=Data!A1+1"]])
    await book.do("calc.rename_sheet", sheet="Data", new_name="Numbers")
    info = await book.read("calc.get_range", range="A1", include_formulas=True)
    assert info["values"] == [[42]] and "Numbers" in info["formulas"]["A1"]


async def test_hide_sheet_and_tab_color(book: Book):
    await book.do("calc.add_sheet", name="B")
    await book.do("calc.set_sheet_properties", sheet="B", visible=False, tab_color="#FF0000")
    sheets = (await book.read("calc.list_sheets"))["sheets"]
    assert sheets[1]["visible"] is False and sheets[1]["tab_color"] == "#FF0000"
    with pytest.raises(OfficeError, match="last visible sheet"):
        await book.do("calc.set_sheet_properties", sheet="Sheet1", visible=False)
    await book.do("calc.set_sheet_properties", sheet="B", tab_color="none", visible=True)
    sheets = (await book.read("calc.list_sheets"))["sheets"]
    assert sheets[1]["visible"] is True and "tab_color" not in sheets[1]
    reply = await book.office.call(
        protocol.KIND_EDIT, book.path,
        [{"op": "calc.set_sheet_properties", "args": {"sheet": "B", "visible": True}}],
    )  # fmt: skip
    assert reply["saved"] is False


async def test_named_ranges(book: Book):
    await book.do("calc.set_range_values", data=[[1], [2], [3]])
    await book.do("calc.add_named_range", name="Nums", range="A1:A3")
    await book.do("calc.set_range_values", data=[["=SUM(Nums)"]], start_cell="C1")
    assert (await book.values("C1"))[0][0] == 6
    listed = await book.read("calc.list_named_ranges")
    assert listed["named_ranges"] == [{"name": "Nums", "refers_to": "Sheet1.A1:A3"}]
    with pytest.raises(OfficeError, match="already exists"):
        await book.do("calc.add_named_range", name="Nums", range="A1")
    with pytest.raises(OfficeError, match="Invalid name"):
        await book.do("calc.add_named_range", name="A1", range="A1")
    await book.do("calc.delete_named_range", name="Nums")
    assert (await book.read("calc.list_named_ranges"))["count"] == 0
    with pytest.raises(OfficeError, match="No named range"):
        await book.do("calc.delete_named_range", name="Nums")


async def test_freeze_panes_roundtrip_through_xlsx(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "freeze.xlsx")
    await book.create()
    await book.do("calc.add_sheet", name="Two")
    await book.do("calc.freeze_panes", sheet="Two", rows=2, columns=1)
    layout = await book.read("calc.get_sheet_layout", sheet="Two")
    assert (layout["frozen_rows"], layout["frozen_columns"]) == (2, 1)
    assert (await book.read("calc.get_sheet_layout", sheet="Sheet1"))["frozen_rows"] == 0
    assert (await book.read("calc.get_sheet_layout", sheet="Two"))[
        "frozen_columns"
    ] == 1  # reloaded
    await book.do("calc.freeze_panes", sheet="Two", rows=0, columns=0)
    assert (await book.read("calc.get_sheet_layout", sheet="Two"))["frozen_rows"] == 0
    noop = await book.office.call(
        protocol.KIND_EDIT, book.path,
        [{"op": "calc.freeze_panes", "args": {"sheet": "Two"}}],
    )  # fmt: skip
    assert noop["saved"] is False


async def test_autofilter_on_and_off(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "filter.xlsx")
    await book.create(("calc.set_range_values", {"data": SALES}))
    await book.do("calc.set_autofilter", range="A1:C5")
    assert (await book.read("calc.get_sheet_layout"))["auto_filter"] == "A1:C5"  # survives xlsx
    await book.do("calc.set_autofilter", range="A1:C4")  # moves the existing filter
    assert (await book.read("calc.get_sheet_layout"))["auto_filter"] == "A1:C4"
    await book.do("calc.set_autofilter", enabled=False)
    assert (await book.read("calc.get_sheet_layout"))["auto_filter"] is None
    with pytest.raises(OfficeError, match="header row plus"):
        await book.do("calc.set_autofilter", range="A1:C1")


async def test_recalculate(book: Book):
    await book.do("calc.set_range_values", data=[[2, "=A1*A1"]])
    assert (await book.do("calc.recalculate"))["recalculated"] is True
    assert (await book.values("B1"))[0][0] == 4


# -- formatting ----------------------------------------------------------------


async def test_format_range_applies_and_get_cell_info_reports(book: Book):
    await book.do("calc.set_range_values", data=[[1234.5, "text"]])
    result = await book.do(
        "calc.format_range", range="A1:B1", font_name="Arial", font_size=14, bold=True,
        italic=True, underline=True, font_color="#FF0000", background="#FFFF00",
        number_format="#,##0.00", h_align="center", v_align="middle", wrap_text=True,
    )  # fmt: skip
    assert "NumberFormat" in result["applied"]
    info = await book.read("calc.get_cell_info", cell="A1")
    assert info["font"] == {
        "name": "Arial", "size": 14.0, "bold": True, "italic": True, "underline": True,
        "color": "#FF0000",
    }  # fmt: skip
    assert info["background"] == "#FFFF00" and info["displayed"] == "1,234.50"
    assert info["h_align"] == "center" and info["v_align"] == "middle" and info["wrap_text"]
    assert info["number_format"] == "#,##0.00"
    await book.do("calc.format_range", range="A1", background="none", bold=False)
    info = await book.read("calc.get_cell_info", cell="A1")
    assert info["background"] is None and not info["font"]["bold"]


async def test_format_range_whole_column_and_cell_style(book: Book):
    await book.do("calc.set_range_values", data=[[1], [2]])
    await book.do("calc.format_range", range="A:A", number_format="0.00")
    assert (await book.read("calc.get_cell_info", cell="A2"))["displayed"] == "2.00"
    await book.do("calc.format_range", range="A1", cell_style="Heading 1")
    assert (await book.read("calc.get_cell_info", cell="A1"))["cell_style"] == "Heading 1"


async def test_format_range_errors(book: Book):
    with pytest.raises(OfficeError, match="Nothing to do"):
        await book.do("calc.format_range", range="A1")
    with pytest.raises(OfficeError, match="Invalid number format"):
        await book.do("calc.format_range", range="A1", number_format="0.0[[x")
    with pytest.raises(OfficeError, match="Invalid colour"):
        await book.do("calc.format_range", range="A1", font_color="pinkish")
    with pytest.raises(OfficeError, match="No cell style"):
        await book.do("calc.format_range", range="A1", cell_style="Nope")


async def test_dates_and_percent_number_formats(book: Book):
    await book.do("calc.set_range_values", data=[["2024-03-05", 0.5]], parse_text=True)
    await book.do("calc.format_range", range="A1", number_format="dd/mm/yyyy")
    await book.do("calc.format_range", range="B1", number_format="0%")
    info = await book.read("calc.get_range", range="A1:B1", formatted=True)
    assert info["formatted"] == [["05/03/2024", "50%"]]


async def test_set_borders_sides(book: Book):
    await book.do("calc.set_range_values", data=SALES)
    await book.do("calc.set_borders", range="A1:C5", sides=["outer"], width_pt=2, color="#0000FF")
    await book.do("calc.set_borders", range="A1:C1", sides=["bottom"], style="double")
    # verify through the document model: not exposed by tools, so use the xlsx round trip
    out = await book.convert(book.path.with_suffix(".xlsx"))
    assert out.stat().st_size > 0
    with pytest.raises(OfficeError, match="Unknown border side"):
        await book.do("calc.set_borders", range="A1", sides=["diagonal"])
    with pytest.raises(OfficeError, match="width_pt"):
        await book.do("calc.set_borders", range="A1", width_pt=50)
    await book.do("calc.set_borders", range="A1:C5", sides=["none"])


async def test_column_width_and_row_height(book: Book):
    await book.do("calc.set_range_values", data=[["a very long piece of text in this cell", 1]])
    result = await book.do("calc.set_column_width", columns="A", width_mm=60)
    assert result["width_mm"] == 60.0
    await book.do("calc.set_column_width", columns="B:C", width_mm=10)
    await book.do("calc.set_column_width", columns="A", auto_fit=True)
    layout = await book.read("calc.get_sheet_layout")
    assert layout["column_widths_mm"]["A"] > 30
    await book.do("calc.set_column_width", columns="B", hidden=True)
    assert (await book.read("calc.get_sheet_layout"))["hidden_columns"] == ["B"]
    r = await book.do("calc.set_row_height", rows="1:2", height_mm=12)
    assert r["height_mm"] == 12.0
    await book.do("calc.set_row_height", rows="3", hidden=True)
    layout = await book.read("calc.get_sheet_layout", sheet="Sheet1")
    assert layout["custom_row_heights_mm"] == {"1": 12.0}
    with pytest.raises(OfficeError, match="Give width_mm"):
        await book.do("calc.set_column_width", columns="A")
    with pytest.raises(OfficeError, match="Invalid column"):
        await book.do("calc.set_column_width", columns="1", width_mm=5)
    with pytest.raises(OfficeError, match="Invalid row"):
        await book.do("calc.set_row_height", rows="x", height_mm=5)


async def test_conditional_formats_add_list_clear(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "cf.xlsx")
    await book.create(("calc.set_range_values", {"data": [[-5], [3], [12], [7]]}))
    await book.do("calc.add_conditional_format", range="A1:A4", rule="less", value=0,
                  background="#FFC7CE", font_color="#9C0006")  # fmt: skip
    await book.do("calc.add_conditional_format", range="A1:A4", rule="between", value=5,
                  value2=10, background="#C6EFCE")  # fmt: skip
    await book.do("calc.add_conditional_format", range="A1:A4", rule="formula", value="=A1>10",
                  bold=True)  # fmt: skip
    await book.do("calc.add_data_bar", range="A1:A4", color="#638EC6")
    listed = await book.read("calc.list_conditional_formats")
    # the .xlsx export folds rules on one range into a single format with several rules
    assert listed["count"] == 1 and len(listed["conditional_formats"][0]["rules"]) == 4
    rules = [r for f in listed["conditional_formats"] for r in f["rules"]]
    kinds = sorted(r["type"] for r in rules)
    assert kinds == ["between", "data_bar", "formula", "less"]
    between = next(r for r in rules if r["type"] == "between")
    assert (between["value"], between["value2"]) == ("5", "10")
    await book.do("calc.clear_conditional_formats", range="A1")
    assert (await book.read("calc.list_conditional_formats"))["count"] == 0
    with pytest.raises(OfficeError, match="needs a value"):
        await book.do("calc.add_conditional_format", range="A1", rule="greater", bold=True)
    with pytest.raises(OfficeError, match="at least one of background"):
        await book.do("calc.add_conditional_format", range="A1", rule="greater", value=1)
    with pytest.raises(OfficeError, match="both value"):
        await book.do("calc.add_conditional_format", range="A1", rule="between", value=1, bold=True)
    noop = await book.office.call(
        protocol.KIND_EDIT, book.path, [{"op": "calc.clear_conditional_formats", "args": {}}]
    )
    assert noop["saved"] is False


async def test_data_validation_list_number_and_removal(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "val.xlsx")
    await book.create()
    await book.do("calc.set_data_validation", range="A1:A5", type="list", items=["Open", "Closed"],
                  error_message="Pick one")  # fmt: skip
    info = await book.read("calc.get_cell_info", cell="A3")
    assert info["validation"]["type"] == "list" and "Open" in info["validation"]["formula1"]
    await book.do("calc.set_data_validation", range="B1:B5", type="whole_number", min_value=1,
                  max_value=10, input_message="1 to 10")  # fmt: skip
    info = await book.read("calc.get_cell_info", cell="B2")
    assert info["validation"]["type"] == "whole"
    assert (info["validation"]["formula1"], info["validation"]["formula2"]) == ("1", "10")
    await book.do("calc.set_data_validation", range="C1", type="date", operator="greater_equal",
                  min_value="2024-01-01")  # fmt: skip
    assert (
        "DATE(2024" in (await book.read("calc.get_cell_info", cell="C1"))["validation"]["formula1"]
    )
    await book.do("calc.set_range_values", data=[["a"], ["b"]], start_cell="E1")
    await book.do("calc.set_data_validation", range="D1", type="list", source_range="E1:E2")
    assert (
        "$E$1:$E$2" in (await book.read("calc.get_cell_info", cell="D1"))["validation"]["formula1"]
    )
    await book.do("calc.set_data_validation", range="A1:A5", type="any")
    assert "validation" not in await book.read("calc.get_cell_info", cell="A3")
    with pytest.raises(OfficeError, match="needs 'items'"):
        await book.do("calc.set_data_validation", range="A1", type="list")
    with pytest.raises(OfficeError, match="min_value"):
        await book.do("calc.set_data_validation", range="A1", type="decimal")
    with pytest.raises(OfficeError, match="needs 'formula'"):
        await book.do("calc.set_data_validation", range="A1", type="custom")


# -- objects -------------------------------------------------------------------


async def _chart_data(book: Book) -> None:
    await book.do("calc.write_table", headers=["Region", "Units", "Price"], rows=SALES[1:])


@pytest.mark.parametrize("kind", ["column", "bar", "line", "pie", "scatter", "area", "radar"])
async def test_add_chart_all_types(book: Book, kind: str):
    await _chart_data(book)
    result = await book.do(
        "calc.add_chart", data_range="A1:B5", chart_type=kind, title=f"{kind} chart",
        x_axis_title="Region", y_axis_title="Units", anchor_cell="E2",
    )  # fmt: skip
    assert result["chart"] == "Chart1" and result["anchor_cell"] == "E2"
    charts = (await book.read("calc.list_charts"))["charts"]
    assert len(charts) == 1
    assert charts[0]["type"] == kind and charts[0]["title"] == f"{kind} chart"
    assert charts[0]["data_ranges"] == ["A1:B5"]


async def test_chart_lifecycle_update_and_delete(book: Book):
    await _chart_data(book)
    await book.do("calc.add_chart", data_range="A1:C5", name="Sales", width_mm=100, height_mm=60)
    await book.do("calc.add_chart", data_range="A1:B5", chart_type="pie")
    names = [c["chart"] for c in (await book.read("calc.list_charts"))["charts"]]
    assert names == ["Sales", "Chart1"]
    await book.do(
        "calc.update_chart", chart="Sales", chart_type="line", title="Revenue", legend=False,
        series_colors=["#FF0000", "#00AA00"],
    )  # fmt: skip
    sales = next(
        c for c in (await book.read("calc.list_charts"))["charts"] if c["chart"] == "Sales"
    )
    assert sales["type"] == "line" and sales["title"] == "Revenue"
    assert sales["size_mm"] == [100.0, 60.0]
    await book.do("calc.update_chart", chart="Sales", title="")
    await book.do("calc.delete_chart", chart="Chart1")
    assert [c["chart"] for c in (await book.read("calc.list_charts"))["charts"]] == ["Sales"]
    with pytest.raises(OfficeError, match=r"No chart named .Chart1.*.Sales."):
        await book.do("calc.delete_chart", chart="Chart1")
    with pytest.raises(OfficeError, match="already exists"):
        await book.do("calc.add_chart", data_range="A1:B5", name="Sales")
    with pytest.raises(OfficeError, match="at least two cells"):
        await book.do("calc.add_chart", data_range="A1")


async def test_charts_survive_xlsx_round_trip(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "chart.xlsx")
    await book.create(("calc.write_table", {"headers": ["k", "v"], "rows": [["a", 1], ["b", 2]]}))
    await book.do("calc.add_chart", data_range="A1:B3", chart_type="bar", title="T")
    charts = (await book.read("calc.list_charts"))["charts"]
    assert len(charts) == 1 and charts[0]["type"] == "bar" and charts[0]["title"] == "T"
    # the name given at creation survives the save and reload (xlsx renames the internal object)
    assert charts[0]["chart"] == "Chart1"
    await book.do("calc.update_chart", chart="Chart1", title="T2")
    assert (await book.read("calc.list_charts"))["charts"][0]["title"] == "T2"


async def test_images(book: Book, tmp_path: Path):
    png = make_png(tmp_path / "logo.png")
    result = await book.do("calc.insert_image", image_path=str(png), anchor_cell="C3", width_mm=40)
    assert result["image"] == "Image1" and result["height_mm"] == 20.0
    listed = await book.read("calc.list_images")
    assert listed["images"][0]["anchor_cell"] == "C3" and listed["images"][0]["size_mm"] == [
        40.0,
        20.0,
    ]
    assert (await book.read("calc.get_sheet_layout"))["images"] == 1
    await book.do("calc.delete_image", image="Image1")
    assert (await book.read("calc.list_images"))["count"] == 0
    with pytest.raises(OfficeError, match="No image named"):
        await book.do("calc.delete_image", image="Image1")
    bad = tmp_path / "not.png"
    bad.write_bytes(b"this is not an image")
    with pytest.raises(OfficeError, match="not an image"):
        await book.do("calc.insert_image", image_path=str(bad))


async def test_hyperlink_and_comments(book: Book):
    await book.do("calc.set_range_values", data=[["docs", "x"]])
    result = await book.do("calc.set_hyperlink", cell="A1", url="https://example.com/docs")
    assert result["text"] == "docs"
    info = await book.read("calc.get_cell_info", cell="A1")
    assert info["hyperlink"] == "https://example.com/docs" and info["displayed"] == "docs"
    await book.do("calc.set_hyperlink", cell="B2", url="https://x.org", text="X site")
    assert (await book.read("calc.get_cell_info", cell="B2"))["displayed"] == "X site"
    await book.do("calc.add_comment", cell="B1", text="check this")
    await book.do("calc.add_sheet", name="S2")
    await book.do("calc.add_comment", sheet="S2", cell="A1", text="other", visible=True)
    comments = (await book.read("calc.list_comments"))["comments"]
    assert {(c["sheet"], c["cell"], c["text"]) for c in comments} == {
        ("Sheet1", "B1", "check this"), ("S2", "A1", "other"),
    }  # fmt: skip
    assert (await book.read("calc.get_cell_info", cell="B1"))["comment"] == "check this"
    again = await book.do("calc.add_comment", cell="B1", text="updated")
    assert again["replaced"] is True
    assert (await book.read("calc.list_comments", sheet="Sheet1"))["comments"][0][
        "text"
    ] == "updated"
    await book.do("calc.delete_comment", cell="B1")
    assert (await book.read("calc.list_comments", sheet="Sheet1"))["count"] == 0
    noop = await book.office.call(
        protocol.KIND_EDIT, book.path, [{"op": "calc.delete_comment", "args": {"cell": "B1"}}]
    )
    assert noop["saved"] is False
    with pytest.raises(OfficeError, match="url must not be empty"):
        await book.do("calc.set_hyperlink", cell="A1", url=" ")
    with pytest.raises(OfficeError, match="text must not be empty"):
        await book.do("calc.add_comment", cell="A1", text=" ")


# -- file formats --------------------------------------------------------------


async def test_csv_files_can_be_created_edited_and_read(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "data.csv")
    await book.create(
        ("calc.write_table", {"headers": ["name", "qty"], "rows": [["a,b", 1], ["c", 2.5]]})
    )
    text = book.path.read_text(encoding="utf-8")
    assert text.splitlines()[0] == "name,qty" and '"a,b",1' in text
    await book.do("calc.set_range_values", data=[["d", 3]], start_cell="A4")
    assert (await book.values()) == [["name", "qty"], ["a,b", 1], ["c", 2.5], ["d", 3]]
    await book.do("calc.replace_text", search="c", replace="C", match_case=True, whole_cell=True)
    assert "\nC,2.5" in book.path.read_text(encoding="utf-8")


async def test_freeze_panes_is_refused_for_opendocument_files(book: Book):
    # LibreOffice cannot write view settings to .ods without a window; say so instead of
    # pretending to succeed.
    with pytest.raises(OfficeError, match=r"cannot be saved in .ods"):
        await book.do("calc.freeze_panes", rows=1)
    result = await book.do("calc.write_table", headers=["a"], rows=[[1]], freeze_header=True)
    assert "freeze_header was not applied" in result["warning"] and "frozen_rows" not in result
    assert (await book.do("calc.freeze_panes"))["frozen_rows"] == 0  # unfreezing is harmless


@pytest.mark.parametrize("ext", ["xlsx", "ods", "xls", "fods"])
async def test_formats_round_trip_data_and_formatting(
    office: OfficeClient, tmp_path: Path, ext: str
):
    book = Book(office, tmp_path / f"rt.{ext}")
    await book.create(
        ("calc.write_table", {
            "headers": ["Region", "Units", "Price", "Total"],
            "rows": [[r, u, p, f"=B{i + 2}*C{i + 2}"] for i, (r, u, p) in enumerate(SALES[1:])],
            "column_formats": {"Price": "0.00"},
        }),
        ("calc.add_sheet", {"name": "Extra"}),
    )  # fmt: skip
    await book.do("calc.add_comment", cell="A1", text="note")
    # reopen from disk
    values = await book.values("A1:D5")
    assert values[0] == ["Region", "Units", "Price", "Total"]
    assert values[1] == ["North", 10, 2.5, 25]
    info = await book.read("calc.get_cell_info", cell="B1")
    assert info["font"]["bold"] and info["background"] == "#D9E2F3"
    assert (await book.read("calc.get_cell_info", cell="D2"))["formula"] == "=B2*C2"
    assert [s["name"] for s in (await book.read("calc.list_sheets"))["sheets"]] == [
        "Sheet1",
        "Extra",
    ]
    assert (await book.read("calc.list_comments"))["count"] == 1


# -- more behaviour ------------------------------------------------------------


async def test_error_cells_can_be_read_back(book: Book):
    await book.do("calc.set_range_values", data=[["=1/0", "ok"]])
    info = await book.read("calc.get_cell_info", cell="A1")
    assert info["type"] == "formula" and info["displayed"] == "#DIV/0!" and info["error_code"]
    values = await book.values("A1:B1")
    assert values[0][1] == "ok"  # the error cell itself must not break the read


async def test_chart_defaults_legend_only_when_useful_and_scatter_is_points(book: Book):
    await book.do(
        "calc.write_table", headers=["x", "a", "b"], rows=[[1, 2, 3], [2, 4, 6], [3, 5, 9]]
    )
    await book.do("calc.add_chart", data_range="A1:B4", chart_type="column", name="single")
    await book.do("calc.add_chart", data_range="A1:C4", chart_type="scatter", name="multi")
    await book.do("calc.add_chart", data_range="A1:B4", chart_type="pie", name="pie")
    await book.do("calc.add_chart", data_range="A1:B4", name="forced", legend=True)
    legends = {c["chart"]: c["legend"] for c in (await book.read("calc.list_charts"))["charts"]}
    assert legends == {"single": False, "multi": True, "pie": True, "forced": True}
    pdf = await book.convert(book.path.with_suffix(".pdf"))
    assert pdf.read_bytes().startswith(b"%PDF")


async def test_get_range_reports_empty_sheet_and_sheet_with_only_an_image(
    book: Book, tmp_path: Path
):
    empty = await book.read("calc.get_range")
    assert empty["values"] == [] and empty["range"] is None
    await book.do("calc.insert_image", image_path=str(make_png(tmp_path / "p.png")), width_mm=20)
    layout = await book.read("calc.get_sheet_layout")
    assert layout["used_range"] is None and layout["images"] == 1


async def test_image_anchor_cell_is_recovered_after_reload(office: OfficeClient, tmp_path: Path):
    book = Book(office, tmp_path / "pic.xlsx")
    await book.create()
    await book.do("calc.set_column_width", columns="A:C", width_mm=50)
    await book.do(
        "calc.insert_image", image_path=str(make_png(tmp_path / "p.png")), anchor_cell="C4",
        height_mm=10,
    )  # fmt: skip
    listed = await book.read("calc.list_images")
    assert listed["images"][0]["anchor_cell"] == "C4" and listed["images"][0]["size_mm"][1] == 10.0


async def test_conditional_format_formula_and_text_rules(book: Book):
    await book.do("calc.set_range_values", data=[["apple", 1], ["pear", 5], ["apple", 9]])
    await book.do(
        "calc.add_conditional_format", range="A1:A3", rule="equal", value="apple", bold=True
    )
    await book.do(
        "calc.add_conditional_format", range="B1:B3", rule="formula", value="=$B1>$B2",
        background="#FFFF00",
    )  # fmt: skip
    await book.do("calc.add_conditional_format", range="A1:A3", rule="duplicate", font_color="red")
    listed = (await book.read("calc.list_conditional_formats"))["conditional_formats"]
    rules = [r for f in listed for r in f["rules"]]
    assert {"equal", "formula", "duplicate"} <= {r["type"] for r in rules}
    equal = next(r for r in rules if r["type"] == "equal")
    assert equal["value"] == '"apple"'


async def test_sheet_names_with_spaces_and_unicode_work_everywhere(book: Book):
    await book.do("calc.add_sheet", name="Q1 Résumé")
    await book.do("calc.set_range_values", sheet="Q1 Résumé", data=[[5]])
    await book.do("calc.set_range_values", data=[["='Q1 Résumé'!A1*2"]])
    assert (await book.values("A1"))[0][0] == 10
    await book.do("calc.add_named_range", name="Five", range="A1", sheet="Q1 Résumé")
    assert (await book.read("calc.list_named_ranges"))["named_ranges"][0]["name"] == "Five"
    found = await book.read("calc.find_text", query="5", sheet="Q1 Résumé")
    assert found["matches"][0]["sheet"] == "Q1 Résumé"


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=SUMIF(A1:A3,">1")', 5),
        ('=COUNTIFS(A1:A3,">0",B1:B3,"<5")', 3),
        ('=TEXT(A1,"0.00")', "1.00"),
        ('=IF(A1>0,"a;b, c","d")', "a;b, c"),
        ("=VLOOKUP(2,Data!A1:B3,2,0)", 20),
        ("=SUM(Data!A1:A3)", 6),
        ("=SUM(Data.A1:A3)", 6),  # LibreOffice reference syntax keeps working
        ("=SUM(A1:A3;B1:B3)", 15),  # and so does the ';' separator
        ("=ROUND(A2/3,1)", 0.7),
        ('=A1&"-"&B1', "1-2"),
        ("=INDEX(A1:B3,2,2)", 3),
    ],
)
async def test_formula_dialects(book: Book, formula: str, expected: Any):
    await book.do("calc.add_sheet", name="Data")
    await book.do("calc.set_range_values", sheet="Data", data=[[1, 10], [2, 20], [3, 30]])
    await book.do("calc.set_range_values", data=[[1, 2], [2, 3], [3, 4]])
    result = await book.do("calc.set_cell", cell="D1", value=formula)
    assert "formula_errors" not in result, result
    assert (await book.values("D1"))[0][0] == expected


async def test_csv_refuses_extra_sheets_and_warns_about_lost_formatting(
    office: OfficeClient, tmp_path: Path
):
    book = Book(office, tmp_path / "plain.csv")
    await book.create(("calc.set_range_values", {"data": [["a", 1], ["b", 2]]}))
    with pytest.raises(OfficeError, match=r"Adding a sheet is not possible in a \.csv"):
        await book.do("calc.add_sheet", name="Two")
    with pytest.raises(OfficeError, match="Copying a sheet"):
        await book.do("calc.copy_sheet", sheet="Sheet1", new_name="Two")
    result = await book.do("calc.format_range", range="A1", bold=True)
    assert "will not be saved" in result["warning"]
    chart = await book.do("calc.add_chart", data_range="A1:B2")
    assert "warning" in chart
    plain = await book.do("calc.set_cell", cell="C1", value="x")
    assert "warning" not in plain


async def test_several_merged_ranges_are_listed(book: Book):
    await book.do("calc.set_range_values", data=[["h1", None, None, "h2"], ["a", "b", "c", "d"]])
    await book.do("calc.format_range", range="A2:D2", italic=True)
    await book.do("calc.merge_cells", range="A1:C1")
    await book.do("calc.merge_cells", range="E1:F2")
    await book.do("calc.set_range_values", data=[["x"]], start_cell="H9")
    layout = await book.read("calc.get_sheet_layout")
    assert sorted(layout["merged_ranges"]) == ["A1:C1", "E1:F2"]


async def test_long_formula_columns_are_exact_even_when_rows_differ(book: Book):
    n = 1200  # above the threshold that switches to "fill down, then repair differences"
    rows = [[i, i * 2, f"=A{i + 1}+B{i + 1}"] for i in range(n)]
    rows[700][2] = '=A701*10&"!"'  # an irregular row inside the run
    rows[701][2] = "=SUM(A1:A3)"
    result = await book.do("calc.set_range_values", data=rows)
    assert result["formulas"] == n and "formula_errors" not in result
    values = [r[0] for r in await book.values("C1:C1200", max_cells=10_000)]
    assert values[0] == 0 and values[5] == 15 and values[1199] == 3597
    assert values[700] == "7000!" and values[701] == 3
    forms = await book.read("calc.get_range", range="C700:C703", include_formulas=True)
    assert forms["formulas"]["C703"] == "=A703+B703"
