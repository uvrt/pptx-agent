"""Tables: cell text, fill and borders, sizing, and merge-aware structural edits.

pptx-svg documents that adding a row or column does not adjust ``gridSpan``/``rowSpan`` on
merged cells.  The tests here check every structural operation against an independent model
of where the merged regions should end up -- including a randomised sweep -- and check the
grid's merge attributes for internal consistency after each step.
"""

from __future__ import annotations

import random

import pytest

from pptx_agent import Document
from pptx_agent.edit.table import Region, _cells
from pptx_agent.oxml.xml import qn


def reopen(document: Document) -> Document:
    return Document.open(document.to_bytes())


def _table_shape(document: Document):
    return next(s for slide in document.slides for s in slide.shapes if s.has_table)


def check_grid(table) -> None:
    """Every invariant PowerPoint relies on, spelled out (see edit/table.py's docstring)."""
    rows = table._rows()
    columns = len(table._grid_columns())
    assert rows, "a table needs rows"
    covered: dict[tuple[int, int], Region] = {}
    for row in rows:
        assert len(_cells(row)) == columns, "every row must have one a:tc per a:gridCol"
    for region in table._regions():
        assert region.bottom < len(rows) and region.right < columns, f"{region} overruns"
        for r in range(region.top, region.bottom + 1):
            for c in range(region.left, region.right + 1):
                assert (r, c) not in covered, f"{(r, c)} is in two regions"
                covered[(r, c)] = region
                tc = _cells(rows[r])[c]
                assert (tc.get("hMerge") == "1") == (c > region.left), (r, c)
                assert (tc.get("vMerge") == "1") == (r > region.top), (r, c)
                expected_rows = str(region.height) if r == region.top and region.height > 1 else None
                expected_cols = str(region.width) if c == region.left and region.width > 1 else None
                assert tc.get("rowSpan") == expected_rows, (r, c)
                assert tc.get("gridSpan") == expected_cols, (r, c)
    assert len(covered) == len(rows) * columns


# -- cells ---------------------------------------------------------------------------------


def test_cell_text_uses_the_text_api(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    cell = table.cell(1, 1)
    run = cell.text_frame.paragraph(0).run(0)
    run.bold = True
    run.color = "accent2"
    cell.text_frame.add_paragraph("second line", italic=True)

    reread = reopen(document).resolve(cell.address)
    assert reread.text.endswith("\nsecond line")
    assert reread.text_frame.paragraph(0).run(0).bold is True
    assert reread.text_frame.paragraph(0).run(0).color == "accent2"
    assert document.resolve(f"{cell.address}/p1/r0").italic is True


def test_cell_text_keeps_formatting(financial_report):
    document = Document.open(str(financial_report))
    cell = _table_shape(document).table.cell(1, 1)
    run = cell.text_frame.paragraph(0).run(0)
    expected = (run.size, run.color)

    cell.text = cell.text.replace("4,285", "4,300")

    reread = reopen(document).resolve(cell.address).text_frame.paragraph(0).run(0)
    assert reread.text.startswith("4,300")
    assert (reread.size, reread.color) == expected


def test_cell_fill(financial_report):
    document = Document.open(str(financial_report))
    cell = _table_shape(document).table.cell(2, 3)
    cell.fill = "accent1 lumMod=20% lumOff=80%"
    assert reopen(document).resolve(cell.address).fill.color == "accent1 lumMod=20000 lumOff=80000"
    cell.set_gradient_fill(["accent1", "accent2"], angle=0)
    assert reopen(document).resolve(cell.address).fill.kind == "gradient"
    cell.fill = "none"
    assert reopen(document).resolve(cell.address).fill.kind == "none"


def test_cell_borders(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    cell = table.cell(1, 1)
    cell.set_border("bottom", width=25400, color="accent1", dash="dash")

    reread = reopen(document).shape(_table_shape(document).id).table
    bottom = reread.cell(1, 1).border("bottom")
    assert (bottom.width, bottom.color, bottom.dash) == (25400, "accent1", "dash")
    # The shared edge is styled on the neighbour too.
    top = reread.cell(2, 1).border("top")
    assert (top.width, top.color) == (25400, "accent1")


def test_border_on_a_cell_with_no_tcpr(product_page, financial_report):
    """Borders must go into tcPr in lnL/lnR/lnT/lnB order even when built from nothing."""
    document = Document.open(str(financial_report))
    cell = _table_shape(document).table.cell(0, 0)
    for side in ("bottom", "top", "right", "left"):
        cell.border(side).width = 12700
    names = [c.tag.rpartition("}")[2] for c in cell._tc().find(qn("a:tcPr"))]
    lines = [n for n in names if n.startswith("ln")]
    assert lines == ["lnL", "lnR", "lnT", "lnB"]
    assert names.index("lnB") < names.index("solidFill")


def test_column_width_and_row_height(financial_report):
    document = Document.open(str(financial_report))
    shape = _table_shape(document)
    table = shape.table
    frame_width, frame_height = shape.width, shape.height
    old_width, old_height = table.column_widths[1], table.row_heights[2]

    table.set_column_width(1, 1500000)
    table.set_row_height(2, 500000)

    reread = reopen(document).shape(shape.id)
    assert reread.table.column_widths[1] == 1500000
    assert reread.table.row_heights[2] == 500000
    assert reread.width == frame_width + 1500000 - old_width
    assert reread.height == frame_height + 500000 - old_height


# -- merges --------------------------------------------------------------------------------


def test_merge_writes_powerpoints_pattern(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    texts = [table.cell(r, c).text for r in (1, 2) for c in (1, 2)]

    origin = table.merge(1, 1, 2, 2)

    check_grid(table)
    rows = table._rows()
    attributes = [[dict(_cells(rows[r])[c].attrib) for c in (1, 2)] for r in (1, 2)]
    assert attributes == [
        [{"gridSpan": "2", "rowSpan": "2"}, {"rowSpan": "2", "hMerge": "1"}],
        [{"gridSpan": "2", "vMerge": "1"}, {"hMerge": "1", "vMerge": "1"}],
    ]
    assert origin.span == (2, 2) and origin.is_merge_origin
    assert table.cell(2, 2).is_spanned and table.cell(2, 2).origin.address == origin.address
    assert origin.text.split("\n") == [t for t in texts if t]
    assert reopen(document).resolve(origin.address).span == (2, 2)


def test_merge_refuses_to_cut_a_region(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.merge(1, 1, 2, 2)
    with pytest.raises(ValueError, match="cut through"):
        table.merge(2, 2, 3, 3)
    table.merge(1, 0, 3, 2)  # a superset is fine
    check_grid(table)
    assert table.merged_regions == [Region(1, 0, 3, 2)]


def test_split_undoes_a_merge(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.merge(1, 1, 3, 2)
    table.split(2, 2)
    check_grid(table)
    assert table.merged_regions == []


def test_insert_row_inside_a_merge_lengthens_it(financial_report):
    document = Document.open(str(financial_report))
    shape = _table_shape(document)
    table = shape.table
    table.merge(1, 1, 3, 2)
    origin_text = table.cell(1, 1).text
    rows, height = table.rows, shape.height

    table.insert_row(2)

    check_grid(table)
    assert table.rows == rows + 1
    assert table.merged_regions == [Region(1, 1, 4, 2)]
    assert table.cell(1, 1).text == origin_text
    assert shape.height == height + table.row_heights[2]
    reread = reopen(document).shape(shape.id).table
    check_grid(reread)
    assert reread.merged_regions == [Region(1, 1, 4, 2)]


@pytest.mark.parametrize("position", [1, 4])
def test_insert_row_at_a_merge_edge_leaves_it_alone(financial_report, position):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.merge(1, 1, 3, 2)
    table.insert_row(position)
    check_grid(table)
    expected = Region(2, 1, 4, 2) if position == 1 else Region(1, 1, 3, 2)
    assert table.merged_regions == [expected]


def test_insert_column_inside_a_merge_widens_it(financial_report):
    document = Document.open(str(financial_report))
    shape = _table_shape(document)
    table = shape.table
    table.merge(0, 1, 1, 3)
    width = shape.width

    table.insert_column(2, width=400000)

    check_grid(table)
    assert table.merged_regions == [Region(0, 1, 1, 4)]
    assert shape.width == width + 400000
    check_grid(reopen(document).shape(shape.id).table)


def test_deleting_a_merges_top_row_hands_the_text_down(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.cell(1, 1).text = "origin text"
    table.cell(1, 1).fill = "accent4"
    table.merge(1, 1, 3, 2)

    table.delete_row(1)

    check_grid(table)
    assert table.merged_regions == [Region(1, 1, 2, 2)]
    assert table.cell(1, 1).text.split("\n")[0] == "origin text"
    assert table.cell(1, 1).fill.color == "accent4"


def test_deleting_a_merges_left_column_hands_the_text_right(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.cell(1, 1).text = "origin text"
    table.merge(1, 1, 2, 3)

    table.delete_column(1)

    check_grid(table)
    assert table.merged_regions == [Region(1, 1, 2, 2)]
    assert table.cell(1, 1).text.split("\n")[0] == "origin text"


def test_shrinking_a_merge_to_one_cell_unmerges_it(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.merge(1, 1, 2, 1)
    table.delete_row(2)
    check_grid(table)
    assert table.merged_regions == []


def test_new_rows_copy_formatting_not_text(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    table.insert_row(1, like=0)
    header = table.cell(0, 0)
    new = table.cell(1, 0)
    assert new.text == ""
    assert new.fill == header.fill
    end = new._tc().find(qn("a:txBody")).find(qn("a:p")).find(qn("a:endParaRPr"))
    assert end is not None  # typing into the new cell continues the header's run style


def test_copied_rows_get_fresh_powerpoint_row_ids(financial_report):
    from lxml import etree

    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    a16 = "http://schemas.microsoft.com/office/drawing/2014/main"
    for index, row in enumerate(table._rows()):
        ext_list = etree.SubElement(row, qn("a:extLst"))
        ext = etree.SubElement(ext_list, qn("a:ext"), uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}")
        etree.SubElement(ext, "{%s}rowId" % a16, val=str(10000 + index))
    table.insert_row(1)
    ids = [node.get("val") for node in table._table().iter("{%s}rowId" % a16)]
    assert len(ids) == len(set(ids)) == table.rows


def test_tables_keep_at_least_one_row_and_column(financial_report):
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    while table.rows > 1:
        table.delete_row(0)
    while table.columns > 1:
        table.delete_column(0)
    with pytest.raises(ValueError):
        table.delete_row(0)
    with pytest.raises(ValueError):
        table.delete_column(0)
    check_grid(table)


def test_structural_edits_are_single_undo_steps(financial_report):
    document = Document.open(str(financial_report))
    original = document.to_bytes()
    table = _table_shape(document).table
    table.merge(1, 1, 3, 2)
    table.insert_row(2)
    table.insert_column(2)
    table.delete_row(1)
    table.delete_column(1)
    for _ in range(5):
        assert document.undo()
    assert document.to_bytes() == original


# -- the randomised sweep ------------------------------------------------------------------


class Model:
    """Where merged regions *should* be, computed without looking at any XML."""

    def __init__(self, rows: int, columns: int) -> None:
        self.rows, self.columns = rows, columns
        self.regions: list[Region] = []

    def merge(self, region: Region) -> None:
        self.regions = [r for r in self.regions if not (
            region.contains(r.top, r.left) and region.contains(r.bottom, r.right))]
        self.regions.append(region)

    def insert_row(self, at: int) -> None:
        self.rows += 1
        self.regions = [self._ins(r, at) for r in self.regions]

    def delete_row(self, at: int) -> None:
        self.rows -= 1
        self.regions = [self._del(r, at) for r in self.regions]

    def insert_column(self, at: int) -> None:
        self.columns += 1
        self.regions = [self._t(self._ins(self._t(r), at)) for r in self.regions]

    def delete_column(self, at: int) -> None:
        self.columns -= 1
        self.regions = [self._t(self._del(self._t(r), at)) for r in self.regions]

    @staticmethod
    def _t(region: Region) -> Region:
        return Region(*_transpose(region))

    @staticmethod
    def _ins(r: Region, at: int) -> Region:
        if r.top >= at:
            return Region(r.top + 1, r.left, r.bottom + 1, r.right)
        if r.bottom >= at:
            return Region(r.top, r.left, r.bottom + 1, r.right)
        return r

    @staticmethod
    def _del(r: Region, at: int) -> Region:
        if r.top > at:
            return Region(r.top - 1, r.left, r.bottom - 1, r.right)
        if r.bottom >= at:
            return Region(r.top, r.left, r.bottom - 1, r.right)
        return r

    def merged(self) -> list[Region]:
        valid = [r for r in self.regions if r.bottom >= r.top and r.right >= r.left]
        return sorted(r for r in valid if r.height > 1 or r.width > 1)


def _transpose(region: Region) -> tuple[int, int, int, int]:
    return region.left, region.top, region.right, region.bottom


@pytest.mark.parametrize("seed", range(40))
def test_random_structural_edits_match_the_model(financial_report, seed):
    rng = random.Random(seed)
    document = Document.open(str(financial_report))
    table = _table_shape(document).table
    model = Model(table.rows, table.columns)

    for _ in range(14):
        operation = rng.choice(["merge", "merge", "insert_row", "insert_column",
                                "delete_row", "delete_column"])
        if operation == "merge":
            top, left = rng.randrange(table.rows), rng.randrange(table.columns)
            bottom = min(table.rows - 1, top + rng.randrange(3))
            right = min(table.columns - 1, left + rng.randrange(3))
            region = Region(top, left, bottom, right)
            try:
                table.merge(top, left, bottom, right)
            except ValueError:
                continue
            if region.height > 1 or region.width > 1:
                model.merge(region)
        elif operation == "insert_row":
            at = rng.randrange(table.rows + 1)
            table.insert_row(at)
            model.insert_row(at)
        elif operation == "insert_column":
            at = rng.randrange(table.columns + 1)
            table.insert_column(at)
            model.insert_column(at)
        elif operation == "delete_row" and table.rows > 1:
            at = rng.randrange(table.rows)
            table.delete_row(at)
            model.delete_row(at)
        elif operation == "delete_column" and table.columns > 1:
            at = rng.randrange(table.columns)
            table.delete_column(at)
            model.delete_column(at)

        check_grid(table)
        assert (table.rows, table.columns) == (model.rows, model.columns)
        assert sorted(table.merged_regions) == model.merged(), operation

    reread = reopen(document).shape(_table_shape(document).id).table
    check_grid(reread)
    assert sorted(reread.merged_regions) == model.merged()
