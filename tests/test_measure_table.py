"""Measuring a table before placing it (T13): ``ppt_add_table`` with ``measure: true`` and
:meth:`Slide.measure_table` add nothing and say each row's drawn height and the pages its
rows split into, the header row repeated atop each -- one measuring model with
:attr:`Table.drawn_row_heights`, so what is measured is what is placed.

The data is made up: a toy grid document's ``t:tasks`` (as :mod:`test_tables_from_data`'s
toy format resolves one) whose every fifth row has a long note that wraps.
"""

from __future__ import annotations

import datetime as dt
import json

import pytest

from ooxml_edit.tools import Dataset, DocumentFormat, ToolError, Toolbox

from pptx_agent import Document, TableMeasure
from pptx_agent.edit.fit import paginate
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

pytest.importorskip("pptx2svg")

CLOCK = lambda: dt.datetime(2026, 10, 15, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
EMU_PER_POINT = 12700

COLUMNS = ("task", "team", "note")
ROWS = [(f"Task {n}", f"Team {n % 4}",
         "wait for the second delivery and check it twice" if n % 5 == 0 else "ok")
        for n in range(1, 41)]


class Grid:
    """Named tables, each a header and rows."""

    def __init__(self, tables: dict) -> None:
        self.tables = tables


def _grid_table(grid: Grid, address: str) -> Dataset:
    name = address[len("t:"):]
    if name not in grid.tables:
        raise ToolError("not_found", f"no table {name!r}", valid_options=sorted(grid.tables))
    header, rows = grid.tables[name]
    sources = [[(f"G!{chr(65 + c)}{r + 2}",) for c in range(len(header))]
               for r in range(len(rows))]
    return Dataset(header, rows, sources=sources, keys=1)


GRID = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table})


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = toolbox.session(clock=CLOCK)
    deck = Document.new()
    deck.add_slide("Title and Content")
    session.open(deck.to_bytes(), "deck.pptx")                                  # d1
    session.adopt(Grid({"tasks": (COLUMNS, ROWS)}), "grid", "book.grid")         # d2
    return session


def call(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
    return result


def failed(result, code):
    assert not result.ok, json.dumps(result.to_json(), indent=1, default=str)
    assert result.error.code == code, result.to_json()
    return result.error


def add(toolbox, session, box=None, **arguments):
    return call(toolbox, session, "ppt_add_table",
                **{"doc": "d1", "slide": "s:256",
                   "box": box or {"x": 40, "y": 120, "w": 600, "h": 0}, **arguments})


def _rows(text: str) -> tuple[int, int]:
    first, last = text.split("-")
    return int(first), int(last)


# -- the tool ------------------------------------------------------------------------------------


def test_a_long_table_from_a_source_is_measured_into_pages(toolbox, session):
    deck = session.entry("d1").document
    before, version = deck.to_bytes(), session.entry("d1").version
    result = ok(add(toolbox, session, **{"from": "t:tasks", "measure": True}))
    data = result.data
    assert result.created == [] and "nothing added" in result.summary
    # The page is the slide's content area when h is 0.
    area_height = deck.slides[0].content_area[3] / EMU_PER_POINT
    assert data["page_height"] == pytest.approx(area_height, abs=0.01)
    assert len(data["heights"]) == 41 and data["fits"] is False
    assert data["height"] == pytest.approx(sum(data["heights"]), abs=0.05)
    header, rows = data["heights"][0], data["heights"][1:]
    assert data["header"] == header
    assert rows[4] > rows[0]                        # row 5's note wraps
    # The pages cover the source rows 1-40 in order, each under the repeated header, as
    # many as fit: the next row would not.
    pages = [_rows(page["rows"]) for page in data["pages"]]
    assert pages[0][0] == 1 and pages[-1][1] == 40 and len(pages) > 1
    for (_, last), (first, _) in zip(pages, pages[1:]):
        assert first == last + 1
    for page, (first, last) in zip(data["pages"], pages):
        height = header + sum(rows[first - 1:last])
        assert page["h"] == pytest.approx(height, abs=0.05)
        assert page["h"] <= data["page_height"]
        if last < 40:
            assert height + rows[last] > data["page_height"]
    # Nothing changed: no version, no undo step, the same bytes.
    assert session.entry("d1").version == version
    assert deck.to_bytes() == before
    assert not call(toolbox, session, "undo", doc="d1").ok           # nothing to undo


def test_the_header_repeats_and_header_none_repeats_nothing(toolbox, session):
    box = {"x": 40, "y": 120, "w": 600, "h": 200}
    headed = ok(add(toolbox, session, box, **{"from": "t:tasks", "measure": True})).data
    bare = ok(add(toolbox, session, box, **{"from": "t:tasks", "header": [],
                                            "measure": True})).data
    assert headed["page_height"] == bare["page_height"] == 200
    assert bare["header"] is None and len(bare["heights"]) == 40
    assert bare["heights"] == headed["heights"][1:]
    # Every page of the headed table carries the header's height on top of its rows.
    for page in headed["pages"]:
        first, last = _rows(page["rows"])
        assert page["h"] == pytest.approx(headed["header"] + sum(
            headed["heights"][first:last + 1]), abs=0.05)
    assert len(headed["pages"]) >= len(bare["pages"])
    assert _rows(bare["pages"][0]["rows"])[1] > _rows(headed["pages"][0]["rows"])[1]


def test_measured_heights_are_the_placed_tables_drawn_heights(toolbox, session):
    measured = ok(add(toolbox, session, **{"from": "t:tasks", "measure": True})).data
    deck = session.entry("d1").document
    # The whole table, placed at its measured heights (a box h of 0).
    whole = ok(add(toolbox, session, **{"from": "t:tasks"}))
    drawn = deck.shape(whole.created[0]).table.drawn_row_heights
    assert [round(h / EMU_PER_POINT, 2) for h in drawn] == measured["heights"]
    # The first page alone, its rows typed in, under the header.
    first, last = _rows(measured["pages"][0]["rows"])
    texts = [list(COLUMNS)] + [list(row) for row in ROWS[first - 1:last]]
    page = ok(add(toolbox, session, data=texts, rows=len(texts), columns=3))
    drawn = deck.shape(page.created[0]).table.drawn_row_heights
    assert [round(h / EMU_PER_POINT, 2) for h in drawn] == measured["heights"][:last + 1]
    assert sum(drawn) / EMU_PER_POINT == pytest.approx(measured["pages"][0]["h"], abs=0.05)


def test_size_and_widths_are_measured_as_they_are_placed(toolbox, session):
    options = {"from": "t:tasks", "size": 11, "widths": [90, 120, 160]}
    small = ok(add(toolbox, session, measure=True, **options)).data
    large = ok(add(toolbox, session, measure=True, **{"from": "t:tasks"})).data
    assert len(small["pages"]) < len(large["pages"])
    placed = ok(add(toolbox, session, **options))
    table = session.entry("d1").document.shape(placed.created[0]).table
    assert [round(w / EMU_PER_POINT, 2) for w in table.column_widths] == [90, 120, 160]
    assert table.cell(1, 0).text_frame.paragraph(0).runs[0].size == 11
    assert [round(h / EMU_PER_POINT, 2) for h in table.drawn_row_heights] == small["heights"]


def test_literal_cells_and_a_box_height(toolbox, session):
    cells = [["Step", "Note"], ["One", "short"], ["Two", ""]]
    data = ok(add(toolbox, session, {"x": 40, "y": 120, "w": 400, "h": 300}, data=cells,
                  rows=4, columns=2, measure=True)).data
    assert data["fits"] is True and data["page_height"] == 300
    assert data["pages"] == [{"rows": "1-3", "h": data["height"]}]
    # Empty rows are drawn a line tall, as the placed table's are.
    assert data["heights"][3] == data["heights"][1]


def test_measuring_refusals(toolbox, session):
    failed(add(toolbox, session, {"x": 0, "y": 0, "w": 0, "h": 0}, **{"from": "t:tasks",
                                                                       "measure": True}),
           "invalid_arguments")
    error = failed(add(toolbox, session, **{"from": "t:tasks", "widths": [100, 100],
                                            "measure": True}), "invalid_arguments")
    assert error.field == "widths"
    error = failed(call(toolbox, session, "ppt_add_table", doc="d1", slide="L2",
                        box={"x": 0, "y": 0, "w": 300, "h": 0}, rows=3, columns=2,
                        measure=True), "invalid_arguments")
    assert error.field == "box.h"


def test_a_layout_is_measured_on_a_slide_built_on_it(toolbox, session):
    data = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="L2",
                   box={"x": 0, "y": 0, "w": 300, "h": 100}, rows=6, columns=2,
                   measure=True)).data
    assert [page["rows"] for page in data["pages"]] == ["1-2", "3-4", "5-5"]
    assert len(session.entry("d1").document.slides) == 1


# -- the library ---------------------------------------------------------------------------------


def test_slide_measure_table_leaves_the_deck_unchanged():
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    before, version = deck.to_bytes(), deck.history.version
    texts = [list(COLUMNS)] + [list(row) for row in ROWS]
    measured = slide.measure_table(texts, size=12)
    assert isinstance(measured, TableMeasure)
    assert deck.to_bytes() == before and deck.history.version == version
    assert measured.page_height == slide.content_area[3] and measured.header_rows == 1
    assert measured.pages[0].start == 1 and measured.pages[-1].stop == 41
    assert all(height <= measured.page_height for height in measured.page_heights)
    assert measured.too_tall == () and not measured.fits
    assert len(measured.pages) < len(slide.measure_table(texts).pages)


def test_paginate_gives_a_row_too_tall_a_page_of_its_own():
    pt = EMU_PER_POINT
    heights = [10 * pt, 20 * pt, 20 * pt, 95 * pt, 20 * pt, 20 * pt]
    assert paginate(heights, 50 * pt) == (range(1, 3), range(3, 4), range(4, 6))
    assert paginate(heights, 50 * pt, header_rows=0) == (range(0, 3), range(3, 4),
                                                         range(4, 6))
    assert paginate([10 * pt], 50 * pt) == ()
    measured = TableMeasure(tuple(heights), 1, 50 * pt, paginate(heights, 50 * pt))
    assert measured.too_tall == (3,)
    assert measured.page_heights == (50 * pt, 105 * pt, 50 * pt)
