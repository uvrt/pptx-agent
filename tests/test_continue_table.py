"""A long table continued on the next slides (T3): ``ppt_add_table`` with ``continue: true``
measures the table into pages as ``measure: true`` does, places the first page on the
target slide and each next one on a new slide after it -- the target's layout and title,
the same box, the header row repeated -- every row at its measured height; each table keeps
the record of its own rows.  And a row range of the source (``d:tasks/#9:#16``) places
fewer rows.

The data is made up: a toy grid document's ``t:tasks`` (as :mod:`test_measure_table`'s)
whose every fifth row has a long note that wraps, and the same rows kept as ``d:tasks``.
"""

from __future__ import annotations

import datetime as dt
import json

import pytest

from ooxml_edit.tools import Dataset, DatasetEntry, DocumentFormat, ToolError, Toolbox
from ooxml_edit.tools.sources import runs

from pptx_agent import Document
from pptx_agent.edit import provenance as table_records
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

pytest.importorskip("pptx2svg")

CLOCK = lambda: dt.datetime(2026, 10, 15, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
EMU_PER_POINT = 12700

COLUMNS = ("task", "team", "note")
ROWS = [(f"Task {n}", f"Team {n % 4}",
         "wait for the second delivery and check it twice" if n % 5 == 0 else "ok")
        for n in range(1, 41)]
BOX = {"x": 40, "y": 120, "w": 600, "h": 0}


class Grid:
    def __init__(self, tables: dict) -> None:
        self.tables = tables


def _grid_table(grid: Grid, address: str) -> Dataset:
    name = address[len("t:"):]
    if name not in grid.tables:
        raise ToolError("not_found", f"no table {name!r}", valid_options=sorted(grid.tables))
    header, rows = grid.tables[name]
    return Dataset(header, rows, sources=_sources(len(rows)), keys=1)


def _sources(count: int):
    return [[(f"G!{chr(65 + c)}{r + 2}",) for c in range(3)] for r in range(count)]


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
    first = deck.add_slide("Title and Content")
    first.title = "Open tasks"
    deck.add_slide("Title Only").title = "Next steps"
    session.open(deck.to_bytes(), "deck.pptx")                                  # d1
    session.adopt(Grid({"tasks": (COLUMNS, ROWS)}), "grid", "book.grid")         # d2
    session.keep_datasets([DatasetEntry("tasks", Dataset(COLUMNS, ROWS,
                                                         sources=_sources(40), keys=1))])
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
                **{"doc": "d1", "slide": "s:256", "box": box or BOX, **arguments})


def _rows(text: str) -> tuple[int, int]:
    first, last = text.split("-")
    return int(first), int(last)


def test_a_long_table_goes_on_over_new_slides_after_the_target(toolbox, session):
    measured = ok(add(toolbox, session, **{"from": "t:tasks", "measure": True})).data
    result = ok(add(toolbox, session, **{"from": "t:tasks", "continue": True}))
    pages = result.data["pages"]
    assert [page["rows"] for page in pages] == [page["rows"] for page in measured["pages"]]
    assert len(pages) > 1 and pages[0]["slide"] == "s:256"
    deck = session.entry("d1").document
    # The new slides come right after the target, before the slide that followed it.
    order = [f"s:{slide.slide_id}" for slide in deck.slides]
    assert order == [page["slide"] for page in pages] + ["s:257"]
    assert set(result.created) == {pages[0]["table"]} | {
        item for page in pages[1:] for item in (page["slide"], page["table"])}
    target = deck.slides[0]
    for page, slide in zip(pages, deck.slides):
        first, last = _rows(page["rows"])
        table = deck.shape(page["table"]).table
        # The header row atop each page, then that page's rows.
        assert table.rows == 1 + last - first + 1
        assert [table.cell(0, c).text for c in range(3)] == list(COLUMNS)
        assert table.cell(1, 0).text == f"Task {first}"
        assert table.cell(table.rows - 1, 0).text == f"Task {last}"
        # Each row at its measured height: the page is drawn as measured.
        heights = [measured["heights"][0]] + measured["heights"][first:last + 1]
        drawn = [round(h / EMU_PER_POINT, 2) for h in table.drawn_row_heights]
        assert drawn == pytest.approx(heights, abs=0.02)
        frame = deck.shape(page["table"])
        assert frame.bounds[1] == 120 * EMU_PER_POINT
        assert frame.bounds[3] / EMU_PER_POINT == pytest.approx(page["h"], abs=0.05)
        # The target's layout and title; no empty placeholder left behind.
        assert slide.layout.name == target.layout.name and slide.title == "Open tasks"
        if slide is not target:
            assert slide.empty_placeholders() == []


def test_each_page_keeps_the_record_of_its_own_rows(toolbox, session):
    pages = ok(add(toolbox, session, **{"from": "t:tasks", "continue": True})).data["pages"]
    records = session.entry("d1").document.provenance()
    for page in pages:
        first, last = _rows(page["rows"])
        placed = page["from"]
        assert placed["header_row"] == 0 and placed["first_row"] == 1
        assert [column["rows"] for column in placed["columns"]] == \
            [runs(range(first - 1, last))] * 3
        assert placed["columns"][0]["sources"] == [f"G!A{row + 1}"
                                                   for row in range(first, last + 1)]
        record = records[page["table"]]
        assert record["source"] == "t:tasks" and record["columns"] == placed["columns"]
    # A page's block is its header and its own rows: a row inserted below it leaves the
    # record, one inside drops it.
    record = records[pages[1]["table"]]
    first, last = _rows(pages[1]["rows"])
    count = last - first + 1
    assert table_records.after_row_insert(record, count + 1) == record
    assert table_records.after_row_insert(record, count) is None


def test_one_undo_takes_back_every_page_and_slide(toolbox, session):
    entry = session.entry("d1")
    before = entry.document.to_bytes()
    ok(add(toolbox, session, **{"from": "t:tasks", "continue": True}))
    assert len(entry.document.slides) > 2
    ok(call(toolbox, session, "undo", doc="d1"))
    assert len(entry.document.slides) == 2
    assert entry.document.to_bytes() == before


def test_a_table_that_fits_stays_on_its_slide(toolbox, session):
    result = ok(add(toolbox, session, **{"from": "d:tasks/#1:#4", "continue": True}))
    assert [page["rows"] for page in result.data["pages"]] == ["1-4"]
    assert result.data["pages"][0]["from"]["source"] == "d:tasks/#1:#4"
    assert len(session.entry("d1").document.slides) == 2
    assert result.data["cells"][4][0] == "Task 4"


def test_a_row_range_continues_from_its_own_rows(toolbox, session):
    box = {"x": 40, "y": 120, "w": 600, "h": 150}
    pages = ok(add(toolbox, session, box, **{"from": "d:tasks/#21:", "continue": True})
               ).data["pages"]
    assert _rows(pages[0]["rows"])[0] == 1 and _rows(pages[-1]["rows"])[1] == 20
    table = session.entry("d1").document.shape(pages[0]["table"]).table
    assert table.cell(1, 0).text == "Task 21"
    assert pages[0]["from"]["columns"][0]["sources"][0].startswith("G!A22")
    assert all(page["h"] <= 150 for page in pages)


def test_literal_cells_continue_under_their_first_row(toolbox, session):
    cells = [["Step", "Note"]] + [[f"Step {n}", "check"] for n in range(1, 13)]
    box = {"x": 40, "y": 120, "w": 400, "h": 120}
    pages = ok(add(toolbox, session, box, data=cells, rows=13, columns=2,
                   **{"continue": True})).data["pages"]
    assert len(pages) > 1 and "from" not in pages[0]
    deck = session.entry("d1").document
    assert all(deck.shape(page["table"]).table.cell(0, 0).text == "Step" for page in pages)


def test_continuing_refusals(toolbox, session):
    error = failed(add(toolbox, session, **{"from": "t:tasks", "rows": 12,
                                            "continue": True}), "invalid_arguments")
    assert error.field == "rows"
    error = failed(call(toolbox, session, "ppt_add_table", doc="d1", slide="L2",
                        box={"x": 0, "y": 0, "w": 300, "h": 100}, rows=6, columns=2,
                        **{"continue": True}), "invalid_arguments")
    assert error.field == "continue"
    failed(add(toolbox, session, {"x": 0, "y": 0, "w": 0, "h": 0},
               **{"from": "t:tasks", "continue": True}), "invalid_arguments")
    assert len(session.entry("d1").document.slides) == 2


def test_fewer_rows_than_the_source_say_how_to_slice(toolbox, session):
    error = failed(add(toolbox, session, **{"from": "t:tasks", "rows": 10}),
                   "invalid_arguments")
    assert "from has 40 rows and a header" in error.message
    assert "d2/t:tasks/#1:#9" in error.message and "continue: true" in error.message
    assert error.valid_options == ["d2/t:tasks/#1:#9"]
    error = failed(add(toolbox, session, **{"from": "d:tasks/#11:#30", "rows": 5}),
                   "invalid_arguments")
    assert error.valid_options == ["d:tasks/#11:#14"]
    error = failed(add(toolbox, session, **{"from": "d:tasks/#41:"}), "invalid_arguments")
    assert error.valid_options == ["d:tasks/#40"] and "nothing is cut" in error.message
