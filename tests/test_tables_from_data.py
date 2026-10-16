"""Slide tables from data (X1): ``ppt_add_table`` and ``ppt_edit_table``'s ``set_cells`` take
their rows from a dataset (``d:<id>``) or a table of another open document (``t:<name>``),
pick its columns by name, and report where every placed value came from.

The toys, all made up: a dataset a toy tool keeps (``d:scores``), and a toy grid format
whose documents are named tables of cells, resolving ``t:<name>`` to rows with a source per
cell (``G!B2``) and merging runs of cells in one column into ranges (``G!B2:B4``).  A test
with xlsx-agent's own ``t:`` waits until xlsx-agent registers its resolver.
"""

from __future__ import annotations

import datetime as dt
import json

import pytest

from ooxml_common.locales import Locale
from ooxml_edit.tools import (Dataset, DocumentFormat, Result, ToolError, Toolbox, string,
                              strict_problems, tool)

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

#: A made-up checks table: a name, a rating in half points, a date as ISO text, a flag, a note.
COLUMNS = ("check", "rating", "due", "done", "note")
ROWS = [("Alpha", 3.5, "2026-01-05", True, "fine"),
        ("Beta", 4, "2026-02-10", False, None),
        ("Gamma", 1.25, "2026-03-15", True, "late")]


#: The checks' column types, as a workbook's date column gives them: ``due`` is a date.
TYPES = ("text", "number", "date", "bool", "text")


def _letter(index: int) -> str:
    return chr(ord("A") + index)


def _sources(rows) -> list:
    return [[(f"G!{_letter(c)}{r + 2}",) for c in range(len(COLUMNS))] for r in range(len(rows))]


class Grid:
    """Named tables, each a header and rows; cell ``B2`` is the second column's first row."""

    def __init__(self, tables: dict) -> None:
        self.tables = tables


def _grid_table(grid: Grid, address: str) -> Dataset:
    """``t:<name>``, or one row of it by key, ``t:<name>/<key>`` (as a workbook's ``t:``
    takes one, for a text fact)."""
    name, _, key = address[len("t:"):].partition("/")
    if name not in grid.tables:
        raise ToolError("not_found", f"no table {name!r}", valid_options=sorted(grid.tables))
    header, rows = grid.tables[name]
    sources = _sources(rows)
    if key:
        found = [i for i, row in enumerate(rows) if row[0] == key]
        if not found:
            raise ToolError("not_found", f"no row {key!r} in {name}")
        rows, sources = [rows[found[0]]], [sources[found[0]]]
    return Dataset(header, rows, sources=sources, keys=1)


def _merge(addresses):
    """Runs of cells in one column as ranges: G!B2, G!B3 -> G!B2:B3."""
    out: list[list] = []
    for address in addresses:
        sheet, cell = address.split("!")
        column, row = cell[0], int(cell[1:])
        last = out[-1] if out else None
        if last and last[:2] == [sheet, column] and last[3] == row - 1:
            last[3] = row
        else:
            out.append([sheet, column, row, row])
    return [f"{s}!{c}{a}" if a == b else f"{s}!{c}{a}:{c}{b}" for s, c, a, b in out]


GRID = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table}, merge_sources=_merge)


@tool("ds_put", "Keep a made-up dataset.",
      {"id": string("Its id."), "size": string("rows x columns of filler.", optional=True),
       "typed": string("Label the columns' types, as a workbook's t: does.", optional=True)},
      documents=())
def ds_put(call, id, size=None, typed=None):
    if typed:
        call.define_dataset(id, Dataset(COLUMNS, ROWS, sources=_sources(ROWS), keys=1,
                                        types=TYPES), source=None)
    elif size:
        rows, columns = (int(n) for n in size.split("x"))
        names = [f"c{i}" for i in range(columns)]
        call.define_dataset(id, Dataset(names, [[r] * columns for r in range(rows)]),
                            source=None)
    else:
        call.define_dataset(id, Dataset(COLUMNS, ROWS, sources=_sources(ROWS), keys=1),
                            source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox, **options):
    session = toolbox.session(clock=CLOCK, **options)
    deck = Document.new()
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")                                  # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")       # d2
    return session


@pytest.fixture
def session(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
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


BOX = {"x": 40, "y": 60, "w": 600, "h": 200}


def add(toolbox, session, **arguments):
    return call(toolbox, session, "ppt_add_table",
                **{"doc": "d1", "slide": "s:256", "box": BOX, **arguments})


# -- the definitions -----------------------------------------------------------------------------


def test_both_tools_declare_from_and_stay_strict_clean(toolbox):
    for name in ("ppt_add_table", "ppt_edit_table"):
        tool_ = toolbox.tools[name]
        assert tool_.sources == ("from",)
        properties = tool_.schema["properties"]
        assert properties["from"]["type"] == "string"
        assert properties["from_columns"]["items"] == {"type": "string"}
        assert properties["header"]["items"] == {"type": "string"}
        assert not {"from", "from_columns", "header"} & set(tool_.schema.get("required", []))
    assert "rows" not in toolbox.tools["ppt_add_table"].schema["required"]
    assert strict_problems([t for t in TOOLS]) == []


# -- ppt_add_table -------------------------------------------------------------------------------


def test_a_table_from_a_dataset_header_rows_and_provenance(toolbox, session):
    result = ok(add(toolbox, session, **{"from": "d:scores",
                                         "from_columns": ["check", "rating", "due", "done"]}))
    data = result.data
    assert (data["rows"], data["columns"]) == (4, 4)
    assert data["cells"] == [["check", "rating", "due", "done"],
                             ["Alpha", "3.5", "2026-01-05", "TRUE"],
                             ["Beta", "4", "2026-02-10", "FALSE"],
                             ["Gamma", "1.25", "2026-03-15", "TRUE"]]
    placed = data["from"]
    assert placed["source"] == "d:scores" and placed["doc"] is None
    assert placed["header_row"] == 0 and placed["first_row"] == 1
    assert placed["columns"][1] == {"col": 1, "column": "rating", "rows": "1-3",
                                    "sources": ["G!B2", "G!B3", "G!B4"]}
    assert [c["column"] for c in placed["columns"]] == ["check", "rating", "due", "done"]
    deck = session.entry("d1").document
    xml = deck.shape(result.created[0]).table
    assert xml.cell(3, 1).text == "1.25"
    # The deck keeps the record too (test_provenance), without the session's handle.
    assert deck.provenance()[result.created[0]]["columns"][1]["sources"] == \
        ["G!B2", "G!B3", "G!B4"]


def test_columns_are_picked_and_ordered_by_name_and_default_to_all(toolbox, session):
    data = ok(add(toolbox, session, **{"from": "d:scores",
                                       "from_columns": ["Rating", "check"]})).data
    assert data["cells"][0] == ["rating", "check"]     # the source's own spelling
    assert data["cells"][1] == ["3.5", "Alpha"]
    everything = ok(add(toolbox, session, **{"from": "d:scores"})).data
    assert everything["columns"] == 5 and everything["cells"][0] == list(COLUMNS)
    assert everything["cells"][2][4] == ""             # a blank is an empty cell


def test_a_header_given_or_none(toolbox, session):
    given = ok(add(toolbox, session, **{"from": "d:scores", "from_columns": ["check", "rating"],
                                        "header": ["Check", "Rating (1-5)"]})).data
    assert given["cells"][0] == ["Check", "Rating (1-5)"] and given["rows"] == 4
    none = ok(add(toolbox, session, **{"from": "d:scores", "from_columns": ["check"],
                                       "header": []})).data
    assert none["rows"] == 3 and none["cells"][0] == ["Alpha"]
    assert none["from"]["header_row"] is None and none["from"]["first_row"] == 0
    error = failed(add(toolbox, session, **{"from": "d:scores", "from_columns": ["check"],
                                            "header": ["a", "b"]}), "invalid_arguments")
    assert error.field == "header"


def test_a_bigger_table_than_the_source_keeps_blank_cells(toolbox, session):
    data = ok(add(toolbox, session, **{"from": "d:scores", "from_columns": ["check"],
                                       "rows": 6, "columns": 2})).data
    assert (data["rows"], data["columns"]) == (6, 2) and data["cells"][5] == ["", ""]
    error = failed(add(toolbox, session, **{"from": "d:scores", "rows": 3}),
                   "invalid_arguments")
    assert error.field == "rows" and "from has 3 rows and a header" in error.message
    # It says how to place fewer: a row range of the source, or continue.
    assert error.valid_options == ["d:scores/#1:#2"] and "continue: true" in error.message
    error = failed(add(toolbox, session, **{"from": "d:scores/#2:", "rows": 2}),
                   "invalid_arguments")
    assert error.valid_options == ["d:scores/#2"]
    error = failed(add(toolbox, session, **{"from": "d:scores", "columns": 1}),
                   "invalid_arguments")
    assert error.field == "columns" and "needs 5" in error.message


def test_numbers_are_written_in_the_sessions_locale(toolbox):
    dutch = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                   language="nl-NL", list_separator=";")
    session = _session(toolbox, locale=dutch)
    data = ok(add(toolbox, session, **{"from": "t:checks",
                                       "from_columns": ["rating", "due"]})).data
    assert [row[0] for row in data["cells"][1:]] == ["3,5", "4", "1,25"]
    assert data["cells"][1][1] == "2026-01-05"          # a date as the source gives it


def test_a_table_from_another_documents_table_merges_its_sources(toolbox, session):
    data = ok(add(toolbox, session, **{"from": "t:checks",
                                       "from_columns": ["check", "note"]})).data
    placed = data["from"]
    assert placed["source"] == "d2/t:checks" and placed["doc"] == "d2"
    assert placed["columns"][0]["sources"] == ["G!A2:A4"]
    assert placed["columns"][1] == {"col": 1, "column": "note", "rows": "1-3",
                                    "sources": ["G!E2:E4"]}


def test_from_with_data_or_without_a_source_is_refused(toolbox, session):
    error = failed(add(toolbox, session, **{"from": "d:scores", "data": [["x"]]}),
                   "invalid_arguments")
    assert error.field == "data"
    error = failed(add(toolbox, session, rows=2, columns=2, from_columns=["check"]),
                   "invalid_arguments")
    assert error.field == "from_columns"
    error = failed(add(toolbox, session, columns=2), "invalid_arguments")
    assert error.field == "rows"
    error = failed(add(toolbox, session, **{"from": "d:scores", "from_columns": ["nope"]}),
                   "not_found")
    assert "rating" in error.valid_options
    error = failed(add(toolbox, session, **{"from": "d:scores",
                                            "from_columns": ["check", "CHECK"]}),
                   "invalid_arguments")
    failed(add(toolbox, session, **{"from": "d:missing"}), "not_found")


def test_a_source_too_big_is_refused_never_cut(toolbox, session):
    ok(call(toolbox, session, "ds_put", id="tall", size="200x2"))
    error = failed(add(toolbox, session, **{"from": "d:tall"}), "limit")
    assert "201 rows with the header" in error.message and "200" in error.message
    ok(add(toolbox, session, **{"from": "d:tall", "header": []}))      # 200 rows fit
    ok(call(toolbox, session, "ds_put", id="wide", size="2x51"))
    error = failed(add(toolbox, session, **{"from": "d:wide"}), "limit")
    assert error.field == "from_columns" and "51 columns" in error.message
    ok(call(toolbox, session, "ds_put", id="big", size="120x45"))
    error = failed(add(toolbox, session, **{"from": "d:big"}), "limit")
    assert "5,000" in error.message
    assert len(session.entry("d1").document.slides[0].shapes) == 1     # the 200-row table


def test_a_value_a_cell_cannot_take_is_refused_naming_it(toolbox, session):
    @tool("ds_odd", "A dataset with a value that is not finite.", {}, documents=())
    def ds_odd(call):
        call.define_dataset("odd", Dataset(("a",), [(1,), (float("nan"),)],
                                           sources=[[("G!A2",)], [("G!A3",)]]), source=None)
        return Result(summary="ok")

    with Toolbox([*TOOLS, ds_odd], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session(clock=CLOCK)
        deck = Document.new()
        deck.add_slide("Blank")
        session.open(deck.to_bytes(), "deck.pptx")
        ok(box.dispatch(session, "ds_odd", {}))
        error = failed(box.dispatch(session, "ppt_add_table", {"doc": "d1", "slide": "s:256",
                                                               "box": BOX, "from": "d:odd"}),
                       "invalid_arguments")
        assert "G!A3" in error.message and "not a finite number" in error.message


# -- ppt_edit_table set_cells --------------------------------------------------------------------


def _table(toolbox, session, rows, columns):
    return ok(add(toolbox, session, rows=rows, columns=columns,
                  data=[["old"] * columns] * rows)).created[0]


def test_set_cells_from_a_source_rewrites_the_block(toolbox, session):
    target = _table(toolbox, session, 4, 2)
    result = ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                     action="set_cells", **{"from": "d:scores",
                                            "from_columns": ["check", "note"]}))
    assert result.data["cells"] == [["check", "note"], ["Alpha", "fine"], ["Beta", ""],
                                    ["Gamma", "late"]]
    assert result.data["from"]["first_row"] == 1 and not result.warnings
    assert f"{target}/cell2,1" in result.changed


def test_set_cells_from_row_index_under_a_kept_header(toolbox, session):
    target = _table(toolbox, session, 5, 3)
    result = ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                     action="set_cells", index=1, header=[],
                     **{"from": "d:scores", "from_columns": ["check", "rating"]}))
    cells = result.data["cells"]
    assert cells[0] == ["old"] * 3 and cells[1] == ["Alpha", "3.5", "old"]
    assert cells[4] == ["old"] * 3
    placed = result.data["from"]
    assert placed["header_row"] is None and placed["first_row"] == 1
    assert result.warnings and "keep their text" in result.warnings[0]


def test_set_cells_from_a_source_too_big_for_the_table_is_refused(toolbox, session):
    target = _table(toolbox, session, 3, 2)
    error = failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                        action="set_cells", **{"from": "d:scores",
                                               "from_columns": ["check"]}), "refused")
    assert "4 rows" in error.message and "insert_row" in error.message
    error = failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                        action="set_cells", header=[], **{"from": "d:scores"}), "refused")
    assert error.field == "from_columns"
    error = failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                        action="merge", **{"from": "d:scores"}), "invalid_arguments")
    assert error.field == "from"
    error = failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                        action="set_cells", cells=[], **{"from": "d:scores"}),
                   "invalid_arguments")
    assert error.field == "cells"
    error = failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                        action="format_cells", fill="accent1", header=[]),
                   "invalid_arguments")
    assert error.field == "header"
    assert session.entry("d1").document.shape(target).table.cell(0, 0).text == "old"


def test_set_cells_from_a_source_is_one_undo_step(toolbox, session):
    target = _table(toolbox, session, 4, 1)
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target, action="set_cells",
            **{"from": "t:checks", "from_columns": ["check"]}))
    ok(call(toolbox, session, "undo", doc="d1"))
    assert session.entry("d1").document.shape(target).table.cell(1, 0).text == "old"


def test_column_names_given_as_columns_are_read_as_from_columns(toolbox, session):
    result = ok(add(toolbox, session, **{"from": "d:scores", "columns": ["rating", "check"]}))
    assert result.data["cells"][0] == ["rating", "check"]
    assert "columns read as from_columns" in result.warnings
    # A count stays a count.
    counted = ok(add(toolbox, session, rows=2, columns=3))
    assert counted.data["columns"] == 3 and not counted.warnings


# -- dates (T18b) --------------------------------------------------------------------------------

DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")
#: The checks' due dates as each locale writes them.
DUE = {"en-US": ["1/5/2026", "2/10/2026", "3/15/2026"],
       "nl-BE": ["5/1/2026", "10/2/2026", "15/3/2026"]}


def _dated(toolbox, locale):
    session = _session(toolbox, locale=locale)
    ok(call(toolbox, session, "ds_put", id="dated", typed="yes"))
    return session


@pytest.mark.parametrize("locale, due", [(None, DUE["en-US"]), (DUTCH, DUE["nl-BE"])])
def test_a_date_columns_iso_text_is_written_in_the_sessions_locale(toolbox, locale, due):
    session = _dated(toolbox, locale)
    result = ok(add(toolbox, session, **{"from": "d:dated", "from_columns": ["check", "due"]}))
    assert [row[1] for row in result.data["cells"][1:]] == due
    deck = session.entry("d1").document
    assert deck.shape(result.created[0]).table.cell(1, 1).text == due[0]
    # The record names the cells, never their values.
    kept = deck.provenance()[result.created[0]]
    assert kept["columns"][1]["sources"] == ["G!C2", "G!C3", "G!C4"]
    assert not any(text in json.dumps(kept) for text in due + ["2026-01-05"])
    # set_cells writes the same.
    target = _table(toolbox, session, 4, 2)
    edited = ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                     action="set_cells", **{"from": "d:dated",
                                            "from_columns": ["check", "due"]}))
    assert [row[1] for row in edited.data["cells"][1:]] == due


def test_an_unlabelled_iso_column_is_written_as_it_is(toolbox):
    session = _dated(toolbox, DUTCH)
    data = ok(add(toolbox, session, **{"from": "t:checks",
                                       "from_columns": ["due"]})).data
    assert data["cells"][1] == ["2026-01-05"]           # no type: text


CHART = {"x": 40, "y": 300, "w": 300, "h": 200}


@pytest.mark.parametrize("locale, key", [(None, "en-US"), (DUTCH, "nl-BE")])
def test_a_chart_from_a_date_column_and_set_values_matched_either_way(toolbox, locale, key):
    session = _dated(toolbox, locale)
    made = ok(call(toolbox, session, "edit_chart", doc="d1", action="add", target="s:256",
                   box=CHART, chart_type="column",
                   **{"from": "d:dated", "columns": ["due", "rating"]})).created[0]
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    assert read.data["categories"] == DUE[key]
    # An older deck's chart, its categories written ISO and in another order.
    older = ok(call(toolbox, session, "edit_chart", doc="d1", action="add", target="s:256",
                    box=CHART, chart_type="column",
                    categories=["2026-03-15", "2026-01-05", "2026-02-10"],
                    data=[{"name": "rating", "values": [0, 0, 0]}])).created[0]
    for target, order in ((made, [3.5, 4, 1.25]), (older, [1.25, 3.5, 4])):
        ok(call(toolbox, session, "edit_chart", doc="d1", target=target, action="set_values",
                series="rating", **{"from": "d:dated", "columns": ["due", "rating"]}))
        read = ok(call(toolbox, session, "edit_chart", doc="d1", target=target,
                       action="read"))
        assert read.data["series"][0]["values"] == order


# -- date_format (round 6, R9) ------------------------------------------------------------------


def _locale(language):
    return Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                  language=language, list_separator=";")


#: The checks' due dates (5 January, 10 February, 15 March 2026) in ``d mmmm yyyy``, with
#: the month names measured in Excel for Mac, 2026-10-16.
@pytest.mark.parametrize("language, due", [
    ("en-US", ["5 January 2026", "10 February 2026", "15 March 2026"]),
    ("nl-BE", ["5 januari 2026", "10 februari 2026", "15 maart 2026"]),
    ("de-DE", ["5 Januar 2026", "10 Februar 2026", "15 März 2026"]),
    ("fr-FR", ["5 janvier 2026", "10 février 2026", "15 mars 2026"]),
])
def test_a_tables_dates_in_a_date_format(toolbox, language, due):
    session = _dated(toolbox, _locale(language))
    result = ok(add(toolbox, session, date_format="d mmmm yyyy",
                    **{"from": "d:dated", "from_columns": ["check", "due", "rating"]}))
    assert [row[1] for row in result.data["cells"][1:]] == due
    assert result.data["cells"][1][2] == "3,5"                   # a number as before
    target = _table(toolbox, session, 4, 2)
    edited = ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                     action="set_cells", date_format="dd/mm/yyyy",
                     **{"from": "d:dated", "from_columns": ["check", "due"]}))
    assert [row[1] for row in edited.data["cells"][1:]] == ["05/01/2026", "10/02/2026",
                                                            "15/03/2026"]


def test_a_date_format_is_checked_for_a_table(toolbox):
    session = _dated(toolbox, _locale("en-AU"))
    error = failed(add(toolbox, session, date_format="d mmmm yyyy",
                       **{"from": "d:dated", "from_columns": ["due"]}), "invalid_arguments")
    assert "no month or day names are measured for en-AU" in error.message
    assert error.field == "date_format"
    ok(add(toolbox, session, date_format="dd-mm-yyyy",
           **{"from": "d:dated", "from_columns": ["due"]}))
    error = failed(add(toolbox, session, date_format="dd/mm/yyyy",
                       **{"from": "d:dated", "from_columns": ["check", "rating"]}),
                   "invalid_arguments")
    assert "date_format is for dates" in error.message
    error = failed(add(toolbox, session, rows=2, columns=2, date_format="dd/mm/yyyy"),
                   "invalid_arguments")
    assert "date_format is for from" in error.message
    error = failed(add(toolbox, session, date_format="hh:mm",
                       **{"from": "d:dated", "from_columns": ["due"]}), "invalid_arguments")
    assert "'h' is no date token" in error.message


@pytest.mark.parametrize("language, code, first", [
    ("en-US", r"[$-409]d\ mmm\ yyyy", 46027), ("nl-BE", r"[$-813]d\ mmm\ yyyy", 46027),
    ("it-IT", r"[$-410]d\ mmm\ yyyy", 46027),
])
def test_a_charts_dates_are_kept_in_the_locales_code(toolbox, language, code, first):
    from lxml import etree

    session = _dated(toolbox, _locale(language))
    made = ok(call(toolbox, session, "edit_chart", doc="d1", action="add", target="s:256",
                   box=CHART, chart_type="line", date_format="d mmm yyyy",
                   **{"from": "d:dated", "columns": ["due", "rating"]})).created[0]
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    assert read.data["categories"] == [first, first + 36, first + 69]
    deck = session.entry("d1").document
    chart = deck.shape(made).chart
    root = etree.fromstring(deck.package.read(chart.part))
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    assert root.xpath("string(//c:cat//c:formatCode)", namespaces=ns) == code
    assert root.xpath("string(//c:catAx/c:numFmt/@formatCode)", namespaces=ns) == code
    assert root.xpath("string(//c:catAx/c:auto/@val)", namespaces=ns) == "0"
    assert chart.workbook_values()["categories"] == [first, first + 36, first + 69]
    # set_values from the same rows matches the dates kept as serials.
    ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="set_values",
            series="rating", **{"from": "d:dated", "columns": ["due", "rating"]}))
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    assert read.data["series"][0]["values"] == [3.5, 4, 1.25]
