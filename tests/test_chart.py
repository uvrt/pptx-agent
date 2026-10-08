"""Charts (ROADMAP E4): data, titles and legend, with the embedded workbook kept in step.

Every edit runs on every chart in the corpus and goes through the same gates:

* **round trip** -- edit, save, reopen, read back;
* **the workbook agrees** -- the saved deck's embedded workbook is opened independently
  (``tests/xlsx.py``, not the library's own reader) and every formula's cells must hold what
  its cache holds, tables must cover the data and be named after their headers, and the
  shared-string counts must be right;
* **undo** -- back to the original bytes, workbook included, and redo to the edited ones;
* **validity** -- the deck-wide checks, plus schema order in the chart part.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile

import pytest
from lxml import etree

from conftest import FIXTURE_DIR
from pptx_agent import Document
from pptx_agent.edit.chart import ChartDataError, ChartDataWarning
from test_validity import assert_valid
from xlsx import Book, check_chart_against_workbook

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _chart_cases() -> list[tuple[str, str]]:
    cases = []
    for path in sorted(FIXTURE_DIR.glob("*.pptx")):
        document = Document.open(str(path))
        for slide in document.slides:
            for shape in slide.shapes:
                if shape.has_chart:
                    cases.append((path.name, shape.id))
    return cases


CASES = _chart_cases()


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()
                if not info.is_dir()}


def check_deck_charts(data: bytes, original: bytes | None = None) -> int:
    """Every chart in a saved deck agrees with its embedded workbook; returns formulas seen.
    A workbook still byte-identical to ``original``'s is held to its generator's standard."""
    document = Document.open(data)
    before = _parts(original) if original is not None else {}
    checked = 0
    for slide in document.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart = shape.chart
            workbook = chart.workbook_part
            data_ = None if workbook is None else document.package.read(workbook)
            checked += check_chart_against_workbook(
                document.package.read(chart.part), data_,
                edited=workbook is None or before.get(workbook) != data_)
            # A generator's own misordering (one fixture's line series) is not the edit's.
            assert chart_order_violations(document.package.read(chart.part)) <= \
                chart_order_violations(before.get(chart.part))
    return checked


def chart_order_violations(data: bytes | None) -> set[tuple[str, str]]:
    """``(parent, child)`` pairs out of their schema order in a chart part."""
    from pptx_agent.core.xml import _ranks, prefixed_name

    found = set()
    if data is None:
        return found
    for parent in etree.fromstring(data).iter():
        ranks = _ranks(prefixed_name(parent))
        if not ranks:
            continue
        last = -1
        for child in parent:
            rank = ranks.get(prefixed_name(child), len(ranks) + 1)
            if rank < last:
                found.add((prefixed_name(parent), prefixed_name(child)))
            last = max(last, rank)
    return found


def test_the_corpus_has_charts():
    assert len(CASES) >= 6


def test_reading(financial_report):
    chart = Document.open(str(financial_report)).shape("257.29").chart
    assert chart.chart_types == ["line"]
    assert chart.categories == ["Q1", "Q2", "Q3"]
    assert [s.name for s in chart.series] == ["営業利益率（%）", "売上総利益率（%）", "純利益率（%）"]
    assert chart.series[1].values == [42, 42.5, 43]
    assert chart.title is None and chart.has_legend
    assert chart.axes == ["category", "value"] and chart.axis_title("value") is None
    assert chart.number_format == "General"
    assert chart.workbook_part == "ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx"


def test_reading_a_titled_chart():
    chart = Document.open(str(FIXTURE_DIR / "authoring-integration.pptx")).shape("256.6").chart
    assert chart.title == "Chart contract"
    assert chart.data == {
        "types": ["bar"], "title": "Chart contract",
        "axis_titles": {"category": None, "value": None}, "legend": True, "format": "General",
        "categories": ["Reader", "Writer", "Renderer"],
        "series": [{"name": "Coverage", "values": [3, 4, 5]}],
    }


def test_every_original_chart_agrees_with_its_workbook(pptx_path):
    """The gate's own baseline: the fixtures' charts are consistent before any edit."""
    data = pptx_path.read_bytes()
    check_deck_charts(data, data)


# ------------------------------------------------------------------------------------------
# Each edit, through every gate
# ------------------------------------------------------------------------------------------


def _edit_value(chart):
    series = chart.series[-1]
    index = len(series.values) - 1
    series.set_value(index, 1234.5)
    return lambda c: c.series[-1].values[index] == 1234.5


def _edit_blank(chart):
    chart.series[0].set_value(0, None)
    return lambda c: c.series[0].values[0] is None


def _edit_values(chart):
    count = len(chart.categories) or chart.point_count
    wanted = [float(10 * (i + 1)) + 0.25 for i in range(count)]
    chart.series[0].set_values(wanted)
    return lambda c: c.series[0].values == wanted


def _edit_category(chart):
    chart.set_category(0, "E4 label")
    return lambda c: c.categories[0] == "E4 label"


def _edit_name(chart):
    chart.series[0].name = "E4 series"
    return lambda c: c.series[0].name == "E4 series"


def _add_category_last(chart):
    count = len(chart.series)
    chart.add_category("E4 new", list(range(7, 7 + count)))
    return lambda c: c.categories[-1] == "E4 new" and c.series[-1].values[-1] == 6 + count


def _add_category_first(chart):
    before = chart.categories
    chart.add_category("E4 first", {chart.series[0].name: 99}, index=0)
    return lambda c: c.categories == ["E4 first"] + before and c.series[0].values[0] == 99 \
        and all(s.values[0] is None for s in c.series[1:])


def _remove_category(chart):
    before = chart.categories
    values = chart.series[0].values
    chart.remove_category(1)
    return lambda c: c.categories == before[:1] + before[2:] \
        and c.series[0].values == values[:1] + values[2:]


def _add_series(chart):
    count = chart.point_count
    names = [s.name for s in chart.series]
    chart.add_series("E4 added", [float(i) + 0.5 for i in range(count)])
    return lambda c: [s.name for s in c.series] == names + ["E4 added"] \
        and c.series[-1].values == [float(i) + 0.5 for i in range(count)]


def _add_series_first(chart):
    names = [s.name for s in chart.series]
    chart.add_series("E4 front", [1] * chart.point_count, index=0)
    return lambda c: [s.name for s in c.series] == ["E4 front"] + names


def _remove_series(chart):
    if len(chart.series) < 2:
        chart.add_series("E4 spare", [1] * chart.point_count)
    names = [s.name for s in chart.series]
    values = [s.values for s in chart.series]
    chart.remove_series(0)
    return lambda c: [s.name for s in c.series] == names[1:] \
        and [s.values for s in c.series] == values[1:]


def _titles(chart):
    chart.set_title("E4 title")
    axes = chart.axes
    for axis in axes:
        chart.set_axis_title(axis, f"E4 {axis}")
    return lambda c: c.title == "E4 title" and all(c.axis_title(a) == f"E4 {a}" for a in axes)


def _legend(chart):
    shown = chart.has_legend
    chart.set_legend(not shown)
    return lambda c: c.has_legend is (not shown)


EDITS = {
    "value": _edit_value, "blank": _edit_blank, "values": _edit_values,
    "category": _edit_category, "name": _edit_name, "add_category_last": _add_category_last,
    "add_category_first": _add_category_first, "remove_category": _remove_category,
    "add_series": _add_series, "add_series_first": _add_series_first,
    "remove_series": _remove_series, "titles": _titles, "legend": _legend,
}


@pytest.mark.parametrize("edit", list(EDITS))
@pytest.mark.parametrize("deck, shape", CASES)
def test_each_edit_round_trips_with_the_workbook_in_step(deck, shape, edit):
    original = (FIXTURE_DIR / deck).read_bytes()
    document = Document.open(original)
    with warnings.catch_warnings():
        warnings.simplefilter("error", ChartDataWarning)  # every fixture chart has a workbook
        holds = EDITS[edit](document.shape(shape).chart)
    edited = document.to_bytes()

    reopened = Document.open(edited)
    assert holds(reopened.shape(shape).chart), reopened.shape(shape).chart.data
    assert check_deck_charts(edited, original) > 0
    assert_valid(edited, original)

    while document.undo():
        pass
    assert _parts(document.to_bytes()) == _parts(original)
    while document.redo():
        pass
    assert document.to_bytes() == edited


@pytest.mark.parametrize("deck, shape", CASES)
def test_all_edits_together(deck, shape):
    original = (FIXTURE_DIR / deck).read_bytes()
    document = Document.open(original)
    for edit in EDITS.values():
        holds = edit(document.shape(shape).chart)
    edited = document.to_bytes()
    check_deck_charts(edited, original)
    assert_valid(edited, original)
    assert holds(Document.open(edited).shape(shape).chart)  # the last edit, at least
    while document.undo():
        pass
    assert _parts(document.to_bytes()) == _parts(original)


def test_an_unchanged_value_changes_nothing(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    chart.series[1].set_value(1, 42.5)
    chart.set_category(0, "Q1")
    chart.series[0].name = "営業利益率（%）"
    chart.set_title(None)
    assert not document.history.can_undo()
    assert document.package.dirty_parts == frozenset()


# ------------------------------------------------------------------------------------------
# The workbook, specifically
# ------------------------------------------------------------------------------------------


def _book(document: Document, shape: str) -> Book:
    chart = document.shape(shape).chart
    return Book(document.package.read(chart.workbook_part))


def test_a_value_reaches_its_cell(financial_report):
    document = Document.open(str(financial_report))
    document.shape("257.29").chart.series[2].set_value(1, 8.25)
    assert _book(document, "257.29").sheet("Sheet1").value("D3") == 8.25


def test_shared_strings_are_reused_and_counted(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    chart.set_category(2, "Q1")  # a string the table already has
    book = _book(document, "257.29")
    assert book.strings.count("Q1") == 1
    assert book.sheet("Sheet1").value("A4") == "Q1"
    assert int(book.string_counts[1]) == len(book.strings)


def test_a_workbook_without_shared_strings_gets_inline_strings():
    document = Document.open(str(FIXTURE_DIR / "authoring-integration.pptx"))
    document.shape("256.6").chart.set_category(1, "Inline")
    book = _book(document, "256.6")
    assert book.strings == []
    assert book.sheet("Sheet1").value("A3") == "Inline"


def test_a_table_follows_the_data(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    chart.add_series("Extra", [1, 2, 3])
    chart.add_category("Q4", [1, 2, 3, 4])
    table = _book(document, "257.29").sheet("Sheet1").tables[0]
    X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    assert table.get("ref") == "A1:E5"
    names = [c.get("name") for c in table.iter(X + "tableColumn")]
    assert names[1:] == ["営業利益率（%）", "売上総利益率（%）", "純利益率（%）", "Extra"]
    chart.remove_series(1)
    chart.remove_category(0)
    table = _book(document, "257.29").sheet("Sheet1").tables[0]
    assert table.get("ref") == "A1:D4"
    assert [c.get("name") for c in table.iter(X + "tableColumn")][1:] == [
        "営業利益率（%）", "純利益率（%）", "Extra"]


def test_formulas_follow_the_data(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    chart.remove_series(0)
    chart.add_category("Q4", [1, 2])
    formulas = re.findall(r"<c:f>([^<]*)</c:f>", document.package.read(chart.part).decode())
    assert formulas == ["Sheet1!$B$1", "Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5",
                        "Sheet1!$C$1", "Sheet1!$A$2:$A$5", "Sheet1!$C$2:$C$5"]
    sheet = _book(document, "257.29").sheet("Sheet1")
    assert sheet.range("B1", "C1") == ["売上総利益率（%）", "純利益率（%）"]
    assert sheet.value("D1") is None and sheet.value("D2") is None


def test_a_point_moves_its_formatting_with_it(financial_report):
    """``c:dPt`` and ``c:dLbl`` follow their point when a category is inserted before it."""
    document = Document.open(str(financial_report))
    chart = document.shape("258.13").chart  # the doughnut, coloured point by point
    before = re.findall(rb'<c:dPt><c:idx val="(\d+)"', document.package.read(chart.part))
    chart.add_category("New", [5], index=0)
    after = re.findall(rb'<c:dPt><c:idx val="(\d+)"', document.package.read(chart.part))
    assert [int(v) for v in after] == [int(v) + 1 for v in before]


def test_a_formula_a_value_replaces_goes_with_the_calculation_chain():
    """A cell computed by a formula becomes a plain value, and the calculation chain that
    named it is removed (Excel rebuilds it; a stale one is a repair)."""
    document = Document.open(str(FIXTURE_DIR / "authoring-integration.pptx"))
    chart = document.shape("256.6").chart
    package = document.package
    part = chart.workbook_part
    nested = package.open_embedded(part)
    sheet = nested.tree("xl/worksheets/sheet1.xml")
    X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    cell = next(c for c in sheet.iter(X + "c") if c.get("r") == "B3")
    formula = etree.Element(X + "f")
    formula.text = "2+2"
    cell.insert(0, formula)
    nested.mark_dirty("xl/worksheets/sheet1.xml")
    nested.add_part("xl/calcChain.xml",
                    b'<calcChain xmlns="%s"><c r="B3" i="1"/></calcChain>' % X[1:-1].encode(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                    override=True)
    nested.add_relationship("xl/workbook.xml", "http://schemas.openxmlformats.org/"
                            "officeDocument/2006/relationships/calcChain", "xl/calcChain.xml")
    package.replace_embedded(part, nested)

    chart.series[0].set_value(1, 40)
    book = _book(document, "256.6")
    assert "xl/calcChain.xml" not in book.archive.namelist()
    assert b"<f>" not in book.archive.read("xl/worksheets/sheet1.xml")
    assert book.sheet("Sheet1").value("B3") == 40
    assert "calcChain" not in book.archive.read("[Content_Types].xml").decode()


def test_untouched_workbook_parts_keep_their_bytes(financial_report):
    original = Book(Document.open(str(financial_report)).package.read(
        "ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx"))
    document = Document.open(str(financial_report))
    document.shape("257.29").chart.series[0].set_value(0, 1)
    edited = _book(document, "257.29")
    # The sheet changed; so did the table on it, whose generator wrote its ref as "A1:D4'"
    # -- touching a table writes that back clean.
    changed = {"xl/worksheets/sheet1.xml", "xl/tables/table1.xml"}
    for name in original.archive.namelist():
        if name not in changed and not name.endswith("/"):
            assert edited.archive.read(name) == original.archive.read(name), name
    assert b'ref="A1:D4"' in edited.archive.read("xl/tables/table1.xml")
    assert [i.compress_type for i in edited.archive.infolist()] == [
        i.compress_type for i in original.archive.infolist() if not i.is_dir()]


# ------------------------------------------------------------------------------------------
# Without a workbook, and with one laid out unusually
# ------------------------------------------------------------------------------------------


def _unlink_workbook(document: Document, shape: str, *, external: bool) -> None:
    chart = document.shape(shape).chart
    rels = "ppt/charts/_rels/" + chart.part.rsplit("/", 1)[1] + ".rels"
    root = document.package.tree(rels)
    node = root[0]
    if external:
        node.set("TargetMode", "External")
        node.set("Target", "file:///elsewhere/book.xlsx")
        node.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                         "oleObject")
    document.package.mark_dirty(rels)


@pytest.mark.parametrize("external", [True], ids=["linked"])
def test_a_linked_workbook_means_cache_only_and_a_warning(financial_report, external):
    document = Document.open(str(financial_report))
    _unlink_workbook(document, "257.29", external=external)
    chart = document.shape("257.29").chart
    assert chart.workbook_part is None
    embedded = document.package.read("ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx")
    with pytest.warns(ChartDataWarning, match="linked from outside"):
        chart.series[0].set_value(0, 77)
    assert chart.series[0].values[0] == 77
    with pytest.warns(ChartDataWarning):
        added = chart.add_series("Cache only", [1, 2, 3])
    assert added.values == [1, 2, 3]
    # The new series cannot point at cells nobody wrote: its data is literal.
    assert b"<c:numLit>" in document.package.read(chart.part)
    assert document.package.read("ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx") == embedded
    with pytest.warns(ChartDataWarning):
        chart.add_category("Q4", [1, 2, 3, 4])
    assert chart.series[-1].values == [1, 2, 3, 4]


def test_a_layout_insertion_cannot_follow_is_refused(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    root = document.package.tree(chart.part)
    formulas = [f for f in root.iter(C + "f") if f.text == "Sheet1!$C$2:$C$4"]
    formulas[0].text = "Sheet1!$C$3:$C$5"  # one series a row lower than the rest
    document.package.mark_dirty(chart.part)
    before = document.to_bytes()
    with pytest.raises(ChartDataError, match="do not line up"):
        chart.add_category("Q4")
    assert document.to_bytes() == before


def test_a_cell_in_the_way_is_refused(financial_report):
    document = Document.open(str(financial_report))
    chart = document.shape("257.29").chart
    package = document.package
    nested = package.open_embedded(chart.workbook_part)
    from pptx_agent.edit.workbook import Cell, Workbook

    book = Workbook(nested)
    book.sheet("Sheet1").set_value(Cell(5, 2), "note")
    package.replace_embedded(chart.workbook_part, nested)
    before = document.to_bytes()
    with pytest.raises(ChartDataError, match="B5 is in the way"):
        chart.add_category("Q4")
    assert document.to_bytes() == before


def test_bad_values_are_refused(financial_report):
    chart = Document.open(str(financial_report)).shape("257.29").chart
    for bad in (float("nan"), float("inf"), "12", True):
        with pytest.raises(ChartDataError):
            chart.series[0].set_value(0, bad)
    with pytest.raises(ChartDataError):
        chart.series[0].set_values([1, 2])
    with pytest.raises(ChartDataError):
        chart.add_series("Short", [1])
    with pytest.raises(IndexError):
        chart.series[0].set_value(3, 1)


def test_the_last_series_and_category_stay(financial_report):
    chart = Document.open(str(financial_report)).shape("258.13").chart
    with pytest.raises(ChartDataError):
        chart.remove_series(0)
    for _ in range(3):
        chart.remove_category(0)
    with pytest.raises(ChartDataError):
        chart.remove_category(0)


# ------------------------------------------------------------------------------------------
# What pptx2svg draws
# ------------------------------------------------------------------------------------------


def _numbers(svg: str) -> list[float]:
    found = []
    for text in re.findall(r">([^<>]+)</tspan>", svg):
        try:
            found.append(float(text.replace(",", "")))
        except ValueError:
            pass
    return found


def test_pptx2svg_draws_the_edited_chart():
    pytest.importorskip("pptx2svg")
    document = Document.open(str(FIXTURE_DIR / "authoring-integration.pptx"))
    chart = document.shape("256.6").chart
    assert max(_numbers(document.slides[0].render_svg())) < 10
    chart.series[0].set_values([300, 400, 520])
    chart.set_category(2, "Painter")
    chart.set_title("Edited chart")
    chart.add_series("Second", [100, 100, 100])
    svg = document.slides[0].render_svg()
    assert "Painter" in svg and "Edited chart" in svg and "Renderer" not in svg
    assert max(_numbers(svg)) >= 520  # the value axis now reaches the new maximum
    # The first series keeps its colour: three bars and its legend key.
    assert svg.count('fill="#f97316"') == 4
