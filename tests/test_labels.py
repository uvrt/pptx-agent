"""Addressing a table cell, a chart series and a category by label -- strictly.

A wrong index writes a plausible number into the wrong place without a word; that was the
end-to-end pilot's main correctness worry.  A label names exactly one thing, or the lookup
raises :class:`LabelError` listing what there is.
"""

from __future__ import annotations

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document, LabelError
from pptx_agent.oxml.xml import qn
from test_chart import check_deck_charts

REPORT = FIXTURE_DIR / "real-financial-report.pptx"


@pytest.fixture
def report() -> Document:
    return Document.open(str(REPORT))


# -- tables ---------------------------------------------------------------------------------


def test_a_cell_is_found_by_its_row_and_column_labels(report):
    table = report.shape("257.3#5").table
    assert table.row_labels() == ["項目", "売上高", "売上総利益", "営業利益", "経常利益", "当期純利益"]
    assert table.column_labels() == ["項目", "当期実績", "前年同期", "増減額", "増減率"]
    cell = table.cell_by_label("営業利益", "当期実績")
    assert (cell.row, cell.column) == (3, 1)
    assert cell.text == "512億円"
    cell.text = "520億円"
    assert table.cell(3, 1).text == "520億円"
    assert table.find_row("当期純利益") == 5
    assert table.find_column("増減率") == 4


def test_a_label_that_names_nothing_raises_and_lists_the_labels(report):
    table = report.shape("257.3#5").table
    with pytest.raises(LabelError) as caught:
        table.cell_by_label("営業利益", "Q5")
    message = str(caught.value)
    assert "no column 'Q5'" in message and "'当期実績'" in message and "'増減率'" in message
    with pytest.raises(KeyError):  # a LabelError is a KeyError
        table.find_row("売上")  # a prefix is not a match


def test_a_label_two_rows_share_is_ambiguous():
    document = Document.new()
    slide = document.add_slide("Title Only")
    table = slide.add_table(4, 3, 914400, 914400, 5486400, 1463040).table
    for row, label in enumerate(["", "Revenue", "Profit", "Revenue"]):
        table.cell(row, 0).text = label
    table.cell(0, 1).text, table.cell(0, 2).text = "Q1", "Q2"
    with pytest.raises(LabelError, match="names 2 rows"):
        table.cell_by_label("Revenue", "Q1")
    assert table.cell_by_label("Profit", "Q2").address.endswith("/cell2,2")


def test_whitespace_width_and_case_are_forgiven_only_when_that_is_unique():
    document = Document.new()
    slide = document.add_slide("Title Only")
    table = slide.add_table(3, 2, 914400, 914400, 5486400, 1463040).table
    table.cell(0, 1).text = "売上高（億円）"
    table.cell(1, 0).text = "Net  profit"
    table.cell(2, 0).text = "Total\vrevenue"  # a line break inside the label
    assert table.cell_by_label("Net profit", "売上高(億円)").row == 1
    assert table.find_row("total revenue") == 2
    assert table.find_row("NET PROFIT") == 1
    table.cell(2, 0).text = "net profit"
    assert table.find_row("net profit") == 2  # the exact match wins
    with pytest.raises(LabelError, match="names 2 rows"):
        table.find_row("NET PROFIT")


def test_merged_header_positions_share_their_label():
    document = Document.new()
    slide = document.add_slide("Title Only")
    table = slide.add_table(3, 3, 914400, 914400, 5486400, 1463040).table
    table.cell(0, 1).text = "FY2025"
    table.merge(0, 1, 0, 2)
    assert table.column_labels() == ["", "FY2025", "FY2025"]
    with pytest.raises(LabelError, match="names 2 columns"):
        table.find_column("FY2025")


def test_other_label_rows_and_columns(report):
    table = report.shape("258.4#3").table
    assert table.cell_by_label("合計", "売上高").text == "4,285億円"
    # The header row's own labels, read from another column.
    assert table.find_row("構成比", column=2) == 0
    assert table.cell_by_label("4,285億円", "利益率", label_column=1).text == "11.9%"


# -- charts ---------------------------------------------------------------------------------


def test_series_by_name_and_values_by_category(report):
    chart = report.shape("257.25").chart
    assert chart.series_names == ["売上高（億円）", "営業利益（億円）"]
    revenue = chart.series["売上高（億円）"]
    assert revenue.index == 0 and chart.series_named("売上高（億円）").index == 0
    assert "営業利益（億円）" in chart.series and "x" not in chart.series
    assert chart.series[1].name == "営業利益（億円）"  # positions still work
    assert chart.category_index("Q3") == 2
    assert revenue.value("Q3") == 4285 == revenue.value(2)
    assert dict(revenue.items()) == {"Q1": 3980, "Q2": 4120, "Q3": 4285}
    revenue.set_value("Q3", 4310)
    chart.series["営業利益（億円）"].set_value("Q3", 520)
    assert [s.values for s in chart.series] == [[3980, 4120, 4310], [465, 488, 520]]
    # The workbook moved with the cache.
    check_deck_charts(report.to_bytes(), REPORT.read_bytes())


def test_a_full_width_name_is_found_by_its_half_width_spelling(report):
    chart = report.shape("257.29").chart
    assert chart.series_named("営業利益率(%)").name == "営業利益率（%）"


def test_chart_labels_that_name_nothing_raise_and_list_the_labels(report):
    chart = report.shape("258.9").chart
    with pytest.raises(LabelError) as caught:
        chart.series["今期"]
    assert "'前年同期', '当期'" in str(caught.value)
    with pytest.raises(LabelError, match="the categories are 'デジタル', 'プラットフォーム'"):
        chart.series["当期"].set_value("Other", 1)
    assert chart.series["当期"].values == [1842, 1285, 814, 344]  # nothing written


def test_two_series_with_one_name_are_ambiguous(report):
    chart = report.shape("257.25").chart
    chart.series[1].name = "売上高（億円）"
    with pytest.raises(LabelError, match="names 2 series"):
        chart.series["売上高（億円）"]


def test_categories_are_relabelled_and_removed_by_label(report):
    chart = report.shape("257.25").chart
    chart.set_category("Q3", "Q3 restated")
    assert chart.categories == ["Q1", "Q2", "Q3 restated"]
    chart.remove_category("Q1")
    assert chart.categories == ["Q2", "Q3 restated"]
    added = chart.add_series("Forecast", [1, 2])
    assert added.name == "Forecast" and added.value("Q2") == 1


def test_number_formats_say_what_decides_the_precision_on_screen(report):
    chart = report.shape("257.29").chart
    assert chart.number_format == "General"
    # Nothing on the slide shows these values as text: the pilot rendered it to be sure.
    assert chart.number_formats == {"values": "General", "data_labels": None,
                                    "value_axis": "General"}
    # The doughnut's series asks for percentages, but every point turns its label off.
    assert report.shape("258.13").chart.number_formats["data_labels"] is None
    # Show the first series' values: its labels' own format, #,##0, is what decides.
    part = chart.part
    root = report.package.tree(part)
    show = root.find(f".//{qn('c:ser')}/{qn('c:dLbls')}/{qn('c:showVal')}")
    show.set("val", "1")
    report.package.mark_dirty(part)
    assert chart.number_formats["data_labels"] == "#,##0"


def test_a_label_error_carries_its_candidates_as_data():
    deck = Document.new()
    table = deck.add_slide("Blank").add_table(3, 3, 0, 0, 3657600, 914400).table
    for column, label in enumerate(["Item", "Q1", "Q2"]):
        table.cell(0, column).text = label
    for row, label in enumerate(["Item", "Revenue", "Revenue"]):
        table.cell(row, 0).text = label
    with pytest.raises(LabelError) as missing:
        table.cell_by_label("Margin", "Q1")
    assert missing.value.label == "Margin" and missing.value.what == "row"
    assert missing.value.candidates == ["Item", "Revenue", "Revenue"]
    assert missing.value.matches == []
    with pytest.raises(LabelError) as twice:
        table.cell_by_label("Revenue", "Q1")
    assert twice.value.matches == [1, 2]
    with pytest.raises(LabelError) as untitled:
        deck.slide_titled("Nothing")
    assert untitled.value.what == "title"
