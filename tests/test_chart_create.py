"""slide.add_chart (LP7) and edit_chart's add, data-label and gap-width actions (LP8).

The chart part and its workbook are ooxml-edit's (``ooxml_edit.charts.create``, tested
there); here: the frame PowerPoint writes, one undo step, the deck staying valid, the
workbook agreeing with the cache through the library's own reader, and the tool calls.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from pptx_agent import ChartDataError, Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

EMU = 12700
REGIONS = [{"name": "North", "values": [12.4, 13.1, 14.0, 15.2]},
           {"name": "South", "values": [9.8, 10.2, 10.9, 11.9]}]


def deck_with_slide():
    deck = Document.new()
    slide = deck.add_slide(deck.layouts[5])
    return deck, slide


@pytest.mark.parametrize("kind", ["column", "stacked_column", "bar", "stacked_bar", "line",
                                  "pie", "scatter", "radar"])
def test_a_new_chart_on_a_slide(kind):
    deck, slide = deck_with_slide()
    categories = [1, 2, 3, 4] if kind == "scatter" else ["Q1", "Q2", "Q3", "Q4"]
    series = REGIONS[:1] if kind == "pie" else REGIONS
    frame = slide.add_chart(kind, categories, series, 40 * EMU, 120 * EMU, 880 * EMU, 380 * EMU,
                            title="Revenue", number_format="#,##0.0")
    assert frame.kind == "graphic_frame" and frame.name.startswith("Chart ")
    chart = deck.shape(frame.id).chart
    assert chart.categories == categories and chart.title == "Revenue"
    book = chart.workbook_values()
    assert book["categories"] == categories
    assert book["series"] == [{"name": s["name"], "values": s["values"]} for s in series]
    assert deck.validate() == []
    reopened = Document.open(deck.to_bytes())
    assert reopened.shape(frame.id).chart.workbook_values() == book


def test_adding_a_chart_is_one_undo_step():
    deck, slide = deck_with_slide()
    before = deck.to_bytes()
    slide.add_chart("column", ["A"], [{"name": "S", "values": [1]}], 0, 0, 100 * EMU, 100 * EMU)
    deck.undo()
    assert deck.to_bytes() == before


def test_bad_data_is_refused_before_anything_changes():
    deck, slide = deck_with_slide()
    before = deck.to_bytes()
    with pytest.raises(ChartDataError):
        slide.add_chart("column", ["A", "B"], [{"name": "S", "values": [1]}], 0, 0, EMU, EMU)
    assert deck.to_bytes() == before


def test_the_frame_is_powerpoint_s():
    deck, slide = deck_with_slide()
    frame = slide.add_chart("line", ["A", "B"], [{"name": "S", "values": [1, 2]}],
                            EMU, 2 * EMU, 300 * EMU, 200 * EMU)
    element = frame._element
    assert element.tag.endswith("graphicFrame")
    assert (frame.left, frame.top, frame.width, frame.height) == (EMU, 2 * EMU, 300 * EMU,
                                                                  200 * EMU)
    assert element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart") \
        is not None


# -- the tool -------------------------------------------------------------------------------------


@pytest.fixture()
def session():
    deck, _ = deck_with_slide()
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        current = box.session()
        current.open(deck.to_bytes(), "deck.pptx")
        yield box, current


ADD = {"doc": "d1", "target": "s:256", "action": "add", "chart_type": "column",
       "categories": ["Q1", "Q2", "Q3", "Q4"], "data": REGIONS,
       "box": {"x": 40, "y": 120, "w": 880, "h": 380}, "text": "North leads", "ref": "c"}


def test_edit_chart_adds_and_formats_a_chart(session):
    box, current = session
    added = box.dispatch(current, "edit_chart", ADD)
    assert added.ok, added.to_json()
    assert added.created == ["256.3"] and added.refs == {"c": "256.3"}
    assert added.data["gap_width"] == 219
    labels = box.dispatch(current, "edit_chart", {"doc": "d1", "target": "$c",
                                                   "action": "show_data_labels",
                                                   "number_format": '"€"#,##0.0"m"'})
    assert labels.ok and labels.data["number_formats"]["data_labels"] == '"€"#,##0.0"m"'
    gap = box.dispatch(current, "edit_chart", {"doc": "d1", "target": "$c",
                                                "action": "set_gap_width", "value": 60})
    assert gap.ok and gap.data["gap_width"] == 60
    hidden = box.dispatch(current, "edit_chart", {"doc": "d1", "target": "$c",
                                                   "action": "hide_data_labels"})
    assert hidden.ok and hidden.data["number_formats"]["data_labels"] is None
    read = box.dispatch(current, "edit_chart", {"doc": "d1", "target": "256.3", "action": "read"})
    assert read.data["workbook"]["series"] == REGIONS


@pytest.mark.parametrize("change, field", [
    ({"box": None}, "box"),
    ({"width": 300}, "width"),
    ({"categories": None}, "categories"),
    ({"chart_type": "scatter"}, "values"),
])
def test_edit_chart_add_names_what_is_missing(session, change, field):
    box, current = session
    arguments = {key: value for key, value in {**ADD, **change}.items() if value is not None}
    result = box.dispatch(current, "edit_chart", arguments)
    assert not result.ok and result.error.code == "invalid_arguments"
    assert result.error.field == field


def test_edit_chart_adds_a_radar_with_powerpoint_s_legend_at_the_top(session):
    box, current = session
    added = box.dispatch(current, "edit_chart", {**ADD, "chart_type": "radar"})
    assert added.ok, added.to_json()
    chart = current.entry("d1").document.shape("256.3").chart
    assert chart.chart_type == "radar" and "gap_width" not in added.data
    legend = chart._root().find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}"
                                "legendPos")
    assert legend.get("val") == "t"
    moved = box.dispatch(current, "edit_chart", {**ADD, "chart_type": "radar", "ref": "r",
                                                  "position": "bottom"})
    assert moved.ok
    assert moved.data["type"] == "radar"


def test_a_gap_width_on_a_line_chart_is_refused(session):
    box, current = session
    box.dispatch(current, "edit_chart", {**ADD, "chart_type": "line"})
    result = box.dispatch(current, "edit_chart", {"doc": "d1", "target": "$c",
                                                   "action": "set_gap_width", "value": 50})
    assert not result.ok and "bar and column" in result.error.message
