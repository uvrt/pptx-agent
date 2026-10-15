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


# -- add with its look (round 5, T2) -------------------------------------------------------------


LOOK = {"data_labels": True, "colors": ["accent2", "1F4E79"], "style": {"size": 9},
        "number_format": "0.0", "axis": "value", "scale": {"min": 0, "max": 20},
        "gridlines": "none"}


def _chart_xml(current, address):
    from lxml import etree

    return etree.tostring(current.entry("d1").document.shape(address).chart._root())


def _two_sessions():
    deck, _ = deck_with_slide()
    data = deck.to_bytes()
    box = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    sessions = []
    for _ in range(2):
        current = box.session()
        current.open(data, "deck.pptx")
        sessions.append(current)
    return box, sessions


def test_add_with_its_look_writes_the_chart_the_separate_calls_write():
    box, (one, separate) = _two_sessions()
    with box:
        added = box.dispatch(one, "edit_chart", {**ADD, **LOOK})
        assert added.ok, added.to_json()
        plain = {**ADD, **{k: LOOK[k] for k in ("number_format", "axis", "scale", "gridlines")}}
        assert box.dispatch(separate, "edit_chart", plain).ok
        for call in ({"action": "show_data_labels"},
                     {"action": "format_series", "series": "North", "style": {"fill": "accent2"}},
                     {"action": "format_series", "series": "South", "style": {"fill": "1F4E79"}},
                     {"action": "set_text_style", "style": {"size": 9}}):
            done = box.dispatch(separate, "edit_chart", {"doc": "d1", "target": "$c", **call})
            assert done.ok, done.to_json()
        written = _chart_xml(one, "256.3")
        assert written == _chart_xml(separate, "256.3")
        for part in (b"c:dLbls", b'val="accent2"', b'val="1F4E79"', b'sz="900"'):
            assert part in written


@pytest.mark.parametrize("kind, data, extra", [
    ("line", REGIONS, {}),
    ("scatter", REGIONS, {"values": [1, 2, 3, 4]}),
    ("pie", REGIONS[:1], {}),
    ("radar_filled", REGIONS, {}),
    ("column", [REGIONS[0], {**REGIONS[1], "type": "line"}], {}),
])
def test_add_colours_each_kind_as_its_separate_calls_do(kind, data, extra):
    from ooxml_edit.charts.actions import add_style_actions

    look = {"data_labels": True, "colors": ["accent2", "accent6"], "style": {"size": 11},
            "part": "legend"}
    arguments = {**ADD, "chart_type": kind, "data": data, **extra}
    if extra:
        arguments.pop("categories")
    box, (one, separate) = _two_sessions()
    with box:
        assert box.dispatch(one, "edit_chart", {**arguments, **look}).ok
        assert box.dispatch(separate, "edit_chart", arguments).ok
        for action, given in add_style_actions(kind, data, **look):
            done = box.dispatch(separate, "edit_chart", {"doc": "d1", "target": "$c",
                                                         "action": action, **given})
            assert done.ok, done.to_json()
        written = _chart_xml(one, "256.3")
        assert written == _chart_xml(separate, "256.3")
        assert b'val="accent6"' in written and b'sz="1100"' in written


@pytest.mark.parametrize("look, words", [
    ({"colors": ["accent1", "accent2", "accent3"]}, "3 colours for 2 series"),
    ({"colors": ["no such colour"]}, "colour"),
    ({"style": {"fill": "accent1"}}, "colour series with colors"),
    ({"part": "legend"}, "give style too"),
])
def test_add_refuses_a_look_it_cannot_give_and_adds_nothing(session, look, words):
    box, current = session
    result = box.dispatch(current, "edit_chart", {**ADD, **look})
    assert not result.ok and result.error.code == "invalid_arguments", result.to_json()
    assert words in result.error.message
    slide = current.entry("d1").document.slides[0]
    assert all(shape.kind != "graphic_frame" for shape in slide.shapes)


def test_add_reads_colours_as_colors(session):
    box, current = session
    added = box.dispatch(current, "edit_chart", {**ADD, "colours": ["accent2"]})
    assert added.ok, added.to_json()
    assert b'val="accent2"' in _chart_xml(current, "256.3")
