"""A minimum text size per kind, grouped tables and chart text (production feedback round 6, R1).

``deck_format(min_text_size=...)`` takes one number or ``{"body": ..., "table": ...,
"chart": ...}``; ``check`` lists a table as one entry and a chart as one entry per part of
its text, measured from the chart part (``pptx_agent.edit.chart_text``).  Every deck here
is made up.
"""

from __future__ import annotations

import datetime as dt
import re

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.chart_text import chart_text
from pptx_agent.edit.effective import TEXT_KINDS, text_minimums
from pptx_agent.tools import GROUPS, TOOLS, deck_format

EMU = 12700
CLOCK = lambda: dt.datetime(2027, 1, 4, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def first_run(frame):
    return frame.paragraphs[0].runs[0]


def chart_deck(title: str | None = None):
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    slide.shapes[0].set_text("Plain title")
    frame = slide.add_chart("column", ["North", "South"],
                            [{"name": "Plan", "values": [3, 4]},
                             {"name": "Actual", "values": [2, 5]}],
                            40 * EMU, 120 * EMU, 400 * EMU, 260 * EMU, title=title)
    return deck, slide, frame


def table_deck():
    """A slide with a 3 x 3 table at 12 pt whose second row is 8, 9 and 14 pt, and a text box at 9."""
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    slide.shapes[0].set_text("Plain title")
    table = slide.add_table(3, 3, 40 * EMU, 120 * EMU, 400 * EMU, 120 * EMU).table
    for row in range(3):
        for column in range(3):
            table.cell(row, column).text = f"r{row}c{column}"
            first_run(table.cell(row, column).text_frame).size = 12
    for column, size in zip((0, 1, 2), (8, 9, 14)):
        first_run(table.cell(1, column).text_frame).size = size
    box = slide.add_textbox(40 * EMU, 300 * EMU, 300 * EMU, 40 * EMU)
    box.set_text("A footnote")
    first_run(box).size = 9
    return deck, slide, table, box


def rewrite_chart(deck, frame, change) -> None:
    part = deck.shape(frame.id).chart.part
    xml = deck.package.read(part).decode("utf-8")
    deck.package.replace_part(part, change(xml).encode("utf-8"))


def _session(deck, minimum):
    toolbox = Toolbox(TOOLS, formats=[deck_format(min_text_size=minimum)], groups=GROUPS,
                      workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


# -- the setting -----------------------------------------------------------------------------


def test_a_number_is_every_kinds_minimum_and_a_mapping_names_kinds():
    assert TEXT_KINDS == ("body", "table", "chart")
    assert text_minimums(12) == {"body": 12.0, "table": 12.0, "chart": 12.0}
    assert text_minimums({"body": 12, "table": 10}) == {"body": 12.0, "table": 10.0}
    assert text_minimums({"body": 12, "chart": None}) == {"body": 12.0}
    assert text_minimums(None) == {}


@pytest.mark.parametrize("value, words", [
    ({"notes": 10}, "no kind 'notes'; the kinds are body, table, chart"),
    ({}, "empty mapping"),
    ({"body": 0}, "between 1 and 400"),
    ({"table": "10"}, "between 1 and 400"),
    (True, "between 1 and 400"),
    (401, "between 1 and 400"),
])
def test_deck_format_refuses_an_unknown_kind_or_size(value, words):
    with pytest.raises(ValueError, match=re.escape(words)):
        deck_format(min_text_size=value)


# -- tables --------------------------------------------------------------------------------


def test_text_below_takes_a_minimum_per_kind_and_skips_kinds_left_out():
    deck, slide, table, box = table_deck()
    found = slide.text_below({"table": 10})
    assert [(t.kind, t.address, t.size, t.shape) for t in found] == [
        ("table", table.cell(1, 0).address, 8.0, table.address),
        ("table", table.cell(1, 1).address, 9.0, table.address)]
    assert [(t.kind, t.address, t.size) for t in slide.text_below({"body": 10})] == [
        ("body", box.id, 9.0)]
    # A plain number is every kind's, as before.
    assert len(slide.text_below(10)) == 3


def test_check_lists_a_table_once_with_its_cells_and_smallest_size():
    deck, slide, table, box = table_deck()
    toolbox, session, doc = _session(deck, {"body": 8, "table": 10})
    data = toolbox.dispatch(session, "check", {"doc": doc}).data
    assert data["small_text"] == [{"table": table.address, "cells": 2, "size": 8.0,
                                   "min": 10.0, "from": "the run"}]
    note = data["small_text_note"]
    assert "minimum (body 8 pt, table 10 pt)" in note and "A table is one entry" in note
    assert "A chart" not in note
    # The checks after a change group it the same way.
    changed = toolbox.dispatch(session, "ppt_format_text", {"doc": doc, "items": [
        {"target": table.cell(1, 0).address, "size": 12}]})
    assert changed.checks["small_text"] == [{"table": table.address, "cells": 1, "size": 9.0,
                                             "min": 10.0, "from": "the run"}]


def test_save_lists_the_grouped_table_as_one_unresolved_entry():
    deck, slide, table, box = table_deck()
    toolbox, session, doc = _session(deck, 10)
    saved = toolbox.dispatch(session, "save_document",
                             {"doc": doc, "name": "out.pptx", "format": "pptx"})
    left = saved.data["unresolved"]["small_text"]
    assert left == [{"table": table.address, "cells": 2, "size": 8.0, "min": 10.0,
                     "from": "the run"},
                    {"shape": box.id, "size": 9.0, "from": "the run"}]
    assert "2 small text" in saved.summary


def test_one_number_keeps_the_old_note():
    deck, slide, table, box = table_deck()
    toolbox, session, doc = _session(deck, 10)
    note = toolbox.dispatch(session, "check", {"doc": doc}).data["small_text_note"]
    assert "minimum of 10 pt" in note and "at least 10 on the shape" in note


# -- charts -------------------------------------------------------------------------------


def test_a_new_charts_text_is_read_from_its_chart_part():
    deck, slide, frame = chart_deck(title="Regions")
    found = [(t.part, t.size, t.source) for t in chart_text(frame)]
    # PowerPoint's new chart (ooxml-edit's POWERPOINT_LOOK): 18.62 pt title, 11.97 the rest.
    assert found == [
        ("title", 18.62, "the title's paragraph"),
        ("axis_labels", 11.97, "the category axis' text properties"),
        ("axis_labels", 11.97, "the value axis' text properties"),
        ("legend", 11.97, "the legend's text properties")]


def test_chart_text_falls_back_to_the_chart_then_ten_points():
    deck, slide, frame = chart_deck()
    # Only the chart space states a size: every label and the legend take it.
    rewrite_chart(deck, frame, lambda xml: re.sub(r' sz="\d+"', "", xml).replace(
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/>",
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/>'))
    assert {(t.part, t.size, t.source) for t in chart_text(frame)} == {
        ("axis_labels", 8.0, "the chart's text properties"),
        ("legend", 8.0, "the chart's text properties")}
    # Nothing states one: ECMA-376's and PowerPoint's 10 pt (ooxml_common.chart.layout).
    rewrite_chart(deck, frame, lambda xml: xml.replace(' sz="800"', ""))
    assert {(t.part, t.size, t.source) for t in chart_text(frame)} == {
        ("axis_labels", 10.0, "the chart default of 10 pt"),
        ("legend", 10.0, "the chart default of 10 pt")}


def test_an_unstyled_title_is_drawn_as_a_text_box_is():
    deck, slide, frame = chart_deck(title="Regions")
    rewrite_chart(deck, frame, lambda xml: re.sub(r' sz="\d+"', "", xml))
    title = [t for t in chart_text(frame) if t.part == "title"]
    # A plain text box's cascade, as pptx2svg resolves a chart title: here the master's.
    assert [(t.size, t.source) for t in title] == [(18.0, "the master's otherStyle")]


def test_hidden_chart_text_is_not_checked():
    deck, slide, frame = chart_deck()
    deck.shape(frame.id).chart.set_data_labels(True)
    parts = [t.part for t in chart_text(frame)]
    assert parts.count("data_labels") == 2 and "legend" in parts
    rewrite_chart(deck, frame, lambda xml: re.sub(
        r"<c:legend>.*?</c:legend>", "", xml, flags=re.S).replace(
        '<c:delete val="0"/>', '<c:delete val="1"/>', 1).replace(
        '<c:showVal val="1"/>', '<c:showVal val="0"/>'))
    assert [t.part for t in chart_text(frame)] == ["axis_labels"]


def test_a_data_labels_own_size_is_read():
    deck, slide, frame = chart_deck()
    deck.shape(frame.id).chart.set_data_labels(True, series=[0])
    rewrite_chart(deck, frame, lambda xml: re.sub(
        r'(<c:dLbls>.*?<a:defRPr[^>]*?)sz="\d+"', r'\1sz="700"', xml, count=1, flags=re.S))
    labels = [(t.size, t.source) for t in chart_text(frame) if t.part == "data_labels"]
    assert labels == [(7.0, "the series' data labels")]


def test_check_lists_a_chart_once_per_part():
    deck, slide, frame = chart_deck(title="Regions")
    toolbox, session, doc = _session(deck, {"body": 12, "chart": 12})
    data = toolbox.dispatch(session, "check", {"doc": doc}).data
    assert data["small_text"] == [{"chart": frame.id, "min": 12.0, "parts": [
        {"part": "axis_labels", "size": 11.97, "from": "the category axis' text properties"},
        {"part": "legend", "size": 11.97, "from": "the legend's text properties"}]}]
    assert "edit_chart set_text_style" in data["small_text_note"]
    fixed = toolbox.dispatch(session, "edit_chart", {
        "doc": doc, "target": frame.id, "action": "set_text_style",
        "part": "axis_labels", "style": {"size": 12}})
    assert fixed.checks["small_text"][0]["parts"] == [
        {"part": "legend", "size": 11.97, "from": "the legend's text properties"}]
    # A chart minimum of 10 passes it; one left out does not check charts at all.
    for minimum in ({"chart": 10}, {"body": 12, "table": 12}):
        toolbox, session, doc = _session(deck, minimum)
        assert "small_text" not in toolbox.dispatch(session, "check", {"doc": doc}).data


def test_a_points_own_label_is_read_over_its_series():
    deck, slide, frame = chart_deck()
    deck.shape(frame.id).chart.set_data_labels(True, series=[0])
    point = ('<c:dLbl><c:idx val="1"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
             '<a:defRPr sz="650"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
             '<c:showVal val="1"/></c:dLbl>')
    rewrite_chart(deck, frame, lambda xml: xml.replace("<c:dLbls>", "<c:dLbls>" + point, 1))
    labels = sorted((t.size, t.source) for t in chart_text(frame) if t.part == "data_labels")
    assert labels == [(6.5, "a point's data label"), (11.97, "the series' data labels")]
