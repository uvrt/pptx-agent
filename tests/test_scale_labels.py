"""A step's label in place of its value, on a deck (round 38; ooxml-edit 0.87.0).

A scale with ``"show": "label"``: ``style_by_value`` writes each value's step's label as the
shape's or cell's text -- filled and coloured as its step -- and keeps, on the shape, the
value it stands for and its reference; ``check(scale=...)`` then says when a label is no
longer its value's: the value changed at its source, the label was edited, or a label has
no value kept.  Made-up decks.
"""

from __future__ import annotations

import pytest

scales = pytest.importorskip("ooxml_edit.tools.scales", reason="needs ooxml-edit 0.87 (scale labels)")
if not hasattr(scales, "SHOW"):
    pytest.skip("needs ooxml-edit 0.87 (scale labels)", allow_module_level=True)

from ooxml_edit.tools import DocumentFormat, Toolbox  # noqa: E402

from pptx_agent import Document  # noqa: E402
from pptx_agent.edit import provenance  # noqa: E402
from pptx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import CLOCK, COLUMNS, ROWS, Grid, _grid_table, call, ok  # noqa: E402

#: The made-up grid, its tables read as values a check may read again.
LIVE = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table},
                      value_sources=lambda grid: [f"t:{name}" for name in grid.tables])
STEPS = [{"min": 1, "fill": "C62828", "label": "Low"},
         {"min": 2.5, "fill": "F9A825", "label": "Middle"},
         {"min": 4, "fill": "2E7D32", "label": "High"}]
BOX = {"x": 40, "y": 40, "w": 300, "h": 60}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT, LIVE], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox, data=None):
    session = toolbox.session(clock=CLOCK)
    if data is None:
        deck = Document.new()
        deck.add_slide("Blank")
        deck.add_slide("Blank")
        data = deck.to_bytes()
    session.open(data, "deck.pptx")                                              # d1
    session.adopt(Grid({"checks": (COLUMNS, list(ROWS))}), "grid", "book.grid")  # d2
    return session


def _build(toolbox):
    session = _session(toolbox)
    shape = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": BOX, "text": "{{t:checks/Alpha/rating}}"}])).created[0]
    lone = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {**BOX, "y": 120}, "text": "High"}])).created[0]
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:257",
                    box={"x": 40, "y": 140, "w": 400, "h": 120}, rows=4, columns=2,
                    data=[["Topic", "Score"], ["Door", "4.5"], ["Roof", "2"], ["Wall", "n/a"]])).created[0]
    return session, shape, lone, table


def labels(*targets):
    return {"name": "Score", "steps": STEPS, "show": "label", "targets": list(targets)}


def _deck(session):
    return session.entry("d1").document


def test_labels_are_the_texts_filled_as_their_steps_in_one_undo_step(toolbox):
    session, shape, _, table = _build(toolbox)
    before = _deck(session).to_bytes()
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}")
    preview = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given, preview=True))
    assert preview.summary == "Score: would label 3 value(s), 0 already in their step, 1 skipped"
    assert _deck(session).to_bytes() == before
    result = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    assert [(t["target"], t.get("value"), t.get("label")) for t in result.data["targets"]] == [
        (f"{table}/cell1,1", "4.5", "High"), (f"{table}/cell2,1", "2", "Low"),
        (f"{table}/cell3,1", "n/a", None), (shape, "3.5", "Middle")]
    deck = _deck(session)
    cell = deck.resolve(f"{table}/cell1,1")
    assert cell.text == "High" and cell.fill.color.value == "2E7D32"
    assert {run.color.value for p in cell.text_frame.paragraphs for run in p.runs} == {"FFFFFF"}
    assert deck.shape(shape).text == "Middle" and deck.shape(shape).fill.color.value == "F9A825"
    kept = provenance.labels_found(deck.shape(shape))
    assert kept == [{"label": "Middle", "value": 3.5, "source": "t:checks/Alpha/rating", "scale": "Score",
                     "step": 2, "text": "Middle", "p": 0, "at": 0}]
    assert provenance.facts_found(deck.shape(shape)) == []      # the value written is a label now
    assert [k["value"] for k in provenance.labels_found(deck.shape(table))] == [4.5, 2]
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (found["right"], found["labels"], found["unrecorded"]) == (3, [], [])
    ok(call(toolbox, session, "undo", doc="d1"))
    assert _deck(session).to_bytes() == before


def test_labels_and_their_values_survive_save_and_reopen(toolbox):
    session, shape, _, table = _build(toolbox)
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}")
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    reopened = _session(toolbox, _deck(session).to_bytes())
    found = ok(call(toolbox, reopened, "check", doc="d1", scale=given)).data["scale"]
    assert (found["right"], found["labels"], found["unrecorded"]) == (3, [], [])


def test_check_flags_a_changed_value_an_edited_label_and_one_with_no_value(toolbox):
    session, shape, lone, table = _build(toolbox)
    ok(call(toolbox, session, "style_by_value", doc="d1",
            scale=labels(f"{table}/col1", "{{t:checks/*/rating}}")))
    header, rows = session.entry("d2").document.tables["checks"]
    rows[0] = ("Alpha", 4.25, *rows[0][2:])                                       # was 3.5
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": f"{table}/cell2,1", "text": "Lowish"}]))
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}", lone)
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert [(e["target"], e["shows"], e["value"], e["label"], e["why"]) for e in found["labels"]] == [
        (f"{table}/cell2,1", "Lowish", "2", "Low", "edited: written 'Low'"),
        (shape, "Middle", "4.25", "High", "its value changed: recorded 3.5, now 4.25")]
    assert found["unrecorded"] == [{"target": lone, "slide": "s:256", "shows": "High", "label_of": 3}]
    assert found["fix"]["arguments"]["scale"]["targets"] == [shape, f"{table}/cell2,1"]
    ok(call(toolbox, session, found["fix"]["tool"], **found["fix"]["arguments"]))
    assert _deck(session).shape(shape).text == "High"
    assert _deck(session).resolve(f"{table}/cell2,1").text == "Low"
    after = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (after["right"], after["labels"], after["wrong"]) == (3, [], [])
    assert [e["target"] for e in after["unrecorded"]] == [lone]


def test_a_row_inserted_moves_the_labels_with_their_cells(toolbox):
    session, _, _, table = _build(toolbox)
    given = labels(f"{table}/col1")
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=table, action="insert_row", index=1))
    kept = provenance.labels_found(_deck(session).shape(table))
    assert [(k["cell"], k["text"]) for k in kept] == [([2, 1], "High"), ([3, 1], "Low")]
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert found["right"] == 2 and found["labels"] == []
