"""Values styled by a scale on a deck (round 36, R31; ooxml-edit 0.85.0).

``style_by_value`` reads each target's value -- a shape, a table cell, a table's column, the
values written by reference -- and fills its shape or cell with its step's colour and colours
its text (a deck has no character styles: direct formatting).  ``check(scale=...)`` lists
the values whose fill is another step's than their value's, or none's.  Made-up decks.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.scales", reason="needs ooxml-edit 0.85 (scales)")

from ooxml_edit.tools import DocumentFormat, Toolbox  # noqa: E402

from pptx_agent import Document  # noqa: E402
from pptx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import CLOCK, COLUMNS, ROWS, Grid, _grid_table, call, failed, ok  # noqa: E402

GRID = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table})
STEPS = [{"min": 1, "fill": "C62828", "label": "Low"},
         {"min": 2.5, "fill": "F9A825"},
         {"min": 4, "fill": "2E7D32", "label": "High"}]
BOX = {"x": 40, "y": 40, "w": 300, "h": 60}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox):
    session = toolbox.session(clock=CLOCK)
    deck = Document.new()
    deck.add_slide("Blank")
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")                                  # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")       # d2
    shape = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": BOX, "text": "{{t:checks/Alpha/rating}}"}])).created[0]
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:257",
                    box={"x": 40, "y": 140, "w": 400, "h": 120}, rows=4, columns=2,
                    data=[["Topic", "Score"], ["Door", "4.5"], ["Roof", "2"], ["Wall", "n/a"]])).created[0]
    return session, shape, table


def scale(*targets):
    return {"name": "Score", "steps": STEPS, "targets": list(targets)}


def _deck(session):
    return session.entry("d1").document


def test_each_value_is_filled_with_its_step_and_its_text_coloured(toolbox):
    session, shape, table = _session(toolbox)
    before = _deck(session).to_bytes()
    result = ok(call(toolbox, session, "style_by_value", doc="d1",
                     scale=scale(f"{table}/col1", "{{t:checks/*/rating}}")))
    assert [(t["target"], t.get("step"), t.get("skipped")) for t in result.data["targets"]] == [
        (f"{table}/cell1,1", 3, None), (f"{table}/cell2,1", 1, None),
        (f"{table}/cell3,1", None, "'n/a' is not one number"), (shape, 2, None)]
    assert result.data["styles"].startswith("none: a deck has no character styles")
    deck = _deck(session)
    cell = deck.resolve(f"{table}/cell1,1")
    assert cell.fill.color.value == "2E7D32"
    assert {run.color.value for p in cell.text_frame.paragraphs for run in p.runs} == {"FFFFFF"}
    assert deck.shape(shape).fill.color.value == "F9A825"
    assert {run.color.value for p in deck.shape(shape).text_frame.paragraphs for run in p.runs} == {"000000"}
    assert deck.resolve(f"{table}/cell0,1").fill is None            # the header row
    ok(call(toolbox, session, "undo", doc="d1"))
    assert _deck(session).to_bytes() == before


def test_preview_and_a_value_already_in_its_step(toolbox):
    session, shape, table = _session(toolbox)
    before = _deck(session).to_bytes()
    preview = ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(shape), preview=True))
    assert preview.summary.startswith("Score: would style 1 value(s)") and _deck(session).to_bytes() == before
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(shape)))
    again = ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(shape)))
    assert again.data["targets"] == [{"target": shape, "value": "3.5", "step": 2, "unchanged": True}]
    assert "nothing to change" in again.warnings[0]


def test_check_reports_a_value_changed_after_it_was_styled_and_slides_scope_it(toolbox):
    session, shape, table = _session(toolbox)
    given = scale(f"{table}/col1", shape)
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1")))
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": f"{table}/cell1,1", "text": "1"}]))
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert found["wrong"] == [{"target": f"{table}/cell1,1", "from": f"{table}/col1", "slide": "s:257",
                               "value": "1", "step": 3, "due": 1}]
    assert found["none"] == [{"target": shape, "slide": "s:256", "value": "3.5", "due": 2}]
    ok(call(toolbox, session, found["fix"]["tool"], **found["fix"]["arguments"]))
    after = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (after["right"], after["wrong"], after["none"]) == (3, [], [])
    scoped = ok(call(toolbox, session, "check", doc="d1", slides=["s:256"], scale=given)).data["scale"]
    assert scoped["checked"] == 1
    refused = failed(call(toolbox, session, "check", doc="d1", pages=[1], scale=given), "invalid_arguments")
    assert refused.field == "pages"


def test_text_coloured_otherwise_is_not_exactly_its_step(toolbox):
    session, shape, _ = _session(toolbox)
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(shape)))
    ok(call(toolbox, session, "ppt_format_text", doc="d1", items=[{"target": shape, "color": "#0000FF"}]))
    found = ok(call(toolbox, session, "check", doc="d1", scale=scale(shape))).data["scale"]
    assert found["wrong"] == [{"target": shape, "slide": "s:256", "value": "3.5", "step": 2, "due": 2,
                               "differs": "text colour 0000FF"}]


def test_targets_that_name_nothing_are_skipped_with_the_way_out(toolbox):
    session, shape, table = _session(toolbox)
    result = ok(call(toolbox, session, "style_by_value", doc="d1",
                     scale=scale(table, f"{table}/col5", f"{shape}/p0", "{{d:none/*/x}}")))
    why = [t["skipped"] for t in result.data["targets"]]
    assert why[0] == f"{table} is a table: name its column ({table}/col1) or its cells"
    assert why[1] == f"no col5: {table} has 2 columns, col0-col1"
    assert why[2].startswith(f"{shape}/p0 is part of a shape's text")
    assert why[3].startswith("{{d:none/*/x}}: no value written by a reference")
