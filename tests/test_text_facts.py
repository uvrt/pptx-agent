"""Text facts in a deck (round 5, T4): ``{{d:scores/Beta/rating}}`` in a text argument is
written as the value, in the session's locale, the result lists it with its sources, and the
deck keeps where it came from in the shape's tags part (``UVRT_TEXT_FACTS``), read back by
``Document.provenance()`` after a save -- while its characters are its text.

The data is made up: the toy dataset and grid of :mod:`test_tables_from_data`.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from ooxml_common.locales import Locale
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit import provenance as kept
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import GRID, _session, add, call, ds_put, failed, ok

NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")
BOX = {"x": 40, "y": 40, "w": 400, "h": 80}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    return session


def deck_of(session):
    return session.entry("d1").document


def box(toolbox, session, text="", **item):
    item = {"preset": "textbox", "box": BOX, **({"text": text} if text else {}), **item}
    return ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256",
                   items=[item])).created[0]


def set_text(toolbox, session, target, **item):
    return ok(call(toolbox, session, "ppt_set_text", doc="d1",
                   items=[{"target": target, **item}]))


def facts_of(session, address):
    return Document.open(deck_of(session).to_bytes()).provenance().get(address, {}).get(
        "facts", [])


# -- writing and keeping -------------------------------------------------------------------------


def test_set_text_writes_the_value_and_the_deck_keeps_where_it_came_from(toolbox, session):
    shape = box(toolbox, session, "Placeholder")
    result = set_text(toolbox, session, shape,
                      text="Scores\nBeta rated {{d:scores/Beta/rating}} of 5")
    assert deck_of(session).shape(shape).text_frame.text == "Scores\nBeta rated 4 of 5"
    assert result.data["facts"] == [{"from": "d:scores/Beta/rating", "text": "4",
                                     "sources": ["G!B3"]}]
    assert facts_of(session, shape) == [{"source": "d:scores/Beta/rating", "text": "4",
                                         "sources": ["G!B3"], "p": 1, "at": 11}]
    # The tag sits in the shape's own nvPr, its part a tags part of the slide.
    package = zipfile.ZipFile(io.BytesIO(deck_of(session).to_bytes()))
    slide = package.read("ppt/slides/slide1.xml").decode()
    assert "<p:nvPr><p:custDataLst><p:tags r:id=" in slide
    tags = [name for name in package.namelist() if name.startswith("ppt/tags/")]
    assert len(tags) == 1 and kept.FACTS_TAG in package.read(tags[0]).decode()


def test_nl_be_writes_decimal_commas_and_its_day_order(toolbox):
    session = _session(toolbox, locale=NL_BE)
    ok(call(toolbox, session, "ds_put", id="typed", typed="yes"))
    shape = box(toolbox, session, "Gamma: {{d:typed/Gamma/rating:0.0}} (exact "
                "{{d:typed/Gamma/rating}}), due {{d:typed/Gamma/due}}")
    assert deck_of(session).shape(shape).text_frame.text == \
        "Gamma: 1,3 (exact 1,25), due 15/3/2026"
    found = facts_of(session, shape)
    assert [(fact["text"], fact["at"]) for fact in found] == [("1,3", 7), ("1,25", 18),
                                                              ("15/3/2026", 29)]
    assert found[0]["format"] == "0.0"


def test_a_table_of_another_document_gives_one_value(toolbox, session):
    shape = box(toolbox, session, "Alpha {{t:checks/Alpha/rating}}")
    assert deck_of(session).shape(shape).text_frame.text == "Alpha 3.5"
    assert facts_of(session, shape)[0]["source"] == "t:checks/Alpha/rating"


def test_runs_of_paragraphs_place_each_value(toolbox, session):
    shape = box(toolbox, session, paragraphs=[
        {"runs": [{"text": "Title"}]},
        {"runs": [{"text": "Gamma ", "bold": True},
                  {"text": "{{d:scores/Gamma/rating}}"}, {"text": " of 5"}]}])
    assert deck_of(session).shape(shape).text_frame.text == "Title\nGamma 1.25 of 5"
    assert [(f["p"], f["at"], f["text"]) for f in facts_of(session, shape)] == [(1, 6, "1.25")]


# -- edits keep it true --------------------------------------------------------------------------


def test_an_edit_of_the_value_drops_it_and_an_edit_elsewhere_moves_it(toolbox, session):
    shape = box(toolbox, session, "Beta {{d:scores/Beta/rating}} and {{d:scores/Alpha/rating}}")
    assert deck_of(session).shape(shape).text_frame.text == "Beta 4 and 3.5"
    set_text(toolbox, session, shape, text="Beta 5 and 3.5")
    assert [f["text"] for f in facts_of(session, shape)] == ["3.5"]
    set_text(toolbox, session, shape, text="Then: Beta 5 and 3.5")
    assert [(f["text"], f["at"]) for f in facts_of(session, shape)] == [("3.5", 17)]
    set_text(toolbox, session, shape, text="Nothing left")
    assert facts_of(session, shape) == []
    assert not [n for n in deck_of(session).package.part_names if n.startswith("ppt/tags/")]


def test_replace_text_writes_values_and_drops_what_it_replaces(toolbox, session):
    shape = box(toolbox, session, "Beta scored X; Alpha {{d:scores/Alpha/rating}}")
    result = ok(call(toolbox, session, "replace_text", doc="d1", find="X",
                     replace="{{d:scores/Beta/rating}}", expect="one"))
    assert result.data["facts"][0]["text"] == "4"
    assert deck_of(session).shape(shape).text_frame.text == "Beta scored 4; Alpha 3.5"
    assert [(f["text"], f["at"]) for f in facts_of(session, shape)] == [("4", 12), ("3.5", 21)]
    ok(call(toolbox, session, "replace_text", doc="d1", find="3.5", replace="n/a",
            expect="one"))
    assert [f["text"] for f in facts_of(session, shape)] == ["4"]


def test_a_text_edited_elsewhere_is_checked_on_reading(toolbox, session):
    shape = box(toolbox, session, "Beta {{d:scores/Beta/rating}}")
    deck = deck_of(session)
    deck.shape(shape).text_frame.paragraph(0).text = "Beta 9"     # not through a tool
    assert deck.provenance().get(shape) is None


def test_undo_takes_the_facts_with_the_text(toolbox, session):
    shape = box(toolbox, session, "Plain")
    set_text(toolbox, session, shape, text="Beta {{d:scores/Beta/rating}}")
    ok(call(toolbox, session, "undo", doc="d1"))
    assert deck_of(session).shape(shape).text_frame.text == "Plain"
    assert facts_of(session, shape) == []


# -- tables --------------------------------------------------------------------------------------


def test_table_cells_keep_their_facts_through_row_edits(toolbox, session):
    frame = ok(add(toolbox, session, rows=2, columns=2, data=[
        ["Check", "Rating"], ["Beta", "{{d:scores/Beta/rating}}"]])).created[0]
    assert facts_of(session, frame) == [{"source": "d:scores/Beta/rating", "text": "4",
                                         "sources": ["G!B3"], "p": 0, "at": 0,
                                         "cell": [1, 1]}]
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=frame, action="set_cells",
            cells=[{"row": 0, "col": 1, "text": "Score {{d:scores/Alpha/rating:0}}"}]))
    assert {tuple(f["cell"]): f["text"] for f in facts_of(session, frame)} == {
        (0, 1): "4", (1, 1): "4"}
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=frame, action="insert_row",
            index=1))
    assert sorted(tuple(f["cell"]) for f in facts_of(session, frame)) == [(0, 1), (2, 1)]
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=frame, action="delete_row",
            index=2))
    assert [tuple(f["cell"]) for f in facts_of(session, frame)] == [(0, 1)]
    set_text(toolbox, session, f"{frame}/cell0,1", text="Score")
    assert facts_of(session, frame) == []


def test_a_table_from_a_source_keeps_its_record_beside_cell_facts(toolbox, session):
    frame = ok(add(toolbox, session, **{"from": "d:scores",
                                        "from_columns": ["check", "rating"]})).created[0]
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=frame, action="insert_row",
            cells=None))
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=frame, action="set_cells",
            cells=[{"row": 4, "col": 1, "text": "{{d:scores/Beta/rating}}"}]))
    record = Document.open(deck_of(session).to_bytes()).provenance()[frame]
    assert record["source"] == "d:scores" and len(record["columns"]) == 2
    assert [tuple(f["cell"]) for f in record["facts"]] == [(4, 1)]


# -- what is no reference, and refusals ----------------------------------------------------------


def test_an_escaped_reference_and_other_braces_are_written_as_they_are(toolbox, session):
    result = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": BOX, "text": "Dear {{name}}: \\{{d:scores/Beta/rating}}"}]))
    shape = result.created[0]
    assert deck_of(session).shape(shape).text_frame.text == \
        "Dear {{name}}: {{d:scores/Beta/rating}}"
    assert result.data is None or "facts" not in result.data
    assert facts_of(session, shape) == []


def test_a_reference_to_many_values_changes_nothing(toolbox, session):
    shape = box(toolbox, session, "Plain")
    error = failed(call(toolbox, session, "ppt_set_text", doc="d1", items=[
        {"target": shape, "text": "{{d:scores/rating}}"}]), "invalid_arguments")
    assert error.field == "items[0].text" and "not one value" in error.message
    assert deck_of(session).shape(shape).text_frame.text == "Plain"


def test_speaker_notes_take_the_value_without_keeping_it(toolbox, session):
    set_text(toolbox, session, "256/notes", text="Say {{d:scores/Beta/rating}}")
    assert deck_of(session).resolve("256/notes").text == "Say 4"
