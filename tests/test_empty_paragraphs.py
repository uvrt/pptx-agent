"""Empty paragraphs between text (production feedback round 18, Q4): ``check`` with
``include: ["layout"]`` lists, per text box and table cell, the blank lines left between
two paragraphs of text, with the ``ppt_set_text`` call that takes them out; applying it
and checking again finds nothing, and the text and its formatting are unchanged."""

from __future__ import annotations

import datetime as dt
import json

import pytest
from lxml import etree

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.emptyparagraphs import _empty, _runs, empty_paragraphs
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

PT = 12700
CLOCK = lambda: dt.datetime(2026, 10, 27, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _deck():
    """A made-up slide: a table whose cell (1, 2) reads Title, a blank line, Body text; a
    text box with an empty first and last paragraph and two runs of blank lines (one a
    space only) between its three paragraphs, its first bold and its second centred."""
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    slide.title = "Orchard notes"
    table = slide.add_table(3, 3, 40 * PT, 250 * PT, 600 * PT, 120 * PT)
    table.table.cell(1, 2).text_frame.set_text("Title\n\nBody text")
    box = slide.add_textbox(40 * PT, 100 * PT, 300 * PT, 120 * PT,
                            "\nApples\n\nPears\n \n\nPlums\n")
    frame = box.text_frame
    frame.paragraphs[1].runs[0].bold = True
    frame.paragraphs[3].alignment = "center"
    return deck, slide, table, box


def _open(toolbox, deck):
    session = toolbox.session(clock=CLOCK)
    return session, session.open(deck.to_bytes(), "notes.pptx")


def _ok(toolbox, session, tool, **arguments):
    result = toolbox.dispatch(session, tool, arguments)
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def _found(toolbox, session, doc, **arguments):
    data = _ok(toolbox, session, "check", doc=doc, include=["layout"], **arguments).data
    return data["layout"]["empty_paragraphs"]


def test_the_repro_lists_the_cell_and_the_text_box_with_their_paragraphs(toolbox):
    deck, slide, table, box = _deck()
    session, doc = _open(toolbox, deck)
    found = {entry["at"]: entry for entry in _found(toolbox, session, doc)}
    cell = f"{table.id}/cell1,2"
    assert found[cell] == {"kind": "empty_paragraphs", "at": cell, "paragraphs": [1],
                           "fix": {"tool": "ppt_set_text",
                                   "items": [{"target": cell, "text": "Title\nBody text"}]}}
    # Leading and trailing empties are not listed; the space-only one is empty; both runs
    # of the box are its one entry, so its one fix takes them all out.
    assert found[box.id]["paragraphs"] == [2, 4, 5]
    assert found[box.id]["fix"]["items"] == [{"target": box.id, "text": "\nApples\nPears\nPlums\n"}]
    assert len(found) == 2
    summary = _ok(toolbox, session, "check", doc=doc, include=["layout"]).summary
    assert "2 text(s) with empty paragraphs" in summary


def test_applying_the_fixes_leaves_nothing_and_keeps_the_text_and_its_formatting(toolbox):
    deck, slide, table, box = _deck()
    session, doc = _open(toolbox, deck)
    for entry in _found(toolbox, session, doc):
        _ok(toolbox, session, entry["fix"]["tool"], doc=doc, items=entry["fix"]["items"])
    assert _found(toolbox, session, doc) == []
    edited = session.entry(doc).document
    frame = edited.shape(box.id).text_frame
    assert frame.text == "\nApples\nPears\nPlums\n"
    assert frame.paragraphs[1].runs[0].bold is True
    assert frame.paragraphs[2].alignment == "center"
    assert edited.resolve(f"{table.id}/cell1,2").text_frame.text == "Title\nBody text"
    # One undo takes the last fix (the box's, listed after the table) back, blank lines and all.
    _ok(toolbox, session, "undo", doc=doc)
    assert [entry["at"] for entry in _found(toolbox, session, doc)] == [box.id]


def test_only_on_request_and_only_on_the_slides_asked_for(toolbox):
    deck, slide, table, box = _deck()
    deck.add_slide("Title Only").title = "Nothing blank"
    session, doc = _open(toolbox, deck)
    assert "layout" not in _ok(toolbox, session, "check", doc=doc).data
    assert _found(toolbox, session, doc, slides=[2]) == []
    assert len(_found(toolbox, session, doc, slides=[1])) == 2


def test_a_group_member_is_listed_and_a_merged_cell_is_passed_over():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    first = slide.add_textbox(40 * PT, 100 * PT, 200 * PT, 80 * PT, "One\n\nTwo")
    second = slide.add_textbox(300 * PT, 100 * PT, 200 * PT, 80 * PT, "Three")
    slide.group([first, second])
    table = slide.add_table(2, 2, 40 * PT, 250 * PT, 400 * PT, 80 * PT).table
    table.cell(0, 0).text_frame.set_text("Left\n\nRight")
    table.merge(0, 0, 0, 1)
    # A cell merged into another keeps its XML text, never drawn: not listed.
    spanned = table.cell(0, 1).text_frame
    spanned._paragraph_elements()[0].addnext(etree.fromstring(f'<a:p xmlns:a="{A}"/>'))
    spanned._paragraph_elements()[0].addprevious(etree.fromstring(
        f'<a:p xmlns:a="{A}"><a:r><a:t>Hidden</a:t></a:r></a:p>'))
    spanned._paragraph_elements()[-1].addnext(etree.fromstring(
        f'<a:p xmlns:a="{A}"><a:r><a:t>Also hidden</a:t></a:r></a:p>'))
    assert _runs("cell0,1", spanned)               # it would be, were it drawn
    found = empty_paragraphs(slide)
    assert [entry["at"] for entry in found] == [first.id, f"{table.address}/cell0,0"]


@pytest.mark.parametrize("inner, empty", [
    ("", True),
    ('<a:endParaRPr lang="en-US"/>', True),
    ('<a:pPr algn="ctr"/><a:r><a:t>  </a:t></a:r>', True),
    ("<a:r><a:t>x</a:t></a:r>", False),
    ("<a:br/>", False),                                      # a line break shows
    ('<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t></a:t></a:fld>',
     False),                                                 # a field is content
    ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/'
     '2006"/>', False),                                      # an equation's wrapper
])
def test_what_makes_a_paragraph_empty(inner, empty):
    paragraph = etree.fromstring(f'<a:p xmlns:a="{A}">{inner}</a:p>')
    assert _empty(paragraph) is empty
