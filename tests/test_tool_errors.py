"""Which error code a tool call's failure gets (review, N5).

The library's own refusals (:class:`~pptx_agent.ArgumentError`, a chart's ``ChartDataError``,
``LabelError``) are the model's to act on; a ``TypeError`` or ``ValueError`` Python raises
itself -- a mistake in a handler -- is not.  That goes to the dispatcher's mapping, the one
docx-agent's and xlsx-agent's tools get: ooxml-edit's ``ArgumentError`` (which the library's
derives from, S9) is ``invalid_arguments``, its ``UnitError`` and any other ``ValueError``
``unit``, anything unexpected ``internal``, logged.
"""

from __future__ import annotations

import logging

import pytest

from conftest import FIXTURE_DIR
from ooxml_edit import errors as edit_errors
from ooxml_edit.charts import ChartDataError
from ooxml_edit.tools import Toolbox

from pptx_agent import ArgumentError, ArgumentTypeError, Document, FullStateError, LabelError
from pptx_agent.edit.layout import LayoutError
from pptx_agent.edit.scales import ScaleError
from pptx_agent.edit.svgprofile import SvgProfileError
from pptx_agent.tools import FORMAT, GROUPS, TOOLS


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = toolbox.session()
    deck = Document.new()
    deck.add_slide("Title Only").title = "Plan"
    session.open(deck.to_bytes(), "plan.pptx")
    return session


def _shape(toolbox, session, **item):
    return toolbox.dispatch(session, "ppt_add_shape", {
        "doc": "d1", "slide": "s:256",
        "items": [{"preset": "rect", "box": {"x": 10, "y": 10, "w": 100, "h": 50}, **item}]})


def test_the_librarys_refusals_are_value_and_type_errors_still():
    assert issubclass(ArgumentError, ValueError)
    assert issubclass(ArgumentTypeError, ArgumentError) and issubclass(ArgumentTypeError, TypeError)
    # One family with ooxml-edit's (S9): the dispatcher's code for either is the same.
    assert issubclass(ArgumentError, edit_errors.ArgumentError)
    assert issubclass(ArgumentTypeError, edit_errors.ArgumentTypeError)
    assert not issubclass(ArgumentError, edit_errors.UnitError)
    for refusal in (ScaleError, LayoutError, SvgProfileError, FullStateError):
        assert issubclass(refusal, ArgumentError)
    deck = Document.new()
    paragraph = deck.add_slide("Blank").add_textbox(0, 0, 914400, 457200, "x") \
        .text_frame.paragraph(0)
    with pytest.raises(ArgumentError, match="Pt"):
        paragraph.space_before = 6                  # 6 EMU: points meant
    with pytest.raises(ArgumentTypeError, match="unknown run formatting"):
        paragraph.run(0).format(sparkle=True)


def test_a_typeerror_in_a_handler_is_internal_and_logged(toolbox, session, monkeypatch, caplog):
    def broken(self, *args, **kwargs):
        return None + 1                             # a mistake in the code, not the call

    monkeypatch.setattr(Document, "add_slide", broken)
    with caplog.at_level(logging.ERROR, logger="ooxml_edit.tools"):
        result = toolbox.dispatch(session, "ppt_add_slide", {"doc": "d1", "layout": "Title Only"})
    assert result.error.code == "internal"
    assert result.error.message.startswith("TypeError: unsupported operand")
    assert any("unexpected TypeError" in record.getMessage() for record in caplog.records)


def test_pythons_own_valueerror_in_a_handler_is_unit_as_in_every_format(toolbox, session,
                                                                         monkeypatch):
    def broken(self, *args, **kwargs):
        return int("not a number")

    monkeypatch.setattr(Document, "add_slide", broken)
    result = toolbox.dispatch(session, "ppt_add_slide", {"doc": "d1", "layout": "Title Only"})
    assert result.error.code == "unit"


def test_a_shared_librarys_unit_error_is_unit(toolbox, session, monkeypatch):
    """A number out of its range is ooxml-edit's ``UnitError``: ``unit``, the code that says
    so, in a deck as in a document (S9)."""
    from ooxml_edit.sheets.workbook import number_text

    def broken(self, *args, **kwargs):
        return number_text(float("inf"))

    monkeypatch.setattr(Document, "add_slide", broken)
    result = toolbox.dispatch(session, "ppt_add_slide", {"doc": "d1", "layout": "Title Only"})
    assert result.error.code == "unit" and "not a finite number" in result.error.message


def test_the_librarys_argument_errors_keep_their_codes(toolbox, session):
    colour = _shape(toolbox, session, fill="notacolour")
    assert colour.error.code == "invalid_arguments"
    assert "neither a hex colour nor a theme colour" in colour.error.message
    made = [_shape(toolbox, session).created[0] for _ in range(2)]
    two = toolbox.dispatch(session, "ppt_arrange", {"doc": "d1", "targets": made,
                                                    "action": "distribute", "axis": "vertical"})
    assert two.error.code == "invalid_arguments" and "three shapes" in two.error.message


def test_a_charts_refusal_is_invalid_arguments(toolbox, session):
    added = toolbox.dispatch(session, "edit_chart", {
        "doc": "d1", "target": "s:256", "action": "add", "chart_type": "bar",
        "categories": ["a", "b"], "data": [{"name": "S", "values": [1, 2]}],
        "box": {"x": 50, "y": 100, "w": 400, "h": 300}})
    chart = added.created[0]
    wrong = toolbox.dispatch(session, "edit_chart", {"doc": "d1", "target": chart,
                                                     "action": "format_series",
                                                     "series": "S"})
    assert issubclass(ChartDataError, ValueError)
    assert wrong.error.code == "invalid_arguments" and "needs style" in wrong.error.message
    label = toolbox.dispatch(session, "edit_chart", {"doc": "d1", "target": chart,
                                                     "action": "rename_series",
                                                     "series": "Nothing", "text": "T"})
    assert label.error.code == "label_not_found" and issubclass(LabelError, KeyError)


def test_a_shared_librarys_refusal_gets_the_code_it_gets_in_a_document(toolbox):
    """ooxml-edit's ``Diagram.remove_node`` keeps the last node with its ``ArgumentError``:
    ``invalid_arguments`` here, as in docx-agent (it was ``unit`` in both before S9)."""
    session = toolbox.session()
    session.open((FIXTURE_DIR / "powerpoint-smartart.pptx").read_bytes(), "s.pptx")
    for _ in range(10):
        result = toolbox.dispatch(session, "edit_smartart", {
            "doc": "d1", "target": "256.3", "action": "remove_node", "node": 0})
        if not result.ok:
            break
    assert result.error.code == "invalid_arguments"
    assert "at least one node" in result.error.message
