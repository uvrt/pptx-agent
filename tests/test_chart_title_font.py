"""A chart title that states no font (round 6, item 11): ``check``, the checks after a change
and ``save_document`` list each such chart once, with the one ``edit_chart`` call that states
it; nothing is changed by itself.

Measured on PowerPoint 16 for Mac, 2026-10-16: a rich title with no face and no size anywhere
is drawn in 18 pt Arial, not the theme's font.  Every deck here is made up.
"""

from __future__ import annotations

import re

from ooxml_edit.tools import Toolbox

from pptx_agent.edit.chart_text import NEW_TITLE_SIZE, title_font_fix
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, chart_deck, rewrite_chart

#: A rich title as the probes wrote it: one run whose ``a:rPr`` states only its language.
UNSTYLED = ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
            '<a:t>Regions</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')
SIZED = UNSTYLED.replace('<a:rPr lang="en-US"/>', '<a:rPr lang="en-US" sz="2400"/>')
AUTOMATIC = '<c:title><c:overlay val="0"/></c:title>'


def _title(markup: str):
    def change(xml: str) -> str:
        return re.sub(r"<c:title>.*?</c:title>", lambda _: markup, xml, count=1, flags=re.S)
    return change


def _deck(markup: str | None):
    deck, slide, frame = chart_deck(title="Regions")
    if markup is not None:
        rewrite_chart(deck, frame, _title(markup))
    return deck, slide, frame


def _session(deck):
    toolbox = Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


def _fix(frame_id: str, size: bool = True) -> dict:
    style = {"font": "+mn-lt", **({"size": NEW_TITLE_SIZE} if size else {})}
    return {"action": "set_text_style", "target": frame_id, "part": "title", "style": style}


def test_check_lists_the_chart_once_with_the_call_that_states_its_font():
    deck, slide, frame = _deck(UNSTYLED)
    toolbox, session, doc = _session(deck)
    data = toolbox.dispatch(session, "check", {"doc": doc}).data
    assert data["chart_title_font"] == [{"chart": frame.id, "fix": _fix(frame.id)}]
    note = data["chart_title_font_note"]
    assert note.startswith("chart_title_font: chart title states no font: PowerPoint draws it "
                           "in Arial 18 pt")
    assert "+mn-lt is the theme's body font" in note and "18.62" in note
    assert NEW_TITLE_SIZE == 18.62


def test_the_fix_settles_it_and_nothing_is_fixed_by_itself():
    deck, slide, frame = _deck(UNSTYLED)
    toolbox, session, doc = _session(deck)
    before = toolbox.dispatch(session, "check", {"doc": doc}).data["chart_title_font"]
    # check changes nothing: the title still states no font.
    again = toolbox.dispatch(session, "check", {"doc": doc}).data
    assert again["chart_title_font"] == before
    fixed = toolbox.dispatch(session, "edit_chart", {"doc": doc, **before[0]["fix"]})
    assert "chart_title_font" not in fixed.checks
    assert "chart_title_font" not in toolbox.dispatch(session, "check", {"doc": doc}).data


def test_the_checks_after_a_change_on_the_slide_list_it():
    deck, slide, frame = _deck(UNSTYLED)
    toolbox, session, doc = _session(deck)
    changed = toolbox.dispatch(session, "ppt_set_text", {"doc": doc, "items": [
        {"target": slide.shapes[0].id, "text": "New title"}]})
    assert changed.checks["chart_title_font"] == [{"chart": frame.id, "fix": _fix(frame.id)}]


def test_save_lists_it_unresolved_with_its_note():
    deck, slide, frame = _deck(UNSTYLED)
    toolbox, session, doc = _session(deck)
    saved = toolbox.dispatch(session, "save_document",
                             {"doc": doc, "name": "out.pptx", "format": "pptx"})
    left = saved.data["unresolved"]
    assert left["chart_title_font"] == [{"chart": frame.id, "fix": _fix(frame.id)}]
    assert "Arial 18 pt" in left["chart_title_font_note"]
    assert "1 chart title font" in saved.summary


def test_a_title_that_states_a_size_but_no_face_is_listed_without_a_size():
    deck, slide, frame = _deck(SIZED)
    assert title_font_fix(frame.id, frame) == _fix(frame.id, size=False)


def _found(markup: str | None):
    deck, slide, frame = _deck(markup)
    return title_font_fix(frame.id, frame)


def test_no_finding_for_a_new_chart_an_automatic_title_or_a_stated_face():
    assert _found(None) is None          # edit_chart add's title: +mn-lt at 18.62 pt
    assert _found(AUTOMATIC) is None     # the minor face, bold (measured): not Arial
    assert _found(UNSTYLED.replace('<a:rPr lang="en-US"/>',
                                   '<a:rPr lang="en-US"><a:latin typeface="Georgia"/></a:rPr>')) is None


def test_no_finding_for_a_deleted_title():
    deck, slide, frame = _deck(UNSTYLED)

    def delete(xml: str) -> str:
        assert '<c:autoTitleDeleted val="0"/>' in xml
        return xml.replace('<c:autoTitleDeleted val="0"/>', '<c:autoTitleDeleted val="1"/>')
    rewrite_chart(deck, frame, delete)
    assert title_font_fix(frame.id, frame) is None
