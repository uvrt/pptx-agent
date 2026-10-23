"""A text box's inherited colour as PowerPoint draws it, where it comes from, and its contrast.

Round 15 (T1-T3).  Measured on PowerPoint 16.106 for Mac, 23 October 2026: a text box or
autoshape takes its run's, paragraph's, own list style's, ``fontRef``'s and then the
presentation's ``defaultTextStyle`` colour at the paragraph's level -- not the master's
``otherStyle`` (``ooxml_common.drawingml.text_chain``, the chain pptx2svg draws with).
The repro: ``defaultTextStyle`` levels white over an ``otherStyle`` of ``tx1``: PowerPoint
draws the text box white.
"""

from __future__ import annotations

import io
import json
import re
import zipfile

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit import effective
from pptx_agent.edit.contrast import luminance, ratio
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK

WHITE_LEVELS = ('<p:defaultTextStyle><a:defPPr><a:defRPr><a:solidFill><a:srgbClr val="000000"/>'
                '</a:solidFill></a:defRPr></a:defPPr>'
                + "".join(f'<a:lvl{n}pPr><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="FFFFFF"/>'
                          f'</a:solidFill></a:defRPr></a:lvl{n}pPr>' for n in range(1, 10))
                + "</p:defaultTextStyle>")


def _deck(default_style: str = WHITE_LEVELS, slides: int = 1) -> Document:
    """A made-up deck whose default text style is ``default_style``; the master's otherStyle
    is the new deck's (``tx1``)."""
    deck = Document.new(size="16:9")
    for _ in range(slides):
        deck.add_slide("Blank")
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(deck.to_bytes())) as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                data = re.sub(rb"<p:defaultTextStyle>.*?</p:defaultTextStyle>",
                              default_style.encode(), data, flags=re.S)
            target.writestr(item, data)
    return Document.open(out.getvalue())


def _session(deck):
    toolbox = Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


def _call(toolbox, session, name, **arguments):
    result = toolbox.dispatch(session, name, arguments)
    assert result.ok, json.dumps(result.to_json())
    return result


def _first_run(shape):
    return shape.text_frame.paragraphs[0].runs[0]


def _origin(run):
    paragraph = run._paragraph
    return effective.color_origin(paragraph._frame._host(), paragraph._element(), run._element())


def test_the_repro_a_text_box_is_white_from_the_default_text_style():
    deck = _deck()
    slide = deck.slides[0]
    box = slide.add_textbox(0, 0, 2286000, 762000, "Hello")
    run = _first_run(box)
    assert run.effective_color.value.upper() == "FFFFFF"
    assert _origin(run) == "defaultTextStyle lvl1"
    assert run.effective_source("color") == "the presentation's default text style"
    # A level-1 paragraph takes lvl2pPr; a placeholder keeps the master's style.
    box.text_frame.paragraphs[0].level = 1
    assert _origin(_first_run(box)) == "defaultTextStyle lvl2"


def test_other_style_is_not_a_text_box_s_but_font_ref_and_the_list_style_are():
    deck = _deck('<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="2000"><a:solidFill>'
                 '<a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:lvl1pPr>'
                 '</p:defaultTextStyle>')
    slide = deck.slides[0]
    box = slide.add_textbox(0, 0, 2286000, 762000, "Hello")
    rect = slide.add_shape("rect", 0, 1000000, 2286000, 762000)   # fontRef lt1, as Office's
    rect.text_frame.text = "Hello"
    assert (_first_run(box).effective_color.value.upper(), _first_run(box).effective_size) \
        == ("FF0000", 20.0)
    assert _first_run(rect).effective_color.value.upper() == "FFFFFF"
    assert _origin(_first_run(rect)) == "style fontRef"
    # a level the default text style lacks: nothing (black, 18 pt), not its defPPr
    box.text_frame.paragraphs[0].level = 2
    assert (_first_run(box).effective_color.value.upper(), _first_run(box).effective_size) \
        == ("000000", 18.0)
    assert _origin(_first_run(box)) == "default"


def test_read_names_the_drawn_colour_and_its_source():
    deck = _deck()
    slide = deck.slides[0]
    box = slide.add_textbox(0, 0, 2286000, 762000, "Hello")
    mixed = slide.add_textbox(0, 1000000, 2286000, 762000, "Hello there")
    paragraph = mixed.text_frame.paragraphs[0]
    paragraph.runs[0].text = "Hello "
    paragraph.add_run("there").color = "#C00000"
    toolbox, session, doc = _session(deck)
    compact = _call(toolbox, session, "ppt_read_slides", doc=doc, detail="compact").data
    rows = {row[0]: dict(zip(compact["columns"], row)) for row in compact["shapes"]}
    assert rows[box.id]["color"] == "#FFFFFF defaultTextStyle lvl1"
    assert rows[mixed.id]["color"] == "mixed"
    geometry = {e["id"]: e for e in _call(toolbox, session, "ppt_read_slides", doc=doc,
                                          detail="geometry").data}
    assert (geometry[box.id]["text_color"], geometry[box.id]["color_source"]) \
        == ("#FFFFFF", "defaultTextStyle lvl1")
    assert geometry[mixed.id]["text_color"] == "mixed" and "color_source" not in geometry[mixed.id]


def test_check_reports_white_on_white_with_its_source():
    deck = _deck()
    toolbox, session, doc = _session(deck)
    slide = f"s:{deck.slides[0].slide_id}"
    added = _call(toolbox, session, "ppt_add_shape", doc=doc, slide=slide, items=[
        {"preset": "textbox", "box": {"x": 40, "y": 40, "w": 200, "h": 40}, "text": "Hello"},
        {"preset": "rect", "box": {"x": 40, "y": 200, "w": 200, "h": 80}, "text": "Readable"}])
    target = added.created[0]
    result = _call(toolbox, session, "check", doc=doc, contrast=4.5)
    assert result.data["contrast"] == [{"shape": target, "text": "Hello", "color": "#FFFFFF",
                                        "behind": "#FFFFFF", "ratio": 1.0,
                                        "source": "defaultTextStyle lvl1",
                                        "over": "background", "large": True}]
    assert "contrast_note" in result.data
    assert "1 text(s) below 4.5:1 contrast" in result.summary
    # Without it, nothing is measured.
    assert "contrast" not in _call(toolbox, session, "check", doc=doc).data


def test_contrast_blends_alpha_takes_a_gradients_worst_stop_and_knows_what_it_cannot_tell():
    deck = _deck('<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1200"><a:solidFill>'
                 '<a:srgbClr val="777777"/></a:solidFill></a:defRPr></a:lvl1pPr>'
                 '</p:defaultTextStyle>')
    slide = deck.slides[0]
    card = slide.add_shape("rect", 0, 0, 4000000, 2000000)
    card.fill = "#000000 alpha=50%"
    over = slide.add_textbox(500000, 500000, 2000000, 600000, "On the card")
    card_entry = {e.shape: e for e in slide.contrast().entries}[over.id]
    assert card_entry.behind == "#808080" and card_entry.over == card.id
    assert card_entry.ratio == pytest.approx(round(ratio("#777777", "#808080"), 2))
    assert card_entry.source == "defaultTextStyle lvl1" and not card_entry.large


def test_wcag_luminance_and_ratio():
    assert luminance("#FFFFFF") == pytest.approx(1.0)
    assert luminance("#000000") == 0.0
    assert ratio("#FFFFFF", "#000000") == pytest.approx(21.0)
    assert ratio("#777777", "#FFFFFF") == pytest.approx(4.48, abs=0.01)


def test_compare_slides_compares_the_drawn_colour():
    deck = _deck(slides=2)
    first = deck.slides[0].add_textbox(0, 0, 2286000, 762000, "Hello")
    second = deck.slides[1].add_textbox(0, 0, 2286000, 762000, "Hello")
    first.name = second.name = "Caption"
    _first_run(second).color = "#000000"
    data = deck.compare_slides(deck.slides).to_json()
    caption = next(e for e in data["shapes"] if e.get("name") == "Caption")
    colours = {caption["differ"][0]["value"]} | {o["value"] for o in caption["differ"][0]["other"]}
    assert colours == {"#FFFFFF", "#000000"}
