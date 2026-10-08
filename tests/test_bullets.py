"""A bullet hangs, as PowerPoint hangs one (the trial's N6).

p7 run 1 found its bullets drawn on top of the text and fixed it with margin_left and
indent by hand; p7 run 2 gave up on real bullets.
"""

from __future__ import annotations

import pytest

from pptx_agent import Document
from pptx_agent.edit.text import BULLET_HANGING, bullet_hanging

EMU = 914400


def _paragraph(size=18):
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, 4 * EMU, EMU, "One\nTwo")
    box.text_frame.format(size=size)
    return deck, box, box.text_frame.paragraph(0)


@pytest.mark.parametrize("size, hanging", [(8, 171450), (12, 171450), (12.5, 285750),
                                           (18, 285750), (19, 342900), (24, 342900),
                                           (28, 457200), (40, 571500), (54, 685800),
                                           (60, 857250), (96, 1143000)])
def test_powerpoints_hanging_indent_by_size(size, hanging):
    """What PowerPoint 16 for Mac wrote when it bulleted a text box at each size."""
    assert bullet_hanging(size) == hanging
    deck, box, paragraph = _paragraph(size)
    paragraph.set_bullet()
    assert (paragraph.margin_left, paragraph.indent) == (hanging, -hanging)


def test_the_steps_only_grow():
    amounts = [amount for _, amount in BULLET_HANGING]
    assert amounts == sorted(amounts)


def test_hanging_can_be_given_or_left_alone():
    deck, box, paragraph = _paragraph()
    paragraph.set_bullet("–", hanging=False)
    assert (paragraph.margin_left, paragraph.indent) == (None, None)
    paragraph.set_bullet("–", hanging=228600)
    assert (paragraph.margin_left, paragraph.indent) == (228600, -228600)
    with pytest.raises(ValueError):
        paragraph.set_bullet("–", hanging=-1)
    second = box.text_frame.paragraph(1)
    second.bullet = "•"                                   # the setter hangs it too
    assert second.indent == -285750


def test_a_paragraph_that_already_hangs_keeps_its_indents():
    deck = Document.new()
    body = deck.add_slide("Title and Content").shapes[1]
    body.set_text("First\nSecond")
    paragraph = body.text_frame.paragraph(0)
    assert paragraph.effective.indent < 0                 # the master's list style
    paragraph.set_bullet("–")
    assert (paragraph.margin_left, paragraph.indent) == (None, None)


def test_the_hanging_indent_narrows_the_lines_it_wraps():
    pytest.importorskip("pptx2svg")
    deck = Document.new()
    words = "Agree the scope with the steering committee before the pilot starts"
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, 3 * EMU, EMU, words)
    plain = box.text_fit()
    box.text_frame.paragraph(0).set_bullet()
    bulleted = box.text_fit()
    assert bulleted.margin_to_wrap != plain.margin_to_wrap
    assert sum(bulleted.lines) >= sum(plain.lines)
