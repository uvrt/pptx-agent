"""Lengths in EMU everywhere, points said with Pt (the trial's N5).

Both p7 agents passed EMU to ``space_after``, which took points, and got
``TextFit(needed=2,904,484,920)`` with no warning -- about twenty minutes lost each.
"""

from __future__ import annotations

import warnings

import pytest

from pptx_agent import Document, Pt, UnitWarning
from pptx_agent.oxml.xml import qn

EMU = 914400


def _paragraph():
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, 3 * EMU, EMU, "One\nTwo")
    return deck, box, box.text_frame.paragraph(0)


def test_pt_is_the_emu_it_names():
    assert Pt(6) == 76200 and isinstance(Pt(6), int) and Pt(6).points == 6.0
    assert Pt(12.5) == 158750 and repr(Pt(12.5)) == "Pt(12.5)"
    assert Pt(1) + 1 == 12701


def test_spacing_is_emu_and_reads_back_as_emu():
    deck, box, paragraph = _paragraph()
    paragraph.space_after = Pt(6)
    paragraph.space_before = 2 * 12700
    assert (paragraph.space_after, paragraph.space_before) == (76200, 25400)
    spacing = paragraph._element().find(f"{qn('a:pPr')}/{qn('a:spcAft')}/{qn('a:spcPts')}")
    assert spacing.get("val") == "600"
    assert paragraph.effective.space_after == 76200 and paragraph.effective.space_before == 25400
    paragraph.space_after = None
    assert paragraph.space_after is None


@pytest.mark.parametrize("value", [6, 12.5, 100])
def test_points_given_as_emu_are_refused(value):
    deck, box, paragraph = _paragraph()
    with pytest.raises(ValueError, match=r"Pt\("):
        paragraph.space_after = value


def test_an_implausible_spacing_is_refused():
    """The trial's value: EMU given where points were taken, then read as points."""
    deck, box, paragraph = _paragraph()
    with pytest.raises(ValueError, match="1,000|EMU"):
        paragraph.space_after = 76200 * 12700
    with pytest.raises(ValueError):
        paragraph.space_before = Pt(1001)
    with pytest.raises(ValueError):
        paragraph.space_before = -1


def test_a_spacing_under_a_point_warns_unless_it_says_pt():
    deck, box, paragraph = _paragraph()
    with pytest.warns(UnitWarning, match="Pt"):
        paragraph.space_after = 600
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        paragraph.space_after = Pt(0.5)
        paragraph.space_after = 0
    assert paragraph.space_after == 0


def test_a_font_size_takes_pt_as_points():
    deck, box, paragraph = _paragraph()
    run = paragraph.run(0)
    run.size = Pt(14)
    assert run.size == 14.0
    paragraph.line_spacing_points = Pt(20)
    assert paragraph.line_spacing_points == 20.0


def test_text_fit_counts_spacing_in_emu():
    pytest.importorskip("pptx2svg")
    deck, box, paragraph = _paragraph()
    before = box.text_fit().needed
    paragraph.space_after = Pt(12)
    assert box.text_fit().needed == before + 152400
