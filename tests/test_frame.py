"""A text frame's insets, vertical anchor and wrapping (the trial's N4).

p7 (both runs), p8 run 2 and p3 run 2 sized their boxes around the default insets and
top anchor because the API had no way to say otherwise.
"""

from __future__ import annotations

import pytest

from pptx_agent import Document

EMU = 914400


def _box(text="Hello world, this is a sentence long enough to wrap in two inches"):
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    return deck, slide, slide.add_textbox(EMU, EMU, 2 * EMU, EMU, text)


def test_the_defaults_are_powerpoints():
    deck, slide, box = _box()
    frame = box.text_frame
    assert frame.insets == (91440, 45720, 91440, 45720)
    assert frame.anchor == "top" and frame.wrap is True


def test_insets_are_set_read_and_inherited_again():
    deck, slide, box = _box()
    frame = box.text_frame
    frame.insets = (0, 12700, 0, 12700)
    assert frame.insets == (0, 12700, 0, 12700)
    body = box._element.find(".//{*}bodyPr")
    assert (body.get("lIns"), body.get("tIns"), body.get("rIns"), body.get("bIns")) == \
        ("0", "12700", "0", "12700")
    frame.set_insets(left=25400)
    assert frame.insets == (25400, 12700, 0, 12700)
    frame.insets = None
    assert frame.insets == (91440, 45720, 91440, 45720) and body.get("lIns") is None
    with pytest.raises(ValueError):
        frame.insets = (0, 0, 0)
    with pytest.raises(ValueError):
        frame.set_insets(top=-1)
    deck.undo()
    assert frame.insets == (25400, 12700, 0, 12700)


def test_a_placeholder_reads_what_it_inherits():
    deck = Document.new()
    title = deck.add_slide("Title and Content").shapes[0]
    assert title.text_frame.anchor == "middle"       # the Office master's title
    title.text_frame.anchor = "b"
    assert title.text_frame.anchor == "bottom"
    title.text_frame.anchor = None
    assert title.text_frame.anchor == "middle"


@pytest.mark.parametrize("given, written, read", [("top", "t", "top"), ("ctr", "ctr", "middle"),
                                                   ("middle", "ctr", "middle"), ("b", "b", "bottom")])
def test_anchor_names(given, written, read):
    deck, slide, box = _box()
    box.text_frame.anchor = given
    assert box._element.find(".//{*}bodyPr").get("anchor") == written
    assert box.text_frame.anchor == read


def test_an_unknown_anchor_is_refused():
    deck, slide, box = _box()
    with pytest.raises(ValueError, match="top"):
        box.text_frame.anchor = "centre"


def test_a_table_cell_frame_uses_its_margins():
    deck = Document.new()
    cell = deck.add_slide("Blank").add_table(2, 2, EMU, EMU, 4 * EMU, EMU).table.cell(0, 0)
    frame = cell.text_frame
    assert frame.insets == (91440, 45720, 91440, 45720) and frame.wrap is True
    frame.insets = (0, 0, 0, 0)
    frame.anchor = "middle"
    properties = cell._tc().find("{*}tcPr")
    assert properties.get("marL") == "0" and properties.get("anchor") == "ctr"
    assert frame.anchor == "middle"
    with pytest.raises(ValueError):
        frame.wrap = False
    assert deck.validate() == []


def test_text_fit_honours_insets_anchor_and_wrap():
    pytest.importorskip("pptx2svg")
    deck, slide, box = _box()
    frame = box.text_frame
    wrapped = box.text_fit()
    assert wrapped.lines[0] > 1
    frame.insets = (0, 0, 0, 0)
    tight = box.text_fit()
    assert tight.needed == wrapped.needed - 2 * 45720 - (wrapped.lines[0] - tight.lines[0]) * \
        (wrapped.needed - 2 * 45720) // wrapped.lines[0]
    frame.wrap = False
    assert box.text_fit().lines == (1,)
    frame.wrap = True
    # The anchor moves where the text is drawn: a label at the bottom of a tall box no
    # longer meets one at its top.
    box.height = 3 * EMU
    frame.set_text("Top")
    frame.anchor = "bottom"
    slide.add_textbox(EMU, EMU, 2 * EMU, EMU // 2, "Label")
    assert slide.collisions() == []
    frame.anchor = "top"
    assert [o.detail for o in slide.collisions()] == ["text"]


# -- each constructor's defaults (the trial's N7) -----------------------------------------------


def test_each_constructor_says_what_its_frame_is():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    shape = slide.add_shape("rect", EMU, EMU, 2 * EMU, EMU, text="Plan").text_frame
    assert (shape.autofit, shape.wrap, shape.insets, shape.anchor) == \
        ("none", True, (91440, 45720, 91440, 45720), "middle")
    box = slide.add_textbox(EMU, 3 * EMU, 2 * EMU, EMU, "Notes").text_frame
    assert (box.autofit, box.wrap, box.insets, box.anchor) == \
        ("shape", True, (91440, 45720, 91440, 45720), "top")
    cell = slide.add_table(2, 2, EMU, 5 * EMU, 4 * EMU, EMU).table.cell(0, 0).text_frame
    assert (cell.wrap, cell.insets, cell.anchor) == (True, (91440, 45720, 91440, 45720), "top")


@pytest.mark.parametrize("mode, tag", [("none", "noAutofit"), ("shape", "spAutoFit"),
                                       ("normal", "normAutofit")])
def test_a_text_box_takes_its_autofit(mode, tag):
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, 2 * EMU, EMU // 2, "Label",
                                               autofit=mode)
    assert box.text_frame.autofit == mode
    assert box._element.find(f".//{{*}}bodyPr/{{*}}{tag}") is not None
    assert deck.validate() == []
    deck.undo()
    assert len(deck.slides[0].shapes) == 0          # one step, autofit included


def test_an_unknown_autofit_is_refused():
    deck = Document.new()
    with pytest.raises(ValueError, match="none"):
        deck.add_slide("Blank").add_textbox(EMU, EMU, EMU, EMU, "x", autofit="grow")


# -- PowerPoint opens all of it ------------------------------------------------------------------


@pytest.mark.oracle
def test_powerpoint_opens_the_design_round_and_hangs_the_bullet():
    """Insets, anchors, wrapping, each autofit, a hanging bullet, spacing in EMU, a new
    theme and a slide title: PowerPoint opens the deck unprompted, and draws the bullet
    left of its text, a hanging indent apart."""
    import os
    from pathlib import Path

    import oracle as oracle_helper
    from pptx_agent import Pt

    if not oracle_helper.available():
        pytest.skip("needs macOS with Microsoft PowerPoint and the pptx2svg oracle script")
    pypdfium2 = pytest.importorskip("pypdfium2")
    deck = Document.new()
    deck.theme.set_colors({"accent1": "#0B6E79", "dk2": "#0B2545"})
    deck.theme.set_fonts(major="Georgia", minor="Arial")
    first = deck.add_slide("Title Only")
    first.title = "Design round"
    label = first.add_textbox(EMU, 2 * EMU, 3 * EMU, EMU, "Centred, no insets",
                              autofit="none")
    label.text_frame.insets = (0, 0, 0, 0)
    label.text_frame.anchor = "middle"
    first.add_textbox(5 * EMU, 2 * EMU, 3 * EMU, EMU, "Shrinks on overflow", autofit="normal")
    run_on = first.add_textbox(EMU, 4 * EMU, 3 * EMU, EMU // 2, "One line", autofit="none")
    run_on.text_frame.wrap = False
    second = deck.add_slide("Blank")
    bulleted = second.add_textbox(EMU, EMU, 6 * EMU, 2 * EMU, "Kick-off\nSteerCo")
    for paragraph in bulleted.text_frame.paragraphs:
        paragraph.set_bullet()
        paragraph.space_after = Pt(12)
    assert deck.validate() == []
    home = Path(os.path.expanduser("~"))
    path, pdf = home / "pptx-agent-design.pptx", home / "pptx-agent-design.pdf"
    deck.save(path)
    try:
        result = oracle_helper.export_pdf(path, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        document = pypdfium2.PdfDocument(str(pdf))
        try:
            page = document[1]
            text = page.get_textpage()
            boxes = [(text.get_text_range(i, 1), text.get_charbox(i)[0])
                     for i in range(text.count_chars()) if text.get_text_range(i, 1).strip()]
        finally:
            document.close()
        bullet_x = next(x for char, x in boxes if char == "•")
        text_x = next(x for char, x in boxes if char == "K")
        hanging = 285750 / 12700                                 # 22.5 pt at 18 pt
        assert abs((text_x - bullet_x) - hanging) < 3, (bullet_x, text_x)
    finally:
        oracle_helper.cleanup(path, pdf)
