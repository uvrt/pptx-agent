"""One way to name a slide everywhere, positions that say which base they use, and edit
calls that return what they edited (the full trial's finding 16)."""

from __future__ import annotations

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document, LineFormat, Paragraph, TableCell, TextFrame

SAMPLE = FIXTURE_DIR / "sample.pptx"           # slides 256..261
REPORT = FIXTURE_DIR / "real-financial-report.pptx"


def _ids(deck: Document) -> list[int]:
    return [slide.slide_id for slide in deck.slides]


def test_a_slide_is_named_by_object_id_or_s_id():
    deck = Document.open(SAMPLE)
    slide = deck.slides[1]
    assert deck.slide(257) is not None
    assert deck.slide("s:257").slide_id == deck.slide("257").slide_id == 257
    assert deck.slide(slide).slide_id == 257
    with pytest.raises(KeyError, match=r"looks like a position: deck\.slides\[2\] is slide s:258"):
        deck.slide(2)
    with pytest.raises(ValueError):
        deck.slide("second")


def test_slide_operations_take_s_ids():
    deck = Document.open(SAMPLE)
    copy = deck.duplicate_slide("s:257")
    assert copy.index == 2
    deck.move_slide("s:261", to=0)
    assert _ids(deck)[0] == 261
    deck.delete_slide(f"s:{copy.slide_id}")
    assert copy.slide_id not in _ids(deck)


def test_move_slide_keeps_its_old_form_and_refuses_a_position_as_the_slide():
    deck = Document.open(SAMPLE)
    moved = deck.move_slide(deck.slides[-1], 0)          # the old positional form
    assert moved.slide_id == 261 and moved.index == 0
    assert deck.move_slide(258, to=5).index == 5           # an int is an sldId
    with pytest.raises(KeyError, match="not a position"):
        deck.move_slide(2, to=0)
    with pytest.raises(TypeError, match="to=n"):
        deck.move_slide(deck.slides[0])
    with pytest.raises(TypeError, match="not both"):
        deck.move_slide(deck.slides[0], 1, to=2)
    with pytest.raises(IndexError, match="0-based"):
        deck.move_slide(deck.slides[0], to=6)


def test_outline_slide_numbers_say_they_are_one_based():
    deck = Document.open(SAMPLE)
    with pytest.raises(IndexError, match=r"1-based, so slide 1 is deck\.slides\[0\]"):
        deck.to_outline(slides=[0])
    with pytest.raises(IndexError, match=r"257 is an sldId: pass 's:257'"):
        deck.to_outline(slides=[257])
    by_number = deck.to_outline(slides=[2])
    assert deck.to_outline(slides=[deck.slides[1]]) == by_number
    assert deck.to_outline(slides=["s:257"]) == by_number
    assert deck.outline_blocks(slides="s:257") == deck.outline_blocks(slides=[2])


def test_rendering_takes_the_same_slide_names():
    pytest.importorskip("pptx2svg")
    deck = Document.open(SAMPLE)
    assert deck.render_svg(["s:257"]) == deck.render_svg([2])
    assert deck._numbers([deck.slides[3], "s:256"]) == [4, 1]


def test_edit_calls_return_what_they_edited():
    deck = Document.open(REPORT)
    frame = deck.shape("256.5").text_frame
    assert isinstance(frame.delete_paragraph(1), TextFrame)
    paragraph = frame.paragraph(0)
    paragraph.add_run("（予定）")
    assert isinstance(paragraph.delete_run(-1), Paragraph)
    assert isinstance(deck.shape("256.3").line.clear(), LineFormat)
    cell = deck.shape("257.3#5").table.cell(1, 1)
    assert isinstance(cell.set_text("4,310億円"), TableCell) and cell.text == "4,310億円"
    assert deck.set_slide_size("4:3") is deck
    assert deck.set_slide_size("4:3") is deck              # unchanged: still the deck


# -- the rest of trial 2's #16 ------------------------------------------------------------------


def test_slide_index_reads_as_a_property_or_a_call():
    from pptx_agent import Document

    deck = Document.new()
    deck.add_slide("Blank")
    second = deck.add_slide("Blank")
    assert second.index == 1 and second.index() == 1 and isinstance(second.index(), int)
    assert deck.slides[second.index] .slide_id == second.slide_id
    assert repr(second.index) == "1"


def test_a_layout_shape_and_a_shape_spell_their_frame_alike():
    from pptx_agent import Document

    deck = Document.new()
    layout = deck.layout("Title and Content")
    placeholder = layout.placeholders[0]
    assert (placeholder.left, placeholder.top, placeholder.width, placeholder.height) == \
        placeholder.bounds
    slide = deck.add_slide(layout)
    title = slide.shapes[0]
    title.left = title.left                                   # materialise its frame
    assert title.bounds == (title.left, title.top, title.width, title.height) == \
        placeholder.bounds
    for path in sorted(FIXTURE_DIR.glob("*.pptx")):
        for each in Document.open(path).layouts:
            for shape in each.shapes:
                frame = (shape.left, shape.top, shape.width, shape.height)
                assert frame == (shape.bounds or (None,) * 4)


def test_format_names_its_keys():
    import pytest

    from pptx_agent import Document
    from pptx_agent.edit.text import FORMAT_KEYS

    deck = Document.new()
    frame = deck.add_slide("Blank").add_textbox(0, 0, 914400, 914400, "x").text_frame
    for target in (frame, frame.paragraph(0), frame.paragraph(0).run(0)):
        with pytest.raises(TypeError, match="bold, italic") as raised:
            target.format(weight="bold")
        assert "space_after" in str(raised.value)               # where paragraph ones go
        for key in FORMAT_KEYS - {"font"}:
            assert f"``{key}``" in type(target).format.__doc__
    empty = deck.add_slide("Blank").add_shape("rect", 0, 0, 914400, 914400).text_frame
    with pytest.raises(TypeError):
        empty.format(colour="accent1")                          # refused with no runs too
    with pytest.raises(TypeError):
        frame.add_paragraph("y", weight="bold")
    frame.paragraph(0).run(0).format(font="Georgia")
    assert frame.paragraph(0).run(0).typeface == "Georgia"
