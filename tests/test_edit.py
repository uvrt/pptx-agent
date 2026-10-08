"""The semantic API: geometry, text and structure."""

from __future__ import annotations

import pytest

from pptx_agent import Document

EMU_PER_INCH = 914400


def test_geometry_reads_back_what_was_written(product_page):
    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[3]

    shape.left = 1_000_000
    shape.top = 2_000_000
    shape.width = 3_000_000
    shape.height = 4_000_000
    shape.rotation = 42.5

    reopened = Document.open(document.to_bytes()).shape(shape.id)
    assert (reopened.left, reopened.top) == (1_000_000, 2_000_000)
    assert (reopened.width, reopened.height) == (3_000_000, 4_000_000)
    assert reopened.rotation == pytest.approx(42.5, abs=0.001)


def test_move_by_is_relative(product_page):
    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[2]
    left, top = shape.left, shape.top

    shape.move_by(dx=EMU_PER_INCH, dy=-EMU_PER_INCH)

    assert shape.left == left + EMU_PER_INCH
    assert shape.top == top - EMU_PER_INCH


def test_rotation_normalises(product_page):
    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[0]
    shape.rotation = 370
    assert shape.rotation == pytest.approx(10, abs=0.001)


def test_setting_one_axis_materialises_both(product_page):
    """A half-written a:off reads as (0, 0) in PowerPoint rather than falling back."""
    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[1]
    top = shape.top

    shape.left = 5000

    reopened = Document.open(document.to_bytes()).shape(shape.id)
    assert reopened.left == 5000
    assert reopened.top == top


def test_text_round_trips(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        for shape in slide.shapes:
            if shape.kind == "shape" and shape.text:
                shape.set_text("replaced")
                assert Document.open(document.to_bytes()).shape(shape.id).text == "replaced"
                return
    pytest.skip("no text-bearing shape in this fixture")


def test_multi_line_text_becomes_multiple_paragraphs(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.text)

    shape.set_text("one\ntwo\nthree")

    assert Document.open(document.to_bytes()).shape(shape.id).text == "one\ntwo\nthree"


def test_set_text_keeps_the_first_run_formatting(product_page):
    from pptx_agent.oxml.xml import qn

    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.text)
    body = shape._element.find(qn("p:txBody"))
    before = body.find(qn("a:p")).find(qn("a:r")).find(qn("a:rPr"))
    assert before is not None, "fixture shape has no run properties to preserve"
    expected = dict(before.attrib)

    shape.set_text("new copy")

    after = body.find(qn("a:p")).find(qn("a:r")).find(qn("a:rPr"))
    assert dict(after.attrib) == expected


def test_text_can_be_added_to_a_shape_that_had_none(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape" and not s.text)

    shape.set_text("added")

    assert Document.open(document.to_bytes()).shape(shape.id).text == "added"


def test_text_on_a_non_text_shape_is_refused(financial_report):
    document = Document.open(str(financial_report))
    frame = next(
        (s for slide in document.slides for s in slide.shapes if s.kind == "graphic_frame"),
        None,
    )
    if frame is None:
        pytest.skip("no graphic frame in this fixture")
    with pytest.raises(ValueError, match="cannot hold text"):
        frame.set_text("nope")


def test_delete_removes_exactly_one_shape(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    before = len(slide.shapes)
    target = slide.shapes[1].id

    slide.shape(target).delete()

    assert len(slide.shapes) == before - 1
    with pytest.raises(KeyError):
        slide.shape(target)


def test_duplicate_offsets_the_copy(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    original = slide.shapes[2]
    left, before = original.left, len(slide.shapes)

    copy = original.duplicate(dx=EMU_PER_INCH)

    assert len(slide.shapes) == before + 1
    assert copy.left == left + EMU_PER_INCH
    assert copy.text == original.text


def test_z_order_moves_the_shape_in_the_tree(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    target = slide.shapes[0].id

    slide.shape(target).bring_to_front()

    assert slide.shapes[-1].id == target


def test_placeholder_inherits_its_position_from_the_layout(pptx_path):
    """A placeholder with no a:xfrm must still report where it will be drawn."""
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        for shape in slide.shapes:
            if shape.placeholder is not None and not shape.has_explicit_transform:
                assert shape.left is not None and shape.width is not None
                return
    pytest.skip("no inheriting placeholder in this fixture")


def test_editing_every_fixture_keeps_it_parseable(pptx_path):
    """Whatever the deck, an edit must produce a file this library can read back."""
    document = Document.open(str(pptx_path))
    edited = 0
    for slide in document.slides:
        for shape in slide.shapes[:3]:
            if shape.left is None:
                continue
            shape.move_by(dx=1000)
            edited += 1
    if not edited:
        pytest.skip("nothing movable in this fixture")

    reopened = Document.open(document.to_bytes())
    assert sum(len(s.shapes) for s in reopened.slides) == sum(
        len(s.shapes) for s in document.slides
    )


def test_flips_read_back_and_undo(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()
    shape = document.slides[0].shapes[2]
    assert (shape.flip_h, shape.flip_v) == (False, False)

    shape.flip_h = True
    shape.flip_v = True
    shape.flip_v = True  # already so: no second undo step

    reopened = Document.open(document.to_bytes()).shape(shape.id)
    assert (reopened.flip_h, reopened.flip_v) == (True, True)
    shape.flip_h = False
    assert document.shape(shape.id).flip_h is False
    while document.undo():
        pass
    assert document.to_bytes() == original


def test_preset_geometry_is_validated_and_drops_adjustments(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.preset not in (None, "custom"))

    shape.preset = "ellipse"

    reopened = Document.open(document.to_bytes()).shape(shape.id)
    assert reopened.preset == "ellipse"
    geometry = reopened._element.find(".//{*}prstGeom")
    assert [child.tag.rpartition("}")[2] for child in geometry] == ["avLst"]
    assert len(geometry.find("{*}avLst")) == 0
    with pytest.raises(ValueError):
        shape.preset = "blob"


def test_groups_have_no_preset(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group([s for s in slide.shapes if s.left is not None][:2])
    assert group.preset is None
    with pytest.raises(ValueError):
        group.preset = "rect"
