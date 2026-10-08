"""Groups: children are addressable, and move within the group's child coordinate space."""

from __future__ import annotations

import pytest

from pptx_agent import Document
from pptx_agent.oxml.xml import qn, set_int

EMU_PER_INCH = 914400


def reopen(document: Document) -> Document:
    return Document.open(document.to_bytes())


def _bounds(document: Document, ids) -> dict[str, tuple]:
    return {i: document.shape(i).slide_bounds for i in ids}


def _scale(group, factor_x: float, factor_y: float) -> None:
    """Make a group's child space differ from its frame, as a resized group's does."""
    xfrm = group._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    ext = xfrm.find(qn("a:ext"))
    set_int(ext, "cx", round(int(ext.get("cx")) * factor_x))
    set_int(ext, "cy", round(int(ext.get("cy")) * factor_y))
    group._slide._touch()


def test_grouping_keeps_everything_in_place(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    members = slide.shapes[4:8]
    ids = [m.id for m in members]
    before = _bounds(document, ids)

    group = slide.group(members, name="Feature card")

    reread = reopen(document)
    assert reread.shape(group.id).name == "Feature card"
    assert [c.id for c in reread.shape(group.id).children] == ids
    assert _bounds(reread, ids) == before
    for identifier in ids:
        assert reread.shape(identifier).parent_group.id == group.id


def test_grouping_keeps_path_derived_ids(financial_report):
    """Duplicate cNvPr ids are addressed by tree path; grouping changes the path, not the id."""
    document = Document.open(str(financial_report))
    slide = document.slide(257)
    duplicated = [s for s in slide.shapes if "#" in s.id]
    assert duplicated
    ids = [s.id for s in duplicated]

    slide.group(duplicated)

    reread = reopen(document)
    for identifier in ids:
        assert reread.shape(identifier).parent_group is not None


def test_group_takes_the_frontmost_members_z_position(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    top_level = [s for s in slide.shapes if s.parent_group is None]
    members = [top_level[1], top_level[3]]
    group = slide.group(members)
    order = [s.id for s in slide.shapes if s.parent_group is None]
    assert order.index(group.id) == 2  # where the frontmost member (index 3) was, less one


def test_grouping_is_one_undo_step(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()
    slide = document.slides[0]
    slide.group(slide.shapes[1:4])
    assert document.undo()
    assert document.to_bytes() == original


def test_grouping_refuses_mixed_containers(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[1:3])
    outsider = next(s for s in slide.shapes if s.parent_group is None and s.id != group.id)
    with pytest.raises(ValueError, match="share a container"):
        slide.group([group.children[0], outsider])


def test_child_moves_in_child_space(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    _scale(group, 2.0, 0.5)
    child = group.children[1]
    left = child.left
    x_before = child.slide_bounds[0]

    child.move_by(dx=100000)  # child units: twice that on the slide

    reread = reopen(document).shape(child.id)
    assert reread.left == left + 100000
    assert reread.slide_bounds[0] - x_before == pytest.approx(200000, abs=2)


def test_child_moves_in_slide_space(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    _scale(group, 2.0, 0.5)
    child = group.children[2]
    x, y = child.slide_bounds[:2]

    child.move_by(dx=EMU_PER_INCH, dy=EMU_PER_INCH, space="slide")

    moved = reopen(document).shape(child.id).slide_bounds
    assert moved[0] - x == pytest.approx(EMU_PER_INCH, abs=2)
    assert moved[1] - y == pytest.approx(EMU_PER_INCH, abs=2)


def test_moving_a_child_refits_the_group_without_moving_siblings(product_page):
    """PowerPoint's behaviour: the group frame grows to the children, nothing else moves."""
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    _scale(group, 1.5, 1.5)
    children = group.children
    siblings = [c.id for c in children[1:]]
    before = _bounds(document, siblings)

    children[0].move_by(dx=-2 * EMU_PER_INCH, space="slide")

    reread = reopen(document)
    after = _bounds(reread, siblings)
    for identifier in siblings:
        assert all(abs(a - b) <= 2 for a, b in zip(after[identifier], before[identifier]))
    frame = reread.shape(group.id)
    moved = reread.shape(children[0].id).slide_bounds
    assert frame.left == pytest.approx(moved[0], abs=2)  # the group grew left to contain it
    offset, extent = frame.child_offset, frame.child_extent
    xs = [reread.shape(c.id).left for c in children]
    rights = [reread.shape(c.id).left + reread.shape(c.id).width for c in children]
    assert offset[0] == min(xs) and offset[0] + extent[0] == max(rights)


def test_nested_groups_refit_upwards(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    inner = slide.group(slide.shapes[4:6])
    outer = slide.group([inner, slide.shape(slide.shapes[7].id)])
    _scale(outer, 0.5, 0.5)
    leaf = document.shape(inner.id).children[0]
    others = [s.id for s in slide.shapes if s.id not in {leaf.id, inner.id, outer.id}]
    before = _bounds(document, others)

    leaf.move_by(dx=3 * EMU_PER_INCH, space="slide")

    reread = reopen(document)
    after = _bounds(reread, others)
    for identifier in others:
        assert all(abs(a - b) <= 2 for a, b in zip(after[identifier], before[identifier])), identifier
    leaf_box = reread.shape(leaf.id).slide_bounds
    outer_box = reread.shape(outer.id).slide_bounds
    assert outer_box[0] + outer_box[2] >= leaf_box[0] + leaf_box[2] - 2


def test_a_rotated_group_is_not_refit(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    group.rotation = 30
    offset = group.child_offset
    group.children[0].move_by(dx=-EMU_PER_INCH)
    assert reopen(document).shape(group.id).child_offset == offset


def test_children_resolve_by_address(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    texts = [c for c in group.children if c.kind == "shape" and c.text]
    run = texts[0].text_frame.paragraph(0).run(0)
    run.bold = True
    assert reopen(document).resolve(run.address).bold is True


def test_group_ids_reach_the_render(product_page):
    pytest.importorskip("pptx2svg")
    import re

    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[4:8])
    group.children[0].move_by(dx=EMU_PER_INCH)
    rendered = set(re.findall(r'data-pptx-id="([^"]*)"', slide.render_svg()))
    known = {s.id for s in slide.shapes}
    assert {c.id for c in group.children} <= rendered
    assert rendered - {i for i in rendered if i.startswith(("lay:", "mst:"))} <= known


def test_group_fill(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    group = slide.group(slide.shapes[1:3])
    group.fill = "accent5"
    assert reopen(document).shape(group.id).fill.color == "accent5"


def test_slide_shapes_lists_members_after_their_group():
    """What Slide.shapes' docstring says: every shape, each group followed by its members."""
    from pptx_agent import Document

    deck = Document.new()
    slide = deck.add_slide("Blank")
    a = slide.add_shape("rect", 0, 0, 914400, 914400)
    b = slide.add_shape("rect", 914400, 0, 914400, 914400)
    c = slide.add_shape("rect", 0, 914400, 914400, 914400)
    group = slide.group([a, b])
    ids = [shape.id for shape in slide.shapes]
    assert ids.index(group.id) + 1 == ids.index(a.id) and ids.index(b.id) == ids.index(a.id) + 1
    assert c.id in ids
    assert {s.id for s in slide.shapes if s.parent_group is None} == {group.id, c.id}
