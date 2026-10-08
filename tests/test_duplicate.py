"""Shape.duplicate on groups: every shape in the copy gets a fresh id, and a connector inside
the copied group stays glued to the copies of its shapes, not to the originals."""

from __future__ import annotations

from pptx_agent import Document, Pt
from pptx_agent.edit.creating import connection
from pptx_agent.edit.ids import cnv_pr
from lxml import etree


def _group_with_connector():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    first = slide.add_shape("rect", Pt(72), Pt(72), Pt(100), Pt(50), text="A")
    second = slide.add_shape("rect", Pt(272), Pt(72), Pt(100), Pt(50), text="B")
    link = slide.add_connector("straight", (first, "right"), (second, "left"))
    group = slide.group([first, second, link])
    return deck, slide, group


def _raw_ids(slide):
    return [node.get("id") for node in slide._sp_tree().iter("{*}cNvPr")]


def test_a_duplicated_group_numbers_every_member_afresh():
    deck, slide, group = _group_with_connector()
    before = _raw_ids(slide)
    copy = group.duplicate(dy=Pt(150))
    after = _raw_ids(slide)
    assert len(after) == len(set(after)), f"repeated cNvPr ids: {after}"
    assert len(after) == len(before) + 4
    originals = {child.id for child in group.children}
    copies = {child.id for child in copy.children}
    assert not originals & copies
    assert all("#" not in shape.id for shape in slide.shapes)


def test_a_connector_inside_a_duplicated_group_is_glued_to_the_copies():
    deck, slide, group = _group_with_connector()
    copy = group.duplicate(dy=Pt(150))
    members = {child.kind: [] for child in copy.children}
    for child in copy.children:
        members[child.kind].append(child)
    (link,) = members["connector"]
    raw = {cnv_pr(child._element).get("id") for child in members["shape"]}
    begin, end = connection(link._element, "begin"), connection(link._element, "end")
    assert begin[0] in raw and end[0] in raw
    assert link.begin_connection[0].id in {child.id for child in members["shape"]}
    # moving a copied box re-routes the copied connector, not the original
    original_link = next(c for c in group.children if c.kind == "connector")
    original_route = original_link.route
    members["shape"][1].move_by(dy=Pt(40))
    assert original_link.route == original_route
    assert link.route != original_route


def test_a_glue_to_a_shape_outside_the_copy_is_kept():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    hub = slide.add_shape("ellipse", Pt(300), Pt(300), Pt(60), Pt(60))
    box = slide.add_shape("rect", Pt(72), Pt(72), Pt(100), Pt(50))
    spoke = slide.add_connector("straight", (box, "bottom"), (hub, "top"))
    group = slide.group([box, spoke])
    copy = group.duplicate(dx=Pt(150))
    link = next(c for c in copy.children if c.kind == "connector")
    assert link.end_connection[0].id == hub.id
    assert link.begin_connection[0].id != box.id


def test_the_duplicate_round_trips_through_validate():
    deck, slide, group = _group_with_connector()
    group.duplicate(dy=Pt(150))
    assert Document.open(deck.to_bytes()).validate() == []
