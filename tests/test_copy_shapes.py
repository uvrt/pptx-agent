"""Slide.copy_shapes (LP23): shapes and groups to another slide or another deck."""

from __future__ import annotations

from pathlib import Path

import pytest

from pptx_agent import Document
from pptx_agent.oxml.xml import qn

from images import png

FIXTURES = Path(__file__).resolve().parent / "fixtures"
E = 12700


def _deck():
    deck = Document.new()
    deck.add_slide("Title Only")
    deck.add_slide("Title Only")
    first = deck.slides[0]
    a = first.add_shape("rect", 100 * E, 150 * E, 100 * E, 50 * E)
    a.text = "A {{n}}"
    b = first.add_shape("rect", 300 * E, 150 * E, 100 * E, 50 * E)
    b.text = "B"
    link = first.add_connector("straight", (a, "right"), (b, "left"))
    picture = first.add_picture(png(4, 4, lambda x, y: (200, 30, 30)), 100 * E, 250 * E,
                                40 * E, 40 * E)
    group = first.group([a, b, link, picture])
    outside = first.add_shape("ellipse", 500 * E, 300 * E, 50 * E, 50 * E)
    member_b = deck.shape(group.id).children[1]
    loose = first.add_connector("straight", (member_b, "bottom"), (outside, "top"))
    return deck, deck.shape(group.id), deck.shape(loose.id)


def _raw_ids(slide):
    return [node.get("id") for node in slide._sp_tree().iter(qn("p:cNvPr"))]


def test_a_group_copied_to_another_slide_gets_fresh_ids_and_keeps_its_inner_glue():
    deck, group, loose = _deck()
    first, second = deck.slides
    result = first.copy_shapes([group, loose], second, at=(50 * E, 120 * E))
    assert set(result.mapping) == {group.id, *(c.id for c in group.children), loose.id}
    copies = [deck.shape(address) for address in result.mapping.values()]
    assert all(c._slide.slide_id == second.slide_id for c in copies)
    ids = _raw_ids(second)
    assert len(ids) == len(set(ids))
    new_group = deck.shape(result.mapping[group.id])
    assert new_group.slide_bounds[:2] == (50 * E, 120 * E)
    inner = deck.shape(result.mapping[group.children[2].id])
    begin, end = inner.begin_connection, inner.end_connection
    assert begin[0].id == result.mapping[group.children[0].id]
    assert end[0].id == result.mapping[group.children[1].id]
    # glue to the ellipse, which is not copied, is dropped and reported
    assert result.dropped_glue == [{"connector": result.mapping[loose.id], "end": "end",
                                    "was_glued_to": result.dropped_glue[0]["was_glued_to"]}]
    assert deck.shape(result.mapping[loose.id]).end_connection is None
    assert deck.validate() == []


def test_a_copy_on_the_same_slide_is_offset_and_its_members_are_renumbered():
    deck, group, _ = _deck()
    first = deck.slides[0]
    result = first.copy_shapes([group], dy=200 * E)
    ids = _raw_ids(first)
    assert len(ids) == len(set(ids))
    old, new = group.slide_bounds, deck.shape(result.mapping[group.id]).slide_bounds
    assert (new[0], new[1] - old[1]) == (old[0], 200 * E)
    # the picture's image is shared, not stored twice
    assert [n for n in deck.package.part_names if n.startswith("ppt/media/")] == \
        ["ppt/media/image1.png"]


def test_a_copy_into_another_deck_imports_the_image_and_is_one_undo_step():
    deck, group, _ = _deck()
    other = Document.new()
    other.add_slide("Title Only")
    before = other.to_bytes()
    source_before = deck.to_bytes()
    result = deck.slides[0].copy_shapes([group], other.slides[0], at=(0, 100 * E))
    picture = other.shape(result.mapping[group.children[3].id])
    assert picture.kind == "picture" and picture.image_part.startswith("ppt/media/")
    assert other.validate() == []
    assert deck.to_bytes() == source_before
    assert Document.open(other.to_bytes()).validate() == []
    other.undo()
    assert other.to_bytes() == before


def test_a_member_copied_out_of_its_group_lands_where_it_was_drawn():
    deck, group, _ = _deck()
    member = group.children[1]
    result = deck.slides[0].copy_shapes([member], deck.slides[1])
    assert deck.shape(result.mapping[member.id]).slide_bounds == member.slide_bounds


def test_a_placeholder_becomes_a_plain_shape_with_its_geometry_and_sizes():
    deck, _, _ = _deck()
    title = deck.slides[0].shapes[0]
    title.text = "Plan"
    assert title.placeholder is not None
    other = Document.new()
    other.add_slide("Blank")
    result = deck.slides[0].copy_shapes([title], other.slides[0])
    copy = other.shape(result.mapping[title.id])
    assert copy.placeholder is None and result.plain_placeholders == [title.id]
    assert copy.slide_bounds == title.slide_bounds
    assert copy.text_frame.paragraphs[0].runs[0].size == title.text_frame.paragraphs[0].runs[0].effective_size


def test_charts_and_smartart_are_copied_with_their_parts_into_another_deck():
    other = Document.new()
    other.add_slide("Blank")
    report = Document.open(FIXTURES / "real-financial-report.pptx")
    chart = report.shape("257.25")
    copied = chart._slide.copy_shapes([chart], other.slides[0])
    smart = Document.open(FIXTURES / "powerpoint-smartart.pptx").shape("256.3")
    diagram = smart._slide.copy_shapes([smart], other.slides[0], at=(0, 0))
    assert other.shape(copied.mapping["257.25"]).has_chart
    nodes = other.shape(diagram.mapping["256.3"]).diagram.nodes
    assert [n.text for n in nodes] == [n.text for n in smart.diagram.nodes]
    assert Document.open(other.to_bytes()).validate() == []
    # within one deck a chart is per-shape state: a second chart part, not a shared one
    again = chart._slide.copy_shapes([chart], dx=E)
    parts = {report.package.related_part(chart._slide.part_path, rel)
             for rel in report.package.relationships(chart._slide.part_path)}
    assert report.shape(again.mapping["257.25"]).has_chart and len(parts) >= 3


def test_a_link_to_another_slide_is_refused_across_decks_and_nothing_changes():
    deck, group, _ = _deck()
    first, second = deck.slides
    shape = first.add_shape("rect", 0, 0, 50 * E, 50 * E)
    rel = deck.package.add_relationship(first.part_path,
                                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                                        second.part_path)
    from pptx_agent.oxml.xml import make
    properties = shape._element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
    properties.append(make("a:hlinkClick", r__id=rel, action="ppaction://hlinksldjump"))
    other = Document.new()
    other.add_slide("Blank")
    before = other.to_bytes()
    with pytest.raises(ValueError, match="another slide"):
        first.copy_shapes([deck.shape(shape.id)], other.slides[0])
    assert other.to_bytes() == before


def test_theme_colours_the_target_draws_differently_are_reported():
    deck, _, _ = _deck()
    shape = deck.slides[0].add_shape("rect", 0, 0, 50 * E, 50 * E)
    shape.fill = "accent2"
    other = Document.new()
    other.add_slide("Blank")
    other.theme.set_colors({"accent2": "#FF0000"})
    result = deck.slides[0].copy_shapes([shape], other.slides[0])
    assert result.theme_changes == [{"color": "accent2", "was": deck.theme.colors["accent2"],
                                     "now": "#FF0000"}]
