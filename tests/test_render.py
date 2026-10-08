"""Rendering, and the identity contract with pptx2svg.

The agent loop only works if the ids in the picture are the ids the API answers to.  These
tests pin that correspondence, since it spans two repositories and would otherwise drift
silently.
"""

from __future__ import annotations

import re

import pytest

from pptx_agent import Document

pptx2svg = pytest.importorskip("pptx2svg", reason="rendering needs pptx-agent[render]")

EMU_PER_INCH = 914400
PX_PER_INCH = 96


def _ids(svg: str) -> list[str]:
    return re.findall(r'data-pptx-id="([^"]*)"', svg)


def _transform(svg: str, identifier: str) -> str | None:
    match = re.search(
        r'data-pptx-id="%s"[^>]*?transform="([^"]*)"' % re.escape(identifier), svg
    )
    return match.group(1) if match else None


def test_every_rendered_id_resolves_to_a_shape(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        known = {shape.id for shape in slide.shapes}
        for identifier in _ids(slide.render_svg()):
            if identifier.startswith(("lay:", "mst:")):
                continue  # inherited decoration, not addressable on the slide
            # A drawn piece of a shape -- SmartArt's cached shapes -- is "<shape>/<path>".
            identifier = identifier.partition("/")[0]
            assert identifier in known, f"{identifier} is not a shape on slide {slide.slide_id}"


def test_rendered_ids_are_unique(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        identifiers = _ids(slide.render_svg())
        assert len(identifiers) == len(set(identifiers))


def test_duplicate_cnvpr_ids_stay_distinct_in_the_svg(financial_report):
    """pptx2svg reports the raw cNvPr id; the rewrite must split the collision apart.

    Both shapes share ``cNvPr@id="3"``, so pptx2svg emits ``257.3`` twice and the rewrite has
    to turn that into two different addresses.
    """
    document = Document.open(str(financial_report))
    slide = document.slide(257)
    identifiers = _ids(slide.render_svg())

    collided = [i for i in identifiers if i.startswith("257.3#")]
    assert len(collided) == 2
    assert len(set(collided)) == 2
    # The raw, ambiguous form must not survive the rewrite.
    assert "257.3" not in identifiers


def test_moving_a_shape_moves_its_group(product_page):
    """The loop an agent actually runs: edit by id, re-render, see the change."""
    document = Document.open(str(product_page))
    slide = document.slides[0]
    shape = slide.shapes[2]

    before = _transform(slide.render_svg(), shape.id)
    shape.move_by(dx=EMU_PER_INCH)
    after = _transform(slide.render_svg(), shape.id)

    assert before is not None and after is not None
    x_before = float(re.search(r"translate\(([-\d.]+)", before).group(1))
    x_after = float(re.search(r"translate\(([-\d.]+)", after).group(1))
    assert x_after - x_before == pytest.approx(PX_PER_INCH, abs=1)


def test_edited_text_reaches_the_render(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    shape = next(s for s in slide.shapes if s.text)

    shape.set_text("Sentinel phrase")

    assert "Sentinel phrase" in slide.render_svg()


def test_raw_ids_are_available_when_asked_for(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    assert _ids(slide.render_svg(rewrite_ids=False)) == _ids(slide.render_svg())


def test_render_produces_well_formed_svg(pptx_path):
    from xml.etree.ElementTree import fromstring

    document = Document.open(str(pptx_path))
    for slide in document.slides:
        fromstring(slide.render_svg())


def test_rendering_does_not_modify_the_document(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        slide.render_svg()
    assert document.package.dirty_parts == frozenset()


def test_ids_still_match_after_the_semantic_edits(pptx_path):
    """After text, fill, table and group edits, the picture and the API still agree on ids."""
    from test_validity import _apply_e1_edits

    document = Document.open(str(pptx_path))
    _apply_e1_edits(document)
    for slide in document.slides:
        known = {shape.id for shape in slide.shapes}
        identifiers = [i for i in _ids(slide.render_svg()) if not i.startswith(("lay:", "mst:"))]
        assert len(identifiers) == len(set(identifiers))
        shapes = {i.partition("/")[0] for i in identifiers}  # "<shape>/<piece>" is drawn in it
        assert shapes <= known, shapes - known


def test_edited_runs_reach_the_render(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    shape = next(s for s in slide.shapes if s.text)
    shape.text_frame.paragraph(0).add_run(" Sentinel run", color="#C00000", bold=True)
    svg = slide.render_svg()
    assert "Sentinel run" in svg or "Sentinel" in svg
    assert "#C00000" in svg.upper()
