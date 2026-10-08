"""Shape identity: unique, stable across edits, and durable across a save."""

from __future__ import annotations

import pytest

from pptx_agent import Document
from pptx_agent.edit.ids import ShapeId


def test_every_shape_has_a_unique_id(pptx_path):
    document = Document.open(str(pptx_path))
    identifiers = [shape.id for slide in document.slides for shape in slide.shapes]
    assert identifiers
    assert len(identifiers) == len(set(identifiers))


def test_duplicate_cnvpr_ids_are_disambiguated(financial_report):
    """The case the obvious implementation gets wrong.

    In this deck a ``p:sp`` and a ``p:graphicFrame`` share ``id="3"`` on three slides.  Keying
    off ``cNvPr@id`` alone would silently collapse them into one address.
    """
    document = Document.open(str(financial_report))
    slide = document.slide(257)

    identifiers = [shape.id for shape in slide.shapes]
    assert len(identifiers) == len(set(identifiers))

    disambiguated = [i for i in identifiers if "#" in i]
    assert len(disambiguated) == 2

    # Both must still resolve, and to *different* shapes.
    shapes = [document.shape(i) for i in disambiguated]
    assert shapes[0].kind != shapes[1].kind


def test_ids_resolve_from_the_document_without_knowing_the_slide(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        for shape in slide.shapes:
            assert document.shape(shape.id).id == shape.id


def test_an_id_survives_deleting_another_shape(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    shapes = slide.shapes
    doomed, survivor = shapes[0], shapes[-1]
    position = survivor.left

    doomed.delete()

    assert document.shape(survivor.id).left == position


def test_an_id_survives_a_save_and_reopen(product_page):
    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[2]
    identifier = shape.id
    shape.move_by(dx=914400)
    expected = shape.left

    reopened = Document.open(document.to_bytes())
    assert reopened.shape(identifier).left == expected


def test_mutation_stamps_a_durable_id(product_page):
    """An edited shape keeps its address even if PowerPoint later renumbers cNvPr ids."""
    from pptx_agent.edit.ids import cnv_pr, read_stamp

    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[1]
    identifier = shape.id
    assert read_stamp(cnv_pr(shape._element)) is None

    shape.rotation = 3

    stamped = read_stamp(cnv_pr(shape._element))
    assert stamped == ShapeId.parse(identifier).local

    # Simulate PowerPoint renumbering: the stamp, not cNvPr@id, now answers.
    cnv_pr(shape._element).set("id", "9999")
    reopened = Document.open(document.to_bytes())
    assert reopened.shape(identifier).rotation == 3


def test_reading_a_deck_never_stamps_it(pptx_path):
    """Stamping happens on first mutation only, or the losslessness gate would fail."""
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        _ = [shape.id for shape in slide.shapes]
    assert document.package.dirty_parts == frozenset()


def test_a_duplicate_gets_a_new_identity(product_page):
    document = Document.open(str(product_page))
    original = document.slides[0].shapes[2]
    copy = original.duplicate(dx=100)

    assert copy.id != original.id
    assert document.shape(original.id).left != copy.left


def test_malformed_ids_are_rejected():
    with pytest.raises(ValueError):
        ShapeId.parse("no-dot")
    with pytest.raises(ValueError):
        ShapeId.parse("notanumber.3")


def test_shape_id_round_trips_through_text():
    assert ShapeId.parse(str(ShapeId(256, "3#1-2"))) == ShapeId(256, "3#1-2")


def test_powerpoint_creation_id_is_preferred(product_page):
    """PowerPoint writes ``a16:creationId/@id``; that GUID addresses the shape, renumbering or not."""
    from lxml import etree

    from pptx_agent.edit.ids import cnv_pr
    from pptx_agent.oxml.xml import CREATION_ID_EXT_URI, PRESENTATION_NAMESPACES, qn

    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[1]
    properties = cnv_pr(shape._element)
    ext_list = properties.find(qn("a:extLst"))
    if ext_list is None:
        ext_list = etree.SubElement(properties, qn("a:extLst"))
    extension = etree.SubElement(ext_list, qn("a:ext"), uri=CREATION_ID_EXT_URI)
    etree.SubElement(extension, qn("a16:creationId"), nsmap={"a16": PRESENTATION_NAMESPACES["a16"]},
                     id="{0E6F5A2B-1C3D-4E5F-8A9B-0C1D2E3F4A5B}")
    properties.set("id", "9999")
    document.package.mark_dirty(document.slides[0].part_path)

    reopened = Document.open(document.to_bytes())
    identifier = f"{ShapeId.parse(shape.id).slide_id}.c0E6F5A2B-1C3D-4E5F-8A9B-0C1D2E3F4A5B"
    assert reopened.shape(identifier).name == shape.name
