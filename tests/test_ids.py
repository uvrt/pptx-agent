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


# ------------------------------------------------------------------------------------------
# New ids are never an earlier shape's or slide's while the deck is open.  The shape rule is
# PowerPoint's, measured with PowerPoint for Mac on 2026-10-15
# (agent-work/feedback-round5/oracle/results.md, A1); the slide rule deliberately is not
# (A2: PowerPoint reuses a deleted top sldId in the session).
# ------------------------------------------------------------------------------------------


def _raw(shape) -> int:
    return int(shape.id.split(".")[1])


def _deck_with(raw_ids: list[int]) -> Document:
    """A deck whose one slide has shapes with exactly these ``cNvPr@id``s, reopened so
    that only the file counts."""
    deck = Document.new()
    slide = deck.add_slide("Blank")
    shapes = [slide.add_shape("rect", 0, 0, 100, 100) for _ in range(max(raw_ids) - 1)]
    for shape in shapes:
        if _raw(shape) not in raw_ids:
            shape.delete()
    deck = Document.open(deck.to_bytes())
    assert sorted(_raw(shape) for shape in deck.slides[0].shapes) == raw_ids
    return deck


def _add(slide, count: int = 1) -> list[int]:
    return [_raw(slide.add_shape("rect", 0, 0, 100, 100)) for _ in range(count)]


@pytest.mark.parametrize("raw_ids, deleted, added, expected", [
    ([2, 3, 4, 5, 6], [], 1, [7]),
    ([2, 3, 4, 5, 6], [6], 1, [7]),
    ([2, 3, 4, 5, 6], [6], 2, [7, 8]),
    ([2, 3, 4, 5, 6], [4], 1, [7]),
    ([2, 3, 7], [], 1, [4]),
    ([2, 3, 7], [], 4, [4, 5, 6, 8]),
    ([2, 3, 7], [3], 1, [4]),
    ([2, 5, 6], [], 2, [3, 4]),
    ([3, 4, 5], [], 2, [2, 6]),
    ([10, 20, 30], [], 2, [2, 3]),
    ([2, 3, 100], [], 4, [4, 5, 6, 7]),
    ([2, 3, 4, 5, 6], [3, 4, 5, 6], 1, [7]),
])
def test_new_shape_ids_follow_powerpoint(raw_ids, deleted, added, expected):
    """Each row is one of the measured runs (results.md, A1)."""
    deck = _deck_with(raw_ids)
    slide = deck.slides[0]
    for raw in deleted:
        deck.shape(f"{slide.slide_id}.{raw}").delete()
    assert _add(slide, added) == expected


def test_after_a_save_and_reopen_only_the_file_counts():
    deck = _deck_with([2, 3, 4, 5, 6])
    slide = deck.slides[0]
    for raw in (3, 4, 5, 6):
        deck.shape(f"{slide.slide_id}.{raw}").delete()
    reopened = Document.open(deck.to_bytes())
    assert _add(reopened.slides[0]) == [3]


def test_an_undone_shape_keeps_its_id():
    deck = _deck_with([2, 3])
    slide = deck.slides[0]
    assert _add(slide) == [4]
    deck.undo()
    assert _add(deck.slides[0]) == [5]
    deck.undo()
    deck.redo()                         # the redone shape is 5 again ...
    assert sorted(_raw(s) for s in deck.slides[0].shapes) == [2, 3, 5]
    assert _add(deck.slides[0]) == [6]  # ... and 4 is still not handed out


def test_a_rolled_back_batch_frees_no_id():
    deck = _deck_with([2, 3])
    with pytest.raises(RuntimeError):
        with deck.batch():
            _add(deck.slides[0], 2)
            raise RuntimeError("the batch fails")
    assert sorted(_raw(s) for s in deck.slides[0].shapes) == [2, 3]
    assert _add(deck.slides[0]) == [6]


def test_a_rolled_back_savepoint_frees_no_id():
    deck = _deck_with([2, 3])
    with deck.batch():
        with pytest.raises(RuntimeError):
            with deck.history.savepoint():
                _add(deck.slides[0])
                raise RuntimeError("one edit fails")
        deck._reset_caches()
        assert _add(deck.slides[0]) == [5]


def test_copies_and_duplicates_take_fresh_ids():
    deck = _deck_with([2, 3, 4])
    slide = deck.slides[0]
    deck.shape(f"{slide.slide_id}.4").delete()
    copy = deck.shape(f"{slide.slide_id}.3").duplicate()
    assert _raw(copy) == 5
    from pptx_agent.edit.document import copy_of

    twin = copy_of(deck)                # a scratch copy carries on the deck's session
    assert _add(twin.slides[0]) == _add(deck.slides[0]) == [6]


def test_layout_shapes_never_reuse_an_id_either():
    deck = Document.new()
    layout = deck.layout("Blank")
    first = layout.add_shape("rect", 0, 0, 100, 100)
    number = int(first.id.split(".")[1])
    first.delete()
    second = layout.add_shape("rect", 0, 0, 100, 100)
    assert int(second.id.split(".")[1]) == number + 1


def test_a_deleted_slides_id_is_not_reused():
    """Unlike PowerPoint (results.md, A2), which hands a deleted top sldId out again."""
    deck = Document.new()
    for _ in range(3):
        deck.add_slide("Blank")
    deck = Document.open(deck.to_bytes())
    assert [s.slide_id for s in deck.slides] == [256, 257, 258]
    deck.delete_slide(258)
    assert deck.add_slide("Blank").slide_id == 259
    deck.delete_slide(257)
    assert deck.add_slide("Blank").slide_id == 260
    deck.undo()                         # the undone slide's id stays taken
    assert deck.add_slide("Blank").slide_id == 261
    copy = deck.duplicate_slide(256)
    assert copy.slide_id == 262
    assert [s.slide_id for s in deck.slides] == [256, 262, 259, 261]


def test_slide_ids_after_a_reopen_follow_the_file():
    deck = Document.new()
    for _ in range(3):
        deck.add_slide("Blank")
    deck.delete_slide(258)
    deck.delete_slide(257)
    reopened = Document.open(deck.to_bytes())
    assert reopened.add_slide("Blank").slide_id == 257   # as PowerPoint after a reopen


def test_a_rolled_back_batch_frees_no_slide_id():
    deck = Document.new()
    deck.add_slide("Blank")
    with pytest.raises(RuntimeError):
        with deck.batch():
            deck.add_slide("Blank")
            raise RuntimeError("the batch fails")
    assert deck.add_slide("Blank").slide_id == 258


def test_slide_ids_stay_in_the_schemas_range():
    from lxml import etree

    from pptx_agent.edit.slides import MAX_SLIDE_ID, new_slide_id

    p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    root = etree.fromstring(
        f'<p:presentation xmlns:p="{p}"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst>'
        '</p:presentation>'.encode())
    assert new_slide_id(root, above=300) == 301
    assert new_slide_id(root, above=MAX_SLIDE_ID) == 257   # past the top: lowest free
    assert MAX_SLIDE_ID == 2**31 - 1                       # ST_SlideId: maxExclusive 2^31
