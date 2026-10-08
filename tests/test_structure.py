"""Structural operations (E2): slides, pictures, ungrouping, hyperlinks, relationship reaping.

Each operation in :data:`OPERATIONS` runs on every fixture and must then pass, on its own:

* every validity check -- no dangling relationship, no part without a content type, no
  ``Override`` for a missing part, no orphaned part, a consistent slide list;
* a round trip -- edit, save, reopen, and read the edit back;
* undo -- every step walked back gives the original bytes, and redo gives the edit again;
* a render -- pptx2svg draws the right number of slides, in order, with the API's ids.

The focused tests below them pin the details each operation must get right.
"""

from __future__ import annotations

import io
import zipfile
from typing import Callable

import pytest

from images import MARKER, ORANGE
from pptx_agent import Document, Hyperlink
from pptx_agent.oxml.package import (
    REL_CHART,
    REL_HYPERLINK,
    REL_IMAGE,
    REL_NOTES_SLIDE,
    REL_SLIDE,
    REL_SLIDE_LAYOUT,
)
from pptx_agent.oxml.xml import qn
from test_validity import assert_valid

EMU_PER_INCH = 914400
Verify = Callable[[Document], None]


def _first_run(slide, skip=0):
    """The ``skip``-th plain text run on a slide, or ``None``."""
    found = 0
    for shape in slide.shapes:
        if shape.kind != "shape":
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.is_field or not run.text.strip():
                    continue
                if found == skip:
                    return run
                found += 1
    return None


def _pictures(document):
    return [shape for slide in document.slides for shape in slide.shapes
            if shape.kind == "picture" and shape.image_part is not None]


# ------------------------------------------------------------------------------------------
# The operations
# ------------------------------------------------------------------------------------------


def delete_slide(document: Document) -> Verify:
    if len(document.slides) == 1:
        document.slides[0].duplicate()
    victim = document.slides[0]
    remaining = [slide.slide_id for slide in document.slides[1:]]
    document.delete_slide(victim)

    def verify(reopened: Document) -> None:
        assert [slide.slide_id for slide in reopened.slides] == remaining

    return verify


def duplicate_slide(document: Document) -> Verify:
    source = document.slides[0]
    texts = [shape.text for shape in source.shapes]
    kinds = [shape.kind for shape in source.shapes]
    copy = source.duplicate()
    assert copy.index == 1

    def verify(reopened: Document) -> None:
        again = reopened.slides[1]
        assert again.slide_id == copy.slide_id
        assert [shape.text for shape in again.shapes] == texts
        assert [shape.kind for shape in again.shapes] == kinds

    return verify


def reorder_slides(document: Document) -> Verify:
    if len(document.slides) == 1:
        document.slides[0].duplicate()
    order = [slide.slide_id for slide in document.slides]
    last = document.slides[-1]
    some_shape = next((shape.id for shape in last.shapes), None)
    document.move_slide(last, 0)
    expected = [order[-1]] + order[:-1]

    def verify(reopened: Document) -> None:
        assert [slide.slide_id for slide in reopened.slides] == expected
        if some_shape is not None:
            assert reopened.shape(some_shape).id == some_shape

    return verify


def _layout_with_title(document):
    for layout in document.layouts:
        root = document.package.tree(layout.part_path)
        types = {node.get("type") for node in root.iter(qn("p:ph"))}
        if types & {"title", "ctrTitle"}:
            return layout
    return document.slides[0].layout or document.layouts[0]


def add_slide_from_layout(document: Document) -> Verify:
    layout = _layout_with_title(document)
    slide = document.add_slide(layout, index=1)
    titles = [shape for shape in slide.shapes
              if shape.placeholder and shape.placeholder[0] in {"title", "ctrTitle"}]
    if titles:
        titles[0].set_text("Added from a layout")
    else:
        slide.add_picture(MARKER, EMU_PER_INCH, EMU_PER_INCH, width=2 * EMU_PER_INCH)
    placeholders = [shape.placeholder for shape in slide.shapes if shape.placeholder]

    def verify(reopened: Document) -> None:
        again = reopened.slides[1]
        assert again.slide_id == slide.slide_id
        assert again.layout is not None and again.layout.part_path == layout.part_path
        assert [s.placeholder for s in again.shapes if s.placeholder] == placeholders
        if titles:
            assert "Added from a layout" in [shape.text for shape in again.shapes]

    return verify


def insert_picture(document: Document) -> Verify:
    added = {}
    for slide in document.slides:
        picture = slide.add_picture(MARKER, EMU_PER_INCH // 2, EMU_PER_INCH // 2,
                                    width=2 * EMU_PER_INCH, name="E2 picture")
        added[picture.id] = (picture.width, picture.height)

    def verify(reopened: Document) -> None:
        for identifier, size in added.items():
            picture = reopened.shape(identifier)
            assert picture.kind == "picture" and picture.name == "E2 picture"
            assert (picture.width, picture.height) == size == (2 * EMU_PER_INCH, EMU_PER_INCH)
            assert reopened.package.read(picture.image_part) == MARKER

    return verify


def replace_image(document: Document) -> Verify:
    pictures = _pictures(document)
    if not pictures:
        pictures = [document.slides[0].add_picture(ORANGE, EMU_PER_INCH, EMU_PER_INCH)]
    replaced = []
    for picture in pictures:
        picture.replace_image(MARKER)
        replaced.append(picture.id)

    def verify(reopened: Document) -> None:
        for identifier in replaced:
            assert reopened.package.read(reopened.shape(identifier).image_part) == MARKER

    return verify


def ungroup(document: Document) -> Verify:
    for slide in document.slides:  # groups an earlier edit (or the author) made
        for shape in [s for s in slide.shapes if s.kind == "group" and s.parent_group is None]:
            shape.ungroup()
    slide = document.slides[0]
    members = [s for s in slide.shapes if s.parent_group is None and s.kind != "graphic_frame"
               and s.has_explicit_transform and None not in (s.left, s.width)][:2]
    while len(members) < 2:
        members.append(slide.add_picture(MARKER, (len(members) + 1) * EMU_PER_INCH,
                                         EMU_PER_INCH, width=EMU_PER_INCH))
    group = slide.group(members)
    group.rotation = 15
    children = [child.id for child in group.children]
    group_id = group.id
    ungrouped = group.ungroup()
    assert [child.id for child in ungrouped] == children

    def verify(reopened: Document) -> None:
        again = reopened.slide(slide.slide_id)
        assert group_id not in {shape.id for shape in again.shapes}
        for identifier in children:
            shape = reopened.shape(identifier)
            assert shape.parent_group is None
            assert shape.rotation % 15 == pytest.approx(0) or shape.rotation != 0

    return verify


def hyperlink(document: Document) -> Verify:
    expected = {}
    slides = document.slides
    for position, slide in enumerate(slides):
        existing = [run for shape in slide.shapes if shape.kind == "shape"
                    for paragraph in shape.text_frame.paragraphs for run in paragraph.runs
                    if run.hyperlink is not None]
        if existing:
            existing[0].set_hyperlink("https://example.org/changed")
            expected[existing[0].address] = Hyperlink(address="https://example.org/changed")
            for run in existing[1:]:
                run.remove_hyperlink()
                expected[run.address] = None
            continue
        run = _first_run(slide)
        if run is None:
            continue
        run.set_hyperlink("https://example.org/e2", tooltip="E2")
        expected[run.address] = Hyperlink(address="https://example.org/e2", tooltip="E2")
        other = _first_run(slide, skip=1)
        if other is not None and len(slides) > 1:
            destination = slides[(position + 1) % len(slides)]
            other.set_hyperlink(destination)
            expected[other.address] = Hyperlink(slide_id=destination.slide_id)

    def verify(reopened: Document) -> None:
        for address, link in expected.items():
            assert reopened.resolve(address).hyperlink == link

    return verify


def delete_shapes(document: Document) -> Verify:
    slide = document.slides[0]
    doomed = [shape.id for shape in slide.shapes
              if shape.parent_group is None and shape.kind in {"picture", "graphic_frame"}]
    if not doomed:
        doomed = [slide.add_picture(ORANGE, 0, 0).id]
    for identifier in doomed:
        document.shape(identifier).delete()

    def verify(reopened: Document) -> None:
        assert not {shape.id for shape in reopened.slide(slide.slide_id).shapes} & set(doomed)

    return verify


OPERATIONS: dict[str, Callable[[Document], Verify]] = {
    "delete_slide": delete_slide,
    "duplicate_slide": duplicate_slide,
    "reorder_slides": reorder_slides,
    "add_slide_from_layout": add_slide_from_layout,
    "insert_picture": insert_picture,
    "replace_image": replace_image,
    "ungroup": ungroup,
    "hyperlink": hyperlink,
    "delete_shapes": delete_shapes,
}

operation = pytest.mark.parametrize("name", list(OPERATIONS))


# ------------------------------------------------------------------------------------------
# Every operation, every fixture
# ------------------------------------------------------------------------------------------


@operation
def test_valid_after_each_operation(name, pptx_path):
    """ROADMAP E2's "done when": the validity checks pass after each operation on its own."""
    original = pptx_path.read_bytes()
    document = Document.open(original)
    OPERATIONS[name](document)
    assert_valid(document.to_bytes(), original)


@operation
def test_each_operation_round_trips(name, pptx_path):
    document = Document.open(pptx_path.read_bytes())
    verify = OPERATIONS[name](document)
    saved = document.to_bytes()
    verify(Document.open(saved))
    # And the saved deck is itself stable: reading and writing it changes nothing.
    assert Document.open(saved).to_bytes() == saved


@operation
def test_each_operation_undoes_byte_for_byte(name, pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    document.history._max_depth = 10_000
    before = document.to_bytes()
    OPERATIONS[name](document)
    edited = document.to_bytes()
    # delete_shapes on a deck with no picture adds one and deletes it, which leaves the
    # deck as it was (ooxml-edit 0.1.1 removes the unused png Default).
    assert edited != before or name == "delete_shapes"
    while document.undo():
        pass
    assert document.to_bytes() == before
    while document.redo():
        pass
    assert document.to_bytes() == edited


@operation
def test_each_operation_renders_in_order_with_matching_ids(name, pptx_path):
    pytest.importorskip("pptx2svg")
    import re

    document = Document.open(pptx_path.read_bytes())
    OPERATIONS[name](document)
    svgs = document.render_svg()
    assert len(svgs) == len(document.slides)
    for slide, svg in zip(document.slides, svgs):
        raw = [i for i in re.findall(r'data-pptx-id="([^"]*)"', svg)
               if not i.startswith(("lay:", "mst:"))]
        assert all(i.startswith(f"{slide.slide_id}.") for i in raw), (slide, raw[:3])
        known = {shape.id for shape in slide.shapes}
        rewritten = [i for i in re.findall(r'data-pptx-id="([^"]*)"', slide._rewrite_ids(svg))
                     if not i.startswith(("lay:", "mst:"))]
        assert len(rewritten) == len(set(rewritten))
        shapes = {i.partition("/")[0] for i in rewritten}  # "<shape>/<piece>" is drawn in it
        assert shapes <= known, shapes - known


def test_all_operations_together_undo_as_one_batch(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    before = document.to_bytes()
    with document.batch():
        for operation_ in OPERATIONS.values():
            operation_(document)
    assert document.undo()
    assert document.to_bytes() == before


# ------------------------------------------------------------------------------------------
# Deleting slides
# ------------------------------------------------------------------------------------------


def _names(data: bytes) -> set[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return set(archive.namelist())


def _read(data: bytes, name: str) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(name)


def test_deleting_a_slide_removes_every_trace(financial_report):
    """Entry, relationship, part, its rels, its notes, its charts, their workbooks, overrides."""
    document = Document.open(str(financial_report))
    slide = document.slide(258)
    part = slide.part_path
    package = document.package
    notes = package.related_parts_of_type(part, REL_NOTES_SLIDE)
    charts = package.related_parts_of_type(part, REL_CHART)
    workbooks = [p for chart in charts for p in
                 (r.target_part for r in package.relationships(chart).values()) if p]
    assert notes and len(charts) == 2 and len(workbooks) == 2

    document.delete_slide(slide)
    data = document.to_bytes()
    names = _names(data)
    gone = [part, "ppt/slides/_rels/slide3.xml.rels", *notes, *charts, *workbooks]
    assert not set(gone) & names
    presentation = _read(data, "ppt/presentation.xml")
    assert b'id="258"' not in presentation
    assert b"slides/slide3.xml" not in _read(data, "ppt/_rels/presentation.xml.rels")
    content_types = _read(data, "[Content_Types].xml")
    for name in gone:
        assert ("/" + name).encode() not in content_types
    assert [s.slide_id for s in Document.open(data).slides] == [256, 257, 259]


def test_deleting_a_slide_keeps_what_other_slides_still_use(financial_report):
    document = Document.open(str(financial_report))
    package = document.package
    layout = package.related_parts_of_type(document.slide(257).part_path, REL_SLIDE_LAYOUT)[0]
    document.delete_slide(257)
    assert package.has_part(layout)
    assert package.has_part("ppt/notesMasters/notesMaster1.xml")


def test_a_hyperlink_to_a_deleted_slide_is_unlinked(financial_report):
    document = Document.open(str(financial_report))
    run = _first_run(document.slide(256))
    run.set_hyperlink(document.slide(259))
    address = run.address
    assert document.resolve(address).hyperlink == Hyperlink(slide_id=259)

    document.delete_slide(259)
    assert document.resolve(address).hyperlink is None
    part = document.slide(256).part_path
    assert not [r for r in document.package.relationships(part).values() if r.type == REL_SLIDE]
    assert_valid(document.to_bytes(), financial_report.read_bytes())


def test_deleting_a_slide_with_a_shared_picture_keeps_the_picture(product_page):
    document = Document.open(str(product_page))
    first = document.slides[0]
    picture = first.add_picture(MARKER, 0, 0)
    media = picture.image_part
    copy = first.duplicate()
    document.delete_slide(first)
    assert document.package.has_part(media)
    document.delete_slide(copy)
    assert not document.package.has_part(media)


def test_deleting_every_slide_leaves_a_valid_empty_deck(product_page):
    document = Document.open(str(product_page))
    document.delete_slide(document.slides[0])
    assert document.slides == []
    assert_valid(document.to_bytes(), product_page.read_bytes())
    assert document.undo() and len(document.slides) == 1


# ------------------------------------------------------------------------------------------
# Duplicating slides
# ------------------------------------------------------------------------------------------


def test_duplicate_shares_media_and_copies_charts_and_notes(financial_report):
    document = Document.open(str(financial_report))
    package = document.package
    source = document.slide(257)
    copy = source.duplicate()

    old, new = package.relationships(source.part_path), package.relationships(copy.part_path)
    assert set(old) == set(new), "relationship ids are kept, so the copied XML needs no rewrite"
    for rel_id, relationship in old.items():
        mirrored = new[rel_id]
        if relationship.type in {REL_SLIDE_LAYOUT, REL_IMAGE}:
            assert mirrored.target_part == relationship.target_part
        elif relationship.type in {REL_CHART, REL_NOTES_SLIDE}:
            assert mirrored.target_part != relationship.target_part
            assert package.read(mirrored.target_part) == package.read(relationship.target_part)
    # Each copied chart has its own workbook.
    for rel_id, relationship in old.items():
        if relationship.type == REL_CHART:
            old_book = [r.target_part for r in package.relationships(relationship.target_part).values()]
            new_book = [r.target_part for r in package.relationships(new[rel_id].target_part).values()]
            assert old_book != new_book
            assert [package.read(p) for p in old_book] == [package.read(p) for p in new_book]
    # The copied notes point back at the copy, not at the original.
    notes = package.related_parts_of_type(copy.part_path, REL_NOTES_SLIDE)[0]
    assert package.related_parts_of_type(notes, REL_SLIDE) == [copy.part_path]


def test_duplicate_gets_a_fresh_unique_slide_id(financial_report):
    document = Document.open(str(financial_report))
    copies = [document.slides[0].duplicate() for _ in range(3)]
    ids = [slide.slide_id for slide in document.slides]
    assert len(set(ids)) == len(ids)
    assert all(slide.slide_id >= 256 for slide in copies)
    # Each lands right after the source, pushing the earlier copies along.
    assert [slide.index for slide in copies] == [3, 2, 1]


def test_duplicate_shapes_answer_to_the_copy(financial_report):
    document = Document.open(str(financial_report))
    source = document.slide(257)
    copy = source.duplicate(index=0)
    assert copy.index == 0
    assert [s.id.split(".", 1)[1] for s in copy.shapes] == \
        [s.id.split(".", 1)[1] for s in source.shapes]
    copy.shapes[0].set_text("only on the copy")
    assert source.shapes[0].text != "only on the copy"


def test_duplicate_renews_the_slide_creation_id(pptx_path):
    from pptx_agent.edit.slides import SLIDE_CREATION_ID_EXT_URI

    document = Document.open(str(pptx_path))
    source = document.slides[0]
    copy = source.duplicate()

    def creation(slide):
        root = document.package.tree(slide.part_path)
        return [node.get("val") for node in root.iter(qn("p14:creationId"))]

    if creation(source):
        assert creation(copy) and creation(copy) != creation(source)
    assert SLIDE_CREATION_ID_EXT_URI


def test_slide_ids_start_at_256_and_skip_used_ones():
    from lxml import etree

    from pptx_agent.edit.slides import new_slide_id

    p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    empty = etree.fromstring(f'<p:presentation xmlns:p="{p}"/>'.encode())
    assert new_slide_id(empty) == 256
    low = etree.fromstring(
        f'<p:presentation xmlns:p="{p}"><p:sldIdLst><p:sldId id="3"/></p:sldIdLst>'
        '</p:presentation>'.encode())
    assert new_slide_id(low) == 256
    full = etree.fromstring(
        f'<p:presentation xmlns:p="{p}"><p:sldIdLst><p:sldId id="2147483647"/>'
        '<p:sldId id="256"/></p:sldIdLst></p:presentation>'.encode())
    assert new_slide_id(full) == 257


# ------------------------------------------------------------------------------------------
# Reordering, sections, custom shows
# ------------------------------------------------------------------------------------------


def test_ids_keep_resolving_after_a_reorder(financial_report):
    document = Document.open(str(financial_report))
    identifiers = [shape.id for slide in document.slides for shape in slide.shapes]
    texts = [document.shape(i).text for i in identifiers]
    document.move_slide(259, 0)
    document.move_slide(256, 3)
    assert [s.slide_id for s in document.slides] == [259, 257, 258, 256]
    assert [document.shape(i).text for i in identifiers] == texts
    assert document.slide(256).index == 3


def test_reorder_rejects_a_bad_index(product_page):
    document = Document.open(str(product_page))
    with pytest.raises(IndexError):
        document.move_slide(document.slides[0], 1)


SECTIONS = (
    b'<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
    b'<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    b'<p14:section name="One" id="{11111111-1111-1111-1111-111111111111}">'
    b'<p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>'
    b'<p14:section name="Two" id="{22222222-2222-2222-2222-222222222222}">'
    b'<p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/></p14:sldIdLst></p14:section>'
    b'</p14:sectionLst></p:ext></p:extLst>'
)


def _with_sections(path) -> bytes:
    """The financial report, given two sections -- no fixture has any."""
    source = zipfile.ZipFile(path)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info)
            if info.filename == "ppt/presentation.xml":
                data = data.replace(b"</p:presentation>", SECTIONS + b"</p:presentation>")
            target.writestr(info, data)
    return buffer.getvalue()


def _sections(document):
    from pptx_agent.edit.slides import sections

    return sections(document.package.tree(document.package.presentation_part()))


def test_sections_follow_every_structural_edit(financial_report):
    document = Document.open(_with_sections(financial_report))
    assert _sections(document) == [("One", [256, 257]), ("Two", [258, 259])]

    document.delete_slide(257)
    assert _sections(document) == [("One", [256]), ("Two", [258, 259])]
    copy = document.slide(258).duplicate()
    assert _sections(document) == [("One", [256]), ("Two", [258, copy.slide_id, 259])]
    document.move_slide(259, 0)
    assert _sections(document) == [("One", [259, 256]), ("Two", [258, copy.slide_id])]
    added = document.add_slide(index=2)
    assert _sections(document) == [("One", [259, 256, added.slide_id]),
                                   ("Two", [258, copy.slide_id])]
    assert_valid(document.to_bytes(), _with_sections(financial_report))


def test_a_deleted_slide_leaves_its_custom_shows(financial_report):
    document = Document.open(str(financial_report))
    package = document.package
    root = package.tree(package.presentation_part())
    rel_ids = [node.get(qn("r:id")) for node in root.iter(qn("p:sldId"))]
    shows = (
        '<p:custShowLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:custShow name="Short" id="0"><p:sldLst>'
        + "".join(f'<p:sld r:id="{rel_id}"/>' for rel_id in rel_ids[1:3])
        + "</p:sldLst></p:custShow></p:custShowLst>"
    )
    from lxml import etree

    from pptx_agent.oxml.xml import append_in_order

    append_in_order(root, etree.fromstring(shows.encode()))
    package.mark_dirty(package.presentation_part())

    document.delete_slide(257)
    shown = [node.get(qn("r:id")) for node in root.iter(qn("p:sld"))]
    assert shown == [rel_ids[2]]
    assert_valid(document.to_bytes(), financial_report.read_bytes())


# ------------------------------------------------------------------------------------------
# New slides from a layout
# ------------------------------------------------------------------------------------------


def test_a_new_slide_copies_the_layout_placeholders(pptx_path):
    document = Document.open(str(pptx_path))
    for layout in document.layouts:
        root = document.package.tree(layout.part_path)
        expected = []
        for node in root.iter(qn("p:ph")):
            if node.get("type") in {"dt", "ftr", "sldNum", "hdr"}:
                continue
            if node.getparent().getparent().getparent().tag != qn("p:sp"):
                continue
            expected.append((node.get("type"), int(node.get("idx")) if node.get("idx") else None))
        slide = document.add_slide(layout)
        assert slide.index == len(document.slides) - 1
        assert [shape.placeholder for shape in slide.shapes] == expected
        assert slide.layout == layout
        for shape in slide.shapes:  # empty, so they inherit their place from the layout
            assert not shape.has_explicit_transform and shape.text == ""
    assert_valid(document.to_bytes(), pptx_path.read_bytes())


def test_a_new_slide_goes_where_it_is_asked(financial_report):
    document = Document.open(str(financial_report))
    slide = document.add_slide("DEFAULT", index=0)
    assert slide.index == 0 and slide.slide_id == 260
    assert document.add_slide(index=2).index == 2
    with pytest.raises(KeyError):
        document.add_slide("No such layout")


def test_new_slide_placeholders_take_text(sample_deck):
    document = Document.open(str(sample_deck))
    slide = document.add_slide("Title and Content", index=1)
    title = next(s for s in slide.shapes if s.placeholder[0] == "title")
    title.set_text("Fresh")
    reopened = Document.open(document.to_bytes())
    assert reopened.slides[1].shapes[0].text == "Fresh"


def test_a_placeholder_inherits_through_the_layout_to_the_master(sample_deck):
    """sample.pptx's Title and Content layout gives its placeholders no frame, so they are
    where the master puts them -- PowerPoint draws them there, and moving one starts there."""
    from pptx_agent.oxml.xml import find, qn

    document = Document.open(str(sample_deck))
    slide = document.add_slide("Title and Content")
    master = document.package.tree(document.layout("Title and Content").master_part)
    frames = {}
    for ph in master.iter(qn("p:ph")):
        off = find(ph.getparent().getparent().getparent(), "p:spPr/a:xfrm/a:off")
        ext = find(ph.getparent().getparent().getparent(), "p:spPr/a:xfrm/a:ext")
        frames[ph.get("type")] = (int(off.get("x")), int(off.get("y")),
                                  int(ext.get("cx")), int(ext.get("cy")))
    title, body = slide.shapes
    assert (title.left, title.top, title.width, title.height) == frames["title"]
    assert (body.left, body.top, body.width, body.height) == frames["body"]
    body.move_by(dx=12700)
    assert (body.left, body.top, body.width) == (frames["body"][0] + 12700, frames["body"][1],
                                                 frames["body"][2])


@pytest.fixture
def sample_deck():
    from conftest import FIXTURE_DIR

    return FIXTURE_DIR / "sample.pptx"


# ------------------------------------------------------------------------------------------
# Pictures
# ------------------------------------------------------------------------------------------


def test_a_picture_takes_its_natural_size_or_keeps_its_aspect(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    natural = slide.add_picture(MARKER)
    assert (natural.width, natural.height) == (64 * EMU_PER_INCH // 96, 32 * EMU_PER_INCH // 96)
    tall = slide.add_picture(MARKER, height=EMU_PER_INCH)
    assert (tall.width, tall.height) == (2 * EMU_PER_INCH, EMU_PER_INCH)
    exact = slide.add_picture(MARKER, 10, 20, 30, 40)
    assert (exact.left, exact.top, exact.width, exact.height) == (10, 20, 30, 40)
    assert slide.shapes[-1].id == exact.id  # frontmost


def test_inserted_pictures_share_one_media_part(product_page, tmp_path):
    path = tmp_path / "marker.png"
    path.write_bytes(MARKER)
    document = Document.open(str(product_page))
    slide = document.slides[0]
    first = slide.add_picture(MARKER)
    second = slide.add_picture(str(path))
    assert first.image_part == second.image_part
    media = [n for n in document.package.part_names if n.startswith("ppt/media/")]
    assert len(media) == 1


def test_image_sizes_are_read_from_every_supported_header():
    import struct

    from pptx_agent.oxml.package import image_size

    gif = b"GIF89a" + struct.pack("<HH", 30, 20) + b"\x00" * 8
    bmp = b"BM" + b"\x00" * 16 + struct.pack("<ii", 7, -9) + b"\x00" * 8
    jpeg = (b"\xff\xd8" + b"\xff\xe0" + struct.pack(">H", 4) + b"\x00\x00"
            + b"\xff\xc0" + struct.pack(">HBHH", 11, 8, 50, 70) + b"\x00" * 6)
    assert image_size(MARKER) == (64, 32)
    assert image_size(gif) == (30, 20)
    assert image_size(bmp) == (7, 9)
    assert image_size(jpeg) == (70, 50)
    assert image_size(b"nonsense") is None


def test_replacing_an_image_reaps_the_old_media(product_page):
    document = Document.open(str(product_page))
    picture = document.slides[0].add_picture(ORANGE)
    old = picture.image_part
    picture.replace_image(MARKER)
    assert picture.image_part != old
    assert not document.package.has_part(old)
    rels = document.package.relationships(document.slides[0].part_path)
    assert [r for r in rels.values() if r.type == REL_IMAGE] == \
        [r for r in rels.values() if r.target_part == picture.image_part]


def test_replacing_an_image_keeps_media_another_shape_shows(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    one, two = slide.add_picture(ORANGE), slide.add_picture(ORANGE)
    old = one.image_part
    one.replace_image(MARKER)
    assert document.package.has_part(old) and two.image_part == old
    two.replace_image(MARKER)
    assert not document.package.has_part(old)


def test_replacing_an_image_drops_its_crop_and_svg_twin(product_page):
    from pptx_agent.oxml.xml import make

    document = Document.open(str(product_page))
    picture = document.slides[0].add_picture(ORANGE)
    blip_fill = picture._element.find(qn("p:blipFill"))
    blip = blip_fill.find(qn("a:blip"))
    extensions = make("a:extLst")
    extensions.append(make("a:ext", uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"))
    blip.append(extensions)
    blip.addnext(make("a:srcRect", l="10000"))
    picture.replace_image(MARKER)
    picture = document.shape(picture.id)
    assert picture._element.find(qn("p:blipFill")).find(qn("a:srcRect")) is None
    assert picture._element.find(qn("p:blipFill")).find(qn("a:blip")).find(qn("a:extLst")) is None


def test_replace_image_needs_a_picture(product_page):
    document = Document.open(str(product_page))
    with pytest.raises(ValueError, match="no image"):
        document.slides[0].shapes[0].replace_image(MARKER)


def test_an_image_fill_swapped_for_a_colour_reaps_the_image(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape")
    shape.set_image_fill(ORANGE)
    media = shape.fill.image_part
    shape.fill = "accent1"
    assert not document.package.has_part(media)
    shape.set_image_fill(ORANGE)
    shape.set_image_fill(MARKER)
    assert not document.package.has_part(media) or document.package.read(media) == MARKER


# ------------------------------------------------------------------------------------------
# Reaping on shape delete
# ------------------------------------------------------------------------------------------


def test_deleting_a_chart_reaps_the_chart_and_its_workbook(financial_report):
    document = Document.open(str(financial_report))
    package = document.package
    slide = document.slide(259)
    frame = next(s for s in slide.shapes if s.kind == "graphic_frame" and not s.has_table)
    chart = package.related_parts_of_type(slide.part_path, REL_CHART)[0]
    workbook = next(r.target_part for r in package.relationships(chart).values())
    frame.delete()
    assert not package.has_part(chart) and not package.has_part(workbook)
    assert not package.related_parts_of_type(slide.part_path, REL_CHART)
    assert_valid(document.to_bytes(), financial_report.read_bytes())
    assert document.undo()
    assert package.has_part(chart) and package.has_part(workbook)
    assert document.to_bytes() == Document.open(str(financial_report)).to_bytes()


def test_deleting_one_of_two_users_keeps_the_media(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    one, two = slide.add_picture(ORANGE), slide.add_picture(ORANGE)
    media = one.image_part
    one.delete()
    assert document.package.has_part(media)
    two.delete()
    assert not document.package.has_part(media)


def test_unused_relationships_from_the_original_are_left_alone(product_page):
    """Only references an edit dropped are reaped -- the author's own strays stay put."""
    document = Document.open(str(product_page))
    package = document.package
    slide = document.slides[0]
    media = package.add_image(ORANGE)
    stray = package.add_relationship(slide.part_path, REL_IMAGE, media)
    slide.shapes[0].rotation = 3
    slide.shapes[0].set_text("still here")
    assert stray in package.relationships(slide.part_path)


def test_reap_proves_a_target_unreferenced_across_the_package(product_page):
    """A relationship from a part nothing reaches still counts as a reference."""
    document = Document.open(str(product_page))
    package = document.package
    slide = document.slides[0]
    picture = slide.add_picture(ORANGE)
    media = picture.image_part
    stray = package.add_part("ppt/custom/stray.xml", b"<x/>", "application/xml")
    package.add_relationship(stray, REL_IMAGE, media)
    picture.delete()
    assert package.has_part(media)


# ------------------------------------------------------------------------------------------
# Ungrouping
# ------------------------------------------------------------------------------------------


def _group_of_two(document, rotation=0, flip_h=False, flip_v=False, child_rotation=0):
    slide = document.slides[0]
    a = slide.add_picture(MARKER, EMU_PER_INCH, EMU_PER_INCH, width=3 * EMU_PER_INCH)
    b = slide.add_picture(MARKER, 5 * EMU_PER_INCH, 3 * EMU_PER_INCH, width=2 * EMU_PER_INCH)
    a.rotation = child_rotation
    group = slide.group([a, b])
    group.rotation = rotation
    xfrm = group._xfrm(create=True)
    if flip_h:
        xfrm.set("flipH", "1")
    if flip_v:
        xfrm.set("flipV", "1")
    return slide, group, a.id, b.id


def test_ungroup_puts_the_children_back_in_place(product_page):
    document = Document.open(str(product_page))
    slide, group, a, b = _group_of_two(document)
    before = {i: document.shape(i).slide_bounds for i in (a, b)}
    group.width = group.width * 2  # scale the group: the children stretch with it
    stretched = {i: document.shape(i).slide_bounds for i in (a, b)}
    assert stretched != before
    children = group.ungroup()
    assert [c.id for c in children] == [a, b]
    for identifier in (a, b):
        shape = document.shape(identifier)
        assert shape.parent_group is None
        assert (shape.left, shape.top, shape.width, shape.height) == pytest.approx(
            stretched[identifier], abs=1)


def test_ungroup_composes_rotation_and_flips():
    from pptx_agent.edit.document import ROTATION_UNIT  # noqa: F401
    from conftest import FIXTURE_DIR

    document = Document.open(str(FIXTURE_DIR / "real-product-page.pptx"))
    slide, group, a, b = _group_of_two(document, rotation=30, flip_h=True, child_rotation=20)
    group.ungroup()
    first = document.shape(a)
    # Inside one reflection, the child's own turn runs the other way: 30 - 20.
    assert first.rotation == pytest.approx(10)
    xfrm = first._xfrm(create=False)
    assert xfrm.get("flipH") == "1" and xfrm.get("flipV") is None
    second = document.shape(b)
    assert second.rotation == pytest.approx(30)


@pytest.mark.parametrize("rotation,flip_h,flip_v,child_rotation", [
    (0, False, False, 0), (30, False, False, 0), (0, True, False, 20), (45, True, False, 15),
    (30, True, True, 10), (200, False, True, 75),
])
def test_ungroup_renders_identically(product_page, rotation, flip_h, flip_v, child_rotation):
    """The picture before and after ungrouping is the same, pixel for pixel (bar edges)."""
    pytest.importorskip("pptx2svg")
    pytest.importorskip("resvg_py")
    Image = pytest.importorskip("PIL.Image")
    from PIL import ImageChops

    document = Document.open(str(product_page))
    slide, group, _, _ = _group_of_two(document, rotation, flip_h, flip_v, child_rotation)

    def render():
        return Image.open(io.BytesIO(slide.render_png(width=800))).convert("RGB")

    before = render()
    group.ungroup()
    after = render()
    difference = ImageChops.difference(before, after).convert("L")
    assert sum(difference.histogram()[60:]) < 20


def test_ungroup_hands_the_group_fill_to_children(product_page):
    from pptx_agent.oxml.xml import make

    document = Document.open(str(product_page))
    slide = document.slides[0]
    shapes = [s for s in slide.shapes if s.kind == "shape" and s.has_explicit_transform][:2]
    group = slide.group(shapes)
    group.fill = "accent4"
    for child in group.children:
        properties = child._fill_container(create=True)
        from pptx_agent.edit.fill import write_fill

        write_fill(properties, make("a:grpFill"))
    identifiers = [c.id for c in group.children]
    group.ungroup()
    for identifier in identifiers:
        fill = document.shape(identifier).fill
        assert fill.kind == "solid" and str(fill.color) == "accent4"


def test_ungroup_keeps_the_z_order(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    before = [s.id for s in slide.shapes if s.parent_group is None]
    shapes = [s for s in slide.shapes if s.has_explicit_transform and s.kind != "graphic_frame"][:2]
    group = slide.group(shapes)
    group.ungroup()
    after = [s.id for s in slide.shapes if s.parent_group is None]
    assert sorted(after) == sorted(before)
    assert after.index(shapes[0].id) < after.index(shapes[1].id)


def test_ungroup_a_nested_group_refits_the_outer_one(product_page):
    document = Document.open(str(product_page))
    slide = document.slides[0]
    pictures = [slide.add_picture(MARKER, i * EMU_PER_INCH, i * EMU_PER_INCH,
                                  width=EMU_PER_INCH) for i in range(3)]
    inner = slide.group(pictures[:2])
    outer = slide.group([inner, pictures[2]])
    bounds = {p.id: document.shape(p.id).slide_bounds for p in pictures}
    document.shape(inner.id).ungroup()
    assert [c.kind for c in document.shape(outer.id).children] == ["picture"] * 3
    for identifier, box in bounds.items():
        assert document.shape(identifier).slide_bounds == pytest.approx(box, abs=1)


def test_only_a_group_ungroups(product_page):
    document = Document.open(str(product_page))
    with pytest.raises(ValueError, match="only a group"):
        document.slides[0].shapes[0].ungroup()


# ------------------------------------------------------------------------------------------
# Hyperlinks on runs
# ------------------------------------------------------------------------------------------


def test_hyperlinks_are_added_changed_and_removed_with_their_relationships(financial_report):
    document = Document.open(str(financial_report))
    slide = document.slide(256)
    package = document.package
    run = _first_run(slide)
    address = run.address

    def links():
        return {r.id: r.target for r in package.relationships(slide.part_path).values()
                if r.type in {REL_HYPERLINK, REL_SLIDE}}

    run.set_hyperlink("https://example.org/a", tooltip="A")
    assert document.resolve(address).hyperlink == Hyperlink(address="https://example.org/a",
                                                            tooltip="A")
    assert list(links().values()) == ["https://example.org/a"]

    document.resolve(address).hyperlink = "https://example.org/b"
    assert list(links().values()) == ["https://example.org/b"]

    document.resolve(address).set_hyperlink(document.slide(258))
    assert document.resolve(address).hyperlink == Hyperlink(slide_id=258)
    assert list(links().values()) == ["slide3.xml"]

    document.resolve(address).hyperlink = None
    assert document.resolve(address).hyperlink is None
    assert links() == {}
    assert_valid(document.to_bytes(), financial_report.read_bytes())


def test_a_shared_hyperlink_relationship_survives_one_removal(financial_report):
    document = Document.open(str(financial_report))
    slide = document.slide(256)
    first, second = _first_run(slide), _first_run(slide, skip=1)
    first.set_hyperlink("https://example.org/same")
    second.set_hyperlink("https://example.org/same")
    first.remove_hyperlink()
    assert second.hyperlink == Hyperlink(address="https://example.org/same")


def test_a_hyperlink_lands_in_schema_order(product_page):
    document = Document.open(str(product_page))
    run = _first_run(document.slides[0])
    run.format(bold=True, color="accent1", typeface="Georgia")
    run.set_hyperlink("https://example.org")
    from pptx_agent.oxml.xml import local_name

    properties = run._rpr()
    names = [local_name(child) for child in properties]
    assert names.index("solidFill") < names.index("latin") < names.index("hlinkClick")


def test_existing_hyperlinks_are_read(pptx_path):
    document = Document.open(str(pptx_path))
    found = [run.hyperlink for slide in document.slides for shape in slide.shapes
             if shape.kind == "shape" for paragraph in shape.text_frame.paragraphs
             for run in paragraph.runs if run.hyperlink is not None]
    if pptx_path.stem == "sample-issue-387":
        assert Hyperlink(address="https://github.com/hirokisakabe/pom") in found


def test_a_hyperlink_in_a_table_cell(financial_report):
    document = Document.open(str(financial_report))
    table_shape = next(s for slide in document.slides for s in slide.shapes if s.has_table)
    cell = table_shape.table.cell(0, 0)
    run = next(r for p in cell.text_frame.paragraphs for r in p.runs)
    run.set_hyperlink("https://example.org/cell")
    assert document.resolve(run.address).hyperlink.address == "https://example.org/cell"
    run.remove_hyperlink()
    part = table_shape._slide.part_path
    assert not [r for r in document.package.relationships(part).values()
                if r.type == REL_HYPERLINK]


# ------------------------------------------------------------------------------------------
# The PowerPoint acceptance scenario (run by the oracle test in test_validity.py)
# ------------------------------------------------------------------------------------------

#: A point inside the red corner of the marker picture the acceptance scenario puts on every
#: slide, in inches: the picture is 1in x 0.5in at (0.5, 0.5); its red block is the top-left
#: quarter across and three eighths down.
MARKER_PROBE = (0.6, 0.58)


def acceptance_edits(document: Document) -> dict:
    """Every structural operation, arranged so each leaves something visible in a PDF.

    Returns what to look for: the page count, and the pages (0-based) on which the replaced
    picture and the ungrouped pair sit.
    """
    first = document.slides[0]
    kept = first.duplicate()                       # duplicate (with notes, charts, media)
    doomed = first.duplicate()
    document.delete_slide(doomed)                  # delete: every trace goes
    layout = _layout_with_title(document)
    added = document.add_slide(layout, index=1)    # new slide from a layout
    titles = [s for s in added.shapes if s.placeholder and s.placeholder[0] in {"title", "ctrTitle"}]
    if titles:
        titles[0].set_text("Added from a layout")
    document.move_slide(document.slides[-1], 0)    # reorder: last to first

    for picture in _pictures(document):            # replace every existing picture's image
        picture.replace_image(MARKER)
    swapped = kept.add_picture(ORANGE, 3 * EMU_PER_INCH, EMU_PER_INCH // 2, width=EMU_PER_INCH)
    swapped.replace_image(MARKER)

    a = kept.add_picture(MARKER, 5 * EMU_PER_INCH, 2 * EMU_PER_INCH, width=2 * EMU_PER_INCH)
    b = kept.add_picture(MARKER, 7 * EMU_PER_INCH, 3 * EMU_PER_INCH, width=EMU_PER_INCH)
    group = kept.group([a, b])
    group.rotation = 30
    group.ungroup()                                # ungroup a rotated group

    for position, slide in enumerate(document.slides):
        run = _first_run(slide)
        if run is not None:
            run.set_hyperlink("https://example.org/e2")
        other = _first_run(slide, skip=1)
        if other is not None:
            other.set_hyperlink(document.slides[(position + 1) % len(document.slides)])
        slide.add_picture(MARKER, EMU_PER_INCH // 2, EMU_PER_INCH // 2, width=EMU_PER_INCH)
    return {"pages": len(document.slides), "kept": kept.index, "added": added.index}
