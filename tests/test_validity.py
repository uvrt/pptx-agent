"""Will PowerPoint open it?

The honest answer to that question is the PowerPoint oracle (the ``oracle``-marked tests at the
bottom), but it needs a Mac with PowerPoint and a granted automation permission, so it cannot be
the only line of defence.  These checks cover the failure modes that actually produce
a repair prompt -- a malformed part, a relationship pointing at nothing, a part with no declared
content type -- and they run everywhere.
"""

from __future__ import annotations

import io
import os
import posixpath
import zipfile
from pathlib import Path

import pytest
from lxml import etree

import oracle as oracle_helper
from pptx_agent import Document

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _edited(path: Path) -> bytes:
    """Exercise each kind of mutation, so the checks see all of them."""
    document = Document.open(str(path))
    _apply_e0_edits(document)
    _apply_e1_edits(document)
    _apply_e2_edits(document)
    return document.to_bytes()


def _apply_e0_edits(document: Document) -> None:
    for slide in document.slides:
        shapes = slide.shapes
        if not shapes:
            continue
        movable = [s for s in shapes if s.left is not None]
        if movable:
            movable[0].move_by(dx=228600, dy=228600)
            movable[0].rotation = 7
        texts = [s for s in shapes if s.kind == "shape"]
        if texts:
            texts[0].set_text("edited")
        if len(movable) > 1:
            movable[1].duplicate(dx=114300)
        if len(shapes) > 2:
            shapes[-1].delete()


def _apply_e1_edits(document: Document) -> None:
    """The semantic API: text runs, fills, outlines, tables across merges, groups."""
    for slide in document.slides:
        shapes = slide.shapes
        texts = [s for s in shapes if s.kind == "shape" and s.text]
        if texts:
            frame = texts[-1].text_frame
            paragraph = frame.paragraph(0)
            if paragraph.runs:
                paragraph.run(0).format(bold=True, italic=True, underline="dbl", size=19,
                                        typeface="Georgia", color="accent2 lumMod=75%")
            paragraph.add_run(" +run", color="#C00000")
            paragraph.alignment = "center"
            paragraph.space_after = 76200
            paragraph.set_bullet("\u2022", color="accent1")
            frame.add_paragraph("added paragraph", index=0, size=12)
            texts[-1].set_text(frame.text.replace("added", "inserted"))
        autoshapes = [s for s in shapes if s.kind == "shape"]
        if autoshapes:
            autoshapes[0].set_gradient_fill([(0, "accent1"), (1, "accent1 lumMod=50%")], angle=30)
            line = autoshapes[0].line
            line.width = 28575
            line.color = "accent6"
            line.dash = "dash"
            line.head = "oval"
            line.set_arrowhead("tail", "triangle", "lg", "lg")
        if len(autoshapes) > 1:
            autoshapes[1].fill = "accent3"
        if len(autoshapes) > 2:
            autoshapes[2].set_image_fill(_PNG)
        tables = [s for s in shapes if s.has_table]
        if tables:
            table = tables[0].table
            if table.rows >= 3 and table.columns >= 2:
                table.merge(0, 0, 1, 1)
                table.insert_row(1)
                table.insert_column(1)
                table.delete_row(table.rows - 1)
            table.cell(0, 0).fill = "accent1"
            table.cell(0, 0).set_border("bottom", width=19050, color="tx1", dash="sysDot")
            table.set_column_width(0, table.column_widths[0] + 100000)
            table.cell(table.rows - 1, 0).text = "edited cell"
        top_level = [s for s in slide.shapes if s.parent_group is None and s.left is not None
                     and s.kind != "graphic_frame"]
        if len(top_level) >= 2:
            group = slide.group(top_level[:2])
            group.children[0].move_by(dx=914400, space="slide")


def _apply_e2_edits(document: Document) -> None:
    """The structural operations, together: see ``test_structure.py`` for each on its own."""
    from test_structure import OPERATIONS

    for operation in OPERATIONS.values():
        operation(document)


def _png() -> bytes:
    import struct
    import zlib

    raw = b"".join(b"\x00" + bytes((40, 120, 200)) * 8 for _ in range(8))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


_PNG = _png()


# -- the checks, on bytes, so the structural tests can run them after every operation ------


def check_well_formed(data: bytes) -> None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if name.endswith((".xml", ".rels")):
                etree.fromstring(archive.read(name))


def check_no_dangling_relationships(data: bytes) -> None:
    """A relationship whose target is missing is the classic cause of a repair prompt."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = set(archive.namelist())
        for name in list(names):
            if not name.endswith(".rels"):
                continue
            base = posixpath.dirname(posixpath.dirname(name))
            for node in etree.fromstring(archive.read(name)):
                if node.get("TargetMode") == "External":
                    continue
                target = node.get("Target") or ""
                resolved = (
                    target[1:]
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join(base, target))
                )
                assert resolved in names, f"{name} -> {target}"


def _stale_overrides(data: bytes) -> set[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = set(archive.namelist())
        root = etree.fromstring(archive.read("[Content_Types].xml"))
        return {(node.get("PartName") or "").lstrip("/") for node in root
                if node.tag == f"{{{CT_NS}}}Override"} - names


def check_content_types(data: bytes, original: bytes) -> None:
    """Every part has a content type, and no ``Override`` names a part that is not there --
    beyond any the original generator left behind (one fixture declares a master it lacks)."""
    assert _stale_overrides(data) <= _stale_overrides(original)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = set(archive.namelist())
        root = etree.fromstring(archive.read("[Content_Types].xml"))
        defaults = {
            node.get("Extension", "").lower()
            for node in root
            if node.tag == f"{{{CT_NS}}}Default"
        }
        overrides = [
            (node.get("PartName") or "").lstrip("/")
            for node in root
            if node.tag == f"{{{CT_NS}}}Override"
        ]
        for name in names:
            if name == "[Content_Types].xml" or name.endswith("/"):
                continue
            extension = name.rsplit(".", 1)[-1].lower()
            assert name in overrides or extension in defaults, name
        assert len(overrides) == len(set(overrides)), "a part has two Overrides"


def orphaned_parts(data: bytes) -> set[str]:
    """Parts no chain of relationships from the package root reaches."""
    from pptx_agent.oxml.package import PresentationPackage

    return PresentationPackage.open(data).unreachable_parts()


def check_no_orphans(data: bytes, original: bytes) -> None:
    """An edit may not strand a part.  (A generator's own strays are not the edit's fault.)"""
    assert orphaned_parts(data) <= orphaned_parts(original)


def check_slide_list(data: bytes) -> None:
    """``sldId`` values unique and in range; each resolves to a distinct slide part."""
    from pptx_agent.oxml.package import PresentationPackage

    package = PresentationPackage.open(data)
    slides = package.slide_parts()
    ids = [slide_id for slide_id, _ in slides]
    root = package.tree(package.presentation_part())
    listed = root.findall(".//{%s}sldId" % "http://schemas.openxmlformats.org/presentationml/2006/main")
    assert len(slides) == len(listed), "a sldId does not resolve to a slide part"
    assert len(set(ids)) == len(ids)
    assert all(256 <= slide_id < 2147483648 for slide_id in ids), ids
    assert len({part for _, part in slides}) == len(slides), "two entries share a slide part"
    from pptx_agent.edit.slides import sections

    in_sections = [slide_id for _, members in sections(root) for slide_id in members]
    if in_sections:
        assert in_sections == ids, "sections no longer partition the slide list in order"


def schema_order_violations(data: bytes) -> set[tuple[str, str]]:
    from pptx_agent.core.xml import _ranks, prefixed_name

    found = set()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
                continue
            for parent in etree.fromstring(archive.read(name)).iter():
                ranks = _ranks(prefixed_name(parent))
                if not ranks:
                    continue
                last = -1
                for child in parent:
                    rank = ranks.get(prefixed_name(child), len(ranks) + 1)
                    if rank < last:
                        found.add((prefixed_name(parent), prefixed_name(child)))
                    last = max(last, rank)
    return found


def check_zip(data: bytes) -> None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert archive.testzip() is None


def assert_valid(data: bytes, original: bytes) -> None:
    """Everything below, at once: what the structural tests run after each operation."""
    check_well_formed(data)
    check_no_dangling_relationships(data)
    check_content_types(data, original)
    check_no_orphans(data, original)
    check_slide_list(data)
    assert schema_order_violations(data) <= schema_order_violations(original)
    check_zip(data)


def test_every_part_stays_well_formed(pptx_path):
    check_well_formed(_edited(pptx_path))


def test_no_relationship_dangles(pptx_path):
    check_no_dangling_relationships(_edited(pptx_path))


def test_every_part_has_a_content_type(pptx_path):
    check_content_types(_edited(pptx_path), pptx_path.read_bytes())


def test_no_part_is_orphaned(pptx_path):
    """Deleting a slide or a picture must take everything only it used along with it."""
    check_no_orphans(_edited(pptx_path), pptx_path.read_bytes())


def test_the_slide_list_stays_consistent(pptx_path):
    check_slide_list(_edited(pptx_path))


def test_children_follow_schema_order(pptx_path):
    """Every child this library could have inserted sits where its schema sequence says.

    Checked only for parents whose order was already valid in the original deck, so a
    generator's own quirks are not blamed on the edit.
    """
    assert schema_order_violations(_edited(pptx_path)) <= schema_order_violations(
        pptx_path.read_bytes())


def test_the_zip_is_intact(pptx_path):
    check_zip(_edited(pptx_path))


def test_an_edited_deck_can_be_rendered(pptx_path):
    pptx2svg = pytest.importorskip("pptx2svg")
    data = _edited(pptx_path)
    assert len(pptx2svg.convert_pptx_to_svg(data)) == len(Document.open(data).slides)


# ------------------------------------------------------------------------------------------
# The oracle
# ------------------------------------------------------------------------------------------

requires_powerpoint = pytest.mark.skipif(
    not oracle_helper.available(),
    reason="needs macOS with Microsoft PowerPoint and the pptx2svg oracle script",
)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_an_edited_deck(product_page):
    """The only authoritative check: PowerPoint reads the file and exports it.

    Not run by default (``-m oracle`` to opt in).  It drives the real application, needs a
    granted automation permission, and both paths must live under the user's home --
    PowerPoint's sandbox rejects temporary directories with error -9074.
    """
    home = Path(os.path.expanduser("~"))
    deck, pdf = home / "pptx-agent-oracle.pptx", home / "pptx-agent-oracle.pdf"
    deck.write_bytes(_edited(product_page))
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_semantically_edited_decks(pptx_path):
    """E1's acceptance test: runs, theme fills, outlines, a row inserted across a merge, an
    image fill and a moved group child -- PowerPoint must export every fixture unprompted."""
    home = Path(os.path.expanduser("~"))
    deck = home / f"pptx-agent-e1-{pptx_path.stem}.pptx"
    pdf = home / f"pptx-agent-e1-{pptx_path.stem}.pdf"
    document = Document.open(str(pptx_path))
    _apply_e1_edits(document)
    deck.write_bytes(document.to_bytes())
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_structurally_edited_decks(pptx_path):
    """E2's acceptance test: slides duplicated, deleted, added from a layout and reordered;
    pictures inserted and replaced; a rotated group ungrouped; runs hyperlinked.  PowerPoint
    must export every fixture unprompted, with one page per slide -- and when the pages can be
    rasterised, the marker picture put on every slide must be there, in place, on each."""
    from test_structure import MARKER_PROBE, acceptance_edits

    home = Path(os.path.expanduser("~"))
    deck = home / f"pptx-agent-e2-{pptx_path.stem}.pptx"
    pdf = home / f"pptx-agent-e2-{pptx_path.stem}.pdf"
    document = Document.open(str(pptx_path))
    expected = acceptance_edits(document)
    deck.write_bytes(document.to_bytes())
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        pages = oracle_helper.pdf_pages(pdf, dpi=72)
        if pages is None:
            assert oracle_helper.pdf_page_count(pdf) == expected["pages"]
            return
        assert len(pages) == expected["pages"]
        x, y = (round(inches * 72) for inches in MARKER_PROBE)
        for number, page in enumerate(pages):
            red, green, blue = page.getpixel((x, y))[:3]
            assert red > 200 and green < 80 and blue < 80, f"no marker on page {number + 1}"
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_svg_edited_decks(pptx_path):
    """E3's acceptance test: every edit made only in each slide's full-state SVG -- a theme
    fill, a retyped run, a moved shape, a table cell, a shadow through the raw floor, and a
    red marker square added from raw XML -- applied back to the deck.  PowerPoint must export
    every fixture unprompted, one page per slide; the marker must be on every page and the
    retyped text in the PDF's text."""
    from svgedit import MARKER_PROBE, SENTINEL, acceptance_edits

    home = Path(os.path.expanduser("~"))
    deck = home / f"pptx-agent-e3-{pptx_path.stem}.pptx"
    pdf = home / f"pptx-agent-e3-{pptx_path.stem}.pdf"
    document = Document.open(str(pptx_path))
    expected = acceptance_edits(document)
    deck.write_bytes(document.to_bytes())
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert oracle_helper.pdf_page_count(pdf) == expected["pages"]
        pages = oracle_helper.pdf_pages(pdf, dpi=72)
        if pages is not None:
            x, y = (round(inches * 72) for inches in MARKER_PROBE)
            for number, page in enumerate(pages):
                red, green, blue = page.getpixel((x, y))[:3]
                assert red > 200 and green < 80 and blue < 80, f"no marker on page {number + 1}"
        texts = oracle_helper.pdf_texts(pdf)
        if texts is not None:
            for page in expected["text_pages"]:
                assert SENTINEL in texts[page], f"page {page + 1}: {texts[page][:200]!r}"
            for page in expected["cell_pages"]:
                assert "SVG-CELL" in texts[page], f"page {page + 1}: {texts[page][:200]!r}"
    finally:
        oracle_helper.cleanup(deck, pdf)


def _has_charts_or_diagrams(document: Document) -> bool:
    return any(shape.has_chart or shape.has_diagram
               for slide in document.slides for shape in slide.shapes)


def test_e4_acceptance_edits_keep_the_deck_valid(pptx_path):
    """The edits the E4 oracle test exports, checked everywhere: validity, and every chart's
    workbook agreeing with its cache."""
    pytest.importorskip("pptx2svg")
    from e4edits import acceptance_edits
    from test_chart import check_deck_charts

    original = pptx_path.read_bytes()
    document = Document.open(original)
    if not _has_charts_or_diagrams(document):
        pytest.skip("no chart or SmartArt in this deck")
    acceptance_edits(document)
    edited = document.to_bytes()
    assert_valid(edited, original)
    check_deck_charts(edited, original)


def _page_numbers(pdf: Path, page: int) -> list[float]:
    """The numbers on a PDF page, one per text line (axis labels are lines of their own)."""
    import pypdfium2

    pdf_document = pypdfium2.PdfDocument(str(pdf))
    try:
        text = pdf_document[page].get_textpage().get_text_range()
    finally:
        pdf_document.close()
    numbers = []
    for token in text.split():
        try:
            numbers.append(float(token.replace(",", "")))
        except ValueError:
            pass
    return numbers


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_chart_and_smartart_edits(pptx_path):
    """E4's acceptance test.  Every chart: a value raised far above its axis, a series and a
    category renamed, a series and a category added, a category removed, the title set
    through the full-state SVG -- cache and workbook together.  Every SmartArt diagram: node
    text edited (in the API and in the SVG), a node added, a node removed.  PowerPoint must
    export the deck unprompted with one page per slide, and the PDF must show the new
    labels, names and node text, the removed node gone, and a value axis reaching the new
    value."""
    from e4edits import BIG, acceptance_edits
    from test_chart import check_deck_charts

    original = pptx_path.read_bytes()
    document = Document.open(original)
    if not _has_charts_or_diagrams(document):
        pytest.skip("no chart or SmartArt in this deck")
    expected = acceptance_edits(document)
    data = document.to_bytes()
    check_deck_charts(data, original)  # what PowerPoint's Edit Data would open agrees
    home = Path(os.path.expanduser("~"))
    deck = home / f"pptx-agent-e4-{pptx_path.stem}.pptx"
    pdf = home / f"pptx-agent-e4-{pptx_path.stem}.pdf"
    deck.write_bytes(data)
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert oracle_helper.pdf_page_count(pdf) == expected["pages"]
        texts = oracle_helper.pdf_texts(pdf)
        if texts is None:
            return
        for page, wanted in expected["texts"].items():
            for text in wanted:
                assert text in texts[page], f"page {page + 1}: no {text!r} in {texts[page]!r}"
        for page, unwanted in expected["absent"].items():
            for text in unwanted:
                assert text not in texts[page], f"page {page + 1}: {text!r} is still there"
        for page in expected["big"]:
            assert max(_page_numbers(pdf, page), default=0) >= BIG, f"page {page + 1}"
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_authored_shapes(pptx_path):
    """E5's acceptance test.  A new slide with presets and adjustments, a text box, three
    connectors (straight, elbow, curved) attached between two shapes -- one shape then
    moved, the other resized -- an edited table and a shape added inside a group.
    PowerPoint must export the deck unprompted with one page per slide; the page must show
    the connectors meeting their shapes (red just outside every site they end on) and the
    new text; and PowerPoint, saving the deck again, must keep every connection."""
    from authoring import SENTINELS, acceptance_edits

    original = pptx_path.read_bytes()
    document = Document.open(original)
    expected = acceptance_edits(document)
    data = document.to_bytes()
    assert_valid(data, original)
    home = Path(os.path.expanduser("~"))
    deck = home / f"pptx-agent-e5-{pptx_path.stem}.pptx"
    pdf = home / f"pptx-agent-e5-{pptx_path.stem}.pdf"
    again = home / f"pptx-agent-e5-{pptx_path.stem}-resaved.pptx"
    deck.write_bytes(data)
    try:
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert oracle_helper.pdf_page_count(pdf) == expected["pages"]
        pages = oracle_helper.pdf_pages(pdf, dpi=144)
        if pages is not None:
            page = pages[expected["page"]]
            for x, y in expected["probes"]:
                assert _red_near(page, round(x * 144), round(y * 144)), (
                    f"no connector at ({x:.2f}, {y:.2f}) in")
        texts = oracle_helper.pdf_texts(pdf)
        if texts is not None:
            for sentinel in SENTINELS:
                assert sentinel in texts[expected["page"]], (sentinel, texts[expected["page"]])

        saved = oracle_helper.save_as_pptx(deck, again)
        assert saved.ok, f"PowerPoint {saved.outcome}: {saved.detail}"
        resaved = Document.open(again.read_bytes())
        for identifier in expected["connectors"]:
            connector = resaved.shape(identifier)
            begin, end = connector.begin_connection, connector.end_connection
            assert begin is not None and end is not None, identifier
            original_connector = document.shape(identifier)
            assert begin[0].id == original_connector.begin_connection[0].id
            assert end[0].id == original_connector.end_connection[0].id
            assert (begin[1], end[1]) == (original_connector.begin_connection[1],
                                          original_connector.end_connection[1])
    finally:
        oracle_helper.cleanup(deck, pdf, again)


def _red_near(image, x: int, y: int, reach: int = 4) -> bool:
    for dx in range(-reach, reach + 1):
        for dy in range(-reach, reach + 1):
            try:
                red, green, blue = image.getpixel((x + dx, y + dy))[:3]
            except IndexError:
                continue
            if red > 190 and green < 90 and blue < 90:
                return True
    return False


@pytest.mark.oracle
@requires_powerpoint
def test_the_oracle_rejects_a_damaged_deck(product_page):
    """Proves the oracle has teeth -- a passing export means something.

    Truncating the slide's root element makes PowerPoint raise its repair dialog, which is the
    failure this whole helper exists to survive: the dialog blocks AppleEvents, so without
    recovery it would fail every later export too.
    """
    home = Path(os.path.expanduser("~"))
    deck, pdf = home / "pptx-agent-damaged.pptx", home / "pptx-agent-damaged.pdf"

    with zipfile.ZipFile(io.BytesIO(product_page.read_bytes())) as source:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "ppt/slides/slide1.xml":
                    data = data.replace(b"</p:sld>", b"")
                target.writestr(info, data)
    deck.write_bytes(buffer.getvalue())

    try:
        # retries=0: the file is meant to be rejected, so do not open it twice.
        result = oracle_helper.export_pdf(deck, pdf, timeout=60, retries=0)
        assert not result.ok
        assert result.outcome == "rejected", result.detail
        # And PowerPoint must be usable again, or every later export would fail too.
        assert result.recovered, "could not clear the repair dialog"
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_still_works_after_a_rejection(product_page):
    """The recovery actually recovered: a good file exports right after a bad one."""
    home = Path(os.path.expanduser("~"))
    deck, pdf = home / "pptx-agent-after.pptx", home / "pptx-agent-after.pdf"
    deck.write_bytes(product_page.read_bytes())
    try:
        assert oracle_helper.export_pdf(deck, pdf).ok
    finally:
        oracle_helper.cleanup(deck, pdf)
