"""The full-state SVG, read back: ROADMAP E3's acceptance gates.

* **Identity** -- emit every slide, apply the SVGs to a fresh copy of the deck, save: every
  part byte-identical.  An SVG that matches the document changes nothing at all.
* **Edits through the SVG** -- a theme fill, a run, a position and a table cell changed in
  the SVG alone arrive as the E1/E2 edits they correspond to; theme colours stay
  ``a:schemeClr``; undo gives back the original bytes.
* **The raw floor** -- a change only ``data-ooxml-xml`` carries arrives intact.
* **Validity** -- the validity checks hold after SVG-applied edits, on every fixture.
"""

from __future__ import annotations

import io
import json
import zipfile

import pytest
from lxml import etree

from pptx_agent import Document
from pptx_agent.fullstate import FullStateError
from svgedit import (
    A,
    EMU_PER_INCH,
    P,
    PML,
    SENTINEL,
    FullStateSvg,
    acceptance_edits,
    red_square,
    with_shadow,
)
from test_validity import assert_valid

pytest.importorskip("pptx2svg", reason="the full-state SVG is drawn by pptx2svg")


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()
                if not info.is_dir()}


def _svg(document: Document, slide_index: int = 0) -> FullStateSvg:
    return FullStateSvg(document.slides[slide_index].render_svg(full_state=True))


def _c14n(element) -> bytes:
    return etree.tostring(etree.fromstring(etree.tostring(element)), method="c14n")


# ------------------------------------------------------------------------------------------
# Identity
# ------------------------------------------------------------------------------------------


def test_applying_its_own_svg_to_a_fresh_copy_changes_no_byte(pptx_path):
    original = pptx_path.read_bytes()
    svgs = [slide.render_svg(full_state=True) for slide in Document.open(original).slides]
    fresh = Document.open(original)

    reports = [fresh.apply_svg(svg) for svg in svgs]

    assert not any(reports), [r for r in reports if r]
    assert fresh.package.dirty_parts == frozenset()
    assert not fresh.history.can_undo()
    assert _parts(fresh.to_bytes()) == _parts(original)


def test_applying_is_idempotent_after_an_edit(pptx_path):
    """Edit through the SVG, re-emit from the result: the new SVG is again a no-op."""
    document = Document.open(pptx_path.read_bytes())
    acceptance_edits(document)
    saved = document.to_bytes()
    reopened = Document.open(saved)
    for slide in Document.open(saved).slides:
        assert not reopened.apply_svg(slide.render_svg(full_state=True))
    assert reopened.to_bytes() == saved


# ------------------------------------------------------------------------------------------
# Edits through the SVG, on every fixture
# ------------------------------------------------------------------------------------------


def test_edits_made_in_the_svg_arrive_in_the_deck(pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    document.history._max_depth = 10_000
    expected = acceptance_edits(document)
    saved = document.to_bytes()
    reopened = Document.open(saved)

    for identifier in expected["fills"]:
        fill = reopened.shape(identifier).fill
        assert fill.kind == "solid" and fill.color == "accent2 lumMod=75%"
        properties = reopened.shape(identifier)._element.find(PML + "spPr")
        solid = properties.find(A + "solidFill")
        assert solid.find(A + "schemeClr").get("val") == "accent2"  # never resolved
        assert solid.find(A + "srgbClr") is None
    for identifier in expected["texts"]:
        assert SENTINEL in reopened.shape(identifier).text
    for identifier, left in expected["moves"]:
        assert reopened.shape(identifier).left == left
    for identifier in expected["cells"]:
        assert reopened.shape(identifier).table.cell(0, 0).text == "SVG-CELL"
    for identifier in expected["markers"]:
        marker = reopened.shape(identifier)
        assert marker.name == "SVG marker" and marker.fill.color == "#FF0000"
    assert len(expected["markers"]) == len(reopened.slides)

    # One undo step per slide, and undoing all of them gives back the original bytes.
    while document.undo():
        pass
    assert _parts(document.to_bytes()) == _parts(original)


def test_svg_edits_keep_the_deck_valid(pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    acceptance_edits(document)
    assert_valid(document.to_bytes(), original)


def test_an_svg_edited_deck_renders_with_matching_ids(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    expected = acceptance_edits(document)
    for slide, marker in zip(document.slides, expected["markers"]):
        svg = FullStateSvg(slide.render_svg(full_state=True))
        assert set(svg.groups) == {shape.id for shape in slide.shapes}
        assert svg.get(marker, "fill-rgb") == "FF0000"


def test_the_raw_floor_carries_what_the_vocabulary_does_not(pptx_path):
    """A shadow is not in the typed vocabulary; edited into data-ooxml-xml it arrives whole."""
    original = pptx_path.read_bytes()
    document = Document.open(original)
    for index, slide in enumerate(document.slides):
        svg = _svg(document, index)
        targets = svg.where(kind="shape")
        if not targets:
            continue
        shadowed = with_shadow(svg.xml(targets[0]))
        svg.set_xml(targets[0], shadowed)

        before = document.package.read(slide.part_path)
        report = slide.apply_svg(str(svg))

        assert report.replaced == [targets[0]] and not report.edited
        saved = Document.open(document.to_bytes()).shape(targets[0])._element
        assert _c14n(saved) == _c14n(shadowed)
        # Namespaces the slide already declares are not repeated on the shape.
        after = document.package.read(slide.part_path)
        assert after.count(b"xmlns") == before.count(b"xmlns")
    assert_valid(document.to_bytes(), original)


# ------------------------------------------------------------------------------------------
# The semantics, in detail
# ------------------------------------------------------------------------------------------


def test_a_raw_only_change_is_not_reverted_by_stale_typed_attributes(product_page):
    """Typed attributes are compared with the document, not with the raw XML they travel
    with -- so editing just the raw fill keeps it, though data-ooxml-fill still says the old."""
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = svg.where(kind="shape", fill="solid")[0]
    raw = svg.xml(target)
    raw.find(PML + "spPr").find(A + "solidFill")[0].set("val", "123456")
    raw.find(PML + "spPr").find(A + "solidFill")[0].tag = A + "srgbClr"
    svg.set_xml(target, raw)

    report = document.apply_svg(str(svg))

    assert report.replaced == [target] and not report.edited
    assert document.shape(target).fill.color == "#123456"


def test_typed_edits_apply_on_top_of_a_raw_change(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = svg.where(kind="shape")[0]
    svg.set_xml(target, with_shadow(svg.xml(target)))
    svg.set(target, "name", "Renamed in the SVG")
    svg.set(target, "rot", "12.5")

    report = document.apply_svg(str(svg))

    shape = document.shape(target)
    assert report.edited[target] == ["name", "rot"]
    assert shape.name == "Renamed in the SVG" and shape.rotation == 12.5
    assert shape._element.find(PML + "spPr").find(A + "effectLst") is not None


def test_geometry_flips_and_preset_through_the_svg(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = next(i for i in svg.where(kind="shape") if svg.get(i, "geom") == "rect")
    svg.set(target, "y", "100000")
    svg.set(target, "cx", "2000000")
    svg.set(target, "flip-h", "1")
    svg.set(target, "geom", "ellipse")

    document.apply_svg(str(svg))

    shape = Document.open(document.to_bytes()).shape(target)
    assert (shape.top, shape.width, shape.flip_h, shape.preset) == (100000, 2000000, True,
                                                                    "ellipse")


def test_inherit_removes_explicit_fills_and_outlines(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = svg.where(kind="shape", fill="solid")[0]
    svg.set_fill(target, "inherit")
    node = svg.group(target)
    for name in [n for n in node.attrib if n.startswith(P + "line")]:
        del node.attrib[name]
    node.set(P + "line", "inherit")

    document.apply_svg(str(svg))

    shape = document.shape(target)
    assert shape.fill is None and not shape.line.exists


def test_outline_and_gradient_through_the_svg(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = svg.where(kind="shape")[0]
    svg.set_fill(target, "gradient", stops="0 scheme:accent1;100000 scheme:accent1 lumMod=50000",
                 angle="90")
    node = svg.group(target)
    for name in [n for n in node.attrib if n.startswith(P + "line")]:
        del node.attrib[name]
    for key, value in {"line": "set", "line-w": "28575", "line-fill": "solid",
                       "line-scheme": "accent6", "line-dash": "dash",
                       "line-tail": "triangle lg lg"}.items():
        node.set(P + key, value)

    document.apply_svg(str(svg))

    shape = Document.open(document.to_bytes()).shape(target)
    assert shape.fill.kind == "gradient" and shape.fill.stops[1].color == "accent1 lumMod=50%"
    assert (shape.line.width, shape.line.dash) == (28575, "dash")
    assert shape.line.color == "accent6" and shape.line.tail.type == "triangle"


def test_runs_split_formatted_and_linked_through_the_svg(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    target = next(i for i in svg.where(kind="shape") if svg.get(i, "text")
                  and len(svg.json(i, "text")["p"][0]["c"]) == 1
                  and len(svg.json(i, "text")["p"][0]["c"][0]["t"]) > 4)
    model = svg.json(target, "text")
    run = model["p"][0]["c"][0]
    head, tail = dict(run, t=run["t"][:3]), dict(run, t=run["t"][3:])
    tail.update({"b": True, "color": "scheme:accent4", "sz": 2400,
                 "link": {"url": "https://example.org/", "tip": "Example"}})
    model["p"][0]["c"] = [head, {"br": True}, tail]
    model["p"][0]["algn"] = "r"
    model["p"].append({"c": [{"t": "A second paragraph"}], "bu": {"char": "-", "size": 80000}})
    svg.set_json(target, "text", model)

    document.apply_svg(str(svg))

    shape = Document.open(document.to_bytes()).shape(target)
    first, second = shape.text_frame.paragraphs
    assert first.text == run["t"][:3] + "\v" + run["t"][3:]
    assert first.alignment == "right"
    styled = first.run(1)
    assert styled.bold and styled.size == 24 and styled.color == "accent4"
    assert styled.hyperlink.address == "https://example.org/"
    assert first.run(0).bold == run.get("b")
    assert second.text == "A second paragraph" and second.bullet.char == "-"


def test_table_rows_and_merges_through_the_svg(financial_report):
    document = Document.open(financial_report.read_bytes())
    slide_index = next(i for i, s in enumerate(document.slides)
                       if any(shape.has_table for shape in s.shapes))
    svg = _svg(document, slide_index)
    frame = next(i for i in svg.groups if svg.get(i, "table"))
    model = svg.json(frame, "table")
    rows_before = len(model["rows"])
    new_row = [{"text": {"p": [{"c": [{"t": f"new {c}"}]}]}} for c in range(len(model["cols"]))]
    model["cells"].insert(1, new_row)
    model["rows"].insert(1, 400000)
    model["merges"] = [[2, 0, 3, 1]]
    model["cells"][0][1]["fill"] = {"kind": "solid", "scheme": "accent5"}
    model["cols"][0] += 50000
    svg.set_json(frame, "table", model)

    document.apply_svg(str(svg))

    table = Document.open(document.to_bytes()).shape(frame).table
    assert table.rows == rows_before + 1
    assert [table.cell(1, c).text for c in range(table.columns)] == \
        [f"new {c}" for c in range(table.columns)]
    assert table.row_heights[1] == 400000
    assert [(r.top, r.left, r.bottom, r.right) for r in table.merged_regions] == [(2, 0, 3, 1)]
    assert table.cell(0, 1).fill.color == "accent5"
    assert table.column_widths[0] == model["cols"][0]
    assert_valid(document.to_bytes(), financial_report.read_bytes())


def test_new_shapes_are_added_where_the_svg_puts_them(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    svg = _svg(document)
    taken = svg.get(next(iter(svg.groups)), "cnvpr-id")
    square = red_square(EMU_PER_INCH, EMU_PER_INCH, EMU_PER_INCH, shape_id=int(taken))
    group = svg.add(square)  # no data-pptx-id: a shape the slide does not have yet
    # Move it to sit second in z-order.
    first = svg.group(next(iter(svg.groups)))
    first.addnext(group)

    report = document.apply_svg(str(svg))

    (added,) = report.added.values()
    shape = document.shape(added)
    assert shape.name == "SVG marker"
    assert [s.id for s in slide.shapes].index(added) == 1
    ids = [s._element.find(f"{PML}nvSpPr/{PML}cNvPr").get("id") for s in slide.shapes
           if s.kind == "shape"]
    assert len(ids) == len(set(ids)), "a taken cNvPr@id was renumbered"


def test_missing_shapes_are_kept_unless_asked(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    doomed = list(svg.groups)[3]
    svg.remove(doomed)

    assert not document.apply_svg(str(svg))
    assert document.shape(doomed)

    report = document.apply_svg(str(svg), delete_missing=True)
    assert report.deleted == [doomed]
    with pytest.raises(KeyError):
        document.shape(doomed)


def test_new_shapes_can_be_refused(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    svg.add(red_square(0, 0, EMU_PER_INCH))
    with pytest.raises(FullStateError, match="add_new"):
        document.apply_svg(str(svg), add_new=False)


def test_a_copied_chart_frame_gets_its_own_chart(financial_report):
    """Copying a shape across the SVG copies per-shape parts, as duplicating a slide does."""
    original = financial_report.read_bytes()
    document = Document.open(original)
    index = next(i for i, s in enumerate(document.slides)
                 if any(sh.kind == "graphic_frame" and not sh.has_table for sh in s.shapes))
    svg = _svg(document, index)
    chart = next(i for i in svg.groups if svg.get(i, "chart"))
    svg.add(svg.xml(chart), rels=json.loads(svg.get(chart, "rels")))

    report = document.apply_svg(str(svg))

    (added,) = report.added.values()
    new_chart = document.package.related_part(document.slides[index].part_path,
                                              document.shape(added)._element.find(
                                                  ".//{*}chart").get(
                                                  "{http://schemas.openxmlformats.org/"
                                                  "officeDocument/2006/relationships}id"))
    assert new_chart != svg.get(chart, "chart") and new_chart.startswith("ppt/charts/")
    assert_valid(document.to_bytes(), original)


def test_a_new_shape_relates_to_media_by_the_description(product_page):
    """Relationship ids in a fragment are re-pointed through data-ooxml-rels, so a picture
    whose id means something else in the slide still shows its own image."""
    from images import MARKER

    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    picture = slide.add_picture(MARKER, 0, 0, width=EMU_PER_INCH)
    svg = FullStateSvg(slide.render_svg(full_state=True))
    raw = svg.xml(picture.id)
    rels = json.loads(svg.get(picture.id, "rels"))
    blip = raw.find(f".//{A}blip")
    old_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
    new_id = "rId999"
    blip.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", new_id)
    rels[0]["id"] = new_id
    svg.add(raw, rels=rels)

    report = document.apply_svg(str(svg))

    (added,) = report.added.values()
    assert document.shape(added).image_part == picture.image_part
    assert new_id not in document.package.relationships(slide.part_path)
    assert old_id


def test_read_only_attributes_are_refused(product_page):
    document = Document.open(product_page.read_bytes())
    before = document.to_bytes()
    svg = _svg(document)
    target = next(iter(svg.groups))
    svg.set(target, "kind", "picture")
    with pytest.raises(FullStateError, match="read-only"):
        document.apply_svg(str(svg))
    assert document.to_bytes() == before


def test_another_state_of_the_deck_is_refused(product_page, financial_report):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    svg.root.set(P + "theme", "ppt/theme/theme9.xml")
    with pytest.raises(FullStateError, match="another"):
        document.apply_svg(str(svg))

    other = Document.open(financial_report.read_bytes())
    foreign = _svg(other, 2)
    with pytest.raises(FullStateError):
        document.apply_svg(str(foreign))


def test_slide_apply_svg_checks_the_slide(financial_report):
    document = Document.open(financial_report.read_bytes())
    svg = document.slides[1].render_svg(full_state=True)
    with pytest.raises(FullStateError, match="slide"):
        document.slides[0].apply_svg(svg)
    assert not document.slides[1].apply_svg(svg)


def test_a_failed_apply_changes_nothing(product_page):
    document = Document.open(product_page.read_bytes())
    before = document.to_bytes()
    svg = _svg(document)
    first, second = svg.where(kind="shape")[:2]
    svg.set(first, "name", "would be renamed")
    svg.set(second, "geom", "custom")  # not expressible as a typed edit: fails mid-apply
    with pytest.raises(FullStateError, match="custom"):
        document.apply_svg(str(svg))
    assert document.to_bytes() == before
    assert not document.history.can_undo()


def test_unknown_groups_without_xml_are_ignored(product_page):
    document = Document.open(product_page.read_bytes())
    svg = _svg(document)
    stray = etree.SubElement(svg.root, "{http://www.w3.org/2000/svg}g")
    stray.set("data-pptx-id", "256.999")
    report = document.apply_svg(str(svg))
    assert not report and report.ignored == ["256.999"]


def _grouped(product_page) -> bytes:
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    slide.group([s for s in slide.shapes if s.left is not None][:2], name="Pair")
    return document.to_bytes()


def test_a_group_shell_change_keeps_its_children(product_page):
    data = _grouped(product_page)
    svg = _svg(Document.open(data))
    group = svg.where(kind="group")[0]
    shell = svg.xml(group)
    properties = shell.find(PML + "grpSpPr")
    fill = etree.SubElement(properties, A + "solidFill")
    etree.SubElement(fill, A + "schemeClr", val="accent3")
    svg.set_xml(group, shell)
    children = [i for i in svg.groups if svg.group(group) in svg.group(i).iterancestors()]

    document = Document.open(data)
    report = document.apply_svg(str(svg))

    assert report.replaced == [group]
    reopened = Document.open(document.to_bytes())
    assert reopened.shape(group).fill.color == "accent3"
    assert [child.id for child in reopened.shape(group).children] == children
    assert_valid(document.to_bytes(), data)


def test_deleting_a_group_takes_its_children_and_a_kept_child_is_refused(product_page):
    data = _grouped(product_page)
    svg = _svg(Document.open(data))
    group = svg.where(kind="group")[0]
    children = [i for i in svg.groups if svg.group(group) in svg.group(i).iterancestors()]

    # A child kept while its group goes: that would be an ungroup, which this does not guess.
    orphaned = FullStateSvg(str(svg))
    child = orphaned.group(children[0])
    orphaned.group(group).addprevious(child)
    orphaned.remove(group)
    document = Document.open(data)
    with pytest.raises(FullStateError, match="ungroup"):
        document.apply_svg(str(orphaned), delete_missing=True)
    assert document.to_bytes() == data

    svg.remove(group)
    report = document.apply_svg(str(svg), delete_missing=True)
    assert report.deleted == [group]
    remaining = {s.id for s in document.slides[0].shapes}
    assert not remaining & {group, *children}


def test_a_new_group_with_new_children(product_page):
    """A group and its children added from raw XML, nested as the SVG nests them."""
    data = _grouped(product_page)
    svg = _svg(Document.open(data))
    group = svg.where(kind="group")[0]
    shell = svg.xml(group)
    children = [i for i in svg.groups if svg.group(group) in svg.group(i).iterancestors()]
    added = svg.add(shell, identifier="new-group")
    for child in children:
        node = svg.group(child)
        svg.add(svg.xml(child), parent="new-group", rels=json.loads(node.get(P + "rels") or "[]"))
    assert added is not None

    document = Document.open(data)
    report = document.apply_svg(str(svg))

    new_group = document.shape(report.added["new-group"])
    assert new_group.kind == "group" and len(new_group.children) == len(children)
    ids = [s._element.find(".//{*}cNvPr").get("id") for s in document.slides[0].shapes]
    assert len(ids) == len(set(ids))
    assert_valid(document.to_bytes(), data)


def test_a_copied_shape_drops_its_creation_id(product_page):
    """A shape copied across the SVG with PowerPoint's ``a16:creationId/@id`` loses it, so the
    copy does not answer to the original's ``c<GUID>`` address."""
    from pptx_agent.edit.ids import cnv_pr
    from pptx_agent.oxml.xml import CREATION_ID_EXT_URI, PRESENTATION_NAMESPACES, qn

    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    properties = cnv_pr(slide.shapes[1]._element)
    ext_list = properties.find(qn("a:extLst"))
    if ext_list is None:
        ext_list = etree.SubElement(properties, qn("a:extLst"))
    extension = etree.SubElement(ext_list, qn("a:ext"), uri=CREATION_ID_EXT_URI)
    etree.SubElement(extension, qn("a16:creationId"), nsmap={"a16": PRESENTATION_NAMESPACES["a16"]},
                     id="{0E6F5A2B-1C3D-4E5F-8A9B-0C1D2E3F4A5B}")
    document.package.mark_dirty(slide.part_path)
    document = Document.open(document.to_bytes())
    original = f"{document.slides[0].slide_id}.c0E6F5A2B-1C3D-4E5F-8A9B-0C1D2E3F4A5B"

    svg = _svg(document)
    svg.add(svg.xml(original))
    report = document.apply_svg(str(svg))

    (added,) = report.added.values()
    assert added != original
    assert document.shape(original).name == document.shape(added).name
    assert b"0E6F5A2B" not in etree.tostring(document.shape(added)._element)
