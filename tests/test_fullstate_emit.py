"""The full-state SVG, written: every shape's state, unresolved, on a picture that is unchanged.

ROADMAP E3.  The picture comes from pptx2svg; this library only adds ``data-ooxml-*``
attributes, read from its own document model.  So the checks are that every shape is there,
that a theme colour is written as the theme reference it is, that the raw XML is the shape's
own, and that nothing about the drawing changed.
"""

from __future__ import annotations

import base64
import json

import pytest
from lxml import etree

from pptx_agent import Document
from pptx_agent.edit.ids import cnv_pr
from pptx_agent.fullstate.emit import fragment_bytes
from svgedit import P, FullStateSvg

pytest.importorskip("pptx2svg", reason="the full-state SVG is drawn by pptx2svg")


def _strip_state(svg: str) -> bytes:
    """The SVG without anything the full-state emitter added, canonicalised."""
    root = etree.fromstring(svg.encode())
    for node in list(root.iter()):
        if not isinstance(node.tag, str):
            continue
        if node.get(P + "unrendered"):
            node.getparent().remove(node)
            continue
        for name in list(node.attrib):
            if name.startswith(P):
                del node.attrib[name]
    return etree.tostring(root, method="c14n")


def test_every_shape_carries_its_state_and_its_xml(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    for slide in document.slides:
        svg = FullStateSvg(slide.render_svg(full_state=True))
        groups = svg.groups
        for shape in slide.shapes:
            assert shape.id in groups, f"{shape.id} is missing from the SVG"
            assert svg.get(shape.id, "kind") == shape.kind
            assert svg.get(shape.id, "name") == (shape.name or "")
            raw = base64.b64decode(svg.get(shape.id, "xml"))
            assert raw == fragment_bytes(shape._element)
        assert set(groups) == {shape.id for shape in slide.shapes}


def test_the_slide_is_described_on_the_root(financial_report):
    document = Document.open(financial_report.read_bytes())
    slide = document.slides[1]
    root = FullStateSvg(slide.render_svg(full_state=True)).root
    attributes = {k[len(P):]: v for k, v in root.attrib.items() if k.startswith(P)}
    assert attributes["vocabulary"] == "pptx-agent/1"
    assert attributes["slide-id"] == str(slide.slide_id)
    assert attributes["slide-index"] == "1"
    assert attributes["slide-part"] == slide.part_path
    assert attributes["layout"] == slide.layout.part_path
    assert attributes["layout-name"] == slide.layout.name
    assert attributes["master"] == slide.layout.master_part
    assert attributes["theme"].startswith("ppt/theme/")
    assert (attributes["slide-cx"], attributes["slide-cy"]) == ("12192000", "6858000")


def test_a_theme_fill_is_written_as_the_theme_reference(product_page):
    """The defect E3 exists to avoid: a resolved hex instead of ``accent1``."""
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    shape = next(s for s in slide.shapes if s.kind == "shape" and s.text)
    shape.fill = "accent1 lumMod=75% lumOff=25%"
    shape.text_frame.paragraph(0).run(0).color = "accent2"
    shape.line.color = "tx1"

    svg = FullStateSvg(slide.render_svg(full_state=True))

    assert svg.get(shape.id, "fill") == "solid"
    assert svg.get(shape.id, "fill-scheme") == "accent1"
    assert svg.get(shape.id, "fill-mods") == "lumMod=75000 lumOff=25000"
    assert svg.get(shape.id, "fill-rgb") is None
    assert svg.get(shape.id, "line-scheme") == "tx1"
    first_run = svg.json(shape.id, "text")["p"][0]["c"][0]
    assert first_run["color"] == "scheme:accent2"


def test_explicit_colours_and_inheritance_are_told_apart(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    shapes = [s for s in slide.shapes if s.kind == "shape"]
    shapes[0].fill = "#4472C4"
    shapes[1].clear_fill()
    shapes[1].line.clear()
    shapes[2].set_gradient_fill([(0, "accent1"), (1, "#FFFFFF")], angle=45)

    svg = FullStateSvg(slide.render_svg(full_state=True))

    assert (svg.get(shapes[0].id, "fill"), svg.get(shapes[0].id, "fill-rgb")) == ("solid", "4472C4")
    assert svg.get(shapes[1].id, "fill") == "inherit"
    assert svg.get(shapes[1].id, "line") == "inherit"
    assert svg.get(shapes[2].id, "fill") == "gradient"
    assert svg.get(shapes[2].id, "fill-stops") == "0 scheme:accent1;100000 rgb:FFFFFF"
    assert svg.get(shapes[2].id, "fill-angle") == "45"


def test_geometry_is_emu_and_an_inherited_position_says_so(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    for slide in document.slides:
        svg = FullStateSvg(slide.render_svg(full_state=True))
        for shape in slide.shapes:
            if shape.has_explicit_transform:
                assert svg.get(shape.id, "rot") == format(shape.rotation, "g")
            else:
                assert svg.get(shape.id, "x") == "inherit"


def test_tables_and_text_are_compact_json(financial_report):
    document = Document.open(financial_report.read_bytes())
    slide = document.slides[1]
    frame = next(s for s in slide.shapes if s.has_table)
    svg = FullStateSvg(slide.render_svg(full_state=True))

    table = svg.json(frame.id, "table")
    assert set(table) == {"cols", "rows", "merges", "cells"}
    assert table["cols"] == frame.table.column_widths
    assert table["rows"] == frame.table.row_heights
    assert len(table["cells"]) == frame.table.rows
    texts = [["".join(i.get("t", "") for p in cell["text"]["p"] for i in p["c"])
              for cell in row] for row in table["cells"]]
    assert texts[0][0] == frame.table.cell(0, 0).text
    assert svg.get(frame.id, "graphic") == "table"
    # Compact and deterministic: no whitespace between tokens.
    assert ", " not in svg.get(frame.id, "table") and ": " not in svg.get(frame.id, "table")


def test_relationships_are_described_by_reference(pptx_path):
    """A picture or chart says which package part it means; the bytes are not embedded."""
    document = Document.open(pptx_path.read_bytes())
    package = document.package
    for slide in document.slides:
        svg = FullStateSvg(slide.render_svg(full_state=True))
        for identifier, group in svg.groups.items():
            raw = svg.xml(identifier)
            used = {v for node in raw.iter() for k, v in node.attrib.items()
                    if k.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/"
                                    "relationships}")}
            described = {entry["id"]: entry for entry in
                         json.loads(group.get(P + "rels") or "[]")}
            relationships = package.relationships(slide.part_path)
            for rel_id in used & set(relationships):
                entry = described[rel_id]
                rel = relationships[rel_id]
                assert entry["type"] == rel.type
                assert entry["target"] == (rel.target if rel.is_external else rel.target_part)
                if not rel.is_external:
                    assert package.has_part(entry["target"])
            if group.get(P + "kind") == "picture" and group.get(P + "image"):
                assert package.has_part(group.get(P + "image"))
            if group.get(P + "chart"):
                assert group.get(P + "chart").startswith("ppt/charts/")


def test_the_full_state_svg_draws_like_the_plain_one(pptx_path):
    """Only attributes are added: the drawing is the same, element for element."""
    document = Document.open(pptx_path.read_bytes())
    for slide in document.slides:
        plain = slide.render_svg()
        full = slide.render_svg(full_state=True)
        assert _strip_state(full) == etree.tostring(etree.fromstring(plain.encode()),
                                                    method="c14n")


def test_the_full_state_svg_renders_the_same_pixels(pptx_path):
    resvg = pytest.importorskip("resvg_py", reason="pixel comparison needs resvg-py")
    document = Document.open(pptx_path.read_bytes())
    for slide in document.slides:
        plain = bytes(resvg.svg_to_bytes(svg_string=slide.render_svg()))
        full = bytes(resvg.svg_to_bytes(svg_string=slide.render_svg(full_state=True)))
        assert plain == full, f"slide {slide.slide_id} draws differently"


def test_an_undrawn_shape_still_gets_a_group_in_its_place(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    hidden = slide.shapes[1]
    cnv_pr(hidden._element).set("hidden", "1")
    slide._touch()

    svg = FullStateSvg(slide.render_svg(full_state=True))

    group = svg.group(hidden.id)
    assert group.get(P + "unrendered") == "1" and len(group) == 0
    order = [i for i in (node.get("data-pptx-id") for node in svg.root.iter()) if i]
    assert order.index(hidden.id) == [s.id for s in slide.shapes].index(hidden.id)


def test_a_group_carries_its_shell_and_nests_its_children(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    members = [s for s in slide.shapes if s.left is not None][:2]
    group = slide.group(members)

    svg = FullStateSvg(slide.render_svg(full_state=True))

    shell = svg.xml(group.id)
    assert not [c for c in shell if c.tag.rpartition("}")[2] in {"sp", "pic", "grpSp"}]
    assert svg.get(group.id, "ch-cx") is not None
    for member in group.children:
        node = svg.group(member.id)
        assert svg.group(group.id) in node.iterancestors()


def test_emitting_never_dirties_the_document(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    for slide in document.slides:
        slide.render_svg(full_state=True)
    assert document.package.dirty_parts == frozenset()
    assert document.to_bytes() == Document.open(pptx_path.read_bytes()).to_bytes()


def test_full_state_needs_this_librarys_ids(product_page):
    document = Document.open(product_page.read_bytes())
    with pytest.raises(ValueError):
        document.slides[0].render_svg(full_state=True, rewrite_ids=False)
