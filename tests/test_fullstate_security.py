"""The reader parses untrusted input, and refuses what it should -- before changing anything.

XXE and entity expansion are refused in the SVG and in the XML inside it; SVG this library did
not write is refused; sizes are bounded; base64, JSON and the shape XML are validated; and
relationships cannot be pointed at the package's skeleton.  Every test also checks that the
document is untouched afterwards.
"""

from __future__ import annotations

import base64

import pytest
from lxml import etree

from pptx_agent import Document
from pptx_agent.fullstate import FullStateError, Limits, NotFullStateSvg, UnsafeInput
from pptx_agent.fullstate.safe import _parser, parse_untrusted
from svgedit import P, FullStateSvg, red_square

pytest.importorskip("pptx2svg", reason="the full-state SVG is drawn by pptx2svg")


@pytest.fixture
def deck(product_page):
    document = Document.open(product_page.read_bytes())
    original = document.to_bytes()
    yield document, FullStateSvg(document.slides[0].render_svg(full_state=True))
    assert document.to_bytes() == original, "a refused SVG changed the document"
    assert not document.history.can_undo()


def _with_doctype(svg: str, doctype: str) -> str:
    return doctype + svg


def test_an_external_entity_is_refused(deck, tmp_path):
    document, svg = deck
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET")
    target = next(iter(svg.groups))
    svg.set(target, "name", "PLACEHOLDER")
    payload = str(svg).replace('data-ooxml-name="PLACEHOLDER"', 'data-ooxml-name="&xxe;"')
    payload = _with_doctype(payload, f'<!DOCTYPE svg [<!ENTITY xxe SYSTEM "file://{secret}">]>')

    with pytest.raises(UnsafeInput):
        document.apply_svg(payload)
    assert "TOP-SECRET" not in document.slides[0].render_svg(full_state=True)


def test_a_billion_laughs_is_refused(deck):
    document, svg = deck
    entities = ['<!ENTITY lol "lol">'] + [
        f'<!ENTITY lol{n} "{("&lol" + (str(n - 1) if n > 1 else "") + ";") * 10}">'
        for n in range(1, 10)]
    payload = _with_doctype(str(svg).replace("<svg ", "<svg data-x=\"&lol9;\" ", 1),
                            "<!DOCTYPE svg [" + "".join(entities) + "]>")
    with pytest.raises(UnsafeInput):
        document.apply_svg(payload)


def test_entities_inside_the_raw_xml_are_refused(deck, tmp_path):
    document, svg = deck
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET")
    target = svg.where(kind="shape")[0]
    raw = etree.tostring(svg.xml(target)).decode()
    raw = raw.replace('name="', 'name="&xxe;', 1)
    raw = f'<!DOCTYPE p:sp [<!ENTITY xxe SYSTEM "file://{secret}">]>' + raw
    svg.set(target, "xml", base64.b64encode(raw.encode()).decode())
    with pytest.raises(UnsafeInput):
        document.apply_svg(str(svg))


def test_the_parser_would_not_expand_entities_even_past_the_precheck(tmp_path):
    """Defence in depth: the parser itself neither expands nor fetches."""
    secret = tmp_path / "secret.txt"
    secret.write_text("TOP-SECRET")
    for document in (b'<!DOCTYPE x [<!ENTITY e "expanded">]><x>&e;</x>',
                     f'<!DOCTYPE x [<!ENTITY e SYSTEM "{secret.as_uri()}">]><x>&e;</x>'.encode()):
        root = etree.fromstring(document, _parser())
        assert root.text is None and len(root) == 1  # an unexpanded entity reference
        assert b"expanded" not in etree.tostring(root) and b"TOP-SECRET" not in etree.tostring(root)
        with pytest.raises(UnsafeInput):
            parse_untrusted(document, limit=1000, what="test")


@pytest.mark.parametrize("doctype", ["<!DOCTYPE svg>", "<!doctype svg>", "<! ENTITY x 'y'>"])
def test_any_document_type_declaration_is_refused(deck, doctype):
    document, svg = deck
    with pytest.raises(UnsafeInput):
        document.apply_svg(doctype + str(svg))


def test_plain_svg_is_not_ours(deck):
    """pptx2svg's own SVG, without the vocabulary, is refused: there is nothing to read."""
    document, _ = deck
    with pytest.raises(NotFullStateSvg):
        document.apply_svg(document.slides[0].render_svg())


def test_foreign_svg_is_not_ours(deck):
    document, _ = deck
    figma = ('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
             '<rect x="10" y="10" width="50" height="50" fill="#FF0000"/></svg>')
    with pytest.raises(NotFullStateSvg):
        document.apply_svg(figma)
    with pytest.raises(NotFullStateSvg):
        document.apply_svg('<html xmlns="http://www.w3.org/1999/xhtml"/>')


def test_another_vocabulary_version_is_refused(deck):
    document, svg = deck
    svg.root.set(P + "vocabulary", "pptx-agent/99")
    with pytest.raises(NotFullStateSvg, match="unsupported"):
        document.apply_svg(str(svg))


def test_malformed_xml_is_refused(deck):
    document, svg = deck
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg)[:-10])


def test_size_limits_are_enforced(deck):
    document, svg = deck
    text = str(svg)
    with pytest.raises(UnsafeInput):
        document.apply_svg(text, limits=Limits(svg_bytes=len(text.encode()) - 1))
    with pytest.raises(UnsafeInput):
        document.apply_svg(text, limits=Limits(xml_attribute=64))
    with pytest.raises(UnsafeInput):
        document.apply_svg(text, limits=Limits(json_attribute=16))
    with pytest.raises(FullStateError, match="more than"):
        document.apply_svg(text, limits=Limits(shapes=3))


@pytest.mark.parametrize("value", ["not base64!", "QUJD\nREVG", "4pyTIMOgIGxhIG1vZGU="])
def test_invalid_base64_or_xml_is_refused(deck, value):
    document, svg = deck
    svg.set(svg.where(kind="shape")[0], "xml", value)
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg))


def test_raw_xml_must_be_a_shape(deck):
    document, svg = deck
    target = svg.where(kind="shape")[0]
    bogus = b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    svg.set(target, "xml", base64.b64encode(bogus).decode())
    with pytest.raises(FullStateError, match="not a shape"):
        document.apply_svg(str(svg))


@pytest.mark.parametrize("feature,value", [
    ("text", "{not json"),
    ("text", '{"p": "nope"}'),
    ("text", '{"p": [{"c": [{"t": "x", "sz": "huge"}]}]}'),
    ("text", '{"p": [{"c": [{"t": "x", "color": "scheme:accent99"}]}]}'),
    ("fill-scheme", "chartreuse"),
    ("fill-mods", "lumMod=lots"),
    ("x", "12.5cm"),
    ("rot", "NaN"),
    ("geom", "blob"),
    ("rels", '[{"id": 1}]'),
])
def test_typed_values_are_validated(deck, feature, value):
    document, svg = deck
    target = next(i for i in svg.where(kind="shape") if svg.get(i, "text"))
    if feature.startswith("fill-"):
        svg.set_fill(target, "solid", scheme="accent1")
    svg.set(target, feature, value)
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg))


def test_relationships_to_the_package_skeleton_are_refused(deck):
    document, svg = deck
    rels = [{"id": "rId1", "type": "http://schemas.openxmlformats.org/officeDocument/2006/"
                                    "relationships/slideLayout",
             "target": "ppt/slideLayouts/slideLayout1.xml"}]
    svg.add(red_square(0, 0, 914400), rels=rels)
    with pytest.raises(FullStateError, match="cannot relate"):
        document.apply_svg(str(svg))


def test_relationships_to_parts_that_are_not_there_are_refused(deck):
    document, svg = deck
    square = red_square(0, 0, 914400)
    blip = etree.SubElement(square.find(".//{*}spPr"), "{http://schemas.openxmlformats.org/"
                            "drawingml/2006/main}blipFill")
    etree.SubElement(blip, "{http://schemas.openxmlformats.org/drawingml/2006/main}blip").set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "rId77")
    rels = [{"id": "rId77", "type": "http://schemas.openxmlformats.org/officeDocument/2006/"
                                    "relationships/image", "target": "../../etc/passwd"}]
    svg.add(square, rels=rels)
    with pytest.raises(FullStateError, match="not a part"):
        document.apply_svg(str(svg))


def test_an_undefined_relationship_id_is_refused(deck):
    document, svg = deck
    square = red_square(0, 0, 914400)
    link = etree.SubElement(square.find(".//{*}cNvPr"), "{http://schemas.openxmlformats.org/"
                            "drawingml/2006/main}hlinkClick")
    link.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "rId404")
    svg.add(square)
    with pytest.raises(FullStateError, match="rId404"):
        document.apply_svg(str(svg))


def test_duplicate_ids_are_refused(deck):
    document, svg = deck
    first = next(iter(svg.groups))
    clone = etree.fromstring(etree.tostring(svg.group(first)))
    svg.root.append(clone)
    with pytest.raises(FullStateError, match="twice"):
        document.apply_svg(str(svg))


def test_rels_json_is_validated(deck):
    document, svg = deck
    svg.add(red_square(0, 0, 914400), rels=[{"id": "rId1", "type": "x", "target": "y",
                                             "external": "yes"}])
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg))


# -- connections and adjust values (E5) -----------------------------------------------------


@pytest.fixture
def wired(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    a = slide.add_shape("roundRect", 914400, 914400, 914400, 914400)
    b = slide.add_shape("ellipse", 4572000, 2743200, 914400, 914400)
    connector = slide.add_connector("elbow", (a, 3), (b, 2))
    original = document.to_bytes()
    depth = len(document.history._undo)
    yield document, FullStateSvg(slide.render_svg(full_state=True)), a, b, connector
    assert document.to_bytes() == original, "a refused SVG changed the document"
    assert len(document.history._undo) == depth


@pytest.mark.parametrize("feature,value,match", [
    ("adj", "adj", "should read"),
    ("adj", "adj=lots", "not an integer"),
    ("adj", "adj=1 adj=2", "twice"),
    ("adj", "adj=99999999999", "out of range"),
    ("adj", "nope=5", "no adjustment"),
    ("adj", "<script>=1", "should read"),
    ("cxn-end", "elsewhere", "not an integer"),
    ("cxn-end", " ", "should be"),
    ("cxn-end", "a b c", "not an integer"),
    ("cxn-end", "999.999 1", "no shape"),
    ("cxn-end", "SHAPE 99999999", "out of range"),
    ("cxn-end", "SHAPE 41", "connection sites"),
    ("cxn-end", "SELF 0", "itself"),
])
def test_connections_and_adjust_values_are_validated(wired, feature, value, match):
    document, svg, a, b, connector = wired
    value = value.replace("SHAPE", b.id).replace("SELF", connector.id)
    target = a.id if feature == "adj" else connector.id
    svg.set(target, feature, value)
    with pytest.raises(FullStateError, match=match):
        document.apply_svg(str(svg))


def test_only_a_connector_has_connections(wired):
    document, svg, a, b, connector = wired
    svg.set(a.id, "cxn-begin", f"{b.id} 0")
    with pytest.raises(FullStateError, match="connector"):
        document.apply_svg(str(svg))


def test_a_custom_geometry_takes_no_typed_adjust_values(wired):
    document, svg, a, b, connector = wired
    svg.set(a.id, "geom", "custom")
    svg.set(a.id, "adj", "adj=5")
    with pytest.raises(FullStateError, match="custom"):
        document.apply_svg(str(svg))


def test_the_number_of_adjust_values_is_bounded(wired):
    document, svg, a, b, connector = wired
    svg.set(a.id, "adj", " ".join(f"g{n}=1" for n in range(9)))
    with pytest.raises(FullStateError, match="more than 8"):
        document.apply_svg(str(svg), limits=Limits(adjustments=8))

