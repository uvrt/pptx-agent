"""Writing the full-state SVG: pptx2svg's picture, annotated from the unresolved document.

pptx2svg draws the slide and marks every shape's group with ``data-pptx-id``.  This module
finds each of those groups and adds the shape's state as ``data-ooxml-*`` attributes, read
from the document model -- never from the picture, which holds resolved values -- plus the
shape's own OOXML as base64, the raw-XML floor that makes the SVG lossless whatever the
typed vocabulary misses.  Only attributes are added, so the SVG draws exactly as before.

A shape pptx2svg did not draw (hidden, or nothing to paint) still gets an empty group, inside
its parent group's, so the SVG lists every shape on the slide: a reader that deletes shapes
missing from an SVG must never mistake "not drawn" for "deleted".
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, NAMESPACES, find, local_name, qn
from .model import PREFIX, SHAPE_TAG_NAMES, VOCABULARY, Resolver, shape_attributes
from .safe import dump_json, encode_base64, parse_untrusted

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Slide

SVG_NS = "http://www.w3.org/2000/svg"
_REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
_REL_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
_SHAPE_TAGS = frozenset(qn(f"p:{name}") for name in SHAPE_TAG_NAMES)


def resolver_for(slide: "Slide") -> Resolver:
    document = slide.document
    return Resolver(document.package.relationships(slide.part_path),
                    {s.part_path: s.slide_id for s in document.slides}, document.package,
                    shape_ids(slide))


def shape_ids(slide: "Slide") -> dict[str, str]:
    """``cNvPr@id`` -> shape id, for the ``cNvPr@id``s that are unique on the slide."""
    from ..edit.ids import cnv_pr

    found: dict[str, str] = {}
    seen: set[str] = set()
    for shape_id, element in slide._ensure_index():
        properties = cnv_pr(element)
        raw = None if properties is None else properties.get("id")
        if not raw:
            continue
        if raw in seen:
            found.pop(raw, None)
        else:
            found[raw] = str(shape_id)
        seen.add(raw)
    return found


def fragment(element: Element) -> Element:
    """The shape's OOXML as a self-contained element, every in-scope namespace declared.

    A group's fragment is its *shell*: the group without its child shapes, which travel as
    their own SVG groups -- so nothing is carried twice, and a child edited in the SVG cannot
    disagree with a stale copy of itself inside its parent.
    """
    if local_name(element) != "grpSp":
        return element
    shell = etree.Element(element.tag, attrib=dict(element.attrib), nsmap=element.nsmap)
    shell.text = element.text
    for child in element:
        if child.tag not in _SHAPE_TAGS:
            shell.append(copy.deepcopy(child))
    return shell


def fragment_bytes(element: Element) -> bytes:
    # Serialising a subelement declares every namespace in scope -- including ones only
    # named inside attribute values (mc:Ignorable, Requires) -- so it parses on its own.
    return etree.tostring(fragment(element), with_tail=False)


def relationship_ids(element: Element) -> list[str]:
    """Relationship ids an element's XML uses (``r:id``, ``r:embed``...), in order.

    Pass a :func:`fragment`, so a group's children's ids are left to the children."""
    namespace = NAMESPACES["r"]
    found: dict[str, None] = {}
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        for name, value in node.attrib.items():
            if name.startswith("{%s}" % namespace) and value:
                found[value] = None
    return list(found)


def describe_relationships(element: Element, resolver: Resolver) -> list[dict]:
    """``[{"id": "rId2", "type": ..., "target": "ppt/media/image1.png"}, ...]``.

    Targets are *references into the package*, not embedded copies: a picture's bytes, a
    chart's part and its workbook stay where they are, and the SVG says which ones it means.
    """
    described = []
    for rel_id in relationship_ids(element):
        rel = resolver.rel(rel_id)
        if rel is None:
            continue
        entry = {"id": rel_id, "type": rel.type,
                 "target": rel.target if rel.is_external else (rel.target_part or "")}
        if rel.is_external:
            entry["external"] = True
        described.append(entry)
    return described


def slide_attributes(slide: "Slide") -> dict[str, str]:
    """The slide-level part of the vocabulary, written on the root ``<svg>``."""
    package = slide.document.package
    attributes = {
        "vocabulary": VOCABULARY,
        "slide-id": str(slide.slide_id),
        "slide-index": str(slide.index),
        "slide-part": slide.part_path,
    }
    layout = slide.layout
    if layout is not None:
        attributes["layout"] = layout.part_path
        attributes["layout-name"] = layout.name
        attributes["master"] = layout.master_part
        themes = package.related_parts_of_type(layout.master_part, _REL_THEME)
        if themes:
            attributes["theme"] = themes[0]
            theme = package.tree(themes[0])
            if theme is not None and theme.get("name") is not None:
                attributes["theme-name"] = theme.get("name")
    presentation = package.tree(package.presentation_part())
    size = find(presentation, "p:sldSz") if presentation is not None else None
    if size is not None:
        attributes["slide-cx"] = size.get("cx", "")
        attributes["slide-cy"] = size.get("cy", "")
    return attributes


def _placeholder(element: Element, groups: dict[int, Element], root: Element) -> Element:
    """An empty group for a shape pptx2svg did not draw, where its z-order puts it."""
    group = etree.Element("{%s}g" % SVG_NS)
    for sibling, attach in ((element.itersiblings(preceding=True), "addnext"),
                            (element.itersiblings(), "addprevious")):
        neighbour = next((groups[id(node)] for node in sibling if id(node) in groups), None)
        if neighbour is not None:
            getattr(neighbour, attach)(group)
            return group
    parent = element.getparent()
    container = groups.get(id(parent), root) if parent is not None else root
    container.append(group)
    return group


def emit_full_state(slide: "Slide", svg: str) -> str:
    """Annotate ``svg`` -- pptx2svg's rendering of ``slide``, ids already rewritten."""
    root = parse_untrusted(svg, limit=1 << 31, what="pptx2svg's SVG")
    for key, value in slide_attributes(slide).items():
        root.set(PREFIX + key, value)

    drawn: dict[str, Element] = {}
    for node in root.iter():
        if isinstance(node.tag, str) and node.get("data-pptx-id"):
            drawn.setdefault(node.get("data-pptx-id"), node)

    resolver = resolver_for(slide)
    shapes = list(slide._ensure_index())
    groups: dict[int, Element] = {  # id(shape element) -> its SVG group
        id(element): drawn[str(shape_id)] for shape_id, element in shapes
        if str(shape_id) in drawn}
    for shape_id, element in shapes:
        identifier = str(shape_id)
        target = groups.get(id(element))
        if target is None:
            target = _placeholder(element, groups, root)
            target.set("data-pptx-id", identifier)
            target.set(PREFIX + "unrendered", "1")
            groups[id(element)] = target
        for key, value in shape_attributes(element, resolver).items():
            target.set(PREFIX + key, value)
        own = fragment(element)
        relationships = describe_relationships(own, resolver)
        if relationships:
            target.set(PREFIX + "rels", dump_json(relationships))
        target.set(PREFIX + "xml", encode_base64(fragment_bytes(element)))
    return etree.tostring(root, encoding="unicode")
