"""Editing a full-state SVG the way an agent or another process would: as SVG, with lxml.

Nothing here goes through pptx-agent's API -- the point of the tests that use it is that the
SVG alone carries enough to describe an edit.
"""

from __future__ import annotations

import base64
import json

from lxml import etree

P = "data-ooxml-"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PML = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
SVG = "{http://www.w3.org/2000/svg}"
EMU_PER_INCH = 914400


class FullStateSvg:
    def __init__(self, svg: str) -> None:
        self.root = etree.fromstring(svg.encode("utf-8"))

    def __str__(self) -> str:
        return etree.tostring(self.root, encoding="unicode")

    # -- finding -----------------------------------------------------------------------------

    @property
    def groups(self) -> dict[str, etree._Element]:
        return {node.get("data-pptx-id"): node for node in self.root.iter()
                if isinstance(node.tag, str) and node.get(P + "xml") is not None}

    def group(self, identifier: str) -> etree._Element:
        return self.groups[identifier]

    def where(self, **wanted) -> list[str]:
        """Ids of groups whose ``data-ooxml-<key>`` equals each value given."""
        return [identifier for identifier, node in self.groups.items()
                if all(node.get(P + key.replace("_", "-")) == value
                       for key, value in wanted.items())]

    # -- attributes --------------------------------------------------------------------------

    def get(self, identifier: str, key: str) -> str | None:
        return self.group(identifier).get(P + key)

    def set(self, identifier: str, key: str, value: str | None) -> None:
        node = self.group(identifier)
        if value is None:
            node.attrib.pop(P + key, None)
        else:
            node.set(P + key, value)

    def json(self, identifier: str, key: str):
        return json.loads(self.get(identifier, key))

    def set_json(self, identifier: str, key: str, value) -> None:
        self.set(identifier, key, json.dumps(value))

    def set_fill(self, identifier: str, kind: str, **fields: str) -> None:
        node = self.group(identifier)
        for name in list(node.attrib):
            if name == P + "fill" or name.startswith(P + "fill-"):
                del node.attrib[name]
        node.set(P + "fill", kind)
        for key, value in fields.items():
            node.set(f"{P}fill-{key}", value)

    # -- the raw floor -----------------------------------------------------------------------

    def xml(self, identifier: str) -> etree._Element:
        return etree.fromstring(base64.b64decode(self.get(identifier, "xml")))

    def set_xml(self, identifier: str, element: etree._Element) -> None:
        self.set(identifier, "xml", base64.b64encode(etree.tostring(element)).decode())

    def add(self, element: etree._Element, *, parent: str | None = None,
            identifier: str | None = None, rels: list | None = None) -> etree._Element:
        """A new shape group, frontmost in ``parent`` (the slide by default)."""
        container = self.root if parent is None else self.group(parent)
        group = etree.SubElement(container, SVG + "g")
        if identifier is not None:
            group.set("data-pptx-id", identifier)
        group.set(P + "xml", base64.b64encode(etree.tostring(element)).decode())
        if rels:
            group.set(P + "rels", json.dumps(rels))
        return group

    def remove(self, identifier: str) -> None:
        node = self.group(identifier)
        node.getparent().remove(node)


def red_square(left: int, top: int, size: int, *, name: str = "SVG marker",
               shape_id: int = 2) -> etree._Element:
    """A red rectangle as raw PresentationML -- a new shape no typed attribute could make."""
    xml = (
        f'<p:sp xmlns:p="{PML[1:-1]}" xmlns:a="{A[1:-1]}">'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{size}" cy="{size}"/>'
        f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'</p:sp>'
    )
    return etree.fromstring(xml)


def with_shadow(shape: etree._Element) -> etree._Element:
    """Give a shape's raw XML an outer shadow -- an effect the typed vocabulary does not know."""
    from pptx_agent.oxml.xml import replace_choice

    properties = shape.find(PML + "spPr")
    effects = etree.Element(A + "effectLst")
    shadow = etree.SubElement(effects, A + "outerShdw", blurRad="50800", dist="38100",
                              dir="2700000", algn="tl", rotWithShape="0")
    color = etree.SubElement(shadow, A + "prstClr", val="black")
    etree.SubElement(color, A + "alpha", val="40000")
    replace_choice(properties, ("a:effectLst", "a:effectDag"), effects)
    return shape


#: Where the marker square sits, and where to probe a rasterised page for it (inches).
MARKER_BOX = (0.15, 0.15, 0.5)
MARKER_PROBE = (0.4, 0.4)
SENTINEL = "SVG-EDIT"


def first_run_path(model: dict) -> tuple[int, int] | None:
    """``(paragraph, item)`` of the first non-empty plain run in a text model."""
    for p, paragraph in enumerate(model["p"]):
        for c, item in enumerate(paragraph["c"]):
            if "t" in item and "fld" not in item and item["t"].strip():
                return p, c
    return None


def acceptance_edits(document) -> dict:
    """E3's acceptance edits, made *only* by editing each slide's full-state SVG.

    On every slide: a theme fill on one shape, the first text run retyped to the sentinel, a
    shape moved, a table cell retyped, a shadow added through the raw floor, and a red marker
    square added as a new shape from raw XML.  Returns what to look for afterwards.
    """
    expected = {"pages": len(document.slides), "fills": [], "texts": [], "moves": [],
                "cells": [], "shadows": [], "markers": [], "text_pages": [], "cell_pages": []}
    for page, slide in enumerate(list(document.slides)):
        svg = FullStateSvg(slide.render_svg(full_state=True))
        autoshapes = svg.where(kind="shape")
        if autoshapes:
            target = autoshapes[0]
            svg.set_fill(target, "solid", scheme="accent2", mods="lumMod=75000")
            expected["fills"].append(target)
        for identifier in autoshapes:
            if svg.get(identifier, "text") is None:
                continue
            model = svg.json(identifier, "text")
            at = first_run_path(model)
            if at is None:
                continue
            model["p"][at[0]]["c"][at[1]]["t"] = SENTINEL
            svg.set_json(identifier, "text", model)
            expected["texts"].append(identifier)
            expected["text_pages"].append(page)
            break
        movable = [i for i, node in svg.groups.items()
                   if node.get(P + "x") not in (None, "inherit")
                   and node.getparent() is svg.root]
        if movable:
            target = movable[-1]
            svg.set(target, "x", str(int(svg.get(target, "x")) + EMU_PER_INCH // 4))
            expected["moves"].append((target, int(svg.get(target, "x"))))
        tables = [i for i in svg.groups if svg.get(i, "table") is not None]
        if tables:
            model = svg.json(tables[0], "table")
            cell = model["cells"][0][0]
            at = first_run_path(cell["text"]) if "text" in cell else None
            if at is None:
                cell["text"] = {"p": [{"c": [{"t": "SVG-CELL"}]}]}
            else:  # retype the run, keeping its formatting
                paragraph = cell["text"]["p"][at[0]]
                paragraph["c"] = [dict(paragraph["c"][at[1]], t="SVG-CELL")]
                cell["text"]["p"] = [paragraph]
            svg.set_json(tables[0], "table", model)
            expected["cells"].append(tables[0])
            expected["cell_pages"].append(page)
        if autoshapes:
            shadowed = autoshapes[-1]
            svg.set_xml(shadowed, with_shadow(svg.xml(shadowed)))
            expected["shadows"].append(shadowed)
        left, top, size = (round(value * EMU_PER_INCH) for value in MARKER_BOX)
        svg.add(red_square(left, top, size), identifier="new-marker")
        report = slide.apply_svg(str(svg))
        expected["markers"].append(report.added["new-marker"])
    return expected
