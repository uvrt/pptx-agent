"""Where a placeholder's geometry and text formatting come from when it does not say.

PowerPoint resolves a slide placeholder through its layout placeholder (matched by ``idx``,
then by type), then the master placeholder of the same kind -- a title from the master's
title, footers from theirs, everything else from the master's body -- and then, for text,
the master's text style for that kind (``titleStyle``, ``bodyStyle`` or ``otherStyle``) and
the presentation's ``defaultTextStyle``.  This is the chain :mod:`pptx_agent.outline.read`
reads bullets through and pptx2svg resolves text with; here it is for sizes, fonts and
spacing that an agent asks about (``Run.effective_size``, ``Layout.placeholders``).
"""

from __future__ import annotations

from ..oxml.package import REL_SLIDE_LAYOUT, REL_SLIDE_MASTER
from ..oxml.xml import Element, find, local_name, qn

TITLE_TYPES = ("title", "ctrTitle")
#: Placeholder types whose list style comes from the master's ``bodyStyle``.
BODY_TYPES = frozenset({"body", "subTitle", "obj"})
#: Placeholder type -> the master placeholder it inherits from.
MASTER_KIND = {"title": "title", "ctrTitle": "title", "dt": "dt", "ftr": "ftr",
               "sldNum": "sldNum", "hdr": "hdr"}
#: What a run is when nothing on the way says: 18 pt.
DEFAULT_SIZE = 18.0
#: Placeholder containers by shape tag.
_NV = {"sp": "p:nvSpPr", "pic": "p:nvPicPr", "graphicFrame": "p:nvGraphicFramePr",
       "grpSp": "p:nvGrpSpPr", "cxnSp": "p:nvCxnSpPr"}


def placeholder_of(shape: Element) -> Element | None:
    """The shape's ``p:ph``, or ``None``."""
    container = _NV.get(local_name(shape))
    return None if container is None else find(shape, f"{container}/p:nvPr/p:ph")


def master_kind(kind: str | None) -> str:
    return MASTER_KIND.get(kind, "body")


def text_style(kind: str | None, is_placeholder: bool = True) -> str:
    """The master text style a placeholder of ``kind`` takes its levels from."""
    if not is_placeholder:
        return "otherStyle"
    if kind in TITLE_TYPES:
        return "titleStyle"
    return "bodyStyle" if (kind or "obj") in BODY_TYPES else "otherStyle"


def layout_of(package, slide_part: str) -> str | None:
    related = package.related_parts_of_type(slide_part, REL_SLIDE_LAYOUT)
    return related[0] if related else None


def master_of(package, layout_part: str | None) -> str | None:
    if layout_part is None:
        return None
    related = package.related_parts_of_type(layout_part, REL_SLIDE_MASTER)
    return related[0] if related else None


def master_placeholder(package, master_part: str | None, kind: str) -> Element | None:
    """The master's placeholder a placeholder of master kind ``kind`` inherits from."""
    root = package.tree(master_part) if master_part else None
    tree = find(root, "p:cSld/p:spTree") if root is not None else None
    for shape in [] if tree is None else tree:
        ph = placeholder_of(shape)
        if ph is not None and master_kind(ph.get("type")) == kind:
            return shape
    return None


def list_styles(package, master_part: str | None, own: Element | None,
                layout_shape: Element | None, kind: str | None,
                is_placeholder: bool) -> list[Element | None]:
    """The list styles a text body's levels inherit from, nearest first: its own (``own``,
    the shape's ``a:lstStyle``), the layout placeholder's, the master placeholder's, the
    master's text style, the presentation's default."""
    chain: list[Element | None] = [own]
    if is_placeholder:
        if layout_shape is not None:
            chain.append(layout_shape.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"))
        master_shape = master_placeholder(package, master_part, master_kind(kind))
        if master_shape is not None:
            chain.append(master_shape.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"))
    root = package.tree(master_part) if master_part else None
    if root is not None:
        chain.append(find(root, f"p:txStyles/p:{text_style(kind, is_placeholder)}"))
    presentation = package.tree(package.presentation_part())
    if presentation is not None:
        chain.append(find(presentation, "p:defaultTextStyle"))
    return chain


def level_properties(chain: list[Element | None], level: int) -> list[Element]:
    """Each style's ``a:lvl<n>pPr`` for a 0-based ``level``, nearest first."""
    found = []
    for style in chain:
        if style is None:
            continue
        node = style.find(qn(f"a:lvl{level + 1}pPr"))
        if node is not None:
            found.append(node)
    return found


def level_size(chain: list[Element | None], level: int) -> float:
    """The size, in points, a level's runs take when they set none."""
    for properties in level_properties(chain, level):
        run = properties.find(qn("a:defRPr"))
        if run is not None and run.get("sz"):
            return int(run.get("sz")) / 100
    return DEFAULT_SIZE


def level_typeface(chain: list[Element | None], level: int, script: str = "a:latin") -> str | None:
    for properties in level_properties(chain, level):
        run = properties.find(qn("a:defRPr"))
        font = run.find(qn(script)) if run is not None else None
        if font is not None and font.get("typeface"):
            return font.get("typeface")
    return None


def frame_of(shape: Element | None) -> tuple[int, int, int, int] | None:
    """``(left, top, width, height)`` from a shape's own ``a:xfrm``, or ``None``."""
    if shape is None:
        return None
    holder = shape.find(qn("p:xfrm")) if local_name(shape) == "graphicFrame" else \
        find(shape, "p:grpSpPr/a:xfrm" if local_name(shape) == "grpSp" else "p:spPr/a:xfrm")
    if holder is None:
        return None
    offset, extent = holder.find(qn("a:off")), holder.find(qn("a:ext"))
    if offset is None or extent is None:
        return None
    return (int(offset.get("x", "0")), int(offset.get("y", "0")),
            int(extent.get("cx", "0")), int(extent.get("cy", "0")))


__all__ = ["BODY_TYPES", "DEFAULT_SIZE", "MASTER_KIND", "TITLE_TYPES", "frame_of",
           "layout_of", "level_properties", "level_size", "level_typeface", "list_styles",
           "master_kind", "master_of", "master_placeholder", "placeholder_of", "text_style"]
