"""A finished slide moved onto another layout, as PowerPoint's Layout menu does it.

What PowerPoint writes was measured in PowerPoint 16 (Mac) on made-up slides
(agent-work/feedback-round7/oracle/results.md, A1-A9, B2 and A2 of the follow-up on the
sample, l2-check/), and :func:`set_layout` writes the same:

1. **Empty placeholders go**; the new layout's placeholders the slide then lacks are added
   empty -- the lowest free shape ids (counted before the empty ones went), type and idx (no
   ``sz``), an empty ``spPr``, a text body with one empty paragraph (none for a picture or
   table placeholder).  PowerPoint names each "<kind> N" in its interface language (and
   gives it an ``a16:creationId``); here it takes the layout placeholder's name, as a new
   slide's placeholders do (:func:`~pptx_agent.edit.slides.slide_from_layout`).  Date,
   footer and slide-number placeholders are never added; the slide's own, empty or not,
   take the layout's of their kind (``sz`` kept), and an empty one with none there goes.
2. **Filled placeholders match by kind**, never by name or "same idx": a title takes a title
   or centred title; text content (a content, body or subtitle placeholder) takes a content,
   body or subtitle placeholder -- the same type first, then a compatible one; picture and
   table placeholders take no text.  A picture takes a picture placeholder, else a content
   one; a table a table or content placeholder; a chart (any other frame) a content, table
   or picture placeholder.
3. **Within a kind** the slide's and the layout's placeholders are each sorted by idx and
   paired in that order; names, stacking order and ``sz`` play no part, and an earlier
   orphan (idx 4294967295) is matched again like any other.
4. **A matched placeholder** takes the layout's type and idx (no ``type`` for content, no
   ``sz``) and loses its own box (``a:xfrm``), also when the layout is the one it is on: it
   moves to the layout's place.  Its text, run formatting, fill, outline, ``bodyPr``,
   paragraph formatting, name and id stay.
   - A picture moving to a picture placeholder (its own again too) is cropped to fill the
     box from the whole image, centred, any crop of its own replaced (``a:srcRect``, empty
     when the shapes agree).  Into a content placeholder it loses its crop and is drawn at
     its own size -- its pixels at the resolution it states, else 144 dpi (measured on a
     Retina Mac) -- made smaller to fit keeping its shape, never larger, centred; with no
     frame of its own when it fills the box exactly.
   - A table takes the box's place and width: its columns and its cells' left and right
     insets (91440 when unstated) scaled by the same factor, its rows kept, the frame as
     tall as they are.  A chart takes the box.
5. **A filled placeholder with no match** stays, an orphan: idx ``4294967295``, its type and
   ``sz`` kept, its box written on it -- its own or the old layout's -- with ``x`` moved to
   the slide's left edge when the box's centre is in the left half (at the middle too), else
   to its right edge (A5); a table or chart keeps its size.
6. **Order**: the placeholders gather into one block where the first placeholder was --
   matched and new ones in the layout's order, orphans last; other shapes keep theirs.
7. Other shapes are not touched; of the relationships only the layout's target changes.

Unmeasured, decided here: a placeholder inside a group is left as it is; a chart does not
take a body or subtitle placeholder; ``orient`` follows the layout, ``hasCustomPrompt``
goes; Windows was not measured.
"""

from __future__ import annotations

import copy
import posixpath
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from ..oxml.package import REL_SLIDE_LAYOUT, normalize_part_path, rels_path_for
from ..oxml.xml import Element, find, local_name, make, qn, remove
from . import inherit

if TYPE_CHECKING:  # pragma: no cover
    from .document import Slide
    from .slides import Layout

#: The idx PowerPoint gives a placeholder its layout has no match for (A5).
ORPHAN_IDX = "4294967295"
#: Content-class types, ``None`` being a content (``obj``) placeholder.
_TEXT = (None, "body", "subTitle")
_TITLES = ("title", "ctrTitle")
_FOOTERS = ("dt", "ftr", "sldNum", "hdr")
#: A picture moved into a content placeholder without a resolution of its own: 144 dpi, as
#: PowerPoint for Mac drew a 4x4 px PNG 2x2 pt (round 7, A2).
_DEFAULT_DPI = 144.0


@dataclass
class LayoutChange:
    """What :meth:`Slide.set_layout <pptx_agent.Slide.set_layout>` did.

    ``matched``: ``(shape id, layout placeholder id)`` for each filled placeholder that
    found a place.  ``added``: the empty placeholders the layout gave.  ``removed``: the
    slide's empty placeholders that went.  ``orphans``: ``(shape id, why)`` for content with
    no placeholder on the layout -- kept where it was drawn, snapped to a side.
    ``boxes_dropped``: ``(shape id, (x, y, w, h) EMU)`` for each matched placeholder whose own
    box was dropped for the layout's, as PowerPoint drops it."""

    slide: str
    layout: str
    matched: list[tuple[str, str]] = field(default_factory=list)
    added: list[str] = field(default_factory=list)
    removed: list[str] = field(default_factory=list)
    orphans: list[tuple[str, str]] = field(default_factory=list)
    boxes_dropped: list[tuple[str, tuple[int, int, int, int]]] = field(default_factory=list)


@dataclass
class _Item:
    element: Element
    ph: Element
    order: int

    @property
    def type(self) -> str | None:
        kind = self.ph.get("type")
        return None if kind == "obj" else kind

    @property
    def idx(self) -> int:
        text = self.ph.get("idx") or "0"
        return int(text) if text.isdigit() else 0

    @property
    def tag(self) -> str:
        return local_name(self.element)


_TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"


def _kind(item: _Item) -> str:
    """The class a slide placeholder's content matches within (A2; A2 of the follow-up)."""
    if item.tag == "pic":
        return "picture"
    if item.tag == "graphicFrame":
        data = find(item.element, "a:graphic/a:graphicData")
        return "table" if data is not None and data.get("uri") == _TABLE_URI else "frame"
    if item.type in _TITLES:
        return "title"
    if item.type in _FOOTERS:
        return item.type
    return "text"


def _accepts(kind: str, item: _Item, offered: _Item) -> bool:
    """Whether layout placeholder ``offered`` takes content of ``kind`` (``item``'s)."""
    if kind == "title":
        return offered.type in _TITLES
    if kind == "text":
        return offered.type in _TEXT or offered.type == item.type
    if kind == "picture":
        return offered.type in ("pic", None)
    if kind == "table":
        return offered.type in ("tbl", None)
    if kind == "frame":
        # A chart takes any placeholder but a title, a text-only or a footer one (measured:
        # table, picture and content placeholders).
        return offered.type not in (*_TITLES, *_FOOTERS, "body", "subTitle")
    return offered.type == kind


_WHAT = {"title": "title", "text": "content or text", "picture": "picture or content",
         "table": "table or content",
         "frame": "content, table or picture", "dt": "date", "ftr": "footer", "sldNum": "slide-number",
         "hdr": "header"}


def _filled(element: Element) -> bool:
    """Whether a slide placeholder holds anything: a picture, a frame, or text (a field
    counts)."""
    if local_name(element) != "sp":
        return True
    body = element.find(qn("p:txBody"))
    if body is None:
        return False
    if any(True for _ in body.iter(qn("a:fld"))):
        return True
    return any(node.text for node in body.iter(qn("a:t")))


def _placeholders(tree: Element) -> list[_Item]:
    out = []
    for order, shape in enumerate(tree):
        ph = inherit.placeholder_of(shape)
        if ph is not None and local_name(shape) in ("sp", "pic", "graphicFrame"):
            out.append(_Item(shape, ph, order))
    return out


def _pair(slide_items: list[_Item], offered: list[_Item]) -> dict[int, _Item]:
    """``{id(slide item): layout item}``: the same type first, then a compatible one, each
    side sorted by idx and paired in order (A2, A3)."""
    pairs: dict[int, _Item] = {}
    taken: set[int] = set()
    for exact in (True, False):
        groups: dict[tuple, list[_Item]] = {}
        for item in slide_items:
            if id(item) not in pairs:
                kind = _kind(item)
                key = (kind, item.type) if exact else (kind, None)
                groups.setdefault(key, []).append(item)
        for (kind, _), group in groups.items():
            group.sort(key=lambda item: (item.idx, item.order))
            free = sorted((place for place in offered if id(place) not in taken
                           and _accepts(kind, group[0], place)
                           and (not exact or place.type == group[0].type)),
                          key=lambda place: (place.idx, place.order))
            for item, place in zip(group, free):
                pairs[id(item)] = place
                taken.add(id(place))
    return pairs


def _xfrm_holder(element: Element) -> Element | None:
    if local_name(element) == "graphicFrame":
        return element
    return element.find(qn("p:spPr"))


def _own_xfrm(element: Element) -> Element | None:
    holder = _xfrm_holder(element)
    if holder is None:
        return None
    return holder.find(qn("p:xfrm" if local_name(element) == "graphicFrame" else "a:xfrm"))


def _box(xfrm: Element | None) -> tuple[int, int, int, int] | None:
    if xfrm is None:
        return None
    offset, extent = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if offset is None or extent is None:
        return None
    return (int(offset.get("x", "0")), int(offset.get("y", "0")),
            int(extent.get("cx", "0")), int(extent.get("cy", "0")))


def _write_xfrm(element: Element, box: tuple[int, int, int, int],
                template: Element | None = None) -> None:
    """Give ``element`` the box ``box`` as its own frame (``template``'s rotation and flips
    kept, when given)."""
    graphic = local_name(element) == "graphicFrame"
    old = _own_xfrm(element)
    if template is not None:
        new = copy.deepcopy(template)
        new.tag = qn("p:xfrm" if graphic else "a:xfrm")
    elif old is not None:
        new = old
    else:
        new = make("p:xfrm" if graphic else "a:xfrm")
    for child in list(new):
        if local_name(child) in ("off", "ext"):
            new.remove(child)
    new.insert(0, make("a:ext", cx=str(box[2]), cy=str(box[3])))
    new.insert(0, make("a:off", x=str(box[0]), y=str(box[1])))
    if old is not None and old is not new:
        old.addprevious(new)
        remove(old)
    elif old is None:
        holder = _xfrm_holder(element)
        if holder is None:
            holder = make("p:spPr")
            nv = element[0] if len(element) else None
            if nv is not None:
                nv.addnext(holder)
            else:
                element.append(holder)
        if graphic:
            nv = find(element, "p:nvGraphicFramePr")
            (nv.addnext(new) if nv is not None else element.insert(0, new))
        else:
            holder.insert(0, new)


def _rewrite_ph(ph: Element, source: Element, *, keep_size: bool = False) -> None:
    """``ph`` takes the layout placeholder's type, orient and idx (no ``type="obj"``, no
    ``sz``: A4; a date, footer or slide number keeps its ``sz``)."""
    for attribute in ("type", "orient", "idx", "hasCustomPrompt",
                      *(() if keep_size else ("sz",))):
        if attribute in ph.attrib:
            del ph.attrib[attribute]
    for attribute in ("type", "orient", "idx"):
        value = source.get(attribute)
        if value is not None and not (attribute == "type" and value == "obj"):
            ph.set(attribute, value)
    # Keep the schema's order of attributes readable: type, orient, sz, idx.
    attributes = dict(ph.attrib)
    ph.attrib.clear()
    for name in ("type", "orient", "sz", "idx", "hasCustomPrompt"):
        if name in attributes:
            ph.set(name, attributes.pop(name))
    for name, value in attributes.items():
        ph.set(name, value)


def _layout_box(document, layout_part: str, item: _Item) -> tuple[int, int, int, int] | None:
    """Where a layout placeholder is drawn: its own frame, else its master's of the kind."""
    own = inherit.frame_of(item.element)
    if own is not None:
        return own
    master = inherit.master_of(document.package, layout_part)
    shape = inherit.master_placeholder(document.package, master, inherit.master_kind(
        item.ph.get("type")))
    return inherit.frame_of(shape)


def _image(slide: "Slide", element: Element) -> bytes | None:
    blip = find(element, "p:blipFill/a:blip")
    if blip is None:
        return None
    part = slide.document.package.related_part(slide.part_path, blip.get(qn("r:embed")))
    if part is None:
        return None
    try:
        return slide.document.package.read(part)
    except KeyError:
        return None


def _picture_into_pic(slide: "Slide", element: Element, box) -> None:
    """A picture moved to a picture placeholder (also its own again): its crop replaced by
    one that fills ``box`` from the whole image, centred -- an empty ``a:srcRect`` when the
    shapes agree -- as PowerPoint writes it (B2; L2b S12; A2 of the follow-up, c)."""
    from ..oxml.package import image_size
    from ..outline.draft import _crop

    blip_fill = element.find(qn("p:blipFill"))
    data = _image(slide, element)
    if blip_fill is None or box is None or data is None:
        return
    rect = blip_fill.find(qn("a:srcRect"))
    new = make("a:srcRect", **_crop(box, image_size(data)))
    if rect is not None:
        rect.addprevious(new)
        remove(rect)
    else:
        blip = blip_fill.find(qn("a:blip"))
        (blip.addnext(new) if blip is not None else blip_fill.insert(0, new))


def native_size(data: bytes) -> tuple[float, float] | None:
    """An image's own size in EMU: its pixels at the resolution it states (a PNG's
    ``pHYs``, a JPEG's density), else at 144 dpi -- measured on a Retina Mac, where
    PowerPoint drew a PNG without ``pHYs`` at half a point per pixel (round 7, A2)."""
    from ..oxml.package import image_dpi, image_size

    pixels = image_size(data)
    if not pixels or not all(pixels):
        return None
    # Whole dots per inch, as PowerPoint draws them: a pHYs of 3780 per metre (96.01 dpi)
    # gave exactly 0.75 pt a pixel.
    dpi = [round(value) or _DEFAULT_DPI for value in image_dpi(data) or ()] or \
        [_DEFAULT_DPI, _DEFAULT_DPI]
    return pixels[0] / dpi[0] * 914400, pixels[1] / dpi[1] * 914400


def _picture_into_content(slide: "Slide", element: Element, box) -> None:
    """A picture moved into a content placeholder: the whole image (its crop removed) at
    its own size, made smaller to fit keeping its shape -- never larger -- centred in
    ``box``; no frame of its own when that fills the box exactly (A2 of the follow-up)."""
    blip_fill = element.find(qn("p:blipFill"))
    rect = blip_fill.find(qn("a:srcRect")) if blip_fill is not None else None
    if rect is not None:
        remove(rect)
    data = _image(slide, element)
    size = native_size(data) if data else None
    if box is None or size is None:
        return
    width, height = size
    scale = min(1.0, box[2] / width, box[3] / height)
    width, height = round(width * scale), round(height * scale)
    if abs(width - box[2]) <= 1 and abs(height - box[3]) <= 1:
        return
    left = box[0] + round((box[2] - width) / 2)
    top = box[1] + round((box[3] - height) / 2)
    _write_xfrm(element, (left, top, width, height))


#: A table cell's left and right inset when its ``tcPr`` states none (ECMA-376, 21.1.3.17).
_CELL_MARGIN = 91440


def _table_into(element: Element, box) -> None:
    """A table moved into a placeholder: at the box's place and width, its columns and the
    cells' left and right insets scaled by the same factor, its rows kept and the frame as
    tall as they are (measured, A2 of the follow-up, e: 91440 x 2.15 = 196596)."""
    grid = list(element.iter(qn("a:gridCol")))
    total = sum(int(col.get("w", "0")) for col in grid)
    if box is None or not total:
        return
    factor = box[2] / total
    for col in grid:
        col.set("w", str(round(int(col.get("w", "0")) * factor)))
    for cell in element.iter(qn("a:tc")):
        properties = cell.find(qn("a:tcPr"))
        if properties is None:
            properties = make("a:tcPr")
            extension = cell.find(qn("a:extLst"))
            (extension.addprevious(properties) if extension is not None
             else cell.append(properties))
        for side in ("marL", "marR"):
            value = properties.get(side)
            margin = int(value) if value and value.lstrip("-").isdigit() else _CELL_MARGIN
            properties.set(side, str(round(margin * factor)))
    height = sum(int(row.get("h", "0")) for row in element.iter(qn("a:tr")))
    _write_xfrm(element, (box[0], box[1], box[2], height))


def _retarget(package, slide_part: str, layout_part: str) -> None:
    """The slide's layout relationship points at ``layout_part``, its id kept (A9)."""
    from ..oxml.package import RELS_NS

    rels_path = rels_path_for(slide_part)
    root = package.tree(rels_path)
    base = posixpath.dirname(normalize_part_path(slide_part))
    target = posixpath.relpath(normalize_part_path(layout_part), base or ".")
    for node in root.findall("{%s}Relationship" % RELS_NS) if root is not None else ():
        if node.get("Type") == REL_SLIDE_LAYOUT:
            node.set("Target", target)
            package.mark_dirty(rels_path)
            return
    package.add_relationship(slide_part, REL_SLIDE_LAYOUT, layout_part)


def set_layout(slide: "Slide", layout: "Layout") -> LayoutChange:
    """See :meth:`Slide.set_layout <pptx_agent.Slide.set_layout>`."""
    from .templates import _language, empty_placeholder

    document = slide.document
    package = document.package
    layout_part = layout.part_path
    layout_tree = find(package.tree(layout_part), "p:cSld/p:spTree")
    if layout_tree is None:
        raise ValueError(f"{layout.id} has no shape tree")
    offered = _placeholders(layout_tree)
    tree = slide._sp_tree()
    slide._ensure_index()
    items = _placeholders(tree)
    slide_w = document.slide_size[0]
    report = LayoutChange(slide=f"s:{slide.slide_id}", layout=layout.id)

    def address(element: Element) -> str:
        return slide._wrap(element).id

    with document.batch():
        document.history.checkpoint()
        # A date, footer or slide number empty or not takes its kind's place on the layout
        # (kept, its id and name too); an empty one with none there goes (A2 follow-up, d).
        footers = [item for item in items if not _filled(item.element)
                   and item.type in _FOOTERS]
        pairs = _pair([item for item in items if _filled(item.element)] + footers, offered)
        empty = [item for item in items if not _filled(item.element)
                 and id(item) not in pairs]
        gone = {id(item) for item in empty}
        filled = [item for item in items if id(item) not in gone]
        missing = [o for o in offered if id(o) not in {id(p) for p in pairs.values()}
                   and local_name(o.element) == "sp" and o.ph.get("type") not in _FOOTERS]
        # New ids before the empty placeholders go: theirs stay taken (A6).
        fresh = iter(slide._next_shape_ids(len(missing)))
        # Effective boxes, read while the slide is still on its old layout.
        boxes = {}
        for item in filled:
            shape = slide._wrap(item.element)
            boxes[id(item)] = (shape.bounds, _own_xfrm(item.element)
                               if _own_xfrm(item.element) is not None
                               else slide._inherited_xfrm(shape.placeholder))
        report.removed = [address(item.element) for item in empty]
        language = _language(document)
        block: dict[int, Element] = {}            # layout order -> element
        for item in filled:
            place = pairs.get(id(item))
            if place is None:
                continue
            target_box = _layout_box(document, layout_part, place)
            own = _box(_own_xfrm(item.element))
            if own is not None:
                report.boxes_dropped.append((address(item.element), own))
            _rewrite_ph(item.ph, place.ph, keep_size=item.type in _FOOTERS)
            if _kind(item) == "table":
                _table_into(item.element, target_box)
            elif item.tag == "graphicFrame":
                if target_box is not None:
                    _write_xfrm(item.element, target_box)
            else:
                xfrm = _own_xfrm(item.element)
                if xfrm is not None:
                    remove(xfrm)
            if item.tag == "pic":
                if place.type == "pic":
                    _picture_into_pic(slide, item.element, target_box)
                else:
                    _picture_into_content(slide, item.element, target_box)
            block[place.order] = item.element
            report.matched.append((address(item.element),
                                   f"{layout.id}.{find_id(place.element)}"))
        orphans = []
        for item in filled:
            if id(item) in pairs:
                continue
            bounds, template = boxes[id(item)]
            item.ph.set("idx", ORPHAN_IDX)
            if bounds is not None:
                x, y, w, h = bounds
                x = 0 if x + w / 2 <= slide_w / 2 else slide_w - w
                _write_xfrm(item.element, (x, y, w, h),
                            template if _own_xfrm(item.element) is None else None)
            orphans.append(item.element)
            report.orphans.append((address(item.element),
                                   f"no {_WHAT[_kind(item)]} placeholder left on {layout.id} "
                                   f"{layout.name!r}"))
        new = []
        for place in missing:
            element = empty_placeholder(place.element, place.ph, next(fresh), language)
            block[place.order] = element
            new.append(element)
        first = items[0].order if items else None
        shapes = [child for child in tree
                  if local_name(child) not in ("nvGrpSpPr", "grpSpPr", "extLst")]
        placeholder_ids = {id(item.element) for item in items}
        others = [child for child in shapes if id(child) not in placeholder_ids]
        before = 0 if first is None else sum(
            1 for child in list(tree)[:first] if id(child) in {id(o) for o in others})
        ordered = (others[:before] + [block[key] for key in sorted(block)] + orphans
                   + others[before:])
        for item in empty:
            remove(item.element)
        anchor = tree.find(qn("p:grpSpPr"))
        if anchor is None:
            anchor = tree.find(qn("p:nvGrpSpPr"))
        for element in ordered:
            if element.getparent() is not None:
                element.getparent().remove(element)
        for element in ordered:
            if anchor is not None:
                anchor.addnext(element)
            else:
                tree.insert(0, element)
            anchor = element
        _retarget(package, slide.part_path, layout_part)
        slide._invalidate()
        document._reset_caches()
    report.added = [slide._wrap(element).id for element in new]
    return report


def find_id(element: Element) -> str:
    node = next(element.iter(qn("p:cNvPr")), None)
    return node.get("id") if node is not None else "?"


__all__ = ["LayoutChange", "ORPHAN_IDX", "set_layout"]
