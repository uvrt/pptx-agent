"""A finished slide moved onto another layout, as PowerPoint's Layout menu does it.

What PowerPoint writes was measured in PowerPoint 16 (Mac) on 53 made-up slides
(agent-work/feedback-round7/oracle/results.md, A1-A9 and B2), and :func:`set_layout`
writes the same:

1. **Empty placeholders go**; the new layout's placeholders the slide then lacks are added
   empty -- the lowest free shape ids (counted before the empty ones went), type and idx (no
   ``sz``), an empty ``spPr``, a text body with one empty paragraph (none for a picture or
   table placeholder).  PowerPoint names each "<kind> N" in its interface language; here it
   takes the layout placeholder's name, as a new slide's placeholders do
   (:func:`~pptx_agent.edit.slides.slide_from_layout`).  Date, footer and slide-number
   placeholders are not added (Header & Footer adds them).
2. **Filled placeholders match by kind**, never by name or "same idx": a title takes a title
   or centred title; text content (a content, body or subtitle placeholder) takes a content,
   body or subtitle placeholder -- the same type first, then a compatible one; picture and
   table placeholders take no text.  A picture in a picture placeholder takes a picture
   placeholder (B2), else a content placeholder (A2).
3. **Within a kind** the slide's and the layout's placeholders are each sorted by idx and
   paired in that order; names, stacking order and ``sz`` play no part, and an earlier
   orphan (idx 4294967295) is matched again like any other.
4. **A matched placeholder** takes the layout's type and idx (no ``type`` for content, no
   ``sz``) and loses its own box (``a:xfrm``), also when the layout is the one it is on: it
   moves to the layout's place.  Its text, run formatting, fill, outline, ``bodyPr``,
   paragraph formatting, name and id stay.  A picture moving to a picture placeholder is
   cropped to fill the new box (``a:srcRect``, empty when the shapes agree: B2 and L2b S12);
   one already cropped keeps its crop (unmeasured) and is listed.  A picture moving into a
   content placeholder is drawn at its own size centred in the box (A2: a 4x4 px image
   without a stated resolution became 2x2 pt on PowerPoint for Mac, so 144 dpi there), made
   smaller to fit when larger (unmeasured).
5. **A filled placeholder with no match** stays, an orphan: idx ``4294967295``, its type and
   ``sz`` kept, its box written on it -- its own or the old layout's -- with ``x`` moved to
   the slide's left edge when the box's centre is in the left half (at the middle too), else
   to its right edge (A5).
6. **Order**: the placeholders gather into one block where the first placeholder was --
   matched and new ones in the layout's order, orphans last; other shapes keep theirs.
7. Other shapes are not touched; of the relationships only the layout's target changes.

Unmeasured, decided here: a placeholder inside a group is left as it is; a table, chart or
diagram frame (``p:graphicFrame``) matches its own kind first, then content, and as its
frame must have a box it takes the layout placeholder's; a date, footer or slide-number
placeholder matches its own kind; ``orient`` follows the layout, ``hasCustomPrompt`` goes.
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
    box was dropped for the layout's, as PowerPoint drops it.  ``crops_kept``: pictures moved
    to a picture placeholder whose own crop was kept."""

    slide: str
    layout: str
    matched: list[tuple[str, str]] = field(default_factory=list)
    added: list[str] = field(default_factory=list)
    removed: list[str] = field(default_factory=list)
    orphans: list[tuple[str, str]] = field(default_factory=list)
    boxes_dropped: list[tuple[str, tuple[int, int, int, int]]] = field(default_factory=list)
    crops_kept: list[str] = field(default_factory=list)


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


def _kind(item: _Item) -> str:
    """The class a slide placeholder's content matches within (A2)."""
    if item.tag == "pic":
        return "picture"
    if item.tag == "graphicFrame":
        return "frame"
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
    if kind == "frame":
        return offered.type in (item.type, None)
    return offered.type == kind


_WHAT = {"title": "title", "text": "content or text", "picture": "picture or content",
         "frame": "content", "dt": "date", "ftr": "footer", "sldNum": "slide-number",
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
                key = (kind, item.type) if exact or kind == "frame" else (kind, None)
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


def _rewrite_ph(ph: Element, source: Element) -> None:
    """``ph`` takes the layout placeholder's type, orient and idx (no ``type="obj"``, no
    ``sz``; A4)."""
    for attribute in ("type", "orient", "sz", "idx", "hasCustomPrompt"):
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


def _picture_into_pic(slide: "Slide", element: Element, box) -> bool:
    """A picture moved to a picture placeholder: cropped to fill ``box`` (an empty
    ``a:srcRect`` when the shapes agree), as PowerPoint writes it (B2; L2b S12).  A crop of
    its own is kept; returns ``True`` then."""
    from ..oxml.package import image_size
    from ..outline.draft import _crop

    blip_fill = element.find(qn("p:blipFill"))
    if blip_fill is None:
        return False
    rect = blip_fill.find(qn("a:srcRect"))
    if rect is not None and any(rect.get(side) not in (None, "0")
                                for side in ("l", "t", "r", "b")):
        return True
    data = _image(slide, element)
    if box is None or data is None:
        return False
    crop = _crop(box, image_size(data))
    new = make("a:srcRect", **crop)
    if rect is not None:
        rect.addprevious(new)
        remove(rect)
    else:
        blip = blip_fill.find(qn("a:blip"))
        (blip.addnext(new) if blip is not None else blip_fill.insert(0, new))
    return False


def _picture_into_content(slide: "Slide", element: Element, box) -> None:
    """A picture moved into a content placeholder: its own size, centred in ``box``, made
    smaller to fit when larger (A2)."""
    from ..oxml.package import image_dpi, image_size

    data = _image(slide, element)
    pixels = image_size(data) if data else None
    if box is None or not pixels or not all(pixels):
        return
    dpi = image_dpi(data) or (_DEFAULT_DPI, _DEFAULT_DPI)
    width = pixels[0] / (dpi[0] or _DEFAULT_DPI) * 914400
    height = pixels[1] / (dpi[1] or _DEFAULT_DPI) * 914400
    scale = min(1.0, box[2] / width if width else 1.0, box[3] / height if height else 1.0)
    width, height = round(width * scale), round(height * scale)
    left = box[0] + round((box[2] - width) / 2)
    top = box[1] + round((box[3] - height) / 2)
    _write_xfrm(element, (left, top, width, height))


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
        empty = [item for item in items if not _filled(item.element)]
        filled = [item for item in items if _filled(item.element)]
        pairs = _pair(filled, offered)
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
            _rewrite_ph(item.ph, place.ph)
            if item.tag == "graphicFrame":
                if target_box is not None:
                    _write_xfrm(item.element, target_box)
            else:
                xfrm = _own_xfrm(item.element)
                if xfrm is not None:
                    remove(xfrm)
            if item.tag == "pic":
                if place.type == "pic":
                    if _picture_into_pic(slide, item.element, target_box):
                        report.crops_kept.append(address(item.element))
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
