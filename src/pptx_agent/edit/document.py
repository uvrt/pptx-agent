"""The semantic editing API: ``Document`` -> ``Slide`` -> ``Shape``.

This is the layer an agent drives.  It never hands out XML; every property reads and writes the
underlying OOXML on demand, so anything this library does not model stays exactly where it was.

Lengths are EMU (914,400 per inch, 12,700 per point) because that is what the file format uses
and round-tripping through pixels would introduce drift.  Rotation is degrees, converted from
the 1/60,000-degree units OOXML stores.
"""

from __future__ import annotations

import copy
import dataclasses
import os
from contextlib import contextmanager
from typing import TYPE_CHECKING, BinaryIO, Iterator, NamedTuple

if TYPE_CHECKING:  # pragma: no cover
    from datetime import datetime

    from ..outline.read import TextBlock
    from .design import DesignFacts
    from .diagram import DiagramNode

from lxml import etree
from ooxml_edit.history import History

from ..oxml.package import (
    CT_SLIDE,
    REL_IMAGE,
    REL_SLIDE_LAYOUT,
    REL_SLIDE_MASTER,
    SHARED_ON_DUPLICATE,
    OoxmlPackage,
    image_dpi,
    image_size,
    normalize_part_path,
)
from ..oxml.xml import (
    FILL_TAGS,
    NAMESPACES,
    PRESET_GEOMETRIES,
    SHAPE_TREE_MEMBERS,
    Element,
    append_in_order,
    find,
    get_int,
    local_name,
    make,
    prefixed_name,
    qn,
    remove,
    replace_choice,
    set_int,
    subelement,
)
from .color import Color
from .fill import (
    Fill,
    LineFormat,
    fill_element,
    gradient_fill,
    image_fill,
    no_fill,
    read_fill,
    solid_fill,
    write_fill,
)
from . import slides as _slides
from .slides import Layout, Placeholder
from .table import Table, TableCell, _table_element
from .text import Paragraph, Run, TextFrame
from .ids import ShapeId, SlideShapeIndex, cnv_pr, read_stamp, write_stamp
from .comments import CommentOps, drop_comments
from . import creating as _creating
from .creating import Adjustments, ShapeFactory

if TYPE_CHECKING:  # pragma: no cover
    from .chart import Chart
    from .diagram import Diagram
    from .fit import TextFit
    from .theme import Theme

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700
ROTATION_UNIT = 60000

#: Shape tag -> the element holding its ``a:xfrm``.  A group's transform lives on its group
#: properties and a graphic frame's on a presentation-namespace ``p:xfrm`` of its own, so the
#: obvious ``p:spPr`` lookup silently fails for two of the five shape kinds.
#: Placeholder type -> the master placeholder a layout placeholder of that type inherits
#: from.  Anything else (body, subtitle, content, picture, chart, table...) is the body.
_MASTER_KIND: dict[str | None, str] = {
    "title": "title", "ctrTitle": "title", "dt": "dt", "ftr": "ftr", "sldNum": "sldNum",
    "hdr": "hdr",
}

_XFRM_PARENT: dict[str, str] = {
    "sp": "p:spPr",
    "pic": "p:spPr",
    "cxnSp": "p:spPr",
    "grpSp": "p:grpSpPr",
    "graphicFrame": "p:xfrm",
}


class Shape(ShapeFactory):
    """One shape on a slide, addressed by a stable id (see :mod:`pptx_agent.edit.ids`).

    A group can have shapes added to it like a slide (``group.add_shape(...)``), in its child
    coordinate space; see :class:`~pptx_agent.edit.creating.ShapeFactory`.

    For example::

        shape = deck.shape("256.29")
        shape.set_text("12.1%")
        shape.move_by(dx=914400)
    """

    def __init__(self, slide: "Slide", element: Element, shape_id: ShapeId) -> None:
        self._slide = slide
        self._element = element
        self._id = shape_id

    # -- identity --------------------------------------------------------------------------

    @property
    def id(self) -> str:
        """The stable id, ``<sldId>.<cNvPr id>`` (``#<n>`` added where the deck repeats one).

        For example::

            shape.id                                     # '256.29'
        """
        return str(self._id)

    @property
    def kind(self) -> str:
        """``"shape"``, ``"picture"``, ``"connector"``, ``"group"`` or ``"graphic_frame"``.

        What each has -- an attribute a kind does not have reads ``None`` (or ``[]``,
        ``""``) rather than raising:

        * ``"shape"`` (an autoshape or text box): ``text_frame``, ``text``, ``fill``,
          ``line``, ``preset``, ``adjustments``, ``connection_sites``, ``text_fit()``;
        * ``"picture"``: ``image_part``, ``image_size``, ``line``, ``replace_image()``;
        * ``"connector"``: ``line``, ``route``, ``begin_connection``, ``end_connection``;
        * ``"group"``: ``children``, ``child_offset``, ``child_extent``;
        * ``"graphic_frame"``: one of ``table``, ``chart`` or ``diagram``.

        All have ``id``, ``name``, ``left``/``top``/``width``/``height``, ``bounds``,
        ``drawn_bounds``, ``rotation`` and the flips.

        For example::

            shape.kind                                   # 'shape', 'picture', 'group'...
        """
        return {
            "sp": "shape",
            "pic": "picture",
            "cxnSp": "connector",
            "grpSp": "group",
            "graphicFrame": "graphic_frame",
        }.get(local_name(self._element), "unknown")

    @property
    def name(self) -> str | None:
        """``cNvPr@name``, the name PowerPoint's selection pane shows.

        For example::

            shape.name = "KPI margin"
        """
        properties = cnv_pr(self._element)
        return None if properties is None else properties.get("name")

    @name.setter
    def name(self, value: str) -> None:
        properties = cnv_pr(self._element)
        if properties is None:
            raise ValueError(f"{self.id}: shape has no cNvPr to name")
        self._before_change()
        properties.set("name", value)
        self._slide._touch()

    @property
    def placeholder(self) -> "Placeholder | None":
        """A :class:`~pptx_agent.Placeholder` -- the named tuple ``(type, idx)`` -- when this
        shape is a placeholder, else ``None``.  ``type`` ``None`` is a content placeholder
        (``"obj"``); ``idx`` is what matches it to its layout's placeholder (see
        :attr:`Layout.placeholders <pptx_agent.Layout.placeholders>`).

        For example::

            shape.placeholder                  # Placeholder(type='title', idx=None) or None
            shape.placeholder.type             # 'title'; ('body', 1) == Placeholder('body', 1)
        """
        node = find(self._element, f"{_nv_container(self._element)}/p:nvPr/p:ph")
        if node is None:
            return None
        return Placeholder(node.get("type"), get_int(node, "idx"))

    # -- geometry --------------------------------------------------------------------------

    @property
    def left(self) -> int | None:
        """Distance from the slide's (or group's) left edge, EMU; ``None`` when unknown.

        For example::

            shape.left = 914400                          # one inch
        """
        return self._offset("x")

    @left.setter
    def left(self, value: int) -> None:
        self._set_offset(x=int(value))

    @property
    def top(self) -> int | None:
        """Distance from the top edge, EMU.

        For example::

            shape.top += 457200
        """
        return self._offset("y")

    @top.setter
    def top(self, value: int) -> None:
        self._set_offset(y=int(value))

    @property
    def width(self) -> int | None:
        """Width, EMU.  Setting a table's width rescales its grid columns in proportion, as
        dragging the frame in PowerPoint does, so the columns fill the new width exactly.

        For example::

            shape.width = 2 * 914400
        """
        return self._extent("cx")

    @width.setter
    def width(self, value: int) -> None:
        with self._slide.document.batch():
            self._set_extent(cx=int(value))
            self._rescale_table("w", int(value))

    @property
    def height(self) -> int | None:
        """Height, EMU.  Setting a table's height rescales its rows in proportion (a row
        still grows to fit its text when PowerPoint draws it).

        For example::

            shape.height = 914400
        """
        return self._extent("cy")

    @height.setter
    def height(self, value: int) -> None:
        with self._slide.document.batch():
            self._set_extent(cy=int(value))
            self._rescale_table("h", int(value))

    def _rescale_table(self, attribute: str, total: int) -> None:
        """Spread ``total`` over a table's grid columns (``"w"``) or rows (``"h"``) in
        proportion to what they are now; the rounding remainder goes to the last."""
        if local_name(self._element) != "graphicFrame" or not self.has_table:
            return
        table = _table_element(self._element)
        nodes = table.findall(f"{qn('a:tblGrid')}/{qn('a:gridCol')}") if attribute == "w" \
            else table.findall(qn("a:tr"))
        current = [int(node.get(attribute) or 0) for node in nodes]
        if not nodes or sum(current) <= 0 or sum(current) == total:
            return
        scaled = [round(value * total / sum(current)) for value in current]
        scaled[-1] += total - sum(scaled)
        for node, value in zip(nodes, scaled):
            node.set(attribute, str(max(0, value)))
        self._slide._touch()

    @property
    def rotation(self) -> float:
        """Rotation in degrees, clockwise, ``0`` when none.

        For example::

            shape.rotation = 15                          # degrees clockwise
        """
        xfrm = self._xfrm(create=False)
        raw = get_int(xfrm, "rot", 0) or 0
        return raw / ROTATION_UNIT

    @rotation.setter
    def rotation(self, degrees: float) -> None:
        self._before_change()
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        set_int(xfrm, "rot", round(degrees * ROTATION_UNIT) % (360 * ROTATION_UNIT))
        self._slide._touch()
        self._slide._reroute_for(self)

    @property
    def flip_h(self) -> bool:
        """Mirrored left to right (``a:xfrm@flipH``).

        For example::

            shape.flip_h = True
        """
        return self._flipped("flipH")

    @flip_h.setter
    def flip_h(self, value: bool) -> None:
        self._set_flip("flipH", value)

    @property
    def flip_v(self) -> bool:
        """Mirrored top to bottom (``a:xfrm@flipV``).

        For example::

            shape.flip_v = True
        """
        return self._flipped("flipV")

    @flip_v.setter
    def flip_v(self, value: bool) -> None:
        self._set_flip("flipV", value)

    @property
    def preset(self) -> str | None:
        """The preset geometry (``"rect"``, ``"ellipse"``...), ``"custom"`` for a freeform,
        ``None`` when the shape has no geometry of its own (groups, graphic frames).

        For example::

            shape.preset = "ellipse"
        """
        properties = self._element.find(qn("p:spPr"))
        if properties is None:
            return None
        node = properties.find(qn("a:prstGeom"))
        if node is not None:
            return node.get("prst")
        return "custom" if properties.find(qn("a:custGeom")) is not None else None

    @preset.setter
    def preset(self, name: str) -> None:
        """Give the shape another preset geometry.  Its adjustments go with the old one."""
        if local_name(self._element) not in {"sp", "pic", "cxnSp"}:
            raise ValueError(f"{self.id}: a {self.kind} has no geometry of its own")
        if name not in PRESET_GEOMETRIES:
            raise ValueError(f"{name!r} is not a preset geometry (ST_ShapeType)")
        if self.preset == name:
            return
        self._before_change()
        geometry = make("a:prstGeom", prst=name)
        geometry.append(make("a:avLst"))
        replace_choice(subelement(self._element, "p:spPr"), ("a:custGeom", "a:prstGeom"),
                       geometry)
        self._slide._touch()
        self._slide._reroute_for(self)

    def _flipped(self, attribute: str) -> bool:
        xfrm = self._xfrm(create=False)
        return xfrm is not None and xfrm.get(attribute) in {"1", "true"}

    def _set_flip(self, attribute: str, value: bool) -> None:
        if bool(value) == self._flipped(attribute):
            return
        self._before_change()
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        # A transform that exists must say where the shape is, or PowerPoint reads (0, 0).
        self._ensure_offset()
        self._ensure_extent()
        if value:
            xfrm.set(attribute, "1")
        else:
            del xfrm.attrib[attribute]
        self._slide._touch()
        self._slide._reroute_for(self)

    @property
    def has_explicit_transform(self) -> bool:
        """``False`` when the position is inherited from the layout or master placeholder.

        Worth checking before moving a placeholder: the first write materialises an ``a:xfrm``
        on the slide, which permanently detaches the shape from the layout's positioning.

        For example::

            shape.has_explicit_transform               # False: it inherits its place
        """
        return self._xfrm(create=False) is not None

    def move_by(self, dx: int = 0, dy: int = 0, *, space: str = "parent") -> "Shape":
        """Move relative to the current position.

        For a shape inside a group, ``left``/``top`` and therefore the default
        ``space="parent"`` are in the group's *child* coordinate space (``chOff``/``chExt``),
        which a scaled group stretches.  ``space="slide"`` takes ``dx``/``dy`` as slide EMU
        instead, so "one inch right" is one inch on the slide however the group is scaled.

        For example::

            shape.move_by(dx=914400, dy=-457200)
        """
        if space not in {"parent", "slide"}:
            raise ValueError("space must be 'parent' or 'slide'")
        left, top = self.left, self.top
        if left is None or top is None:
            raise ValueError(f"{self.id}: cannot move a shape with no resolvable position")
        if space == "slide":
            scale_x, scale_y = self._parent_scale()
            dx, dy = round(dx / scale_x), round(dy / scale_y)
        with self._slide.document.batch():
            self.left = left + dx
            self.top = top + dy
        return self

    # -- groups ----------------------------------------------------------------------------

    @property
    def parent_group(self) -> "Shape | None":
        """The group directly containing this shape, or ``None`` on the slide itself.

        For example::

            shape.parent_group                           # the group, or None
        """
        parent = self._element.getparent()
        if parent is None or parent.tag != qn("p:grpSp"):
            return None
        return self._slide._wrap(parent)

    @property
    def children(self) -> list["Shape"]:
        """A group's direct children, in z-order (back to front).  Empty for other shapes.

        For example::

            [child.id for child in group.children]
        """
        if local_name(self._element) != "grpSp":
            return []
        return [self._slide._wrap(child) for child in self._element
                if local_name(child) in _XFRM_PARENT and isinstance(child.tag, str)]

    @property
    def child_offset(self) -> tuple[int, int] | None:
        """A group's ``chOff``: where its child coordinate space starts.

        For example::

            group.child_offset
        """
        node = self._xfrm_child("a:chOff")
        return None if node is None else (get_int(node, "x", 0) or 0, get_int(node, "y", 0) or 0)

    @property
    def child_extent(self) -> tuple[int, int] | None:
        """A group's ``chExt``: how large its child coordinate space is.

        For example::

            group.child_extent
        """
        node = self._xfrm_child("a:chExt")
        return None if node is None else (get_int(node, "cx", 0) or 0, get_int(node, "cy", 0) or 0)

    @property
    def slide_bounds(self) -> tuple[int, int, int, int] | None:
        """``(left, top, width, height)`` on the slide, through every enclosing group.

        Group rotation and flips are not applied -- this is the unrotated frame, the same
        thing ``left``/``top`` report for a shape on the slide itself.

        For example::

            left, top, width, height = shape.slide_bounds
        """
        left, top, width, height = self.left, self.top, self.width, self.height
        if None in (left, top, width, height):
            return None
        x, y, cx, cy = float(left), float(top), float(width), float(height)
        group = self.parent_group
        while group is not None:
            mapped = group._map_child_rect(x, y, cx, cy)
            if mapped is None:
                return None
            x, y, cx, cy = mapped
            group = group.parent_group
        return round(x), round(y), round(cx), round(cy)

    @property
    def bounds(self) -> tuple[int, int, int, int] | None:
        """``(left, top, width, height)``, EMU, in the parent's space: the frame, as
        ``left``/``top``/``width``/``height`` give it one by one -- the same tuple a
        :class:`~pptx_agent.LayoutShape` and a :class:`~pptx_agent.LayoutPlaceholder`
        call ``bounds``.  ``None`` when any of the four is unknown.

        For example::

            left, top, width, height = shape.bounds
        """
        values = (self.left, self.top, self.width, self.height)
        return None if None in values else values  # type: ignore[return-value]

    @property
    def drawn_bounds(self) -> tuple[int, int, int, int] | None:
        """``(left, top, width, height)`` on the slide of what is *drawn*: the frame turned
        by its rotation and flips and those of every enclosing group, and for a connector
        the box around its route (:attr:`route`) -- not its frame, which PowerPoint writes
        rotated and can reach far past the line.  A table's height is its rows' as drawn,
        each grown to fit its text (:attr:`Table.drawn_row_heights`; the stored heights
        without pptx2svg).  :meth:`Document.overflows` checks these.  See :mod:`.geometry`.

        For example::

            left, top, width, height = connector.drawn_bounds
        """
        from .geometry import drawn_bounds

        return drawn_bounds(self)

    @property
    def route(self) -> list[tuple[float, float]] | None:
        """A connector's (or a line's) drawn path on the slide, read-only: the points of
        its polyline in slide EMU, from its begin to its end, bends included -- a curved
        connector's curve sampled.  ``None`` for any other shape.

        For example::

            points = connector.route            # [(x0, y0), (x1, y0), (x1, y1), ...]
        """
        from .geometry import route

        return route(self)

    def fit_to_children(self) -> "Shape":
        """Re-fit a group's frame to its children without moving anything on the slide.

        PowerPoint does this whenever a child is moved or resized: the group's child space
        (``chOff``/``chExt``) becomes the children's bounding box and its frame
        (``off``/``ext``) moves and resizes by the same amount through the group's scale, so
        every child stays exactly where it was drawn.  Called automatically after a child's
        geometry changes; rotated or flipped groups are left alone, because re-fitting one
        moves its rotation centre and so would move the children.

        For example::

            group.fit_to_children()
        """
        if local_name(self._element) != "grpSp":
            raise ValueError(f"{self.id}: only a group has children to fit")
        self._before_change()
        self._refit()
        self._after_change()
        return self

    def _xfrm_child(self, tag: str) -> Element | None:
        xfrm = self._xfrm(create=False)
        return None if xfrm is None else xfrm.find(qn(tag))

    def _map_child_rect(self, x: float, y: float, cx: float, cy: float):
        """Map a rectangle from this group's child space into its parent's space."""
        off, ext = self._xfrm_child("a:off"), self._xfrm_child("a:ext")
        ch_off, ch_ext = self._xfrm_child("a:chOff"), self._xfrm_child("a:chExt")
        if off is None or ext is None:
            return None
        ox, oy = get_int(off, "x", 0) or 0, get_int(off, "y", 0) or 0
        ex, ey = get_int(ext, "cx", 0) or 0, get_int(ext, "cy", 0) or 0
        cox = get_int(ch_off, "x", ox) if ch_off is not None else ox
        coy = get_int(ch_off, "y", oy) if ch_off is not None else oy
        cex = get_int(ch_ext, "cx", ex) if ch_ext is not None else ex
        cey = get_int(ch_ext, "cy", ey) if ch_ext is not None else ey
        sx = ex / cex if cex else 1.0
        sy = ey / cey if cey else 1.0
        return ox + (x - (cox or 0)) * sx, oy + (y - (coy or 0)) * sy, cx * sx, cy * sy

    def _parent_scale(self) -> tuple[float, float]:
        """Slide EMU per unit of this shape's own coordinate space, through all groups."""
        scale_x = scale_y = 1.0
        group = self.parent_group
        while group is not None:
            ext, ch_ext = group._xfrm_child("a:ext"), group._xfrm_child("a:chExt")
            if ext is not None and ch_ext is not None:
                cex, cey = get_int(ch_ext, "cx", 0) or 0, get_int(ch_ext, "cy", 0) or 0
                if cex:
                    scale_x *= (get_int(ext, "cx", 0) or 0) / cex
                if cey:
                    scale_y *= (get_int(ext, "cy", 0) or 0) / cey
            group = group.parent_group
        return (scale_x or 1.0), (scale_y or 1.0)

    def _refit(self) -> None:
        xfrm = self._xfrm(create=False)
        if xfrm is None or (get_int(xfrm, "rot", 0) or 0) or xfrm.get("flipH") in {"1", "true"} \
                or xfrm.get("flipV") in {"1", "true"}:
            return
        boxes = []
        for child in self.children:
            left, top, width, height = child.left, child.top, child.width, child.height
            if None in (left, top, width, height):
                return
            boxes.append((left, top, left + width, top + height))
        if not boxes:
            return
        min_x = min(box[0] for box in boxes)
        min_y = min(box[1] for box in boxes)
        max_x = max(box[2] for box in boxes)
        max_y = max(box[3] for box in boxes)
        off, ext = subelement(xfrm, "a:off"), subelement(xfrm, "a:ext")
        ox, oy = get_int(off, "x", 0) or 0, get_int(off, "y", 0) or 0
        ex, ey = get_int(ext, "cx", 0) or 0, get_int(ext, "cy", 0) or 0
        ch_off, ch_ext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
        cox = get_int(ch_off, "x", ox) if ch_off is not None else ox
        coy = get_int(ch_off, "y", oy) if ch_off is not None else oy
        cex = get_int(ch_ext, "cx", ex) if ch_ext is not None else ex
        cey = get_int(ch_ext, "cy", ey) if ch_ext is not None else ey
        if (cox, coy, cex, cey) == (min_x, min_y, max_x - min_x, max_y - min_y):
            return
        sx = ex / cex if cex else 1.0
        sy = ey / cey if cey else 1.0
        set_int(off, "x", round(ox + (min_x - cox) * sx))
        set_int(off, "y", round(oy + (min_y - coy) * sy))
        set_int(ext, "cx", round((max_x - min_x) * sx))
        set_int(ext, "cy", round((max_y - min_y) * sy))
        ch_off = subelement(xfrm, "a:chOff")
        ch_ext = subelement(xfrm, "a:chExt")
        set_int(ch_off, "x", min_x)
        set_int(ch_off, "y", min_y)
        set_int(ch_ext, "cx", max_x - min_x)
        set_int(ch_ext, "cy", max_y - min_y)
        parent = self.parent_group
        if parent is not None:
            parent._refit()

    def _after_geometry_change(self) -> None:
        parent = self.parent_group
        if parent is not None:
            parent._refit()
        # Connectors attached to this shape (or, for a group, to anything in it) follow it.
        self._slide._reroute_for(self)

    # -- text ------------------------------------------------------------------------------

    @property
    def text(self) -> str:
        r"""All text in the shape, raw -- not escaped, no Markdown: paragraphs joined by
        ``"\n"``, and a line break inside one read as ``"\n"`` too (E0's reading)::

            deck.shape("256.29").text            # '11.9%'

        ``""`` for a shape without text (a picture, a table's frame).  To edit from, when
        the text may have line breaks, read ``shape.text_frame.text``: it gives a break as
        ``"\v"``, so ``shape.set_text(shape.text_frame.text)`` changes nothing.  This, not
        :meth:`Document.to_outline` (a Markdown view, escaped), is the text to edit from.
        """
        body = self._element.find(qn("p:txBody"))
        if body is None:
            return ""
        paragraphs = []
        for paragraph in body.findall(qn("a:p")):
            pieces = []
            for node in paragraph.iter():
                name = local_name(node)
                if name == "t":
                    pieces.append(node.text or "")
                elif name == "br":
                    pieces.append("\n")
            paragraphs.append("".join(pieces))
        return "\n".join(paragraphs)

    @text.setter
    def text(self, value: str) -> None:
        self.set_text(value)

    def set_text(self, value: "str | TextSpec") -> "Shape":
        r"""Replace the shape's text, keeping the formatting of every character that survives.

        ``value`` may also be a :class:`~pptx_agent.TextSpec` -- paragraphs, runs,
        bullets and frame properties together, one undo step; see
        :meth:`TextFrame.set_text`.

        The new text is diffed against the old, paragraph by paragraph and then character by
        character, so changing one figure in a sentence of mixed formatting keeps the rest of
        the sentence -- and the new figure -- formatted as before.  ``"\n"`` separates
        paragraphs and ``"\v"`` is a line break inside one.  See :mod:`pptx_agent.edit.text`.

        The text is written exactly as given: take it from :attr:`text`,
        :meth:`Document.outline_blocks` or :meth:`Document.find_text`, never from
        :meth:`Document.to_outline`, whose Markdown is escaped (a
        :class:`~pptx_agent.MarkdownEscapeWarning` says when it looks that way)::

            deck.shape("256.29").set_text("12.1%")      # still bold, same colour
        """
        frame = self.text_frame
        if frame is None:
            raise ValueError(f"{self.id}: a {self.kind} cannot hold text directly")
        frame.set_text(value)
        return self

    @property
    def text_frame(self) -> "TextFrame | None":
        """The shape's paragraphs and runs.  Only autoshapes and text boxes (``kind``
        ``"shape"``) carry text; ``None`` for a picture, connector, group or graphic frame
        (a table's text is in its cells' frames).

        For example::

            shape.text_frame.paragraph(0).run(0).bold = True
        """
        if local_name(self._element) != "sp":
            return None
        document, identifier = self._slide.document, self.id
        return TextFrame(lambda: document.shape(identifier), identifier)

    def text_fit(self) -> "TextFit":
        """Whether the shape's text fits it, as PowerPoint will draw it: the height the text
        needs and the height there is (EMU), whether and by how much it overflows, each
        paragraph's effective font sizes and line count, and the autofit mode and stored
        font scale (:class:`~pptx_agent.TextFit`).  Needs pptx2svg (``pptx-agent[render]``).

        The text is laid out with the measurement pptx2svg draws with, at the **stored**
        autofit scale: PowerPoint shrinks text only when it edits it, so ``normAutofit``
        with no stored ``fontScale`` is drawn full size, overflowing or not -- whatever a
        render that re-derives a scale may show.  See :mod:`pptx_agent.edit.fit`.

        For example::

            fit = deck.shape("258.3").text_fit()
            fit.overflows, fit.needed, fit.available, fit.sizes[0]
        """
        _require_renderer()
        from .fit import measure

        return measure(self)

    def fit_height(self) -> int:
        """The height, EMU, this shape needs for its text at its present width -- what
        PowerPoint grows a ``"shape"``-autofit box to -- from the same layout as
        :meth:`text_fit`: its frame's height plus what the text lacks (or less what it
        does not use).  A geometry whose text area is a share of the shape (an ellipse, a
        chevron) is grown in proportion.  The shape is not changed.  Needs pptx2svg.

        For example::

            box.height = box.fit_height()        # exactly as tall as its text
        """
        _require_renderer()
        from .fit import fit_height

        return fit_height(self)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The text frame's paragraphs.

        For example::

            [paragraph.text for paragraph in shape.paragraphs]
        """
        frame = self.text_frame
        return [] if frame is None else frame.paragraphs

    # -- fill and outline ------------------------------------------------------------------

    @property
    def fill(self) -> Fill | None:
        """The explicit fill, or ``None`` when it comes from the style, layout or theme --
        as a new :meth:`add_shape` shape's does: it is drawn in ``accent1`` through its
        shape style, and :attr:`effective_fill` says so.

        For example::

            shape.fill = "accent1 lumMod=75%"            # None inherits; "none" removes
        """
        fill = read_fill(self._fill_container(create=False))
        if fill is not None and fill.kind == "image":
            part = self._slide.document.package.related_part(self._slide.part_path,
                                                             fill.image_rel_id)
            fill = dataclasses.replace(fill, image_part=part)
        return fill

    @property
    def effective_fill(self) -> Fill | None:
        """The fill as drawn: the explicit :attr:`fill`, else the one the shape's style
        names (``p:style/a:fillRef``: the theme's fill style in the reference's colour) --
        ``Fill("solid", color=Color("accent1"))`` for a new :meth:`add_shape` shape,
        ``Fill("none")`` for a style that names none.  ``None`` when neither says: a
        placeholder's fill comes from its layout.

        For example::

            box = slide.add_shape("rect", 0, 0, 914400, 914400)
            box.fill, box.effective_fill.color           # None, Color('accent1')
        """
        explicit = self.fill
        if explicit is not None:
            return explicit
        reference = find(self._element, "p:style/a:fillRef")
        if reference is None:
            return None
        index = int(reference.get("idx") or 0)
        if index == 0:
            return Fill("none")
        color = Color.from_element(reference)
        from .theme import theme_fill_style

        style = theme_fill_style(self._slide, index)
        if style is None or local_name(style) == "solidFill":
            return Fill("solid", color=color)
        styled = copy.deepcopy(style)
        replacement = next((child for child in reference if isinstance(child.tag, str)), None)
        if replacement is not None:
            for placeholder in list(styled.iter(qn("a:schemeClr"))):
                if placeholder.get("val") == "phClr":
                    substitute = copy.deepcopy(replacement)
                    for modifier in placeholder:
                        substitute.append(copy.deepcopy(modifier))
                    placeholder.getparent().replace(placeholder, substitute)
        holder = make("a:spPr")
        holder.append(styled)
        return read_fill(holder) or Fill("solid", color=color)

    @fill.setter
    def fill(self, value: "str | Color | None") -> None:
        """A colour (``"accent1"``, ``"#4472C4"``, a :class:`Color`) for a solid fill,
        ``"none"`` for no fill, ``None`` to inherit again."""
        self._write_fill(fill_element(value))

    def set_solid_fill(self, color: "str | Color") -> "Shape":
        """Fill with one colour: a theme colour, ``#RRGGBB``, modifiers allowed.

        For example::

            shape.set_solid_fill("#1F4E79")
        """
        self._write_fill(solid_fill(color))
        return self

    def set_gradient_fill(self, stops, *, angle: float | None = 90.0,
                          path: str | None = None) -> "Shape":
        """``stops``: ``[(0, "accent1"), (1, "accent1 lumMod=50%")]`` or bare colours.

        For example::

            shape.set_gradient_fill([(0, "accent1"), (1, "accent2")], angle=45)
        """
        self._write_fill(gradient_fill(stops, angle=angle, path=path))
        return self

    def set_image_fill(self, image: "bytes | str | os.PathLike[str]", *,
                       stretch: bool = True) -> "Shape":
        """Fill the shape with a picture, stretched to its frame.

        ``image`` is the image's bytes or a path to it (PNG, JPEG, GIF, BMP or TIFF).  The
        picture is stored once under ``ppt/media/`` -- an identical one already in the deck
        is reused -- related from this slide, and declared in ``[Content_Types].xml``; undo
        takes all of that back out.

        For example::

            shape.set_image_fill("texture.png")
        """
        data, suffix = _read_image(image)
        package = self._slide.document.package
        with self._slide.document.batch():
            self._before_change()
            media = package.add_image(data, suffix)
            rel_id = package.add_relationship(self._slide.part_path, REL_IMAGE, media)
            container = self._fill_container(create=True)
            assert container is not None
            write_fill(container, image_fill(rel_id, stretch=stretch))
            self._after_change()
        return self

    def set_no_fill(self) -> "Shape":
        """No fill at all (transparent), whatever the style says.

        For example::

            shape.set_no_fill()
        """
        self._write_fill(no_fill())
        return self

    def clear_fill(self) -> "Shape":
        """Remove the explicit fill, so the style, layout or theme decides again.

        For example::

            shape.clear_fill()                           # inherit again
        """
        self._write_fill(None)
        return self

    @property
    def line(self) -> "LineFormat | None":
        """The outline (``a:ln``); ``None`` for a group or a graphic frame, which have none.

        For example::

            shape.line.width = 12700; shape.line.color = "accent2"
        """
        if local_name(self._element) not in {"sp", "pic", "cxnSp"}:
            return None
        document, identifier = self._slide.document, self.id

        def locate(create: bool) -> Element | None:
            shape = document.shape(identifier)
            container = shape._fill_container(create=create)
            if container is None:
                return None
            return subelement(container, "a:ln") if create else container.find(qn("a:ln"))

        return LineFormat(
            locate,
            lambda: document.shape(identifier)._before_change(),
            lambda: document.shape(identifier)._after_change(),
        )

    # -- adjustments and connections -------------------------------------------------------

    @property
    def adjustments(self) -> Adjustments:
        """The preset geometry's adjust values (``a:avLst``) by name: ``shape.adjustments
        ["adj"] = 30000``.  See :class:`~pptx_agent.edit.creating.Adjustments`.

        For example::

            shape.adjustments["adj"] = 30000
        """
        if local_name(self._element) not in {"sp", "pic", "cxnSp"}:
            raise ValueError(f"{self.id}: a {self.kind} has no geometry to adjust")
        return Adjustments(self)

    @property
    def connection_sites(self) -> list[tuple[int, int]]:
        """Where connectors can attach, as ``(x, y)`` on the slide, in ``stCxn@idx`` order.

        From the preset's (or a custom geometry's) ``cxnLst``, for this shape's size,
        adjustments, flips and rotation.  Empty for groups, graphic frames and presets
        without sites.

        The order is the preset's own, as PowerPoint numbers it (measured presets,
        unrotated)::

            rect, roundRect, diamond, flowChartProcess,
            flowChartDecision, flowChartTerminator,
            rightArrow, trapezoid ............ 0 top, 1 left, 2 bottom, 3 right
            ellipse ........................ 0 top, then counter-clockwise every 45 degrees:
                                             1 top-left, 2 left, 3 bottom-left, 4 bottom,
                                             5 bottom-right, 6 right, 7 top-right
            triangle ....................... 0 apex, 1 left side, 2 bottom-left corner,
                                             3 bottom, 4 bottom-right corner, 5 right side
            hexagon ........................ 0 right, 1 bottom-right, 2 bottom-left,
                                             3 left, 4 top-left, 5 top-right
            pentagon, star5 ................ 0 top, then counter-clockwise
            can ............................ 0 the lid's centre, 1 top, 2 left,
                                             3 bottom, 4 right

        Rather than counting, name the side: :meth:`connection_site`.

        For example::

            shape.connection_sites                      # [(x, y), ...], slide EMU
        """
        return [(round(end.point[0]), round(end.point[1])) for end in _creating.shape_sites(self)]

    def connection_site(self, which: "str | int") -> "tuple[Shape, int]":
        """A connector end on this shape, ``(shape, index)``, as :meth:`Slide.add_connector`
        takes it: ``which`` is ``"top"``, ``"right"``, ``"bottom"`` or ``"left"`` -- the
        site furthest that way *as drawn* (rotation and flips included), the one nearest
        the middle of that side when several tie -- or an index into
        :attr:`connection_sites`.

        For example::

            slide.add_connector("elbow", box.connection_site("right"),
                                goal.connection_site("left"), line={"end": "triangle"})
        """
        sites = self.connection_sites
        if not sites:
            raise ValueError(f"{self.id}: a {self.kind} has no connection sites")
        if isinstance(which, int) and not isinstance(which, bool):
            if not 0 <= which < len(sites):
                raise IndexError(f"{self.id} has connection sites 0..{len(sites) - 1}")
            return (self, which)
        axes = {"top": (1, -1), "bottom": (1, 1), "left": (0, -1), "right": (0, 1)}
        if which not in axes:
            raise ValueError("a connection site is 'top', 'right', 'bottom', 'left' or an index")
        axis, sign = axes[which]
        middle = sum(point[1 - axis] for point in sites) / len(sites)
        best = min(range(len(sites)),
                   key=lambda k: (-sign * sites[k][axis], abs(sites[k][1 - axis] - middle), k))
        return (self, best)

    @property
    def is_connector(self) -> bool:
        """``True`` for a connector (``p:cxnSp``).

        For example::

            shape.is_connector
        """
        return local_name(self._element) == "cxnSp"

    @property
    def begin_connection(self) -> "tuple[Shape, int] | None":
        """``(shape, site)`` the connector starts on, or ``None`` when its start is free.

        For example::

            arrow.begin_connection                       # (box, 3)
        """
        return self._connection("begin")

    @property
    def end_connection(self) -> "tuple[Shape, int] | None":
        """``(shape, site)`` the connector ends on, or ``None`` when its end is free.

        For example::

            arrow.end_connection                         # (goal, 1)
        """
        return self._connection("end")

    def connect(self, begin=None, end=None) -> "Shape":
        """Attach (or move) a connector's ends: each a ``(shape, site)`` pair or a point
        ``(x, y)`` in the connector's parent space, ``None`` to leave that end as it is.
        The connector is re-routed, as one undo step.

        For example::

            arrow.connect(begin=(box, 3), end=(goal, 1))
        """
        self._require_connector()
        with self._slide.document.batch():
            self._before_change()
            container = self._element.getparent()
            references = {}
            for which, spec in (("begin", begin), ("end", end)):
                if spec is None:
                    references[which] = _creating.connection(self._element, which)
                    continue
                end_point, reference = _creating._resolve_end(self._slide, container, spec, which)
                references[which] = reference
                if reference is None:  # a free point: put the end there first
                    self._move_end(which, end_point.point)
            _creating._write_connections(self._element, references["begin"], references["end"])
            self.reroute()
        return self

    def disconnect(self, which: str = "both") -> "Shape":
        """Detach ``"begin"``, ``"end"`` or ``"both"`` ends; the geometry stays put.

        For example::

            arrow.disconnect("end")
        """
        self._require_connector()
        if which not in {"begin", "end", "both"}:
            raise ValueError("which must be 'begin', 'end' or 'both'")
        references = {w: _creating.connection(self._element, w) for w in ("begin", "end")}
        for name in ("begin", "end"):
            if which in (name, "both"):
                references[name] = None
        if references == {w: _creating.connection(self._element, w) for w in ("begin", "end")}:
            return self
        self._before_change()
        _creating._write_connections(self._element, references["begin"], references["end"])
        self._after_change()
        return self

    def reroute(self) -> "Shape":
        """Re-compute the connector's geometry from the shapes it is attached to, keeping
        their sites -- what happens automatically when one of them moves.

        For example::

            arrow.reroute()
        """
        self._require_connector()
        if _creating.connector_kind(self._element) is None:
            raise ValueError(f"{self.id}: only straight, elbow and curved connectors are routed")
        with self._slide.document.batch():
            self._before_change()
            _creating.route_connector(self._slide, self._element)
            self._after_geometry_change()
            self._slide._touch()
        return self

    def _connection(self, which: str) -> "tuple[Shape, int] | None":
        if local_name(self._element) != "cxnSp":
            return None
        reference = _creating.connection(self._element, which)
        if reference is None:
            return None
        target = self._slide._element_by_raw_id(reference[0])
        return None if target is None else (self._slide._wrap(target), reference[1])

    def _move_end(self, which: str, point: tuple[float, float]) -> None:
        """Put one end of a free connector at a slide point, the other end staying."""
        container = self._element.getparent()
        ends = _creating.current_ends(self._element)
        if ends is None:
            return
        begin, end = (_creating.container_to_slide(container, *p) for p in ends)
        if which == "begin":
            begin = point
        else:
            end = point
        kind = _creating.connector_kind(self._element) or "straight"
        frame = _creating._frame_in(self._slide, container, kind, _creating.End(begin),
                                    _creating.End(end))
        _creating.write_frame(self._element, frame)

    def _require_connector(self) -> None:
        if local_name(self._element) != "cxnSp":
            raise ValueError(f"{self.id}: a {self.kind} is not a connector")

    def _factory(self):
        if local_name(self._element) != "grpSp":
            raise ValueError(f"{self.id}: shapes can be added to a slide or a group, not to "
                             f"a {self.kind}")
        return self._slide, self._element, self

    # -- tables ----------------------------------------------------------------------------

    @property
    def has_table(self) -> bool:
        """``True`` for a graphic frame holding a table.

        For example::

            if shape.has_table: print(shape.table.rows)
        """
        return local_name(self._element) == "graphicFrame" and _table_element(self._element) is not None

    @property
    def table(self) -> "Table | None":
        """The table in a graphic frame; ``None`` for any other shape (see
        :attr:`kind`).

        For example::

            deck.shape("257.3#5").table.cell_by_label("売上高", "当期実績")
        """
        if not self.has_table:
            return None
        document, identifier = self._slide.document, self.id
        return Table(lambda: document.shape(identifier))

    # -- charts and diagrams ---------------------------------------------------------------

    @property
    def has_chart(self) -> bool:
        """``True`` for a graphic frame holding a chart.

        For example::

            if shape.has_chart: print(shape.chart.series_names)
        """
        return local_name(self._element) == "graphicFrame" and self._element.find(
            f"{qn('a:graphic')}/{qn('a:graphicData')}/{qn('c:chart')}") is not None

    @property
    def chart(self) -> "Chart | None":
        """The chart in a graphic frame: its series, categories, values, titles and legend,
        edited in the chart's cache and its embedded workbook together.  ``None`` for any
        other shape.

        For example::

            deck.shape("257.25").chart.series["売上高（億円）"].set_value("Q3", 4310)
        """
        if not self.has_chart:
            return None
        from .chart import Chart

        document, identifier = self._slide.document, self.id
        return Chart(lambda: document.shape(identifier))

    @property
    def has_diagram(self) -> bool:
        """``True`` for a graphic frame holding SmartArt.

        For example::

            if shape.has_diagram: print(shape.diagram.texts)
        """
        return local_name(self._element) == "graphicFrame" and self._element.find(
            f"{qn('a:graphic')}/{qn('a:graphicData')}/{qn('dgm:relIds')}") is not None

    @property
    def diagram(self) -> "Diagram | None":
        """The SmartArt in a graphic frame: its nodes' text, and adding or removing nodes.
        ``None`` for any other shape.

        For example::

            shape.diagram.set_text(1, "Ship it")
        """
        if not self.has_diagram:
            return None
        from .diagram import Diagram

        document, identifier = self._slide.document, self.id
        return Diagram(lambda: document.shape(identifier))

    # -- host protocol (text frames, line formats) -----------------------------------------

    @property
    def address(self) -> str:
        """The shape's id, as an address (see :meth:`Document.resolve`).

        For example::

            shape.address                                # '256.29'
        """
        return self.id

    def _text_body(self, create: bool) -> Element | None:
        body = self._element.find(qn("p:txBody"))
        if body is None and create:
            # An autoshape may legitimately have no txBody yet.  A new one needs a:bodyPr and
            # a:lstStyle before its paragraphs, in that order -- the sequence is schema-
            # enforced, and PowerPoint reports a repair rather than a parse error if it is wrong.
            body = subelement(self._element, "p:txBody")
            subelement(body, "a:bodyPr")
            subelement(body, "a:lstStyle")
        return body

    def _after_change(self) -> None:
        self._slide._touch()
        self._slide._settle()

    def _batch(self):
        return self._slide.document.batch()

    def _list_styles(self) -> list:
        """The list styles this shape's text inherits from, nearest first (:mod:`.inherit`)."""
        from . import inherit

        package = self._slide.document.package
        master = inherit.master_of(package, inherit.layout_of(package, self._slide.part_path))
        placeholder = self.placeholder
        own = self._element.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")
        layout_shape = self._slide._layout_placeholder(placeholder) if placeholder else None
        kind = placeholder.type if placeholder else None
        if layout_shape is not None:
            node = inherit.placeholder_of(layout_shape)
            kind = node.get("type") if node is not None else kind
        return inherit.list_styles(package, master, own, layout_shape, kind,
                                   placeholder is not None)

    def _body_chain(self) -> list:
        """``a:bodyPr`` of the shape, its layout placeholder and its master placeholder."""
        from . import inherit

        def body(shape):
            return None if shape is None else find(shape, "p:txBody/a:bodyPr")

        chain = [body(self._element)]
        placeholder = self.placeholder
        if placeholder is None:
            return chain
        package = self._slide.document.package
        layout_shape = self._slide._layout_placeholder(placeholder)
        chain.append(body(layout_shape))
        kind = placeholder.type
        if layout_shape is not None:
            node = inherit.placeholder_of(layout_shape)
            kind = node.get("type") if node is not None else kind
        master = inherit.master_of(package, inherit.layout_of(package, self._slide.part_path))
        chain.append(body(inherit.master_placeholder(package, master, inherit.master_kind(kind))))
        return chain

    def _theme_source(self):
        return self

    @property
    def _plain_shape(self) -> bool:
        return self.placeholder is None

    def _fill_container(self, *, create: bool) -> Element | None:
        name = local_name(self._element)
        if name == "grpSp":
            tag = "p:grpSpPr"
        elif name in {"sp", "pic", "cxnSp"}:
            tag = "p:spPr"
        else:
            if create:
                raise ValueError(f"{self.id}: a {self.kind} has no fill of its own")
            return None
        if create:
            return subelement(self._element, tag)
        return self._element.find(qn(tag))

    def _write_fill(self, element: Element | None) -> None:
        if element is None and self._fill_container(create=False) is None:
            return
        self._before_change()
        container = self._fill_container(create=True)
        assert container is not None
        write_fill(container, element)
        self._after_change()

    # -- structure -------------------------------------------------------------------------

    def delete(self) -> None:
        """Remove the shape from its slide.

        Relationships only this shape used -- its picture, a chart, a hyperlink -- go with
        it, and so do their parts once nothing else in the package refers to them (see
        :meth:`~ooxml_edit.opc.OpcPackage.reap`).  A picture another shape or slide
        still shows stays.

        For example::

            shape.delete()
        """
        with self._slide.document.batch():
            parent = self.parent_group
            self._before_change()
            gone = _creating.raw_ids_within(self._element)
            remove(self._element)
            self._slide._detach(gone)
            self._slide._invalidate()
            self._slide._settle()
            if parent is not None:
                parent._refit()

    # -- pictures --------------------------------------------------------------------------

    @property
    def image_part(self) -> str | None:
        """The package part a picture shows (``ppt/media/image3.png``), else ``None``.

        For example::

            picture.image_part                            # 'ppt/media/image1.png'
        """
        blip = self._blip()
        if blip is None:
            return None
        return self._slide.document.package.related_part(self._slide.part_path,
                                                         blip.get(qn("r:embed")))

    @property
    def image_size(self) -> "ImageSize | None":
        """The picture's image as stored: its size in pixels and the resolution it states
        (``dpi``, ``None`` when it states none), as an :class:`~pptx_agent.ImageSize`;
        ``None`` for a shape that is not a picture or an image whose header cannot be read.

        For example::

            picture.image_size                    # ImageSize(width=600, height=240, dpi=(72.0, 72.0))
            picture.image_size.aspect             # 2.5
        """
        part = self.image_part
        if part is None:
            return None
        data = self._slide.document.package.read(part)
        pixels = image_size(data)
        if pixels is None or not all(pixels):
            return None
        return ImageSize(pixels[0], pixels[1], image_dpi(data))

    def replace_image(self, image: "bytes | str | os.PathLike[str]", *, keep: str = "frame",
                      anchor: str = "top_left") -> "Shape":
        """Show a different image in this picture, keeping its effects and outline.

        ``keep`` says what of the frame stays: ``"frame"`` (the default) its position and
        size, so a new image of another shape is stretched into it; ``"height"`` the
        height, the width following the new image's proportions; ``"width"`` the width;
        ``"none"`` neither -- the new image at its pixel size at 96 dpi, as
        :meth:`Slide.add_picture` places one.  When the size changes, ``anchor`` is the
        point of the frame that stays put: ``"top_left"`` (the default), ``"top"``,
        ``"top_right"``, ``"left"``, ``"center"``, ``"right"``, ``"bottom_left"``,
        ``"bottom"`` or ``"bottom_right"``.

        The crop (``a:srcRect``) and any SVG alternative belonged to the old image and are
        dropped.  The old image's relationship is removed, and its media part too unless
        something else in the package still shows it.  One undo step.

        For example::

            picture.replace_image("new-logo.png")                       # same frame
            logo.replace_image("new-logo.png", keep="height", anchor="top_right")
        """
        if self._blip() is None:
            raise ValueError(f"{self.id}: a {self.kind} has no image to replace")
        if keep not in ("frame", "height", "width", "none"):
            raise ValueError("keep is 'frame', 'height', 'width' or 'none'")
        if anchor not in _ANCHORS:
            raise ValueError(f"anchor is one of {sorted(_ANCHORS)}")
        data, suffix = _read_image(image)
        frame = None
        if keep != "frame":
            pixels = image_size(data)
            if pixels is None or not all(pixels):
                raise ValueError("cannot read this image's size; keep='frame' needs none")
            left, top, width, height = self.left, self.top, self.width, self.height
            if None in (left, top, width, height):
                raise ValueError(f"{self.id}: the picture has no frame to keep")
            if keep == "height":
                new_w, new_h = round(height * pixels[0] / pixels[1]), height
            elif keep == "width":
                new_w, new_h = width, round(width * pixels[1] / pixels[0])
            else:
                new_w, new_h = (dimension * EMU_PER_INCH // 96 for dimension in pixels)
            fx, fy = _ANCHORS[anchor]
            frame = (round(left + (width - new_w) * fx), round(top + (height - new_h) * fy),
                     new_w, new_h)
        package = self._slide.document.package
        with self._slide.document.batch():
            self._before_change()
            media = package.add_image(data, suffix)
            rel_id = package.add_relationship(self._slide.part_path, REL_IMAGE, media)
            blip = self._blip()
            assert blip is not None
            blip.set(qn("r:embed"), rel_id)
            blip.attrib.pop(qn("r:link"), None)
            extensions = blip.find(qn("a:extLst"))
            if extensions is not None:
                for extension in list(extensions):
                    if extension.get("uri") == _SVG_BLIP_EXT_URI:
                        remove(extension)
                if len(extensions) == 0:
                    remove(extensions)
            crop = blip.getparent().find(qn("a:srcRect"))
            if crop is not None:
                remove(crop)
            self._after_change()
            if frame is not None:
                self._set_offset(x=frame[0], y=frame[1])
                self._set_extent(cx=frame[2], cy=frame[3])
        return self

    def _blip(self) -> Element | None:
        if local_name(self._element) != "pic":
            return None
        return find(self._element, "p:blipFill/a:blip")

    # -- ungrouping ------------------------------------------------------------------------

    def ungroup(self) -> list["Shape"]:
        """Dissolve a group, putting its children back where the group drew them.

        Each child's transform is composed with the group's -- child-space scale, then the
        group's flips and rotation about its centre -- so nothing moves on the slide.  The
        children take the group's place in the z-order and keep their ids.  A child filled
        "like its group" (``a:grpFill``) gets a copy of the group's fill.  Returns the children.

        For example::

            members = group.ungroup()
        """
        if local_name(self._element) != "grpSp":
            raise ValueError(f"{self.id}: only a group can be ungrouped")
        with self._slide.document.batch():
            self._before_change()
            children = self.children
            for child in children:
                child._before_change()  # freeze ids: a path-derived id changes with the path
            xfrm = self._xfrm(create=False)
            rotation = (get_int(xfrm, "rot", 0) or 0) / ROTATION_UNIT if xfrm is not None else 0.0
            flip_h = xfrm is not None and xfrm.get("flipH") in {"1", "true"}
            flip_v = xfrm is not None and xfrm.get("flipV") in {"1", "true"}
            off = self._xfrm_child("a:off")
            ext = self._xfrm_child("a:ext")
            centre_x = (get_int(off, "x", 0) or 0) + (get_int(ext, "cx", 0) or 0) / 2
            centre_y = (get_int(off, "y", 0) or 0) + (get_int(ext, "cy", 0) or 0) / 2
            group_fill = next((c for c in _children_of(self._fill_container(create=False))
                               if c.tag in {qn(t) for t in FILL_TAGS}), None)

            elements = [child._element for child in children]
            for child in children:
                child._compose_into_parent(self, rotation, flip_h, flip_v, centre_x, centre_y)
                child._inherit_group_fill(group_fill, self.parent_group is not None)
            for element in elements:
                remove(element)
                element.tail = None
                self._element.addprevious(element)
            parent = self.parent_group
            remove(self._element)
            self._slide._invalidate()
            if parent is not None:
                parent._refit()
            return [self._slide._wrap(element) for element in elements]

    def _compose_into_parent(self, group: "Shape", rotation: float, flip_h: bool, flip_v: bool,
                             centre_x: float, centre_y: float) -> None:
        import math

        left, top, width, height = self.left, self.top, self.width, self.height
        if None in (left, top, width, height):
            return
        mapped = group._map_child_rect(float(left), float(top), float(width), float(height))
        if mapped is None:
            return
        x, y, w, h = mapped
        cx, cy = x + w / 2, y + h / 2
        if flip_h:
            cx = 2 * centre_x - cx
        if flip_v:
            cy = 2 * centre_y - cy
        if rotation:
            angle = math.radians(rotation)
            dx, dy = cx - centre_x, cy - centre_y
            cx = centre_x + dx * math.cos(angle) - dy * math.sin(angle)
            cy = centre_y + dx * math.sin(angle) + dy * math.cos(angle)
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        offset, extent = subelement(xfrm, "a:off"), subelement(xfrm, "a:ext")
        set_int(offset, "x", round(cx - w / 2))
        set_int(offset, "y", round(cy - h / 2))
        set_int(extent, "cx", round(w))
        set_int(extent, "cy", round(h))
        # A reflection reverses the sense of a rotation inside it; two cancel out.
        own = self.rotation
        sign = -1 if flip_h != flip_v else 1
        combined = round((rotation + sign * own) * ROTATION_UNIT) % (360 * ROTATION_UNIT)
        if combined or xfrm.get("rot") is not None:
            set_int(xfrm, "rot", combined or None)
        for attribute, flipped in (("flipH", flip_h), ("flipV", flip_v)):
            if flipped:
                mine = xfrm.get(attribute) in {"1", "true"}
                if mine:
                    del xfrm.attrib[attribute]
                else:
                    xfrm.set(attribute, "1")

    def _inherit_group_fill(self, group_fill: Element | None, nested: bool) -> None:
        """Replace ``a:grpFill`` (fill like the group) with what the group's fill was."""
        container = self._fill_container(create=False) if self.kind != "graphic_frame" else None
        if container is None or container.find(qn("a:grpFill")) is None:
            return
        if nested and group_fill is not None and local_name(group_fill) == "grpFill":
            return  # the group itself filled like *its* group, which is now this one's parent
        replacement = copy.deepcopy(group_fill) if group_fill is not None \
            and local_name(group_fill) != "grpFill" else make("a:noFill")
        write_fill(container, replacement)

    def duplicate(self, dx: int = 0, dy: int = 0) -> "Shape":
        """Copy the shape into the same container, offset by ``(dx, dy)``.

        The copy gets a fresh ``cNvPr@id`` and loses any inherited identity -- a stamped id or
        a ``creationId`` -- because it is a different shape and must not answer to the
        original's address.  A group's copy renumbers every shape inside it too, and a
        connector inside it that was glued to a shape inside it is glued to that shape's
        copy; one glued to a shape outside the group stays glued to that shape.

        For example::

            copy = shape.duplicate(dx=457200)
        """
        parent = self._element.getparent()
        if parent is None:
            raise ValueError(f"{self.id}: shape is not attached to a slide")

        self._before_change()
        clone = copy.deepcopy(self._element)
        _strip_identity(clone)
        renumber_copy(clone, self._slide._next_shape_ids)
        append_in_order(parent, clone)
        self._slide._invalidate()

        duplicated = self._slide._wrap(clone)
        if dx or dy:
            left, top = duplicated.left, duplicated.top
            if left is not None and top is not None:
                duplicated.left = left + dx
                duplicated.top = top + dy
        return duplicated

    def bring_to_front(self) -> "Shape":
        """Move the shape in front of everything in its container.

        For example::

            shape.bring_to_front()
        """
        return self._reorder(-1)

    def bring_forward(self) -> "Shape":
        """One step forward: in front of the next shape in its container (PowerPoint steps
        past the next shape whether or not the two overlap -- measured).

        For example::

            shape.bring_forward()
        """
        return self._step(1)

    def send_backward(self) -> "Shape":
        """One step back: behind the previous shape in its container.

        For example::

            shape.send_backward()
        """
        return self._step(-1)

    def _step(self, direction: int) -> "Shape":
        parent = self._element.getparent()
        if parent is None:
            raise ValueError(f"{self.id}: shape is not attached to a slide")
        members = [child for child in parent if prefixed_name(child) in SHAPE_TREE_MEMBERS]
        position = members.index(self._element)
        neighbour = position + direction
        if not 0 <= neighbour < len(members):
            return self  # already frontmost (or backmost): nothing to do, nothing to undo
        self._before_change()
        other = members[neighbour]
        remove(self._element)  # leaves the surrounding whitespace where it was
        self._element.tail = None
        if direction > 0:
            other.addnext(self._element)
        else:
            other.addprevious(self._element)
        self._slide._invalidate()
        return self

    def send_to_back(self) -> "Shape":
        """Move the shape behind everything in its container.

        For example::

            shape.send_to_back()
        """
        return self._reorder(0)

    def _reorder(self, position: int) -> "Shape":
        parent = self._element.getparent()
        if parent is None:
            raise ValueError(f"{self.id}: shape is not attached to a slide")
        self._before_change()
        siblings = [child for child in parent if local_name(child) in _XFRM_PARENT]
        remove(self._element)
        if position == 0 and siblings:
            first = siblings[0] if siblings[0] is not self._element else None
            if first is not None:
                first.addprevious(self._element)
            else:
                append_in_order(parent, self._element)
        else:
            # In order, not appended: a shape tree may end with an extLst the shape must
            # stay in front of.
            append_in_order(parent, self._element)
        self._slide._invalidate()
        return self

    # -- internals -------------------------------------------------------------------------

    def _before_change(self) -> None:
        """Checkpoint for undo, then freeze this shape's id so later edits cannot move it."""
        self._slide.document.history.checkpoint()
        self._slide._watch()
        properties = cnv_pr(self._element)
        if properties is not None and read_stamp(properties) is None:
            write_stamp(properties, self._id.local)
        self._slide._touch()

    def _xfrm(self, *, create: bool) -> Element | None:
        parent_tag = _XFRM_PARENT.get(local_name(self._element))
        if parent_tag is None:
            return None
        if parent_tag == "p:xfrm":  # graphicFrame carries its transform directly
            if create:
                return subelement(self._element, "p:xfrm")
            return self._element.find(qn("p:xfrm"))
        if create:
            return subelement(subelement(self._element, parent_tag), "a:xfrm")
        container = self._element.find(qn(parent_tag))
        return None if container is None else container.find(qn("a:xfrm"))

    def _offset(self, attribute: str) -> int | None:
        xfrm = self._xfrm(create=False)
        if xfrm is not None:
            value = get_int(xfrm.find(qn("a:off")), attribute)
            if value is not None:
                return value
        return self._inherited_offset(attribute)

    def _extent(self, attribute: str) -> int | None:
        xfrm = self._xfrm(create=False)
        if xfrm is not None:
            value = get_int(xfrm.find(qn("a:ext")), attribute)
            if value is not None:
                return value
        return self._inherited_extent(attribute)

    def _set_offset(self, **values: int) -> None:
        self._before_change()
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        node = subelement(xfrm, "a:off")
        # Materialising a transform means filling in both axes, or PowerPoint reads a partial
        # a:off as (0, 0) rather than falling back to the inherited position.
        for attribute, inherited in (("x", self._inherited_offset("x")),
                                     ("y", self._inherited_offset("y"))):
            if attribute in values:
                set_int(node, attribute, values[attribute])
            elif node.get(attribute) is None:
                set_int(node, attribute, inherited or 0)
        self._ensure_extent()
        self._after_geometry_change()
        self._slide._touch()

    def _set_extent(self, **values: int) -> None:
        self._before_change()
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        node = subelement(xfrm, "a:ext")
        for attribute, inherited in (("cx", self._inherited_extent("cx")),
                                     ("cy", self._inherited_extent("cy"))):
            if attribute in values:
                set_int(node, attribute, values[attribute])
            elif node.get(attribute) is None:
                set_int(node, attribute, inherited or 0)
        self._ensure_offset()
        self._after_geometry_change()
        self._slide._touch()

    def _ensure_extent(self) -> None:
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        if xfrm.find(qn("a:ext")) is None:
            node = subelement(xfrm, "a:ext")
            set_int(node, "cx", self._inherited_extent("cx") or 0)
            set_int(node, "cy", self._inherited_extent("cy") or 0)

    def _ensure_offset(self) -> None:
        xfrm = self._xfrm(create=True)
        assert xfrm is not None
        if xfrm.find(qn("a:off")) is None:
            node = subelement(xfrm, "a:off")
            set_int(node, "x", self._inherited_offset("x") or 0)
            set_int(node, "y", self._inherited_offset("y") or 0)

    def _inherited_offset(self, attribute: str) -> int | None:
        xfrm = self._slide._inherited_xfrm(self.placeholder)
        return get_int(xfrm.find(qn("a:off")), attribute) if xfrm is not None else None

    def _inherited_extent(self, attribute: str) -> int | None:
        xfrm = self._slide._inherited_xfrm(self.placeholder)
        return get_int(xfrm.find(qn("a:ext")), attribute) if xfrm is not None else None

    def __repr__(self) -> str:
        return f"<Shape {self.id} {self.kind} name={self.name!r}>"


class Slide(ShapeFactory):
    """One slide, and the shapes on it.

    New shapes: :meth:`add_shape`, :meth:`add_textbox`, :meth:`add_connector`,
    :meth:`add_table` (from :class:`~pptx_agent.edit.creating.ShapeFactory`), and
    :meth:`add_picture`.

    For example::

        slide = deck.slides[0]
        slide.shape("256.29").set_text("12.1%")
        png = slide.render_png(width=1280)
    """

    def __init__(self, document: "Document", slide_id: int, part_path: str) -> None:
        self.document = document
        self.slide_id = slide_id
        self.part_path = normalize_part_path(part_path)
        self._index: SlideShapeIndex | None = None
        #: Relationship ids the part referenced when the current edit began; see _settle.
        self._baseline: set[str] | None = None

    # -- structure -------------------------------------------------------------------------

    @property
    def layout(self) -> Layout | None:
        """The layout the slide is built on, or ``None``.

        For example::

            slide.layout.name                            # 'Title and Content'
        """
        part = _slides.slide_layout(self.document.package, self.part_path)
        return next((layout for layout in self.document.layouts if layout.part_path == part),
                    None)

    def delete(self) -> None:
        """Remove this slide.  See :meth:`Document.delete_slide`.

        For example::

            slide.delete()
        """
        self.document.delete_slide(self)

    def duplicate(self, index: int | None = None, *, notes: str | None = None) -> "Slide":
        """Copy this slide; the copy goes right after it unless ``index`` says otherwise.
        ``notes`` replaces the copy's speaker notes (see :meth:`Document.duplicate_slide`).

        For example::

            copy = slide.duplicate(notes="Engineering works from the Utrecht office.")
        """
        return self.document.duplicate_slide(self, index=index, notes=notes)

    @property
    def content_area(self) -> tuple[int, int, int, int]:
        """Where this slide's content goes, EMU ``(left, top, width, height)``: its
        layout's body and content placeholders together, or -- for a layout without one
        (Title Only, Blank) -- the band below the title and above any footer the layout
        draws (see :func:`pptx_agent.edit.area.content_area`).  A fact about the layout,
        to place a graphic into; not a rule.

        For example::

            left, top, width, height = slide.content_area
            slide.add_shape("rect", left, top, width, height // 2)
        """
        from .area import content_area

        return content_area(self)

    @property
    def theme(self) -> "Theme":
        """The theme this slide draws with -- its master's, through its layout's and its
        own colour-map overrides (:class:`~pptx_agent.Theme`).

        For example::

            slide.theme.colors["tx1"]                    # '#000000'
        """
        from .theme import theme_for

        return theme_for(self)

    # -- speaker notes ---------------------------------------------------------------------

    @property
    def notes(self) -> str:
        r"""The speaker notes, raw: paragraphs joined by ``"\n"``, a line break as ``"\v"``;
        ``""`` when the slide has none.  Settable: the new text is diffed against the old,
        as :meth:`Shape.set_text` does, so formatting that survives stays.  A slide without
        notes gets a notes page (and the deck a notes master, as PowerPoint adds them);
        one undo step.  The address is ``"<sldId>/notes"`` (:meth:`Document.resolve`)::

            slide.notes = "Lead with the margin.\nThen the outlook."
            print(deck.slide(257).notes)
        """
        return self.notes_frame.text

    @notes.setter
    def notes(self, value: str) -> None:
        self.set_notes(value)

    def set_notes(self, value: str) -> "Slide":
        """Set the speaker notes (as :attr:`notes`); returns the slide.

        For example::

            deck.slide(261).set_notes("Engineering works from the Utrecht office.")
        """
        value = "" if value is None else str(value)
        if not value and not self.has_notes:
            return self
        self.notes_frame.set_text(value)
        return self

    @property
    def has_notes(self) -> bool:
        """Whether the slide has a notes page (it may be empty).

        For example::

            [slide.slide_id for slide in deck.slides if slide.has_notes]
        """
        from . import notes as _notes

        return _notes.notes_part(self.document.package, self.part_path) is not None

    @property
    def notes_frame(self) -> TextFrame:
        """The speaker notes as a :class:`TextFrame`, addressed ``"<sldId>/notes"``: its
        paragraphs and runs, formatting and all.  Writing to it gives a slide without
        notes a notes page.

        For example::

            frame = slide.notes_frame
            frame.paragraph(0).run(0).bold = True
            frame.add_paragraph("Ask for the decision.")
        """
        document, slide_id = self.document, self.slide_id
        address = f"{slide_id}/notes"
        return TextFrame(lambda: _NotesHost(document.slide(slide_id), address), address)

    def move_to(self, index: int) -> "Slide":
        """Move this slide to 0-based position ``index``.  Its id, and its shapes', stay.

        For example::

            slide.move_to(0)
        """
        return self.document.move_slide(self, index)

    # -- pictures --------------------------------------------------------------------------

    def add_picture(self, image: "bytes | str | os.PathLike[str]", left: int = 0, top: int = 0,
                    width: int | None = None, height: int | None = None, *,
                    name: str | None = None, description: str | None = None) -> Shape:
        """Insert a picture shape, frontmost; returns it.

        ``image`` is the image's bytes or a path (PNG, JPEG, GIF, BMP or TIFF).  Without a size
        the picture is placed at its pixel size at 96 dpi; with only one of ``width`` and
        ``height`` the other keeps the aspect ratio.  The image is stored once -- an identical
        one already in the deck is reused -- and undo removes it again.

        For example::

            slide.add_picture("logo.png", left=914400, top=457200, width=1828800)
        """
        data, suffix = _read_image(image)
        pixels = image_size(data)
        if width is None or height is None:
            if pixels is None or not all(pixels):
                raise ValueError("cannot read this image's size; pass width and height")
            natural_w, natural_h = (dimension * EMU_PER_INCH // 96 for dimension in pixels)
            if width is None and height is None:
                width, height = natural_w, natural_h
            elif width is None:
                width = round(height * natural_w / natural_h)
            else:
                height = round(width * natural_h / natural_w)
        package = self.document.package
        with self.document.batch():
            self.document.history.checkpoint()
            self._watch()
            media = package.add_image(data, suffix)
            rel_id = package.add_relationship(self.part_path, REL_IMAGE, media)
            identifier = self._next_shape_id()
            picture = make("p:pic")
            nv = make("p:nvPicPr")
            properties = make("p:cNvPr", id=str(identifier),
                              name=name or f"Picture {identifier - 1}")
            if description:
                properties.set("descr", description)
            nv.append(properties)
            locks = make("p:cNvPicPr")
            locks.append(make("a:picLocks", noChangeAspect="1"))
            nv.append(locks)
            nv.append(make("p:nvPr"))
            picture.append(nv)
            blip_fill = make("p:blipFill")
            blip_fill.append(make("a:blip", r__embed=rel_id))
            stretch = make("a:stretch")
            stretch.append(make("a:fillRect"))
            blip_fill.append(stretch)
            picture.append(blip_fill)
            shape_properties = make("p:spPr")
            xfrm = make("a:xfrm")
            xfrm.append(make("a:off", x=str(int(left)), y=str(int(top))))
            xfrm.append(make("a:ext", cx=str(int(width)), cy=str(int(height))))
            shape_properties.append(xfrm)
            geometry = make("a:prstGeom", prst="rect")
            geometry.append(make("a:avLst"))
            shape_properties.append(geometry)
            picture.append(shape_properties)
            append_in_order(self._sp_tree(), picture)
            self._invalidate()
            self._settle()
            return self._wrap(picture)

    # -- relationship upkeep ---------------------------------------------------------------

    def _watch(self) -> None:
        """Before an edit: note which relationship ids the part references now.

        Accumulated until the next :meth:`_settle`, so an id referenced at *any* point since
        then -- including one an earlier, unsettled edit added -- is a candidate.
        """
        root = self.document.package.tree(self.part_path)
        current = _relationship_references(root) if root is not None else set()
        self._baseline = current if self._baseline is None else self._baseline | current

    def _settle(self) -> None:
        """After an edit: release relationships it stopped referencing.

        Only ids the part referenced before the edit are candidates, so a relationship that
        was already unused in the original file is left exactly as it was.  The package then
        proves each one unreferenced before removing it, and reaps its target if nothing else
        in the package uses it.
        """
        baseline, self._baseline = self._baseline, None
        if not baseline:
            return
        root = self.document.package.tree(self.part_path)
        dropped = baseline - (_relationship_references(root) if root is not None else set())
        if dropped:
            self.document.package.release(self.part_path, sorted(dropped))

    # -- shapes ----------------------------------------------------------------------------

    @property
    def shapes(self) -> list[Shape]:
        """Every shape on the slide, groups' members included, in document order -- back
        to front, each group followed at once by its members (and theirs, depth first).
        For the top level alone, keep those whose ``parent_group`` is ``None``; a group's
        own members are its ``children``.

        For example::

            [(shape.id, shape.kind, shape.name) for shape in slide.shapes]
            top = [shape for shape in slide.shapes if shape.parent_group is None]
        """
        index = self._ensure_index()
        return [Shape(self, element, shape_id) for shape_id, element in index]

    def shape(self, shape_id: str | ShapeId) -> Shape:
        """A shape on this slide by id.

        For example::

            slide.shape("256.29")
        """
        parsed = ShapeId.parse(shape_id)
        element = self._ensure_index().get(parsed)
        if element is None:
            raise KeyError(f"no shape {parsed} on slide {self.slide_id}")
        return Shape(self, element, parsed)

    def find_shapes(self, *, name: str | None = None, kind: str | None = None) -> list[Shape]:
        """The top-level shapes with this ``name`` and/or ``kind``.

        For example::

            slide.find_shapes(kind="picture")
        """
        return [
            shape
            for shape in self.shapes
            if (name is None or shape.name == name) and (kind is None or shape.kind == kind)
        ]

    def copy_shapes(self, shapes: "list[Shape | str]", to_slide: "Slide | None" = None, *,
                    at: "tuple[int, int] | None" = None, dx: int = 0, dy: int = 0):
        """Copy shapes of this slide -- loose shapes, groups, connectors between them,
        pictures, charts, SmartArt -- to ``to_slide`` (this slide by default), which may be a
        slide of another deck.  Returns a :class:`~pptx_agent.edit.copying.CopyResult`:
        the copies, ``mapping`` from every copied shape's address to its copy's (group
        members included), the connector ends that were glued outside the set and are now
        loose, and the theme colours the target slide draws differently.

        Every copied shape gets a fresh id; glue inside the set follows the copies; media
        is shared in one deck and imported once into another; charts and diagrams are
        copied with their parts; a placeholder becomes a plain shape.  ``at`` (EMU) puts the
        top-left corner of the set's bounds there; otherwise the copy is offset by
        ``(dx, dy)``.  One undo step on the target deck.

        For example::

            result = slide.copy_shapes(["256.5", "256.9"], other_deck.slides[1],
                                       at=(914400, 1371600))
            result.mapping                               # {'256.5': '257.12', ...}
        """
        from .copying import copy_shapes

        resolved = [self.shape(s) if isinstance(s, str) else s for s in shapes]
        return copy_shapes(resolved, to_slide, at=at, dx=dx, dy=dy)

    def group(self, shapes: "list[Shape | str]", *, name: str | None = None) -> Shape:
        """Group shapes that share a container; returns the new group.

        The group's child space starts out identical to its frame (``chOff`` = ``off``,
        ``chExt`` = ``ext`` = the children's bounding box), so nothing moves.  It takes the
        z-position of the frontmost member.  Members keep their ids: each is stamped before
        it moves, so even a path-derived id survives the change of path.

        For example::

            group = slide.group([box, "256.5"], name="Header")
        """
        members = [self.shape(s) if isinstance(s, str) else s for s in shapes]
        if len(members) < 2:
            raise ValueError("a group needs at least two shapes")
        parent = members[0]._element.getparent()
        if parent is None or any(m._element.getparent() is not parent for m in members):
            raise ValueError("grouped shapes must share a container (the slide or one group)")
        if len({id(m._element) for m in members}) != len(members):
            raise ValueError("a shape was listed twice")

        with self.document.batch():
            boxes = []
            for member in members:
                member._before_change()
                # Materialise inherited placeholder geometry: inside a group it must be explicit.
                left, top, width, height = member.left, member.top, member.width, member.height
                if None in (left, top, width, height):
                    raise ValueError(f"{member.id}: cannot group a shape with no position")
                if not member.has_explicit_transform:
                    member.left = left
                    member.width = width
                boxes.append((left, top, left + width, top + height))
            min_x = min(b[0] for b in boxes)
            min_y = min(b[1] for b in boxes)
            width = max(b[2] for b in boxes) - min_x
            height = max(b[3] for b in boxes) - min_y

            identifier = self._next_shape_id()
            group = make("p:grpSp")
            nv = make("p:nvGrpSpPr")
            nv.append(make("p:cNvPr", id=str(identifier), name=name or f"Group {identifier - 1}"))
            nv.append(make("p:cNvGrpSpPr"))
            nv.append(make("p:nvPr"))
            group.append(nv)
            properties = make("p:grpSpPr")
            xfrm = make("a:xfrm")
            for tag, attributes in (("a:off", {"x": min_x, "y": min_y}),
                                    ("a:ext", {"cx": width, "cy": height}),
                                    ("a:chOff", {"x": min_x, "y": min_y}),
                                    ("a:chExt", {"cx": width, "cy": height})):
                xfrm.append(make(tag, **{k: str(v) for k, v in attributes.items()}))
            properties.append(xfrm)
            group.append(properties)

            ordered = [child for child in parent if any(child is m._element for m in members)]
            ordered[-1].addnext(group)
            for element in ordered:
                remove(element)  # keeps the surrounding whitespace where it was
                element.tail = None
                group.append(element)
            self._invalidate()
            return self._wrap(group)

    @property
    def title(self) -> str | None:
        """The text of the slide's title placeholder (``title``, or a title slide's
        ``ctrTitle``), raw; ``None`` when the slide has none.  Setting it replaces that
        text, keeping its formatting (:meth:`Shape.set_text`); a slide with no title
        placeholder raises ``ValueError``.

        For example::

            slide.title = "Budget"
            deck.slide_titled("Budget").slide_id == slide.slide_id   # True
        """
        shape = self._title_shape()
        return None if shape is None else shape.text

    @title.setter
    def title(self, value: str) -> None:
        shape = self._title_shape()
        if shape is None:
            raise ValueError(f"slide {self.slide_id} has no title placeholder; its layout "
                             f"({self.layout.name if self.layout else 'none'}) gives none, or "
                             f"it was deleted")
        shape.set_text(value)

    def _title_shape(self) -> "Shape | None":
        for shape in self.shapes:
            placeholder = shape.placeholder
            if placeholder is not None and placeholder.type in ("title", "ctrTitle"):
                return shape
        return None

    def collisions(self, *, boxes: bool = False) -> list:
        """What collides on this slide -- the ``"overlap"`` items of
        :meth:`Document.overflows` (:class:`~pptx_agent.Overflow`), each with a ``detail``:

        * ``"text"``: two shapes holding text (autoshapes, text boxes, tables) overlap --
          their drawn boxes, or for a text box with neither fill nor outline the text
          itself, placed by its alignment and anchor;
        * ``"line"``: a line or connector crosses a text-bearing shape's text area (the
          text itself, for a text box with neither fill nor outline) -- unless the shape is
          opaque and in front, hiding the line; ``amount`` is the length crossed, EMU.  So
          a marker line (a "today" or "board update" line) drawn in front of bars is
          reported where it crosses their text, and the same line sent behind them
          (:meth:`Shape.send_to_back`) is not: it is hidden there, and there is no need to
          split it into pieces between the bars;
        * ``"placeholder"``: a shape crosses the text (or, for a picture, chart or table,
          the frame) of a title or another placeholder holding something.

        With ``boxes=True``, also

        * ``"box"``: two shapes holding text whose boxes overlap although their text does
          not -- labels side by side whose frames run into each other, which PowerPoint
          shows as soon as a box is selected or its text grows.  Off by default: the text
          is what a reader sees.

        Intended layering is not a collision: a shape lying wholly on an opaque shape
        behind it (a label on a bar, a title on a banner, text on a picture).  Groups are
        looked into, so a group is never compared with its own members.  Overlaps under
        a point (lines: crossings under a point) are not reported.  Needs pptx2svg.

        For example::

            for problem in slide.collisions():
                print(problem.detail, problem.shape, problem.other, problem.amount)
        """
        _require_renderer()
        from .fit import slide_problems

        return slide_problems(self, collisions_only=True, boxes=boxes)

    def design_facts(self, *, region=None, within: int = 25400, include=None,
                     **tolerances) -> "DesignFacts":
        """Measurable design facts of this slide -- palette, like shapes and the colours
        they carry (and any legend-like group), empty regions of the content area,
        alignment and near-misses, shape vocabulary, text sizes, lines over text -- with no
        verdicts; see :func:`pptx_agent.edit.design.design_facts`.  ``within`` is EMU
        (default 2 pt); ``region`` is ``(left, top, width, height)`` EMU.

        For example::

            facts = slide.design_facts()
            [group.accent_hues for group in facts.color_groups]
        """
        from .design import design_facts

        return design_facts(self, region=region, within=within, include=include,
                            **tolerances)

    def facts(self) -> dict:
        """The problem facts beyond overflows and collisions: colours that are not theme
        colours, and title and near-wrap margins, in points; see
        :func:`pptx_agent.edit.design.slide_facts`.  Needs pptx2svg.

        For example::

            slide.facts()["non_theme_colors"]
        """
        _require_renderer()
        from .design import slide_facts

        return slide_facts(self)

    # -- rendering -------------------------------------------------------------------------

    @property
    def index(self) -> int:
        """0-based position in presentation order.  Read it as a property or call it --
        ``slide.index`` and ``slide.index()`` are both the position, an ``int``.

        For example::

            deck.slides[2].index                         # 2
        """
        for position, slide in enumerate(self.document.slides):
            if slide.slide_id == self.slide_id:
                return _Position(position)
        raise ValueError(f"slide {self.slide_id} is no longer in the presentation")

    def render_svg(self, *, rewrite_ids: bool = True, full_state: bool = False,
                   agent: bool = False, **options) -> str:
        """The slide as SVG, via pptx2svg, with this library's shape ids on the groups.

        ``agent=True`` gives pptx2svg's compact agent view instead (LR3): user units are
        points, every shape carries ``data-id`` -- its address here, which every tool
        accepts -- colours carry their theme names, text is plain ``<text>`` with its runs,
        and pictures, charts and tables are placeholders: no fonts, images or glyphs.  The
        layout's and master's own shapes come first, marked ``data-layer``.

        ``full_state=True`` adds the ``data-ooxml-*`` vocabulary: every shape's state as
        typed attributes read from the unresolved document, and its OOXML as base64 -- an
        SVG that :meth:`Document.apply_svg` can read back losslessly.  Only attributes are
        added, so it draws exactly like the plain one.

        For example::

            svg = slide.render_svg(full_state=True)
        """
        if agent:
            if full_state:
                raise ValueError("the agent view is read-only: no full state")
            svg = self.document.render_svg(slides=[self.index + 1], agent=True)[0]
            return self._rewrite_ids(svg).replace('data-pptx-id="', 'data-id="')
        svg = self.document.render_svg(slides=[self.index + 1], **options)[0]
        if full_state:
            if not rewrite_ids:
                raise ValueError("a full-state SVG is addressed by this library's ids")
            from ..fullstate import emit_full_state

            return emit_full_state(self, self._rewrite_ids(svg))
        return self._rewrite_ids(svg) if rewrite_ids else svg

    def apply_svg(self, svg: "str | bytes", **options):
        """Apply a full-state SVG of *this* slide.  See :meth:`Document.apply_svg`.

        For example::

            slide.apply_svg(edited_svg)
        """
        from ..fullstate.apply import read_svg
        from ..fullstate.safe import FullStateError

        named, _ = read_svg(svg, options.get("limits") or _default_limits())
        if named.get("slide-id") != str(self.slide_id):
            raise FullStateError(f"the SVG is of slide {named.get('slide-id')}, not "
                                 f"{self.slide_id}")
        return self.document.apply_svg(svg, **options)

    def render_png(self, **options) -> bytes:
        """Render this slide to PNG via pptx2svg; ``width`` and ``height`` in pixels.

        For example::

            png = slide.render_png(width=1280)
        """
        return self.document.render_png(slides=[self.index + 1], **options)[0]

    def _rewrite_ids(self, svg: str) -> str:
        """Replace pptx2svg's ``data-pptx-id`` with this library's canonical shape ids.

        pptx2svg addresses a shape by its raw ``cNvPr@id``, which is not unique in real decks;
        this library's ids are.  Rather than join on the index path -- the two walks skip
        different nodes, so the paths are not guaranteed to correspond -- duplicates are
        matched in document order, which both sides preserve.
        """
        import re
        from collections import defaultdict

        index = self._ensure_index()
        pending: dict[str, list[str]] = defaultdict(list)
        for shape_id, element in index:
            properties = cnv_pr(element)
            raw = properties.get("id") if properties is not None else None
            if raw is not None:
                pending[raw].append(str(shape_id))

        cursor: dict[str, int] = defaultdict(int)
        prefix = f"{self.slide_id}."

        def replace(match: re.Match[str]) -> str:
            value = match.group(1)
            if not value.startswith(prefix):
                return match.group(0)  # inherited from the layout or master: leave it alone
            raw, slash, piece = value[len(prefix):].partition("/")
            candidates = pending.get(raw)
            if not candidates:
                return match.group(0)
            if slash:
                # A drawn piece of a shape (SmartArt's cached shapes, "<frame>/0/3"): it
                # follows its frame, so it takes the id the frame was just given.
                position = min(max(cursor[raw] - 1, 0), len(candidates) - 1)
                return f'data-pptx-id="{candidates[position]}/{piece}"'
            position = min(cursor[raw], len(candidates) - 1)
            cursor[raw] += 1
            return f'data-pptx-id="{candidates[position]}"'

        return re.sub(r'data-pptx-id="([^"]*)"', replace, svg)

    # -- internals -------------------------------------------------------------------------

    def _sp_tree(self) -> Element:
        root = self.document.package.tree(self.part_path)
        if root is None:
            raise ValueError(f"slide part {self.part_path} is missing")
        tree = find(root, "p:cSld/p:spTree")
        if tree is None:
            raise ValueError(f"slide part {self.part_path} has no shape tree")
        return tree

    def _ensure_index(self) -> SlideShapeIndex:
        if self._index is None:
            self._index = SlideShapeIndex(self.slide_id, self._sp_tree())
        return self._index

    def _invalidate(self) -> None:
        """Structure changed: index paths may have shifted, so rebuild on next access."""
        self._index = None
        self._touch()

    def _touch(self) -> None:
        self.document.package.mark_dirty(self.part_path)

    def _wrap(self, element: Element) -> Shape:
        index = self._ensure_index()
        shape_id = index.id_of(element)
        if shape_id is None:
            raise ValueError("shape is not present in this slide's index")
        return Shape(self, element, shape_id)

    def _next_shape_id(self) -> int:
        """A ``cNvPr@id`` not used anywhere in this slide part -- nor as another shape's
        local id (a stamp can outlive the ``cNvPr@id`` it froze), so a new shape's id is
        its ``cNvPr@id`` without a disambiguating path."""
        return self._next_shape_ids(1)[0]

    def _next_shape_ids(self, count: int) -> list[int]:
        """``count`` fresh ids at once, as :meth:`_next_shape_id` picks one: for a copy whose
        shapes all need one before any of them is in the tree."""
        used = set()
        tree = self._sp_tree()
        for node in tree.iter(qn("p:cNvPr")):
            raw = node.get("id")
            if raw and raw.isdigit():
                used.add(int(raw))
        taken = {shape_id.local for shape_id, _ in SlideShapeIndex(self.slide_id, tree)}
        fresh: list[int] = []
        candidate = 2  # 1 is conventionally the shape tree itself
        while len(fresh) < count:
            if candidate not in used and str(candidate) not in taken:
                fresh.append(candidate)
            candidate += 1
        return fresh

    def _factory(self):
        return self, self._sp_tree(), None

    def _element_by_raw_id(self, raw: str) -> Element | None:
        """The one shape whose ``cNvPr@id`` is ``raw`` -- what ``stCxn@id`` names -- or
        ``None`` when there is none, or more than one."""
        found = [node for node in self._sp_tree().iter(qn("p:cNvPr")) if node.get("id") == raw]
        if len(found) != 1:
            return None
        shape = found[0].getparent().getparent()
        return shape if local_name(shape) in _XFRM_PARENT else None

    def _reroute_for(self, shape: "Shape") -> None:
        """Re-route every connector attached to ``shape`` or to anything inside it."""
        raw_ids = _creating.raw_ids_within(shape._element)
        for element in _creating.attached_connectors(self, raw_ids):
            if element is shape._element:
                continue
            if _creating.route_connector(self, element):
                self._touch()
                parent = element.getparent()
                if parent is not None and parent.tag == qn("p:grpSp"):
                    self._wrap(parent)._refit()

    def _detach(self, raw_ids: set[str]) -> None:
        """Shapes with these ``cNvPr@id``s are gone: connectors ending on them let go."""
        present = {node.get("id") for node in self._sp_tree().iter(qn("p:cNvPr"))}
        gone = raw_ids - present
        for element in _creating.attached_connectors(self, gone):
            references = {}
            for which in ("begin", "end"):
                reference = _creating.connection(element, which)
                references[which] = None if reference is None or reference[0] in gone \
                    else reference
            _creating._write_connections(element, references["begin"], references["end"])

    def _inherited_xfrm(self, placeholder: tuple[str | None, int | None] | None) -> Element | None:
        """The ``a:xfrm`` a placeholder without its own is drawn with: its layout
        placeholder's, or -- when that has none either -- the master placeholder of the same
        kind, as PowerPoint resolves it (a title from the master's title, a body, subtitle
        or content placeholder from the master's body, the footers from theirs)."""
        layout_shape = self._layout_placeholder(placeholder)
        if layout_shape is None:
            return None
        xfrm = find(layout_shape, "p:spPr/a:xfrm")
        if xfrm is not None and xfrm.find(qn("a:off")) is not None:
            return xfrm
        node = find(layout_shape, "p:nvSpPr/p:nvPr/p:ph")
        kind = _MASTER_KIND.get(node.get("type") if node is not None else None, "body")
        package = self.document.package
        layouts = package.related_parts_of_type(self.part_path, REL_SLIDE_LAYOUT)
        masters = package.related_parts_of_type(layouts[0], REL_SLIDE_MASTER) \
            if layouts else []
        master = package.tree(masters[0]) if masters else None
        tree = find(master, "p:cSld/p:spTree") if master is not None else None
        for shape in [] if tree is None else tree:
            if local_name(shape) != "sp":
                continue
            ph = find(shape, "p:nvSpPr/p:nvPr/p:ph")
            if ph is not None and _MASTER_KIND.get(ph.get("type"), "body") == kind:
                return find(shape, "p:spPr/a:xfrm")
        return None

    def _layout_placeholder(self, placeholder: tuple[str | None, int | None] | None) -> Element | None:
        """The layout shape a placeholder inherits from, matched by idx then by type."""
        if placeholder is None:
            return None
        wanted_type, wanted_idx = placeholder

        layout_path = self.document.package.related_parts_of_type(self.part_path, REL_SLIDE_LAYOUT)
        if not layout_path:
            return None
        layout = self.document.package.tree(layout_path[0])
        tree = find(layout, "p:cSld/p:spTree") if layout is not None else None
        if tree is None:
            return None

        by_idx: list[Element] = []
        by_type: list[Element] = []
        for shape in tree:
            if local_name(shape) not in _XFRM_PARENT:
                continue
            node = find(shape, f"{_nv_container(shape)}/p:nvPr/p:ph")
            if node is None:
                continue
            if get_int(node, "idx", 0) == (wanted_idx or 0):
                by_idx.append(shape)
            if wanted_type is not None and node.get("type") == wanted_type:
                by_type.append(shape)

        # Match by idx only when it is unambiguous; an ambiguous idx match is worse than none.
        if len(by_idx) == 1:
            return by_idx[0]
        if len(by_type) == 1:
            return by_type[0]
        return None

    def __repr__(self) -> str:
        return f"<Slide {self.slide_id} {self.part_path}>"


class _NotesHost:
    """A slide's speaker notes as a text frame's host: the notes page's notes placeholder,
    made (with a notes master, when the deck has none) the first time it is written."""

    def __init__(self, slide: Slide, address: str) -> None:
        self._slide = slide
        self.address = address

    def _part(self) -> str | None:
        from . import notes as _notes

        return _notes.notes_part(self._slide.document.package, self._slide.part_path)

    def _text_body(self, create: bool) -> Element | None:
        from . import notes as _notes

        package = self._slide.document.package
        part = self._part()
        if part is None:
            if not create:
                return None
            shape = _notes.add_notes(self._slide.document, self._slide.part_path,
                                     self._slide.index + 1)
        else:
            shape = _notes.notes_body(package, part)
            if shape is None:
                if not create:
                    return None
                raise ValueError(f"{part} has no notes placeholder to write to")
        body = shape.find(qn("p:txBody"))
        if body is None and create:
            body = subelement(shape, "p:txBody")
            subelement(body, "a:bodyPr")
            subelement(body, "a:lstStyle")
        return body

    def _before_change(self) -> None:
        self._slide.document.history.checkpoint()

    def _after_change(self) -> None:
        part = self._part()
        if part is not None:
            self._slide.document.package.mark_dirty(part)

    def _batch(self):
        return self._slide.document.batch()

    def _shape(self) -> Element | None:
        from . import notes as _notes

        part = self._part()
        return None if part is None else _notes.notes_body(self._slide.document.package, part)

    def _list_styles(self) -> list:
        from ..outline.read import notes_chain

        shape = self._shape()
        return [] if shape is None else notes_chain(self._slide.document, shape)

    def _body_chain(self) -> list:
        shape = self._shape()
        return [None if shape is None else find(shape, "p:txBody/a:bodyPr")]

    def _theme_source(self):
        return self._slide


class Document(CommentOps):
    """An open .pptx, editable and re-renderable.

    For example::

        deck = Document.open("results.pptx")
        deck.shape("256.29").set_text("12.1%")
        deck.save("out.pptx")
    """

    def __init__(self, package: OoxmlPackage) -> None:
        self.package = package
        self.history = History(package)
        self._slides: list[Slide] | None = None

    # -- lifecycle -------------------------------------------------------------------------

    @classmethod
    def open(cls, source: str | os.PathLike[str] | bytes | BinaryIO) -> "Document":
        """Open a deck from a path, bytes or a binary file.

        A template (``.potx``) opens as the template itself -- with a
        :class:`~pptx_agent.TemplateOpenedWarning`, because a deck *from* a template is
        ``Document.new(template=...)``, which drops the template's sample slides.  Saved as
        ``.pptx``, an opened template becomes a presentation (:meth:`save`).

        For example::

            deck = Document.open("results.pptx")
        """
        document = cls(OoxmlPackage.open(source))
        from .deck import TEMPLATE_TYPES, TemplateOpenedWarning, main_content_type

        if main_content_type(document.package) in TEMPLATE_TYPES:
            import warnings

            name = os.fspath(source) if isinstance(source, (str, os.PathLike)) else "this file"
            warnings.warn(
                f"{name} is a PowerPoint template: Document.open edits the template itself. "
                "To make a deck from it, use Document.new(template=...); saving this one as "
                ".pptx declares it a presentation.",
                TemplateOpenedWarning, stacklevel=2)
        return document

    @classmethod
    def new(cls, *, size: "str | tuple[int, int] | None" = None,
            template: "str | os.PathLike[str] | bytes | None" = None,
            title: str | None = None, author: str | None = None,
            language: str | None = None, created: "datetime | None" = None) -> "Document":
        """A new deck with no slides; add them with :meth:`add_slide`.

        Without a ``template`` it is the deck PowerPoint makes for File > New: the Office
        theme, one master and its eleven layouts (Title Slide, Title and Content, Section
        Header, Two Content, Comparison, Title Only, Blank, Content with Caption, Picture
        with Caption, Title and Vertical Text, Vertical Title and Text), at ``size`` --
        ``"16:9"`` (the default), ``"4:3"`` or ``(width, height)`` in EMU, laid out for it
        as PowerPoint lays its master out for another size.  ``language`` is the default
        text language (``"en-US"``).

        With a ``template`` -- a ``.pptx`` or ``.potx``, as a path or bytes -- the deck
        starts from its masters, layouts and theme, without its slides; ``size`` and
        ``language``, when given, are then applied as :meth:`set_slide_size` and
        :attr:`language` would.

        ``title`` and ``author`` go into the core properties; ``created`` (default: now)
        is both the creation and modification time.  Creating the deck is the base state:
        there is nothing to undo.

        For example::

            deck = Document.new(size="16:9", title="Plan")    # or template="brand.potx"
        """
        from . import deck

        return deck.new(cls, size=size, template=template, title=title, author=author,
                        language=language, created=created)

    def save(self, target: str | os.PathLike[str]) -> None:
        """Write the deck.  The extension decides what the file declares itself to be:
        ``.pptx`` (and ``.pptm``) a presentation, ``.potx`` (and ``.potm``) a template --
        PowerPoint refuses a file whose extension and declared type disagree, so a template
        opened with :meth:`open` and saved as ``.pptx`` is saved as a presentation.  A
        macro-enabled deck saves only as ``.pptm`` or ``.potm``.  A path with another
        extension keeps the type the deck has.

        For example::

            deck.save("out.pptx")
        """
        from .deck import target_kind

        kind = target_kind(target)
        macro, template = (None, None) if kind is None else kind
        self.package.save(target, self._replacements(template=template, macro=macro))

    def save_as_template(self, target: str | os.PathLike[str]) -> None:
        """Write the deck as a PowerPoint template: the same parts, with the main part
        declared a template.  The path ends in ``.potx`` (``.potm`` for a macro-enabled
        deck); a ``.pptx`` path is refused, since PowerPoint would refuse the file.

        For example::

            deck.save_as_template("brand.potx")
        """
        from .deck import target_kind

        kind = target_kind(target)
        if kind is not None and not kind[1]:
            raise ValueError(f"a template is saved as .potx (or .potm), not "
                             f"{os.path.splitext(os.fspath(target))[1]}; use save() for a deck")
        self.package.save(target, self._replacements(
            template=True, macro=None if kind is None else kind[0]))

    def to_bytes(self, *, template: bool | None = None) -> bytes:
        """The deck's bytes, as :meth:`save` would write them: ``template=True`` declares it a
        template (``.potx``), ``False`` a presentation (``.pptx``), ``None`` (the default)
        keeps what it is::

            Document.open(deck.to_bytes())        # a fresh copy
        """
        return self.package.to_bytes(self._replacements(template=template))

    def validate(self, target: "str | os.PathLike[str] | None" = None) -> list:
        """The structural problems PowerPoint would repair or refuse
        (:mod:`pptx_agent.validate`): parts that do not parse, content types,
        relationships that dangle or that the XML names but the part lacks, orphaned
        parts, the slide list and its sections, schema order in slides, layouts and
        masters, table grids.  An empty list means none were found::

            problems = deck.validate()
            assert not problems, "\n".join(map(str, problems))

        With a ``target`` -- the file name it is going to -- the main part's declared type
        is checked against the extension too (``package-type``): a template's in a
        ``.pptx``, a presentation's in a ``.potx``, a macro-enabled deck's in either.
        :meth:`save` writes the right type for its target, so the problem only stands for
        bytes written another way -- or, for a macro-enabled deck, refused outright::

            deck.validate(target="out.pptx")

        Each is a :class:`pptx_agent.validate.Problem` (``code``, ``part``, ``detail``).
        No edit adds one; a deck may arrive with some, so compare before and after.  Only
        PowerPoint can say for certain that a deck opens.
        """
        from ..validate import check, package_type_problems

        problems = check(self.package)
        if target is not None:
            problems = sorted(set(problems) | set(package_type_problems(self.package, target)))
        return problems

    def _replacements(self, *, template: bool | None,
                      macro: bool | None = None) -> dict[str, bytes]:
        from . import deck

        return deck.write_replacements(self, template=template, macro=macro)

    # -- metadata --------------------------------------------------------------------------

    @property
    def title(self) -> str | None:
        """``dc:title`` in ``docProps/core.xml``.

        For example::

            deck.title = "Q3 results"
        """
        return self._core().get("dc:title")

    @title.setter
    def title(self, value: str | None) -> None:
        self._set_core("dc:title", value)

    @property
    def author(self) -> str | None:
        """``dc:creator`` in ``docProps/core.xml``.

        For example::

            deck.author = "Finance"
        """
        return self._core().get("dc:creator")

    @author.setter
    def author(self, value: str | None) -> None:
        self._set_core("dc:creator", value)

    @property
    def created(self) -> "datetime | None":
        """When the deck was created, as its core properties say.

        For example::

            deck.created = datetime(2026, 1, 30, tzinfo=timezone.utc)
        """
        from .properties import parse_w3cdtf

        return parse_w3cdtf(self._core().get("dcterms:created"))

    @created.setter
    def created(self, value: "datetime | None") -> None:
        self._set_core("dcterms:created", value)

    @property
    def modified(self) -> "datetime | None":
        """When the deck was last modified, as its core properties say.  Saving does not
        change it: set it when that is wanted (``deck.modified = datetime.now()``).

        For example::

            deck.modified = datetime.now(timezone.utc)
        """
        from .properties import parse_w3cdtf

        return parse_w3cdtf(self._core().get("dcterms:modified"))

    @modified.setter
    def modified(self, value: "datetime | None") -> None:
        self._set_core("dcterms:modified", value)

    def _core(self):
        from .properties import CoreProperties

        return CoreProperties(self.package)

    def _set_core(self, tag: str, value) -> None:
        from datetime import datetime as _datetime

        from .properties import w3cdtf

        dated = tag.startswith("dcterms:")
        if value is not None and dated:
            if not isinstance(value, _datetime):
                raise TypeError(f"{tag} is a datetime, not {type(value).__name__}")
            value = w3cdtf(value)
        elif value is not None and not isinstance(value, str):
            raise TypeError(f"{tag} is a string, not {type(value).__name__}")
        if self._core().get(tag) == value:
            return
        with self.batch():
            self.history.checkpoint()
            self._core().set(tag, value, dated=dated)

    @property
    def language(self) -> str:
        """The default text language (``"en-US"``): what new text and new shapes take.

        For example::

            deck.language = "ja-JP"
        """
        from .authoring import default_language

        return default_language(self.package)

    @language.setter
    def language(self, value: str) -> None:
        from . import deck

        if not isinstance(value, str) or not value.strip():
            raise ValueError("a language is a tag like 'en-US'")
        with self.batch():
            self.history.checkpoint()
            deck.set_language(self, value.strip())

    @property
    def theme(self) -> "Theme":
        """The deck's theme (its first master's): ``colors`` -- every scheme name to
        ``"#RRGGBB"`` -- and ``fonts`` -- the major (headings) and minor (body) faces.  A
        slide whose master differs has its own: :attr:`Slide.theme`.

        For example::

            deck.theme.colors["accent1"]                 # '#4472C4'
            deck.theme.fonts.major, deck.theme.fonts.minor
        """
        from .theme import theme_for

        return theme_for(self)

    @property
    def slide_size(self) -> tuple[int, int]:
        """``(width, height)`` of the slides, EMU.

        For example::

            width, height = deck.slide_size              # (12192000, 6858000) for 16:9
        """
        root = self.package.tree(self.package.presentation_part())
        node = find(root, "p:sldSz") if root is not None else None
        if node is None:
            return (9144000, 6858000)  # the schema's default
        return get_int(node, "cx", 9144000), get_int(node, "cy", 6858000)

    @slide_size.setter
    def slide_size(self, size: "str | tuple[int, int]") -> None:
        self.set_slide_size(size)

    def set_slide_size(self, size: "str | tuple[int, int]", *,
                       scale: bool = True) -> "Document":
        """Change the slide size: ``"16:9"``, ``"4:3"`` or ``(width, height)`` in EMU.

        With ``scale`` (the default) the deck is laid out for it as PowerPoint's "Ensure
        Fit" does (measured): masters and layouts are stretched axis by axis, slide content
        is scaled uniformly by the smaller ratio and centred, and text sizes, indents and
        spacing scale by the smaller ratio.  Without, only the size changes.  One undo
        step; returns the deck.

        For example::

            deck.set_slide_size("4:3")                   # shapes scaled to fit
        """
        from . import blank, deck

        width, height = blank.slide_size(size)
        if (width, height) == self.slide_size:
            return self
        with self.batch():
            self.history.checkpoint()
            deck.resize(self, width, height, scale=scale)
            self._reset_caches()
        return self

    # -- slides ----------------------------------------------------------------------------

    @property
    def slides(self) -> list[Slide]:
        """Every slide, in presentation order.

        For example::

            for slide in deck.slides:
                print(slide.slide_id, len(slide.shapes))
        """
        if self._slides is None:
            self._slides = [
                Slide(self, slide_id, path) for slide_id, path in self.package.slide_parts()
            ]
        return self._slides

    def slide(self, slide_id: "int | str | Slide") -> Slide:
        """A slide by its ``sldId`` (the first part of its shapes' ids): ``257``, or
        ``"s:257"`` as :meth:`to_outline` writes it.  Not a position -- for that,
        ``deck.slides[i]`` (0-based).

        For example::

            deck.slide(257)                              # by sldId: shapes on it are 257.<n>
            deck.slide("s:257")                          # the same slide
        """
        wanted = _slide_id_of(slide_id)
        for slide in self.slides:
            if slide.slide_id == wanted:
                return slide
        ids = [slide.slide_id for slide in self.slides]
        hint = ""
        if 0 <= wanted < len(ids):
            hint = (f"; {wanted} looks like a position: deck.slides[{wanted}] is slide "
                    f"s:{ids[wanted]}")
        raise KeyError(f"no slide with id {wanted} (the slide ids are {ids}){hint}")

    # -- slide structure -------------------------------------------------------------------

    def slide_titled(self, title: str) -> Slide:
        """The one slide whose title (:attr:`Slide.title`) is ``title`` -- whitespace
        collapsed; failing that, case and full-width forms ignored, as a table label is
        matched.  No slide, or more than one, raises :class:`~pptx_agent.LabelError`
        listing the titles there are.

        For example::

            budget = deck.slide_titled("Budget")
        """
        from .labels import find_label

        return find_label(title, ((slide, slide.title) for slide in self.slides),
                          what="title", where="slide titles")

    @property
    def layouts(self) -> list[Layout]:
        """Every slide layout, master by master.

        For example::

            [layout.name for layout in deck.layouts]
        """
        return _slides.layouts(self.package)

    def layout(self, name: str) -> Layout:
        """A layout by name (``"Title and Content"``) or part path.

        For example::

            deck.layout("Title and Content")
        """
        for layout in self.layouts:
            if name in (layout.name, layout.part_path):
                return layout
        raise KeyError(f"no layout {name!r}; have {[layout.name for layout in self.layouts]}")

    def add_slide(self, layout: "Layout | str | None" = None, *, index: int | None = None) -> Slide:
        """A new slide from ``layout``, at 0-based ``index`` (default: last).

        The slide gets the layout's placeholders -- title, body, picture and so on, not the
        date, footer and slide number, which PowerPoint also leaves off -- empty, so they
        inherit position and formatting from the layout.  Without a layout, the last slide's
        is used (or the first layout of the first master).

        For example::

            slide = deck.add_slide("Title and Content", index=1)
        """
        if layout is None:
            last = self.slides[-1].layout if self.slides else None
            layout = last or (self.layouts[0] if self.layouts else None)
            if layout is None:
                raise ValueError("the presentation has no slide layouts")
        elif isinstance(layout, str):
            layout = self.layout(layout)
        package = self.package
        with self.batch():
            self.history.checkpoint()
            part = package.unused_part_name("ppt/slides/slide{n}.xml")
            package.add_part(part, _slides.slide_from_layout(package, layout.part_path),
                             CT_SLIDE, override=True)
            package.add_relationship(part, REL_SLIDE_LAYOUT, layout.part_path)
            slide_id = _slides.insert_entry(package, part, index)
            self._reset_caches()
        return self.slide(slide_id)

    def delete_slide(self, slide: "Slide | int | str") -> None:
        """Remove a slide and every trace of it.

        Gone are its ``sldIdLst`` entry, its section and custom-show entries, the
        presentation's relationship to it, its part and relationships part, its content-type
        override, and whatever only it used -- its notes slide, its charts and their embedded
        workbooks, media no other slide shows.  Hyperlinks elsewhere that jumped to it are
        removed, as PowerPoint does.  Undo puts all of it back.

        For example::

            deck.delete_slide(deck.slides[-1])          # or deck.delete_slide("s:263")
        """
        slide = self._as_slide(slide, "delete_slide")
        package = self.package
        with self.batch():
            self.history.checkpoint()
            rel_id = _slides.remove_entry(package, slide.slide_id)
            going_too = _slides.reachable_from(package, slide.part_path)
            _slides.unlink_references(package, slide.part_path,
                                      keep=going_too | {package.presentation_part()})
            package.release(package.presentation_part(), [rel_id])
            self._reset_caches()

    def duplicate_slide(self, slide: "Slide | int | str", *, index: int | None = None,
                        notes: str | None = None) -> Slide:
        """Copy a slide; the copy goes right after the original unless ``index`` is given.

        Parts are shared where sharing is safe -- the layout, media, linked slides -- and
        copied where they are per-slide state: the notes slide, charts with their embedded
        workbooks, diagrams, OLE objects.  The copy has a new ``sldId``, so all its shape ids
        are new too.  ``notes``, when given, replaces the copy's speaker notes (as
        :attr:`Slide.notes` is set); the copy and its notes are one undo step.

        For example::

            copy = deck.duplicate_slide(deck.slides[0], notes="Thank the sponsors.")
        """
        slide = self._as_slide(slide, "duplicate_slide")
        position = slide.index + 1 if index is None else index
        package = self.package
        with self.batch():
            self.history.checkpoint()
            part = package.copy_part(slide.part_path,
                                     share=lambda rel: rel.type in SHARED_ON_DUPLICATE)
            slide_id = _slides.insert_entry(package, part, position)
            _slides.renew_creation_id(package, part, slide_id)
            drop_comments(package, part)           # the review is of the original
            self._reset_caches()
            if notes is not None:
                self.slide(slide_id).set_notes(notes)
        return self.slide(slide_id)

    def move_slide(self, slide: "Slide | int | str", index: int | None = None, *,
                   to: int | None = None) -> Slide:
        """Move a slide to a new 0-based position; returns it.  Ids do not change.

        ``slide`` says *which* slide -- a :class:`Slide`, ``"s:<sldId>"`` or an ``sldId``
        (``257``), never a position -- and ``to`` (or ``index``, positionally) *where*: the
        0-based position it will have.  Passing the slide as an object or ``"s:"`` id and
        the position as ``to=`` leaves nothing to mix up.

        For example::

            deck.move_slide(deck.slides[-1], to=0)       # the last slide to the front
            deck.move_slide("s:259", to=2)
        """
        if index is not None and to is not None and index != to:
            raise TypeError("move_slide takes the new position once: index or to=, not both")
        position = to if to is not None else index
        if position is None:
            raise TypeError("move_slide needs the new 0-based position: move_slide(slide, to=n)")
        if isinstance(position, bool) or not isinstance(position, int):
            raise TypeError(f"the new position is a 0-based int, not {position!r}")
        slide = self._as_slide(slide, "move_slide")
        count = len(self.slides)
        if not 0 <= position < count:
            raise IndexError(f"slide position {position} out of range 0..{count - 1} "
                             "(positions are 0-based, as in deck.slides)")
        if slide.index == position:
            return slide
        with self.batch():
            self.history.checkpoint()
            _slides.move_entry(self.package, slide.slide_id, position)
            self._reset_caches()
        return self.slide(slide.slide_id)

    def _as_slide(self, slide: "Slide | int | str", method: str) -> Slide:
        """``slide`` as a :class:`Slide`: itself, or the slide an ``sldId`` or ``"s:<sldId>"``
        names -- with an error that says a position is not an id."""
        if isinstance(slide, Slide):
            return slide
        try:
            return self.slide(slide)
        except KeyError as error:
            raise KeyError(f"{method}: the slide is a Slide, 's:<sldId>' or an sldId, not a "
                           f"position -- {error.args[0]}") from None

    def shape(self, shape_id: str | ShapeId) -> Shape:
        """Resolve a deck-wide shape id without knowing which slide it is on.

        For example::

            table = deck.shape("257.3#5").table
        """
        parsed = ShapeId.parse(shape_id)
        return self.slide(parsed.slide_id).shape(parsed)

    def __iter__(self) -> Iterator[Slide]:
        return iter(self.slides)

    def resolve(self, address: str) -> "Shape | TableCell | Paragraph | Run | DiagramNode | TextFrame":
        """Look up any addressable thing by its address string.

        ``"256.5"`` is a shape, ``"256.5/p1"`` its second paragraph, ``"256.5/p1/r0"`` that
        paragraph's first run.  Table cells insert ``cell<row>,<column>``:
        ``"256.7/cell2,0/p0/r1"``; a SmartArt node is ``"258.3/node<k>"`` (the k-th of
        ``diagram.nodes``); a slide's speaker notes are ``"256/notes"`` (a
        :class:`TextFrame`; ``"256/notes/p0"`` its first paragraph).  Every address is
        re-resolved from the current document, so it names a *position* -- after an
        insertion, the same address may name a different run.  Each has a ``text`` to read
        and set::

            deck.resolve("257.3#5/cell1,1").text = "4,310億円"
            deck.resolve("257/notes").text = "Lead with the margin."
        """
        head, *steps = address.split("/")
        if steps and steps[0] == "notes" and "." not in head:
            target: object = self.slide(head).notes_frame
            steps = steps[1:]
        else:
            target = self.shape(head)
        for step in steps:
            target = _step(target, step, address)
        return target  # type: ignore[return-value]

    # -- history ---------------------------------------------------------------------------

    @contextmanager
    def batch(self) -> Iterator[None]:
        """Collapse every edit inside into one undo step; roll all of them back on error.

        For example::

            with deck.batch():
                deck.shape("256.13").set_text("4,310")
                deck.shape("256.17").set_text("520")
        """
        try:
            with self.history.batch():
                yield
        except Exception:
            # The rollback reparsed the edited parts, so cached elements are stale.
            if not self.history.in_batch:
                self._reset_caches()
            raise

    def undo(self) -> bool:
        """Undo the last edit or batch; ``False`` when there is nothing to undo.

        For example::

            deck.undo()
        """
        changed = self.history.undo()
        self._reset_caches()
        return changed

    def redo(self) -> bool:
        """Redo what :meth:`undo` undid; ``False`` when there is nothing to redo.

        For example::

            deck.redo()
        """
        changed = self.history.redo()
        self._reset_caches()
        return changed

    def _reset_caches(self) -> None:
        # restore() reparsed the trees, so every cached Element reference is stale.
        self._slides = None

    # -- the full-state SVG ----------------------------------------------------------------

    def apply_svg(self, svg: "str | bytes", *, add_new: bool = True,
                  delete_missing: bool = False, limits=None):
        """Read a full-state SVG (``Slide.render_svg(full_state=True)``) back in.

        The slide is the one the SVG names.  Typed ``data-ooxml-*`` attributes that differ
        from the document are applied as the E1/E2 edits they correspond to; a changed
        ``data-ooxml-xml`` replaces the shape's XML; a group with XML but an id the slide does
        not know is added (``add_new``); and with ``delete_missing=True`` -- off by default,
        because it is destructive -- shapes absent from the SVG are deleted.  All of it is one
        undo step, and an SVG that matches the document changes nothing at all.

        The SVG is untrusted input: an SVG not written by this library, or one with a DTD,
        an entity, invalid base64 or XML, or beyond the size ``limits``, is refused before
        anything changes.  Returns an :class:`~pptx_agent.fullstate.ApplyReport`.

        For example::

            report = deck.apply_svg(edited_svg)          # report.edited, .added, ...
        """
        from ..fullstate import apply_full_state

        return apply_full_state(self, svg, add_new=add_new, delete_missing=delete_missing,
                                limits=limits or _default_limits())

    # -- the Markdown outline --------------------------------------------------------------

    def to_outline(self, slides=None, *, ids: bool = True, notes: bool = True) -> str:
        """The deck as a compact Markdown reading view; it changes no byte.

        Each slide is a ``#`` heading -- its title placeholder's text -- under a comment
        with its id and layout (``<!-- s:256 layout: Title and Content -->``), the title's
        shape id at the heading's end.  Then each placeholder and text shape is a block
        under a comment with its shape id (``<!-- 256.3 ph: obj 1 -->``): bullets as lists
        that keep their levels, bold, italic, strikethrough, code and links as Markdown
        marks; a table as a GFM table (its cells are ``<id>/cell<row>,<column>``); a picture
        as ``![alt text](media part)``; a chart as one line (``[chart: column; series: ...;
        categories: ...]``); SmartArt as a list of its node text.  Speaker notes follow
        ``Notes:``, under a comment with their address (``<!-- 256/notes -->``).  See :mod:`pptx_agent.outline` for the whole format.

        ``slides``: ``None`` for all, or 1-based slide numbers (as :meth:`render_svg`
        takes), :class:`Slide` objects or ``"s:<sldId>"`` ids.  ``ids=False`` writes plain
        Markdown, the blocks of a slide apart by ``---``: a separator before every block
        but the slide's first, so it depends on the slide's structure only, never on its
        text.  ``notes=False`` leaves the speaker notes out.

        **The text is Markdown, escaped: do not feed it to** ``set_text``.  Markdown
        punctuation in the text is backslash-escaped so it reads back as text (``a\\*b``,
        ``\\- item``, ``1\\. one``), and bold, italic and code are ``**``, ``*`` and
        backticks.  ``set_text`` writes what it is given, so a string copied from here
        would put those backslashes and asterisks on the slide (``set_text`` warns with a
        :class:`~pptx_agent.MarkdownEscapeWarning` when it sees one).  For text to edit
        from, use :meth:`outline_blocks` (every block's raw text and address),
        :meth:`find_text`, or ``shape.text`` -- or :meth:`insert_outline`, which reads this
        Markdown back::

            print(deck.to_outline(slides=[1]))        # read; then edit with set_text
        """
        from ..outline.read import to_outline

        return to_outline(self, slides, ids=ids, notes=notes)

    def overflows(self, slides=None, *, boxes: bool = False) -> list:
        """What will not look right in PowerPoint, slide by slide
        (:class:`~pptx_agent.Overflow`): text taller than its shape (``"text"``, with its
        :meth:`Shape.text_fit`), collisions (``"overlap"``, :meth:`Slide.collisions`) and
        a shape drawn past the slide's edge (``"off_slide"``, by its
        :attr:`Shape.drawn_bounds`).  ``slides`` as :meth:`to_outline`; an empty list
        means none.  ``boxes=True`` adds overlapping text boxes whose text does not
        overlap (see :meth:`Slide.collisions`).  Needs pptx2svg.

        For example::

            problems = deck.overflows()
            assert not problems, "\n".join(map(str, problems))
        """
        _require_renderer()
        from .fit import overflows

        return overflows(self, slides, boxes=boxes)

    def outline_blocks(self, slides=None) -> "list[TextBlock]":
        r"""Every piece of text in the deck, raw, with its address -- the outline's blocks
        as data, for editing rather than reading.

        In the outline's order: per slide its title, then each text shape in z-order
        (groups flattened), each table cell with text (``<id>/cell<row>,<column>``), each
        SmartArt node (``<id>/node<k>``) and last its speaker notes (``<sldId>/notes``,
        kind ``"notes"``).  Each is a :class:`~pptx_agent.TextBlock`
        ``(address, kind, text, slide)`` whose ``text`` is not escaped and carries no
        marks: it is what ``deck.resolve(address).text`` reads (``text_frame.text``, for a
        shape whose text has line breaks) and what its ``text`` setter takes back.
        ``slides`` as :meth:`to_outline`::

            for block in deck.outline_blocks(slides=[1]):
                if block.text == "11.9%":
                    deck.resolve(block.address).text = "12.1%"
        """
        from ..outline.read import outline_blocks

        return outline_blocks(self, slides)

    def find_text(self, needle, slides=None) -> "list[TextBlock]":
        r"""The :meth:`outline_blocks` whose text contains ``needle`` -- a substring, or a
        compiled :func:`re.compile` pattern (searched with ``pattern.search``).  Speaker
        notes are searched too (kind ``"notes"``, address ``"<sldId>/notes"``)::

            for block in deck.find_text("4,285"):
                target = deck.resolve(block.address)   # a Shape, TableCell or DiagramNode
                target.text = block.text.replace("4,285", "4,310")   # formatting kept
        """
        from ..outline.read import find_text

        return find_text(self, needle, slides)

    def insert_outline(self, markdown: str, *, at: int | None = None,
                       layout_map: "dict | None" = None, images=None) -> list[Slide]:
        """Draft slides from Markdown; returns them.  One undo step.

        ``#`` starts a slide and its text goes to the title placeholder.  The layout is the
        deck's own, found by type and then name (so a localised template works), chosen
        by what the slide holds: nothing but the heading -- Section Header, or Title Slide
        when it will be the first slide; a heading and text -- Title and Content (Title
        Slide, the text its subtitle, when first and not a list); two columns (``---``
        between them, or two lists) -- Two Content; tables or pictures and no text -- Title
        Only.  A layout the outline names (``<!-- layout: Comparison -->`` above the
        heading, as :meth:`to_outline` writes) wins, and ``layout_map`` overrides either:
        keys are the roles (``"title"``, ``"secHead"``, ``"obj"``, ``"twoObj"``,
        ``"titleOnly"``, ``"blank"``, or their English names) or layout names, values a
        layout of this deck (name, type or :class:`Layout`).

        Lists fill the body placeholder with their levels, and bold, italic,
        strikethrough, code and links become runs; a GFM table becomes a table, an image a
        picture (in a picture placeholder when the layout has a free one); text after a
        ``Notes:`` line becomes the speaker notes.  ``images`` finds the pictures: a
        directory, a mapping of source to bytes or path, a callable, or the
        :class:`Document` an outline was read from (its media parts); by default a source is
        a path.  ``at`` is the 0-based position of the first new slide (default: last).
        Charts and SmartArt are not drafted (an :class:`~pptx_agent.outline.OutlineWarning`
        says so).  See :mod:`pptx_agent.outline`.

        For example::

            slides = deck.insert_outline(Path("outline.md").read_text())
        """
        from ..outline.draft import insert_outline

        return insert_outline(self, markdown, at=at, layout_map=layout_map, images=images)

    # -- rendering -------------------------------------------------------------------------

    def render_svg(self, slides=None, **options) -> list[str]:
        """Render via pptx2svg.  Requires ``pip install pptx-agent[render]``.  ``slides`` as
        :meth:`to_outline` takes them: 1-based numbers, :class:`Slide` objects or
        ``"s:<sldId>"`` ids.

        For example::

            svgs = deck.render_svg([1, 2])               # 1-based slide numbers
            svgs = deck.render_svg(["s:257"])            # or by id

        ``agent=True`` gives pptx2svg's compact agent view (see :meth:`Slide.render_svg`),
        with pptx2svg's own ids.
        """
        pptx2svg = _require_renderer()
        if options.pop("agent", False):
            convert = getattr(pptx2svg, "convert_pptx_to_agent_svg", None)
            if convert is None:
                raise RuntimeError("this pptx2svg has no agent view "
                                   "(convert_pptx_to_agent_svg); upgrade pptx2svg")
            return convert(self.to_bytes(),
                           pptx2svg.ConvertOptions(slide_numbers=self._numbers(slides)))
        convert_options = pptx2svg.ConvertOptions(slide_numbers=self._numbers(slides),
                                                  **options)
        return pptx2svg.convert_pptx_to_svg(self.to_bytes(), convert_options)

    def render_png(self, slides=None, **options) -> list[bytes]:
        """Render to PNG via pptx2svg.  Requires ``pip install pptx-agent[png]``.  ``slides``
        as :meth:`render_svg`.

        For example::

            pngs = deck.render_png([1], width=1280)
        """
        pptx2svg = _require_renderer()
        render_keys = {"width", "height", "font_mapping", "measurer"}
        convert_options = pptx2svg.ConvertOptions(
            slide_numbers=self._numbers(slides),
            **{k: v for k, v in options.items() if k in render_keys}
        )
        return pptx2svg.convert_pptx_to_png(
            self.to_bytes(),
            convert_options,
            **{k: v for k, v in options.items() if k not in render_keys},
        )

    def _numbers(self, slides) -> "list[int] | None":
        """1-based slide numbers from what :meth:`to_outline` accepts."""
        if slides is None:
            return None
        from ..outline.read import select

        return [slide.index + 1 for slide in select(self, slides)]

    def __repr__(self) -> str:
        return f"<Document {len(self.slides)} slides>"


def _slide_id_of(slide: object) -> int:
    """An ``sldId`` from a :class:`Slide`, an ``int``, ``"s:257"`` or ``"257"``."""
    if hasattr(slide, "slide_id"):
        return slide.slide_id  # type: ignore[union-attr]
    if isinstance(slide, bool):
        raise TypeError(f"not a slide: {slide!r}")
    if isinstance(slide, int):
        return slide
    if isinstance(slide, str):
        text = slide.strip()
        text = text[2:] if text.startswith("s:") else text
        if text.isdigit():
            return int(text)
        raise ValueError(f"a slide is a Slide, an sldId or 's:<sldId>', not {slide!r}")
    raise TypeError(f"not a slide: {slide!r}; pass a Slide, an sldId or 's:<sldId>'")


def _step(target: object, step: str, address: str) -> object:
    import re

    cell = re.fullmatch(r"cell(\d+),(\d+)", step)
    if cell is not None and isinstance(target, Shape):
        return target.table.cell(int(cell.group(1)), int(cell.group(2)))
    paragraph = re.fullmatch(r"p(\d+)", step)
    if paragraph is not None and isinstance(target, (Shape, TableCell)):
        return target.text_frame.paragraph(int(paragraph.group(1)))
    if paragraph is not None and isinstance(target, TextFrame):
        return target.paragraph(int(paragraph.group(1)))
    run = re.fullmatch(r"r(\d+)", step)
    if run is not None and isinstance(target, Paragraph):
        return target.run(int(run.group(1)))
    node = re.fullmatch(r"node(\d+)", step)
    if node is not None and isinstance(target, Shape):
        nodes = target.diagram.nodes
        k = int(node.group(1))
        if k >= len(nodes):
            raise IndexError(f"{target.id} has {len(nodes)} SmartArt nodes; no node{k}")
        return nodes[k]
    raise KeyError(f"cannot resolve {step!r} in address {address!r}")


class _Position(int):
    """A slide's position: an ``int`` that may also be called, so ``slide.index`` and
    ``slide.index()`` agree (trial 2's p3 called it)."""

    def __call__(self) -> int:
        return int(self)

    def __repr__(self) -> str:
        return repr(int(self))


class ImageSize(NamedTuple):
    """A picture's image as stored: ``width`` and ``height`` in pixels, and the resolution
    it states, ``dpi`` ``(x, y)``, or ``None`` when it states none.

    For example::

        size = picture.image_size
        size.width / size.height == size.aspect
    """

    width: int
    height: int
    dpi: tuple[float, float] | None

    @property
    def aspect(self) -> float:
        """Width over height.

        For example::

            new_width = round(picture.height * picture.image_size.aspect)
        """
        return self.width / self.height


#: Where a frame's fixed point is, as fractions of its width and height.
_ANCHORS = {"top_left": (0, 0), "top": (0.5, 0), "top_right": (1, 0), "left": (0, 0.5),
            "center": (0.5, 0.5), "right": (1, 0.5), "bottom_left": (0, 1),
            "bottom": (0.5, 1), "bottom_right": (1, 1)}

#: ``a:ext`` URI of an SVG alternative to a bitmap blip (``asvg:svgBlip``).
_SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"


def _read_image(image: "bytes | str | os.PathLike[str]") -> tuple[bytes, str | None]:
    """An image's bytes, and the extension its path gave, if it came from a file."""
    if isinstance(image, (str, os.PathLike)):
        with open(os.fspath(image), "rb") as handle:
            data = handle.read()
        return data, os.path.splitext(os.fspath(image))[1].lstrip(".") or None
    return bytes(image), None


def _children_of(element: Element | None) -> list[Element]:
    return [] if element is None else list(element)


def _relationship_references(root: Element) -> set[str]:
    """Values of every relationship-namespace attribute (``r:id``, ``r:embed``...) in a part."""
    return set(_R_ATTRIBUTES(root, ns=NAMESPACES["r"]))


#: Compiled once: every relationship-namespace attribute value, collected in C.
_R_ATTRIBUTES = etree.XPath("//@*[namespace-uri() = $ns]")


def _default_limits():
    from ..fullstate.safe import DEFAULT_LIMITS

    return DEFAULT_LIMITS


def _require_renderer():
    try:
        import pptx2svg
    except ImportError as error:  # pragma: no cover - exercised by the extras-less install
        raise ImportError(
            "rendering needs pptx2svg: pip install 'pptx-agent[render]' "
            "(or 'pptx-agent[png]' for PNG output)"
        ) from error
    return pptx2svg


def _nv_container(shape: Element) -> str:
    return {
        "sp": "p:nvSpPr",
        "pic": "p:nvPicPr",
        "cxnSp": "p:nvCxnSpPr",
        "grpSp": "p:nvGrpSpPr",
        "graphicFrame": "p:nvGraphicFramePr",
    }.get(local_name(shape), "p:nvSpPr")


_SHAPE_TAGS = ("sp", "pic", "cxnSp", "grpSp", "graphicFrame", "contentPart")


def renumber_copy(clone: Element, next_ids) -> dict[str, str]:
    """Give every shape in a copied subtree a fresh ``cNvPr@id`` (``next_ids(n)`` hands out
    ``n``), and glue each connector inside it that was glued to a shape inside it to that
    shape's copy.  Returns ``{old id: new id}``."""
    properties = [node for node in clone.iter(qn("p:cNvPr"))
                  if local_name(node.getparent().getparent()) in _SHAPE_TAGS]
    fresh = next_ids(len(properties))
    mapping: dict[str, str] = {}
    for node, number in zip(properties, fresh):
        old = node.get("id")
        if old is not None and old not in mapping:
            mapping[old] = str(number)
        node.set("id", str(number))
    for tag in ("a:stCxn", "a:endCxn"):
        for glue in clone.iter(qn(tag)):
            target = glue.get("id")
            if target in mapping:
                glue.set("id", mapping[target])
    return mapping


def _strip_identity(shape: Element) -> None:
    """Remove stamped ids and creationIds from a copied subtree, including nested shapes."""
    from ..oxml.xml import CREATION_ID_EXT_URI, PA_ID_EXT_URI

    for properties in shape.iter(qn("p:cNvPr")):
        ext_list = properties.find(qn("a:extLst"))
        if ext_list is None:
            continue
        for extension in list(ext_list.findall(qn("a:ext"))):
            if extension.get("uri") in {PA_ID_EXT_URI, CREATION_ID_EXT_URI}:
                remove(extension)
        if len(ext_list) == 0:
            remove(ext_list)
