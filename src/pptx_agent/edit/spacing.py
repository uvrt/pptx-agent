"""The tightest spacings on a slide, measured: :meth:`Slide.spacing <pptx_agent.Slide.spacing>`.

Measurements only, nearest first, as ``check`` reports overlaps: what a gap *means* -- a
heading crowding the panel above it, chips pressed to their card's edge, or a tight row on
purpose -- is the application's to say.  Four kinds, each below a threshold the caller sets
(``below``, default 8 pt):

* **between** -- two neighbours whose drawn extents face each other on one axis and overlap
  on the other: the gap between them.  Only each shape's nearest neighbour below it and to
  its right (and above and to its left), and only neighbours inside the same container (or
  both on the slide), so a chip in a card is not measured to a shape beyond the card.
* **inside** -- content inside a filled or outlined shape behind it (a card, panel, badge,
  chip): the padding from the content to each of the container's edges, the nearest
  content per edge (``also`` the others within half a point of it).  The innermost
  container only.
* **clearance to the frame** -- the nearest content to each footer, slide-number and date
  placeholder the slide holds (their boxes); and content that passes the edge of the
  layout's content area (:meth:`Slide.content_area <pptx_agent.Slide.content_area>`), its
  ``gap`` negative: how far past.  Placeholders where their layout puts them (no position
  of their own) are the area's own and not measured to it.

**What is drawn, not the boxes.**  Shapes are measured as ``check``'s collisions measure
them (:mod:`.fit`): rotation and groups resolved (:attr:`Shape.drawn_bounds
<pptx_agent.Shape.drawn_bounds>`), and text that has no fill or outline of its own by where
its lines are drawn -- pptx2svg's layout, the overflow measurement's: the line boxes'
width, the top of the first and about the ink's bottom of the last -- not its text box.
A filled or outlined shape, a picture, a table or a chart is its box.  Without pptx2svg
every shape is its box (:attr:`SpacingReport.drawn` is false).

**Left out**: pairs that overlap (``check`` reports overlaps as collisions), lines and
connectors, shapes that draw nothing (an empty unfilled box, an empty placeholder), a
shape covering the whole slide to within 2% of each side (a decorative background), and
a shape's own text to its own edges (its insets).  Alignment near-misses are the design
facts' (:meth:`Slide.design_facts <pptx_agent.Slide.design_facts>`, ``alignment``).

For example::

    report = slide.spacing(below=Pt(8))
    for entry in report.entries:
        print(entry.kind, entry.shapes, entry.where, entry.gap / 12700)
    # between ('257.6', '257.34') y 4.0
    # inside ('257.25', '257.41') bottom 4.0
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover
    from .document import Slide

EMU_PER_POINT = 12700
#: The default threshold: gaps below 8 pt.
DEFAULT_BELOW = 8 * EMU_PER_POINT
#: Below this (EMU, 0.1 pt) two edges touch: a gap this small is no gap, an overlap this
#: small no overlap -- the collision measurement's noise (:data:`.fit._TOLERANCE`).
_TOUCH = 1270
#: Content this close (EMU, half a point) to an edge's nearest is listed with it (``also``).
_TIE = 6350
#: A shape covering the slide to within this share of its width and height on every side
#: is a background.
_BLEED = 0.02

#: What a frame placeholder's type is called in an entry.
FRAMES = {"ftr": "footer", "sldNum": "slide_number", "dt": "date"}


@dataclass(frozen=True)
class Spacing:
    """One measured spacing, EMU.

    ``kind`` is ``"between"`` (``shapes`` the two neighbours, ``where`` the axis ``"x"`` or
    ``"y"``), ``"inside"`` (``shapes`` the container and the content, ``where`` the
    container's edge, ``also`` other content as near that edge) or ``"clearance"``
    (``shapes`` the content alone, ``frame`` ``"footer"``, ``"slide_number"``, ``"date"``
    or ``"content_area"``, ``where`` the content area's edge it passes).  ``gap`` is the
    distance; negative only past the content area's edge."""

    kind: str
    shapes: tuple[str, ...]
    gap: int
    where: str | None = None
    frame: str | None = None
    also: tuple[str, ...] = ()

    def to_json(self) -> dict[str, Any]:
        gap = round(self.gap / EMU_PER_POINT, 1)
        if self.kind == "between":
            return {"between": list(self.shapes), "axis": self.where, "gap": gap}
        if self.kind == "inside":
            data: dict[str, Any] = {"inside": self.shapes[0], "shape": self.shapes[1],
                                    "edge": self.where, "padding": gap}
            if self.also:
                data["also"] = list(self.also)
            return data
        data = {"clearance": self.frame, "shape": self.shapes[0]}
        if self.where is not None:
            data["edge"] = self.where
        data["gap"] = gap
        return data


@dataclass(frozen=True)
class SpacingReport:
    """:meth:`Slide.spacing <pptx_agent.Slide.spacing>`: the spacings below the threshold,
    nearest first (``entries``); ``drawn`` false when pptx2svg was not there to lay the
    text out and every shape was measured by its box."""

    entries: tuple[Spacing, ...]
    below: int
    drawn: bool = True


@dataclass
class _Box:
    address: str
    order: int
    extent: tuple[float, float, float, float]     # what is drawn: x0, y0, x1, y1
    box: tuple[float, float, float, float]
    visible: bool                                  # filled or outlined: a container
    container: bool
    frame: str | None = None
    own_place: bool = True                         # not a placeholder where its layout puts it
    parent: "_Box | None" = None


def spacing(slide: "Slide", *, below: int = DEFAULT_BELOW) -> SpacingReport:
    """See :meth:`Slide.spacing <pptx_agent.Slide.spacing>`."""
    below = int(below)
    try:
        import pptx2svg  # noqa: F401
    except ImportError:                       # pragma: no cover - the extras-less install
        boxes, drawn = _boxes_alone(slide), False
    else:
        boxes, drawn = _drawn_boxes(slide), True
    width, height = slide.document.slide_size
    shown = [b for b in boxes if not _background(b.box, width, height)]
    frames = [b for b in shown if b.frame is not None]
    content = [b for b in shown if b.frame is None]
    _nest(content)
    found: list[Spacing] = []
    found += _between(content, below)
    found += _padding(content, below)
    found += _frames(frames, content, below)
    found += _area(slide, content)
    found.sort(key=lambda s: (s.gap, s.kind, s.shapes))
    return SpacingReport(tuple(found), below, drawn)


# -- what is drawn -----------------------------------------------------------------------------


def _drawn_boxes(slide: "Slide") -> list[_Box]:
    """The slide's shapes as ``check``'s collisions measure them (:func:`.fit._items`)."""
    from .fit import _items, _outlined

    measured_ref: list = [None]
    out = []
    for item in _items(slide, measured_ref, {}):
        shape = item.shape
        if item.route is not None or shape.kind == "connector":
            continue
        element = None
        if shape.kind == "shape" and measured_ref[0] is not None:
            element, _ = measured_ref[0].element(shape)
        visible = item.opaque or (element is not None and _outlined(element))
        if not visible and not item.text and shape.kind == "shape":
            continue                          # draws nothing: an empty unfilled box
        extent = item.region if item.text and item.region is not None and not visible \
            else item.box
        out.append(_make(shape, item.order, extent, item.box, visible))
    return out


def _boxes_alone(slide: "Slide") -> list[_Box]:    # pragma: no cover - without pptx2svg
    """Every shape by its box, its fill and outline read from its own XML."""
    from .fit import _box, _flatten, _holds_text

    out = []
    for order, shape in enumerate(_flatten(slide.shapes)):
        bounds = shape.drawn_bounds
        if bounds is None or shape.kind == "connector":
            continue
        if shape.kind == "shape" and _is_line(shape):
            continue
        if shape.kind == "shape":
            fill = shape.effective_fill
            visible = (fill is not None and fill.type != "none") or _has_outline(shape)
            if not visible and not _holds_text(shape):
                continue
        else:
            visible = True
        box = _box(bounds)
        out.append(_make(shape, order, box, box, visible))
    return out


def _is_line(shape) -> bool:
    from .design import _LINE_PRESETS

    return getattr(shape, "preset", None) in _LINE_PRESETS


def _has_outline(shape) -> bool:          # pragma: no cover - without pptx2svg
    line = shape._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ln")
    if line is not None:
        return any(child.tag.rpartition("}")[2] in ("solidFill", "gradFill", "pattFill")
                   for child in line)
    ref = shape._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}lnRef")
    return ref is not None and ref.get("idx") not in (None, "0")


def _make(shape, order, extent, box, visible) -> _Box:
    placeholder = shape.placeholder
    frame = FRAMES.get(placeholder.type) if placeholder is not None else None
    own_place = placeholder is None or shape._xfrm(create=False) is not None
    return _Box(shape.id, order, tuple(extent), tuple(box), visible,
                visible and shape.kind in ("shape", "picture"), frame, own_place)


def _background(box, width: int, height: int) -> bool:
    x0, y0, x1, y1 = box
    return (x0 <= width * _BLEED and y0 <= height * _BLEED
            and x1 >= width * (1 - _BLEED) and y1 >= height * (1 - _BLEED))


def _inside(inner, outer) -> bool:
    return (inner[0] >= outer[0] - _TOUCH and inner[1] >= outer[1] - _TOUCH
            and inner[2] <= outer[2] + _TOUCH and inner[3] <= outer[3] + _TOUCH)


def _area_of(box) -> float:
    return (box[2] - box[0]) * (box[3] - box[1])


def _nest(boxes: list[_Box]) -> None:
    """Each shape's innermost container: a filled or outlined shape (or picture) behind it
    whose box holds what it draws."""
    for item in boxes:
        holders = [other for other in boxes
                   if other is not item and other.container and other.order < item.order
                   and _inside(item.extent, other.box)
                   and _area_of(other.box) > _area_of(item.extent)]
        item.parent = min(holders, key=lambda h: _area_of(h.box), default=None)


# -- the kinds ---------------------------------------------------------------------------------


def _facing(a, b) -> tuple[str, float] | None:
    """``(axis, gap)`` when ``a`` and ``b`` face each other on one axis and overlap on the
    other; ``None`` when they overlap or face on neither (diagonal)."""
    over_x = min(a[2], b[2]) - max(a[0], b[0])
    over_y = min(a[3], b[3]) - max(a[1], b[1])
    if over_x > _TOUCH and over_y > _TOUCH:
        return None                           # an overlap: check's collisions
    if over_x > _TOUCH:
        return "y", -over_y
    if over_y > _TOUCH:
        return "x", -over_x
    return None


def _between(boxes: list[_Box], below: int) -> list[Spacing]:
    nearest: dict[tuple, tuple[str, float]] = {}
    for item in boxes:
        best: dict[tuple[str, int], tuple[float, _Box]] = {}
        for other in boxes:
            if other is item or other.parent is not item.parent:
                continue
            facing = _facing(item.extent, other.extent)
            if facing is None:
                continue
            axis, gap = facing
            index = 0 if axis == "x" else 1
            side = 1 if other.extent[index] >= item.extent[index] else -1
            key = (axis, side)
            if key not in best or gap < best[key][0]:
                best[key] = (gap, other)
        for (axis, _), (gap, other) in best.items():
            if gap < below:
                first, second = sorted((item, other), key=lambda b: (b.extent[1], b.extent[0])
                                       if axis == "y" else (b.extent[0], b.extent[1]))
                nearest[(first.address, second.address)] = (axis, gap)
    return [Spacing("between", pair, round(max(0.0, gap)), axis)
            for pair, (axis, gap) in nearest.items()]


_EDGES = ("left", "top", "right", "bottom")


def _padding(boxes: list[_Box], below: int) -> list[Spacing]:
    out = []
    for holder in boxes:
        held = [item for item in boxes if item.parent is holder]
        if not held:
            continue
        x0, y0, x1, y1 = holder.box
        for edge in _EDGES:
            gaps = sorted(((_edge_gap(item.extent, holder.box, edge), item.order, item)
                           for item in held), key=lambda g: (g[0], g[1]))
            gap, _, item = gaps[0]
            if gap >= below:
                continue
            also = tuple(other.address for g, _, other in gaps[1:] if g - gap <= _TIE)
            out.append(Spacing("inside", (holder.address, item.address), round(max(0.0, gap)),
                               edge, also=also))
    return out


def _edge_gap(inner, outer, edge: str) -> float:
    if edge == "left":
        return inner[0] - outer[0]
    if edge == "top":
        return inner[1] - outer[1]
    if edge == "right":
        return outer[2] - inner[2]
    return outer[3] - inner[3]


def _frames(frames: list[_Box], content: list[_Box], below: int) -> list[Spacing]:
    out = []
    for frame in frames:
        best = None
        for item in content:
            facing = _facing(item.extent, frame.box)
            if facing is not None and (best is None or facing[1] < best[0]):
                best = (facing[1], item)
        if best is not None and best[0] < below:
            out.append(Spacing("clearance", (best[1].address,), round(max(0.0, best[0])),
                               frame=frame.frame))
    return out


def _area(slide: "Slide", content: list[_Box]) -> list[Spacing]:
    left, top, width, height = slide.content_area
    area = (left, top, left + width, top + height)
    out = []
    for item in content:
        if item.parent is not None or not item.own_place:
            continue
        over_x = min(item.extent[2], area[2]) - max(item.extent[0], area[0])
        over_y = min(item.extent[3], area[3]) - max(item.extent[1], area[1])
        if over_x <= _TOUCH or over_y <= _TOUCH:
            continue                          # beside the area, not in it: a title, a note
        for edge in _EDGES:
            gap = _edge_gap(item.extent, area, edge)
            if gap < -_TOUCH:
                out.append(Spacing("clearance", (item.address,), round(gap), edge,
                                   frame="content_area"))
    return out


__all__ = ["DEFAULT_BELOW", "FRAMES", "Spacing", "SpacingReport", "spacing"]
