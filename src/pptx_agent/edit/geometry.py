"""Where a shape is drawn on the slide: rotation, flips and groups applied.

``left``/``top``/``width``/``height`` are a shape's *frame* in its parent's space, before
its rotation; :attr:`Shape.slide_bounds` maps that frame through the enclosing groups'
offsets and scales, and no more.  What is *drawn* can be elsewhere: a shape rotated by 90
degrees occupies its frame turned about its centre, and a connector draws its route --
the frame's preset path -- through the frame's flips and rotation.  PowerPoint writes an
elbow connector that leaves its start downwards with ``rot="5400000"`` and its width and
height swapped, so the frame alone can reach far past the slide while the line stays
between its two shapes.

**The transform, as OOXML defines it.**  A point of the shape's own frame (0..cx,
0..cy) is mirrored within the frame by ``flipH``/``flipV``, then turned clockwise by
``rot`` about the frame's centre; inside a group the result is in the group's child space,
which maps onto the group's frame (``chOff``/``chExt`` onto ``off``/``ext``), where the
group's own flips and rotation apply about the group's centre -- and so on outwards.

**A connector's route** is its preset's path: ``straightConnector1`` (and ``line``) the
diagonal, ``bentConnector2``..``5`` the elbow through their adjustments (``adj1`` the
first bend's x, ``adj2`` the second's y, ``adj3`` the third's x, as fractions of the
frame), ``curvedConnector2`` and ``3`` the Béziers ECMA-376 gives them, sampled; the
longer curved connectors are taken along their elbow, which their curves follow.
"""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING

from ..oxml.xml import Element, find, get_int, local_name, qn

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape

Point = tuple[float, float]

_XFRM_PATH = {"sp": "p:spPr/a:xfrm", "pic": "p:spPr/a:xfrm", "cxnSp": "p:spPr/a:xfrm",
              "grpSp": "p:grpSpPr/a:xfrm", "graphicFrame": "p:xfrm"}


def _xfrm(element: Element) -> Element | None:
    path = _XFRM_PATH.get(local_name(element))
    return None if path is None else find(element, path)


def _flag(xfrm: Element | None, name: str) -> bool:
    return xfrm is not None and xfrm.get(name) in {"1", "true"}


def _turn(points: list[Point], frame: tuple[float, float, float, float],
          xfrm: Element | None) -> list[Point]:
    """Points of a frame ``(x, y, cx, cy)`` through its flips and rotation."""
    x, y, cx, cy = frame
    centre_x, centre_y = x + cx / 2, y + cy / 2
    angle = math.radians((get_int(xfrm, "rot", 0) or 0) / 60000) if xfrm is not None else 0.0
    cos, sin = math.cos(angle), math.sin(angle)
    flip_h, flip_v = _flag(xfrm, "flipH"), _flag(xfrm, "flipV")
    out = []
    for px, py in points:
        if flip_h:
            px = 2 * centre_x - px
        if flip_v:
            py = 2 * centre_y - py
        dx, dy = px - centre_x, py - centre_y
        out.append((centre_x + dx * cos - dy * sin, centre_y + dx * sin + dy * cos))
    return out


def _frame(xfrm: Element | None) -> tuple[float, float, float, float] | None:
    if xfrm is None:
        return None
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return (float(get_int(off, "x", 0) or 0), float(get_int(off, "y", 0) or 0),
            float(get_int(ext, "cx", 0) or 0), float(get_int(ext, "cy", 0) or 0))


def to_slide(element: Element, points: list[Point]) -> list[Point]:
    """Points in ``element``'s parent space carried out through every enclosing group --
    its child-space mapping, then its own flips and rotation -- onto the slide."""
    for group in element.iterancestors(qn("p:grpSp")):
        xfrm = _xfrm(group)
        frame = _frame(xfrm)
        if frame is None:
            continue
        ox, oy, ex, ey = frame
        ch_off = xfrm.find(qn("a:chOff"))
        ch_ext = xfrm.find(qn("a:chExt"))
        cox = get_int(ch_off, "x", 0) if ch_off is not None else ox
        coy = get_int(ch_off, "y", 0) if ch_off is not None else oy
        cex = get_int(ch_ext, "cx", 0) if ch_ext is not None else ex
        cey = get_int(ch_ext, "cy", 0) if ch_ext is not None else ey
        sx = ex / cex if cex else 1.0
        sy = ey / cey if cey else 1.0
        mapped = [(ox + (px - (cox or 0)) * sx, oy + (py - (coy or 0)) * sy)
                  for px, py in points]
        points = _turn(mapped, frame, xfrm)
    return points


def frame_corners(shape: "Shape") -> list[Point] | None:
    """The four corners of the shape's frame as drawn on the slide."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if None in (left, top, width, height):
        return None
    frame = (float(left), float(top), float(width), float(height))
    corners = [(left, top), (left + width, top), (left + width, top + height),
               (left, top + height)]
    return to_slide(shape._element, _turn(corners, frame, shape._xfrm(create=False)))


def route(shape: "Shape") -> list[Point] | None:
    """A connector's drawn path on the slide, as a polyline; ``None`` for other shapes."""
    if local_name(shape._element) != "cxnSp" and _line_preset(shape._element) is None:
        return None
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if None in (left, top, width, height):
        return None
    local = local_route(shape._element, float(width), float(height))
    if local is None:
        return None
    frame = (float(left), float(top), float(width), float(height))
    points = [(left + x, top + y) for x, y in local]
    return to_slide(shape._element, _turn(points, frame, shape._xfrm(create=False)))


def _line_preset(element: Element) -> str | None:
    node = find(element, "p:spPr/a:prstGeom")
    preset = None if node is None else node.get("prst")
    if preset in ("line", "straightConnector1") or (
            preset and re.fullmatch(r"(bent|curved)Connector[2-5]", preset)):
        return preset
    return None


def _adjustments(element: Element) -> dict[str, float]:
    values = {}
    for gd in element.iterfind(f"{qn('p:spPr')}/{qn('a:prstGeom')}/{qn('a:avLst')}/{qn('a:gd')}"):
        match = re.fullmatch(r"\s*val\s+(-?\d+)\s*", gd.get("fmla") or "")
        if match and gd.get("name"):
            values[gd.get("name")] = int(match.group(1)) / 100000
    return values


def _bezier(p0: Point, p1: Point, p2: Point, p3: Point, steps: int = 16) -> list[Point]:
    out = []
    for index in range(1, steps + 1):
        t = index / steps
        u = 1 - t
        out.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                    u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return out


def local_route(element: Element, w: float, h: float) -> list[Point] | None:
    """The preset path of a connector's frame, in the frame (before flips and rotation)."""
    preset = _line_preset(element)
    if preset is None:
        return None
    adjust = _adjustments(element)
    a1, a2, a3 = (adjust.get(name, 0.5) for name in ("adj1", "adj2", "adj3"))
    if preset in ("line", "straightConnector1"):
        return [(0.0, 0.0), (w, h)]
    count = int(preset[-1])
    x1, y2, x3 = w * a1, h * a2, w * a3
    elbow = {2: [(0.0, 0.0), (w, 0.0), (w, h)],
             3: [(0.0, 0.0), (x1, 0.0), (x1, h), (w, h)],
             4: [(0.0, 0.0), (x1, 0.0), (x1, y2), (w, y2), (w, h)],
             5: [(0.0, 0.0), (x1, 0.0), (x1, y2), (x3, y2), (x3, h), (w, h)]}[count]
    if preset.startswith("bent"):
        return elbow
    if count == 2:
        return [(0.0, 0.0)] + _bezier((0.0, 0.0), (w / 2, 0.0), (w, h / 2), (w, h))
    if count == 3:
        first = _bezier((0.0, 0.0), (x1 / 2, 0.0), (x1, h / 4), (x1, h / 2))
        second = _bezier((x1, h / 2), (x1, h * 3 / 4), ((w + x1) / 2, h), (w, h))
        return [(0.0, 0.0)] + first + second
    return elbow


def bounding_box(points: list[Point]) -> tuple[int, int, int, int]:
    """``(left, top, width, height)`` of points, rounded to EMU."""
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    left, top = min(xs), min(ys)
    return round(left), round(top), round(max(xs) - left), round(max(ys) - top)


def drawn_bounds(shape: "Shape", *, row_heights: "list[int] | None" = None
                 ) -> tuple[int, int, int, int] | None:
    """See :attr:`Shape.drawn_bounds`.  A table's rows are ``row_heights`` when given, else
    measured (:func:`.fit.table_heights`)."""
    points = route(shape)
    if points is None:
        points = frame_corners(shape)
    if points is None:
        return None
    left, top, width, height = bounding_box(points)
    if shape.kind == "graphic_frame" and shape.has_table:
        # PowerPoint draws a table from its grid, not its frame: the columns' widths and
        # the rows' heights, each grown to fit its text -- ``a:tr@h`` is only a minimum.
        if row_heights is None:
            from .fit import table_heights

            row_heights = table_heights(shape)
        columns, rows = sum(shape.table.column_widths), sum(row_heights)
        width, height = columns or width, rows or height
    return left, top, width, height


# -- crossings --------------------------------------------------------------------------------


def _clip(a: Point, b: Point, box: tuple[float, float, float, float]) -> float:
    """How much of segment ``a``-``b`` lies inside ``box`` (x0, y0, x1, y1): Liang-Barsky."""
    x0, y0, x1, y1 = box
    dx, dy = b[0] - a[0], b[1] - a[1]
    low, high = 0.0, 1.0
    for p, q in ((-dx, a[0] - x0), (dx, x1 - a[0]), (-dy, a[1] - y0), (dy, y1 - a[1])):
        if p == 0:
            if q < 0:
                return 0.0
            continue
        t = q / p
        if p < 0:
            low = max(low, t)
        else:
            high = min(high, t)
        if low > high:
            return 0.0
    return (high - low) * math.hypot(dx, dy)


def length_inside(points: list[Point], box: tuple[float, float, float, float]) -> float:
    """The length of a polyline inside an axis-aligned box ``(x0, y0, x1, y1)``."""
    return sum(_clip(points[index], points[index + 1], box) for index in range(len(points) - 1))


__all__ = ["bounding_box", "drawn_bounds", "frame_corners", "length_inside", "local_route",
           "route", "to_slide"]
