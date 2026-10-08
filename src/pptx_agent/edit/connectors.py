"""Connector geometry: where an elbow bends, and how PowerPoint writes the frame.

A connector is a ``p:cxnSp`` whose preset draws from its frame's top-left corner to its
bottom-right one; flips and a rotation put those corners on the two ends.  Straight
connectors are ``straightConnector1``.  Elbow and curved connectors are ``bentConnector2``
to ``5`` and ``curvedConnector2`` to ``5`` -- one to four bends -- and their adjustments say
where the bends are, as fractions (1/100,000) of the frame.

**What PowerPoint does, measured.**  Connectors were attached between shapes through
PowerPoint's AppleScript (``begin connect``/``end connect``), the shapes moved, and the
saved XML read back: 180 elbow connectors over fourteen arrangements of two shapes, every
pairing of the four sides of a rectangle, plus ellipses.  An elbow and a curved connector
between the same sites always got the same frame and adjustments; only the preset differs.

* An end leaves its site along the site's angle (``a:cxn@ang``, turned with the shape).
* Ends facing each other with room between them: one bend each way, halfway
  (``bentConnector3``, ``adj1`` 50000) -- halfway between the *ends*, not the shapes.
* Ends at right angles, the end ahead and on the side it faces away from: one bend
  (``bentConnector2``).
* Otherwise the route steps out of the start by a fixed 228,600 EMU (a quarter inch) --
  or, when the end shape lies wholly ahead, to the middle of the gap between the shapes --
  and runs past the end shape in the same way, so it never doubles back through either.
  A run that would cross the end shape's span goes beyond it, a quarter inch clear.
* An extent of zero is written as 12,700 EMU (one point): the adjustments are fractions of
  it, so PowerPoint never lets it vanish.  The far end is then up to a point off its site,
  exactly as in PowerPoint's own files.
* The frame's rotation follows the direction the connector leaves its start: 0 to the
  right, 90 down, 180 left, 270 up; flips then put the far corner on the end.  (PowerPoint
  spells two of the vertical cases differently -- ``rot=270 flipH`` and ``rot=90 flipH
  flipV`` -- and so does this module; every spelling of the same frame draws the same.)

The router reproduces 167 of the 180 measured routes to within a point.  The rest are a
detour around overlapping shapes taken on the other side, and degenerate layouts (both ends
level, or the shapes overlapping) where PowerPoint's own route runs through a shape.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

#: How far a route steps out of a site before it turns (measured: 0.25 in, at any size).
STUB = 228600
#: The smallest extent PowerPoint writes for an elbow or curved connector's frame.
MIN_EXTENT = 12700
QUARTER = 5400000

#: ``kind`` -> the preset family.
KINDS = {"straight": "straightConnector1", "elbow": "bentConnector", "curved": "curvedConnector"}

Point = tuple[float, float]
Box = tuple[float, float, float, float]


@dataclass(frozen=True)
class End:
    """One end of a connector, in slide EMU."""

    point: Point
    #: The way a route leaves this end: a unit axis vector, or ``None`` for a free end.
    direction: tuple[int, int] | None = None
    #: The bounding box of the shape the end is attached to, if any.
    box: Box | None = None


@dataclass(frozen=True)
class Frame:
    """A connector's ``a:xfrm`` and geometry."""

    preset: str
    x: int
    y: int
    cx: int
    cy: int
    rot: int = 0
    flip_h: bool = False
    flip_v: bool = False
    #: ``(name, value)`` for each adjustment, in the preset's order.
    adjustments: tuple[tuple[str, int], ...] = ()


def axis(angle: float) -> tuple[int, int]:
    """The axis direction nearest an angle (1/60,000 degree, clockwise from +x)."""
    quarter = round(angle / QUARTER) % 4
    return ((1, 0), (0, 1), (-1, 0), (0, -1))[quarter]


# -- routing ------------------------------------------------------------------------------


def _to_normal(direction: tuple[int, int]):
    """A rotation taking ``direction`` onto +x, and its inverse."""
    if direction == (1, 0):
        return (lambda p: (p[0], p[1])), (lambda p: (p[0], p[1]))
    if direction == (0, 1):
        return (lambda p: (p[1], -p[0])), (lambda p: (-p[1], p[0]))
    if direction == (-1, 0):
        return (lambda p: (-p[0], -p[1])), (lambda p: (-p[0], -p[1]))
    return (lambda p: (-p[1], p[0])), (lambda p: (p[1], -p[0]))


def _box_through(f, box: Box, origin: Point) -> Box:
    corners = [f((box[0], box[1])), f((box[2], box[3]))]
    return (min(c[0] for c in corners) - origin[0], min(c[1] for c in corners) - origin[1],
            max(c[0] for c in corners) - origin[0], max(c[1] for c in corners) - origin[1])


def _touches(box: Box, a: Point, b: Point) -> bool:
    """Does the axis-aligned segment ``a``-``b`` touch the closed ``box``?"""
    low_x, high_x = sorted((a[0], b[0]))
    low_y, high_y = sorted((a[1], b[1]))
    return low_x <= box[2] and high_x >= box[0] and low_y <= box[3] and high_y >= box[1]


def free_directions(start: End, end: End) -> tuple[tuple[int, int], tuple[int, int]]:
    """Directions for ends that are not attached: they face each other along the axis
    the attached end uses, or horizontally when neither is attached."""
    ds, de = start.direction, end.direction
    if ds is None and de is None:
        ds = (1, 0) if end.point[0] >= start.point[0] else (-1, 0)
    if ds is None:
        ds = (-de[0], -de[1])
    if de is None:
        de = (-ds[0], -ds[1])
    return ds, de


def route(start: End, end: End) -> list[Point]:
    """The elbow's corner points from ``start`` to ``end``, both ends included.

    Always three to six points (one to four bends; a zero-length middle run is kept, as
    PowerPoint keeps it), in slide EMU.
    """
    ds, de = free_directions(start, end)
    f, back = _to_normal(ds)
    origin = f(start.point)
    target = f(end.point)
    facing = f(de)
    X, Y = target[0] - origin[0], target[1] - origin[1]
    a = _box_through(f, start.box, origin) if start.box else (0.0, 0.0, 0.0, 0.0)
    b = _box_through(f, end.box, origin) if end.box else (X, Y, X, Y)
    points = _route_normal(X, Y, facing, a, b)
    return [back((x + origin[0], y + origin[1])) for x, y in points]


def _route_normal(X: float, Y: float, facing: Point, a: Box, b: Box) -> list[Point]:
    """The route with the start at the origin leaving along +x; ``facing`` is the end's."""
    ahead = (a[2] + b[0]) / 2 if b[0] > a[2] else None
    behind = (b[2] + a[0]) / 2 if b[2] < a[0] else None
    stub = max(a[2], 0.0) + STUB
    overlap = a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]

    def depart(y2: float) -> float:
        """Where the first bend goes, for a route that next runs to ``y2``."""
        if ahead is not None:
            return ahead
        low, high = sorted((0.0, y2))
        if low <= b[3] and b[1] <= high:  # the run crosses the end shape's span
            return max(a[2], b[2], 0.0) + STUB
        return stub

    if facing == (-1, 0):  # the end faces back at the start
        if X > 0:
            return [(0, 0), (X / 2, 0), (X / 2, Y), (X, Y)]
        y2 = _between(a, b, Y)
        x1 = depart(y2)
        x3 = min(b[0], X) - STUB
        if _touches(a, (x3, y2), (x3, Y)):
            x3 = min(a[0], X) - STUB
        return [(0, 0), (x1, 0), (x1, y2), (x3, y2), (x3, Y), (X, Y)]

    if facing == (1, 0):  # both ends face the same way: a U, unless something is in it
        turn = max(a[2], b[2], X, 0.0) + STUB
        candidate = [(0, 0), (turn, 0), (turn, Y), (X, Y)]
        segments = list(zip(candidate, candidate[1:]))
        by_a = not overlap and any(_touches(a, p, q) for p, q in segments[1:])
        by_b = not overlap and any(_touches(b, p, q) for p, q in segments[:-1])
        if by_a:
            y2 = a[3] + STUB
            x3 = behind if behind is not None else max(b[2], X) + STUB
            return [(0, 0), (stub, 0), (stub, y2), (x3, y2), (x3, Y), (X, Y)]
        if by_b:
            # Round the end shape on its nearer side.
            y2 = b[3] + STUB if b[3] <= -b[1] else b[1] - STUB
            x1 = depart(y2)
            x3 = max(b[2], X) + STUB
            return [(0, 0), (x1, 0), (x1, y2), (x3, y2), (x3, Y), (X, Y)]
        return candidate

    # The end is at right angles: it must be reached travelling against the way it faces.
    sign = facing[1]
    if X > 0 and Y != 0 and (1 if Y > 0 else -1) == -sign:
        return [(0, 0), (X, 0), (X, Y)]
    if sign < 0:
        y2 = (a[3] + b[1]) / 2 if a[3] < b[1] else min(a[1], b[1], Y) - STUB
    else:
        y2 = (b[3] + a[1]) / 2 if a[1] > b[3] else max(a[3], b[3], Y) + STUB
    x1 = depart(y2)
    return [(0, 0), (x1, 0), (x1, y2), (X, y2), (X, Y)]


def _between(a: Box, b: Box, Y: float) -> float:
    """A run across the route: midway between the shapes, or clear of both."""
    if b[1] > a[3]:
        return (a[3] + b[1]) / 2
    if a[1] > b[3]:
        return (b[3] + a[1]) / 2
    return max(a[3], b[3]) + STUB if Y >= 0 else min(a[1], b[1]) - STUB


# -- frames -------------------------------------------------------------------------------


def _matrix(rot: int, flip_h: bool, flip_v: bool) -> tuple[float, float, float, float]:
    """The linear map from the frame's local axes to the slide's: rotation after flips."""
    angle = math.radians(rot / 60000)
    cos, sin = round(math.cos(angle)), round(math.sin(angle))
    fx, fy = (-1 if flip_h else 1), (-1 if flip_v else 1)
    return (cos * fx, -sin * fy, sin * fx, cos * fy)


def _spelling(direction: tuple[int, int], dx: float, dy: float) -> tuple[int, bool, bool]:
    """``(rot, flipH, flipV)`` as PowerPoint writes a connector leaving along ``direction``.

    The rotation is the start's direction; the flips put the frame's far corner in the
    end's quadrant.  An end level with the start on an axis takes the side the unflipped
    rotation gives, as PowerPoint does.
    """
    if direction[1] != 0:
        rot = 5400000 if direction[1] > 0 else 16200000
    else:
        rot = 0 if direction[0] > 0 else 10800000
    m = _matrix(rot, False, False)
    sx = (1 if dx > 0 else -1) if dx else (1 if m[0] + m[1] > 0 else -1)
    sy = (1 if dy > 0 else -1) if dy else (1 if m[2] + m[3] > 0 else -1)
    if direction[1] != 0:  # vertical start: PowerPoint's two fixed spellings
        if (sx, sy) == (1, -1):
            return 5400000, True, True
        if (sx, sy) == (1, 1):
            return 16200000, True, False
    for flip_h in (False, True):
        for flip_v in (False, True):
            m = _matrix(rot, flip_h, flip_v)
            # The local diagonal (+,+) must land in the quadrant of the end.
            if (1 if m[0] + m[1] > 0 else -1) == sx and (1 if m[2] + m[3] > 0 else -1) == sy:
                return rot, flip_h, flip_v
    raise AssertionError("no spelling found")  # pragma: no cover


def straight_frame(start: Point, end: Point) -> Frame:
    """A ``straightConnector1`` from ``start`` to ``end``: flips, never a rotation."""
    (x1, y1), (x2, y2) = start, end
    return Frame("straightConnector1", round(min(x1, x2)), round(min(y1, y2)),
                 round(abs(x2 - x1)), round(abs(y2 - y1)), 0, x2 < x1, y2 < y1)


def bent_frame(points: list[Point], family: str = "bentConnector") -> Frame:
    """The frame and adjustments that draw ``points`` (from :func:`route`) with a
    ``bentConnector<n>`` or ``curvedConnector<n>``."""
    count = len(points) - 1
    if not 2 <= count <= 5:
        raise ValueError(f"an elbow has two to five segments, not {count}")
    preset = f"{family}{count}"
    (sx, sy), (ex, ey) = points[0], points[-1]
    first = (points[1][0] - sx, points[1][1] - sy)
    if abs(first[0]) >= abs(first[1]):
        direction = (1 if first[0] >= 0 else -1, 0)
    else:
        direction = (0, 1 if first[1] >= 0 else -1)
    if first == (0, 0):  # a zero-length first run: fall back to where the end is
        direction = (1 if ex >= sx else -1, 0)
    dx, dy = ex - sx, ey - sy
    rot, flip_h, flip_v = _spelling(direction, dx, dy)
    sideways = rot in (5400000, 16200000)
    width, height = (abs(dy), abs(dx)) if sideways else (abs(dx), abs(dy))
    if sideways:
        centre_x, centre_y = (sx + ex) / 2, (sy + ey) / 2
        x, y = centre_x - width / 2, centre_y - height / 2
    else:
        x, y = min(sx, ex), min(sy, ey)
    cx, cy = max(width, MIN_EXTENT), max(height, MIN_EXTENT)

    a, b, c, d = _matrix(rot, flip_h, flip_v)
    determinant = a * d - b * c

    def local(point: Point) -> Point:
        vx, vy = point[0] - sx, point[1] - sy
        return ((d * vx - b * vy) / determinant, (-c * vx + a * vy) / determinant)

    def fraction(value: float, extent: float) -> int:
        return round(value / extent * 100000)

    names = ("adj1", "adj2", "adj3")
    values: list[int] = []
    if count >= 3:
        values.append(fraction(local(points[1])[0], cx))
    if count >= 4:
        values.append(fraction(local(points[2])[1], cy))
    if count == 5:
        values.append(fraction(local(points[3])[0], cx))
    return Frame(preset, round(x), round(y), round(cx), round(cy), rot, flip_h, flip_v,
                 tuple(zip(names, values)))


def frame_for(kind: str, start: End, end: End) -> Frame:
    """The frame of a connector of ``kind`` (``straight``, ``elbow``, ``curved``)."""
    if kind not in KINDS:
        raise ValueError(f"connector kind must be one of {sorted(KINDS)}")
    if kind == "straight":
        return straight_frame(start.point, end.point)
    return bent_frame(route(start, end), KINDS[kind])


def endpoints(x: float, y: float, cx: float, cy: float, rot: float = 0.0,
              flip_h: bool = False, flip_v: bool = False) -> tuple[Point, Point]:
    """Where a connector's frame puts its two ends: the local top-left and bottom-right
    corners, through the flips and the rotation about the frame's centre."""
    centre_x, centre_y = x + cx / 2, y + cy / 2
    angle = math.radians(rot / 60000)
    cos, sin = math.cos(angle), math.sin(angle)
    ends = []
    for lx, ly in ((0.0, 0.0), (cx, cy)):
        if flip_h:
            lx = cx - lx
        if flip_v:
            ly = cy - ly
        vx, vy = lx - cx / 2, ly - cy / 2
        ends.append((centre_x + vx * cos - vy * sin, centre_y + vx * sin + vy * cos))
    return ends[0], ends[1]
