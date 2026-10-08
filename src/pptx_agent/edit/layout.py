"""Layout helpers: shapes in a row, a column or a grid, and labels beside their anchors
(LP21, LP22).

They place *given* shapes in a box the caller names -- they know no diagram type, and
draw nothing but the leader lines :func:`place_labels` is asked for.  Groups move as one
and rotated shapes are placed by what is drawn (:mod:`.arrange`).  Lengths are EMU, and
every call is one undo step.

* :func:`stack` -- left to right with a gap, aligned top, middle or bottom, packed at
  the start, centre or end of the box or spread across it; optionally equal widths;
* :func:`column` -- the same, top to bottom;
* :func:`grid` -- ``rows`` x ``columns`` cells with gutters, each shape centred in its
  cell or resized to it;
* :func:`place_labels` -- each label beside its anchor (a shape or a point) on the first
  side, in order of preference, where it collides with nothing, within a stated distance
  of the anchor; the labels that cannot be placed are returned with the reason.

For example::

    from pptx_agent.edit.layout import column, place_labels

    column([heading, body, footer], box=(x, y, w, h), gap=Pt(6), fit_text=True)
    report = place_labels(slide, [{"shape": label, "anchor": bubble}, ...],
                          sides=("right", "left"), distance=Pt(4), max_center=Pt(72))
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Iterable, Mapping, Sequence

from .arrange import Box, drawn_box, move_to, outermost, place, resize

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

ALIGN_ROW = ("top", "middle", "bottom", "keep")
ALIGN_COLUMN = ("left", "center", "right", "keep")
JUSTIFY = ("start", "center", "end", "spread")
SIDES = ("right", "left", "above", "below")
#: A point, in EMU.
_PT = 12700


class LayoutError(ValueError):
    """The shapes do not fit their box: ``needed`` and ``available`` say by how much."""

    def __init__(self, message: str, *, needed: float | None = None,
                 available: float | None = None) -> None:
        super().__init__(message)
        self.needed = needed
        self.available = available


def _fit_text(shape: "Shape") -> None:
    from .fit import _converged_height

    frame = shape.text_frame if shape.kind == "shape" else None
    if frame is not None and shape.text.strip():
        _converged_height(shape)


def _line(shapes: Sequence["Shape"], axis: int, *, box: Box, gap: float, align: str,
          justify: str, equal: bool, fit_text: bool) -> list[Box]:
    shapes = outermost(shapes)
    if not shapes:
        raise ValueError("no shapes to arrange")
    count = len(shapes)
    main = box[2] if axis == 0 else box[3]
    cross = box[3] if axis == 0 else box[2]
    with shapes[0]._slide.document.batch():
        if equal:
            size = (main - gap * (count - 1)) / count
            if size <= 0:
                raise LayoutError(f"{count} shapes with a {gap / _PT:g} pt gap leave no room "
                                  f"in {main / _PT:.2f} pt", needed=gap * (count - 1),
                                  available=main)
            for shape in shapes:
                resize(shape, *((size, None) if axis == 0 else (None, size)))
        if fit_text:
            for shape in shapes:
                _fit_text(shape)
        boxes = [drawn_box(shape) for shape in shapes]
        sizes = [b[2 + axis] for b in boxes]
        total = sum(sizes) + gap * (count - 1)
        step = gap
        if justify == "spread" and count > 1:
            step = (main - sum(sizes)) / (count - 1)
            total = main
        if total > main + _PT / 100 or step < 0:
            raise LayoutError(f"the shapes need {total / _PT:.2f} pt and the box is "
                              f"{main / _PT:.2f} pt", needed=total, available=main)
        start = box[axis] + {"start": 0, "spread": 0, "center": (main - total) / 2,
                             "end": main - total}[justify]
        placed = []
        position = start
        for shape, current, size in zip(shapes, boxes, sizes):
            other = current[3 - axis]
            offset = {"top": 0, "left": 0, "middle": (cross - other) / 2,
                      "center": (cross - other) / 2, "bottom": cross - other,
                      "right": cross - other}.get(align)
            across = None if align == "keep" else box[1 - axis] + offset
            if axis == 0:
                move_to(shape, position, across)
            else:
                move_to(shape, across, position)
            placed.append(drawn_box(shape))
            position += size + step
    return placed


def stack(shapes: Sequence["Shape"], *, box: Box, gap: float = 0, align: str = "top",
          justify: str = "start", equal: bool = False, fit_text: bool = False) -> list[Box]:
    """Place ``shapes`` left to right in ``box``, in the order given, ``gap`` apart.

    ``align`` places each across the box (top, middle, bottom; keep leaves it), ``justify``
    packs the row at the box's start, centre or end, or spreads it so the first starts and
    the last ends at the box's edges.  ``equal`` gives every shape the same width, filling
    the box; ``fit_text`` gives each text shape the height its text needs first.  A row
    that does not fit raises :class:`LayoutError`.  Returns the drawn boxes, in order.
    """
    _check(align, ALIGN_ROW, "align")
    _check(justify, JUSTIFY, "justify")
    return _line(shapes, 0, box=box, gap=gap, align=align, justify=justify, equal=equal,
                 fit_text=fit_text)


def column(shapes: Sequence["Shape"], *, box: Box, gap: float = 0, align: str = "left",
           justify: str = "start", equal: bool = False, fit_text: bool = False) -> list[Box]:
    """:func:`stack`, top to bottom: ``align`` is left, center or right (or keep), and
    ``equal`` gives every shape the same height."""
    _check(align, ALIGN_COLUMN, "align")
    _check(justify, JUSTIFY, "justify")
    return _line(shapes, 1, box=box, gap=gap, align=align, justify=justify, equal=equal,
                 fit_text=fit_text)


def grid(shapes: Sequence["Shape"], *, box: Box, rows: int | None = None,
         columns: int | None = None, gutter: tuple[float, float] = (0, 0),
         order: str = "rows", fit: str = "keep") -> list[Box]:
    """Place ``shapes`` in a grid of ``rows`` x ``columns`` equal cells filling ``box``,
    ``gutter`` (x, y) apart, filling ``order`` rows first (or columns first).  ``fit``:
    keep each shape's size and centre it in its cell, or resize it to the cell.  Either
    count may be left out (it follows from the other); both out makes the grid as square
    as can be.  Returns the drawn boxes, in order."""
    _check(order, ("rows", "columns"), "order")
    _check(fit, ("keep", "cell"), "fit")
    shapes = outermost(shapes)
    count = len(shapes)
    if not count:
        raise ValueError("no shapes to arrange")
    if rows is None and columns is None:
        columns = math.ceil(math.sqrt(count))
    if columns is None:
        columns = math.ceil(count / rows)
    if rows is None:
        rows = math.ceil(count / columns)
    if rows < 1 or columns < 1:
        raise ValueError("rows and columns are 1 or more")
    if rows * columns < count:
        raise LayoutError(f"{count} shapes do not fit {rows} x {columns} cells",
                          needed=count, available=rows * columns)
    gx, gy = gutter
    width = (box[2] - gx * (columns - 1)) / columns
    height = (box[3] - gy * (rows - 1)) / rows
    if width <= 0 or height <= 0:
        raise LayoutError("the gutters leave no room for the cells",
                          needed=gx * (columns - 1) + gy * (rows - 1), available=box[2])
    placed = []
    with shapes[0]._slide.document.batch():
        for index, shape in enumerate(shapes):
            row, col = divmod(index, columns) if order == "rows" else divmod(index, rows)[::-1]
            x = box[0] + col * (width + gx)
            y = box[1] + row * (height + gy)
            if fit == "cell":
                place(shape, x, y, width, height)
            else:
                _, _, w, h = drawn_box(shape)
                if w > width + _PT / 100 or h > height + _PT / 100:
                    raise LayoutError(f"{shape.id} ({w / _PT:.2f} x {h / _PT:.2f} pt) is "
                                      f"larger than a cell ({width / _PT:.2f} x "
                                      f"{height / _PT:.2f} pt); use fit='cell'",
                                      needed=max(w, h), available=min(width, height))
                move_to(shape, x + (width - w) / 2, y + (height - h) / 2)
            placed.append(drawn_box(shape))
    return placed


def _check(value: str, allowed: Sequence[str], name: str) -> None:
    if value not in allowed:
        raise ValueError(f"{name} must be one of {', '.join(allowed)}, not {value!r}")


# -- labels --------------------------------------------------------------------------------------


@dataclass
class Placed:
    """A label :func:`place_labels` placed: where, on which side, how far from its anchor
    (centre to centre and edge to edge, EMU), and its leader line's address if it got one."""

    label: str
    box: Box
    side: str
    center_distance: float
    edge_distance: float
    leader: str | None = None


@dataclass
class Unplaced:
    """A label :func:`place_labels` could not place, and why: what the candidate positions
    ran into most (addresses, or ``"bounds"``), within the distance rule."""

    label: str
    reason: str
    blockers: list[str] = field(default_factory=list)


@dataclass
class LabelReport:
    placed: list[Placed]
    unplaced: list[Unplaced]
    #: The slide's collisions afterwards (``boxes=True``) that involve a label or leader.
    collisions: list


def _rect(box: Box) -> tuple[float, float, float, float]:
    return box[0], box[1], box[0] + box[2], box[1] + box[3]


def _overlap(a, b, slack: float) -> bool:
    return min(a[2], b[2]) - max(a[0], b[0]) > slack and min(a[3], b[3]) - max(a[1], b[1]) > slack


def _inside(inner, outer, slack: float) -> bool:
    return (inner[0] >= outer[0] - slack and inner[1] >= outer[1] - slack
            and inner[2] <= outer[2] + slack and inner[3] <= outer[3] + slack)


def _gap(a, b) -> float:
    dx = max(0.0, max(a[0], b[0]) - min(a[2], b[2]))
    dy = max(0.0, max(a[1], b[1]) - min(a[3], b[3]))
    return math.hypot(dx, dy)


@dataclass
class _Obstacle:
    address: str
    rect: tuple[float, float, float, float]
    order: int
    background: bool          # opaque, no text: a label may lie wholly on it, in front
    route: list | None = None


def _obstacles(slide: "Slide", skip: set[str]) -> list[_Obstacle]:
    from . import geometry
    from .fit import _flatten

    out = []
    for order, shape in enumerate(_flatten(slide.shapes)):
        if shape.kind == "group" or shape.id in skip:
            continue
        if any(parent.id in skip for parent in _parents(shape)):
            continue
        bounds = shape.drawn_bounds
        if bounds is None:
            continue
        route = geometry.route(shape)
        text = shape.kind == "shape" and bool(shape.text.strip())
        filled = shape.kind == "shape" and _filled(shape)
        background = (filled and not text) or shape.kind == "picture"
        out.append(_Obstacle(shape.id, _rect(bounds), order, background, route))
    return out


def _parents(shape: "Shape"):
    parent = shape.parent_group
    while parent is not None:
        yield parent
        parent = parent.parent_group


def _filled(shape: "Shape") -> bool:
    fill = shape.effective_fill
    return fill is not None and getattr(fill, "kind", None) not in (None, "none")


def _candidates(anchor, size, side: str, distance: float, step: float, reach: float):
    """Positions on ``side`` of ``anchor`` (a rect), nearest first: the base spot, then
    slid along the side and pushed further out, as far as ``reach``."""
    ax0, ay0, ax1, ay1 = anchor
    w, h = size
    cx, cy = (ax0 + ax1) / 2, (ay0 + ay1) / 2
    spots = []
    pushes = int(reach // step) + 1
    slides = int(reach // step) + 1
    for push in range(pushes):
        out = distance + push * step
        for k in range(slides):
            for sign in ((0,) if k == 0 else (1, -1)):
                shift = sign * k * step
                if side == "right":
                    x, y = ax1 + out, cy - h / 2 + shift
                elif side == "left":
                    x, y = ax0 - out - w, cy - h / 2 + shift
                elif side == "above":
                    x, y = cx - w / 2 + shift, ay0 - out - h
                else:
                    x, y = cx - w / 2 + shift, ay1 + out
                spots.append((push * step + abs(shift), push, k, (x, y, x + w, y + h)))
    spots.sort(key=lambda spot: (spot[0], spot[1], spot[2]))
    return [spot[3] for spot in spots]


def place_labels(slide: "Slide", labels: Sequence[Mapping[str, Any]], *,
                 sides: Sequence[str] = ("right", "left", "above", "below"),
                 distance: float = 4 * _PT, max_center: float | None = None,
                 max_edge: float | None = None, leader: str = "none",
                 avoid: Iterable["Shape"] = (), box: Box | None = None,
                 step: float = 2 * _PT, align_text: bool = True) -> LabelReport:
    """Put each label beside its anchor where it collides with nothing.

    ``labels`` are mappings: ``shape`` (the label, a text shape already on ``slide``) and
    ``anchor`` (a shape) or ``point`` (``(x, y)`` slide EMU).  Labels are placed in order;
    each tries ``sides`` in order of preference -- its base spot ``distance`` from the
    anchor's edge, centred on it, then slid along that side and pushed further out --
    before the next side.  A spot is taken when it lies within ``box`` (default the
    slide's content area), keeps to the distance rule (``max_center``: the label's centre
    within that of the anchor's; ``max_edge``: the gap between their edges at most
    that), and overlaps no other shape or placed label and crosses no line.  Shapes in
    ``avoid`` may be overlapped (a background band); so may a filled shape without text
    the label lies wholly on (a quadrant panel), as :meth:`Slide.collisions` allows.

    ``leader``: ``"line"`` joins every label to its anchor with a straight connector,
    ``"auto"`` only labels moved off their base spot, ``"none"`` none.  ``align_text``
    aligns a label's text toward its anchor (left on the right side, and so on).

    Returns a :class:`LabelReport`: what was placed, what could not be (with the
    shapes in the way), and the collisions afterwards that involve a label.  One undo
    step.
    """
    from .area import content_area
    from .geometry import length_inside

    for side in sides:
        _check(side, SIDES, "side")
    _check(leader, ("none", "line", "auto"), "leader")
    if max_center is None and max_edge is None:
        reach = 2 * 72 * _PT
    else:
        reach = max(v for v in (max_center, max_edge) if v is not None)
    step = max(step, reach / 24)
    bounds = _rect(box if box is not None else content_area(slide))
    label_ids = {item["shape"].id for item in labels}
    avoid_ids = {shape.id for shape in avoid}
    obstacles = _obstacles(slide, label_ids | avoid_ids)
    orders = {shape.id: order for order, shape in enumerate(_flatten_all(slide))}
    slack = _PT / 20
    placed: list[Placed] = []
    unplaced: list[Unplaced] = []
    taken: list[tuple[str, tuple]] = []
    leaders: list[str] = []
    with slide.document.batch():
        for item in labels:
            label = item["shape"]
            anchor_shape = item.get("anchor")
            if anchor_shape is not None:
                anchor = _rect(drawn_box(anchor_shape))
            else:
                x, y = item["point"]
                anchor = (x, y, x, y)
            _, _, w, h = drawn_box(label)
            ax, ay = (anchor[0] + anchor[2]) / 2, (anchor[1] + anchor[3]) / 2
            order = orders.get(label.id, 1 << 30)
            hits: dict[str, int] = {}
            found = None
            for side in sides:
                spots = _candidates(anchor, (w, h), side, distance, step, reach)
                for index, rect in enumerate(spots):
                    centre = math.hypot((rect[0] + rect[2]) / 2 - ax, (rect[1] + rect[3]) / 2 - ay)
                    edge = _gap(rect, anchor)
                    if max_center is not None and centre > max_center + slack:
                        continue
                    if max_edge is not None and edge > max_edge + slack:
                        continue
                    if not _inside(rect, bounds, slack):
                        hits["bounds"] = hits.get("bounds", 0) + 1
                        continue
                    blocker = _blocked(rect, obstacles, taken, order, slack, length_inside,
                                       anchor_shape)
                    if blocker is not None:
                        hits[blocker] = hits.get(blocker, 0) + 1
                        continue
                    found = (side, rect, index, centre, edge)
                    break
                if found:
                    break
            if found is None:
                ranked = sorted(hits, key=lambda key: -hits[key])[:5]
                rule = (f"within {max_center / _PT:g} pt centre to centre" if max_center
                        else f"within {max_edge / _PT:g} pt edge to edge" if max_edge
                        else f"within {reach / _PT:g} pt")
                unplaced.append(Unplaced(label.id, f"every spot on {', '.join(sides)} "
                                         f"{rule} is blocked", ranked))
                continue
            side, rect, index, centre, edge = found
            move_to(label, rect[0], rect[1])
            if align_text and label.text_frame is not None:
                alignment = {"right": "left", "left": "right"}.get(side, "center")
                for paragraph in label.text_frame.paragraphs:
                    paragraph.alignment = alignment
            taken.append((label.id, rect))
            line = None
            if leader == "line" or (leader == "auto" and index > 0):
                line = _leader(slide, anchor_shape, anchor, label, side)
                if line is not None:
                    leaders.append(line.id)
                    taken.append((line.id, ("route", line)))
            placed.append(Placed(label.id, drawn_box(label), side, centre, edge,
                                 line.id if line is not None else None))
    mine = label_ids | set(leaders)
    collisions = [c for c in slide.collisions(boxes=True)
                  if c.shape in mine or c.other in mine]
    return LabelReport(placed, unplaced, collisions)


def _flatten_all(slide: "Slide"):
    from .fit import _flatten

    return _flatten(slide.shapes)


def _blocked(rect, obstacles, taken, order, slack, length_inside, anchor_shape) -> str | None:
    for other_id, other in taken:
        if other[0] == "route":
            continue
        if _overlap(rect, other, slack):
            return other_id
    for obstacle in obstacles:
        if obstacle.route is not None:
            crossed = length_inside(obstacle.route, rect)
            if crossed > _PT / 2:
                return obstacle.address
            continue
        if not _overlap(rect, obstacle.rect, slack):
            continue
        if (obstacle.background and obstacle.order < order
                and _inside(rect, obstacle.rect, slack)
                and (anchor_shape is None or obstacle.address != anchor_shape.id)):
            continue
        return obstacle.address
    return None


def _leader(slide: "Slide", anchor_shape, anchor, label, side):
    """A straight connector from the anchor (glued, or from the point) to the label's
    facing side."""
    facing = {"right": "left", "left": "right", "above": "bottom", "below": "top"}[side]
    near = {"right": "right", "left": "left", "above": "top", "below": "bottom"}[side]
    try:
        begin = (anchor_shape, near) if anchor_shape is not None else \
            (round(anchor[0]), round(anchor[1]))
        return slide.add_connector("straight", begin, (label, facing))
    except (ValueError, KeyError):
        return None


__all__ = ["ALIGN_COLUMN", "ALIGN_ROW", "JUSTIFY", "LabelReport", "LayoutError", "Placed",
           "SIDES", "Unplaced", "column", "grid", "place_labels", "stack"]
